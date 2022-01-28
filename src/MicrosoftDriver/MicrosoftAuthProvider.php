<?php


namespace DynamicScreen\Microsoft\MicrosoftDriver;

use DynamicScreen\SdkPhp\Interfaces\IModule;
use Illuminate\Support\Arr;
use Illuminate\Support\Facades\Session;
use DynamicScreen\SdkPhp\Handlers\OAuthProviderHandler;
use Microsoft\Graph\Graph;

class MicrosoftAuthProvider extends OAuthProviderHandler
{
    protected static string $provider = 'microsoft';

    public function __construct(IModule $module, $config = null)
    {
        parent::__construct($module, $config);
    }

    public function testConnection($config = null)
    {
        $config = $config ?? $this->default_config;
        try {
            $this->getUserInfos($config);
            return response('', 200);
        } catch (\Exception $e) {
            return response('Connection failed', 403);
        }
    }

    public function signin($callbackUrl = null)
    {
        $data = Session::get('auth_provider');
        $data = json_encode($data);

        $oauthClient = $this->getOAuthClient();

        $authUrl = $oauthClient->getAuthorizationUrl();

        $oauthState = $oauthClient->getState();
        Session::put($oauthState, $data);

        return $authUrl;

    }

    public function callback($request, $redirectUrl = null)
    {
        abort_unless(Session::has($request->get('state')), 400, 'No state');
        $state = Session::get($request->get('state'));
        Session::forget($request->get('state'));

        $authCode = $request->get('code');
        abort_unless($authCode, 400, 'No code');

        $oauthClient = $this->getOAuthClient();

        try {
            $options = $oauthClient->getAccessToken('authorization_code', ['code' => $authCode])->jsonSerialize();
            $options = array_merge($options, ['deltaLinks' => [$this->getNewPersonalDeltaLink($options)]]);
        } catch (\Exception $e) {
            dd('Error in callback microsoft driver', $e);
        }

        $data = $this->processOptions($options);
        $dataStr = json_encode($data);

        return redirect()->away($redirectUrl ."&data=$dataStr");
    }

    public function getUserInfos($config = null)
    {
        $config = $config ?? $this->default_config;

        $graph = new Graph();
        $graph->setAccessToken(Arr::get($config, 'access_token'));

        return $graph->createRequest('GET', '/me?$select=displayName,mail,userPrincipalName')->execute()->getBody();
    }

    public function getDriveItem($file_id, $site_id = null, $config = null)
    {
        $config = $config ?? $this->default_config;

        try {
            $graph = new Graph();
            $graph->setAccessToken(Arr::get($config, 'access_token'));

            $endpoint = "/drive/items/$file_id";
            if($site_id) {
                $endpoint = "/sites/$site_id" . $endpoint;
            } else {
                $endpoint = "/me$endpoint";
            }

            return $graph->createRequest('GET', $endpoint)->execute()->getBody();
        } catch (\Exception $e) {
            return false;
        }
    }

    public function getAvailableSites($config = null)
    {
        $config = $config ?? $this->default_config;

        try {
            $graph = new Graph();
            $graph->setAccessToken(Arr::get($config, 'access_token'));

            return $graph->createRequest('GET', '/sites?search=*')->execute()->getBody()['value'];

        } catch(\Exception $e) {
            return false;
        }

    }

    public function getUserPhoto($config = null)
    {
        $config = $config ?? $this->default_config;

        try {

            $graph = new Graph();
            $graph->setAccessToken(Arr::get($config, 'access_token'));

            $photo = $graph->createRequest("GET", "/me/photo/\$value")->execute()->getRawBody();
            $meta = $graph->createRequest("GET", "/me/photo")->execute()->getBody();

            return 'data:' . $meta["@odata.mediaContentType"] . ';base64,' . base64_encode($photo);

        } catch(\Exception $e) {
            return false;
        }
    }

    public function getNewPersonalDeltaLink($config = null)
    {
        $config = $config ?? $this->default_config;

        $graph = new Graph();
        $graph->setAccessToken(Arr::get($config, 'access_token'));

        $resp = $graph->createRequest('GET', '/me/drive/root/delta')->execute()->getBody();

        return $resp['@odata.deltaLink'];
    }

    public function getNewSiteDeltaLink($site_id, $config = null)
    {
        $config = $config ?? $this->default_config;

        $graph = new Graph();
        $graph->setAccessToken(Arr::get($config, 'access_token'));

        $resp = $graph->createRequest('GET', "/sites/$site_id/drive/root/delta")->execute()->getBody();

        return $resp['@odata.deltaLink'];
    }

    public function getRecentFilesChanges($config = null)
    {
        $config = $config ?? $this->default_config;

        $graph = new Graph();
        $graph->setAccessToken(Arr::get($config, 'access_token'));

        $deltaLinks = Arr::get($config, 'deltaLinks');

        if(empty($deltaLinks)) {
            return false;
        }

        $files_changes = collect();

        foreach($deltaLinks as $delta) {
            try {
                $resp = $graph->createRequest('GET', $delta)->execute()->getBody();
                $files_changes = $files_changes->merge(collect($resp['value'])->pluck('id'));
            } catch(\Exception $e) {
                continue;
            }
        }

        return $files_changes;
    }

    public function downloadPdf($file_id, $site_id = null, $config = null)
    {
        $config = $config ?? $this->default_config;

        $graph = new Graph();
        $graph->setAccessToken(Arr::get($config, 'access_token'));

        $endpoint = "drive/items/$file_id/content?format=pdf";
        if(empty($site_id)) {
            $endpoint = "/me/$endpoint";
        } else {
            $endpoint = "/sites/$site_id/$endpoint";
        }

        return $graph->createRequest('GET', $endpoint)->execute()->getRawBody();
    }

    public function refreshToken($config = null): array
    {
        $config = $config ?? $this->default_config;

        $oauthClient = $this->getOAuthClient();

        return array_merge($config, $oauthClient->getAccessToken('refresh_token', ['refresh_token' => Arr::get($config, 'refresh_token')])->jsonSerialize());
    }

    public function getOAuthClient()
    {
        return new \League\OAuth2\Client\Provider\GenericProvider([
            'clientId'                => config('azure.APP_ID'),
            'clientSecret'            => config('azure.APP_SECRET'),
            'redirectUri'             => route('api.oauth.callback'),
            'urlAuthorize'            => config('azure.AUTHORITY').config('azure.AUTHORIZE_ENDPOINT'),
            'urlAccessToken'          => config('azure.AUTHORITY').config('azure.TOKEN_ENDPOINT'),
            'urlResourceOwnerDetails' => '',
            'scopes'                  => 'openid profile offline_access user.read mailboxsettings.read calendars.readwrite files.readwrite.all sites.readwrite.all'
        ]);
    }
}
