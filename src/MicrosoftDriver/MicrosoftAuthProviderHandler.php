<?php


namespace DynamicScreen\Microsoft\MicrosoftDriver;

use DynamicScreen\SdkPhp\Interfaces\IModule;
use Illuminate\Support\Arr;
use Illuminate\Support\Str;
use Illuminate\Support\Facades\Session;
use DynamicScreen\SdkPhp\Handlers\OAuthProviderHandler;
use Microsoft\Graph\Graph;

class MicrosoftAuthProviderHandler extends OAuthProviderHandler
{
    protected static string $provider = 'microsoft';

    public function __construct(IModule $module, $config = null)
    {
        parent::__construct($module, $config);
    }

    public function getScopes()
    {
        return [
            "openid",
            "profile",
            "offline_access",
            "user.read",
            "mailboxsettings.read",
            "calendars.readwrite",
            "files.readwrite.all",
            "sites.readwrite.all",
        ];
    }

    public function getSharepointScopes()
    {
        return [
            ".default",
        ];
    }

    public function getOAuthClient(array $overwrite = [])
    {
        return new \League\OAuth2\Client\Provider\GenericProvider(array_merge([
            'clientId'                => config('azure.APP_ID'),
            'clientSecret'            => config('azure.APP_SECRET'),
            'redirectUri'             => route('api.oauth.callback'),
            'urlAuthorize'            => config('azure.AUTHORITY') . config('azure.AUTHORIZE_ENDPOINT'),
            'urlAccessToken'          => config('azure.AUTHORITY') . config('azure.TOKEN_ENDPOINT'),
            'urlResourceOwnerDetails' => '',
            'scopes'                  => implode(" ", $this->getScopes()),
        ], $overwrite));
    }

    public function getSharepointOAuthClient($tenantId, $tenantUrl, array $overwrite = [])
    {
        $urls = collect($tenantUrl);

        return $this->getOAuthClient(array_merge([
            'urlAuthorize'   => "https://login.microsoftonline.com/" . $tenantId . config('azure.AUTHORIZE_ENDPOINT'),
            'urlAccessToken' => "https://login.microsoftonline.com/" . $tenantId . config('azure.TOKEN_ENDPOINT'),
            'scopes'         => collect($this->getSharepointScopes())
                ->map(fn ($scope) => $urls->map(fn ($url) => \Str::finish($url, "/") . $scope))
                ->collapse()
                ->implode(" "),
        ], $overwrite));
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
        Session::put("{$oauthState}_step", 1);

        return $authUrl;

    }

    public function callback($request, $redirectUrl = null)
    {
        $state = $request->get('state');
        logs()->info("=========================");
        logs()->info("State: " . $state);
        abort_unless(Session::has($request->get('state')), 400, 'No state');
        $stateData = Session::get($request->get('state'));
        $step = Session::get($request->get('state') . "_step", 1);

        logs()->info("Microsoft callback, step: " . $step);

        if ($step === 1) {
            // Step 1 : We authenticated using the "public" authority, now we can use the Microsoft Graph
            // to find the Sharepoint URL and make a new authentication process using Sharepoint tenant

            // First, we get the access token from the auth code we just got
            $authCode = $request->get('code');
            abort_unless($authCode, 400, 'No code');

            $oauthClient = $this->getOAuthClient();

            try {
                $auth = $oauthClient->getAccessToken('authorization_code', [ 'code' => $authCode ])->jsonSerialize();
                $auth = array_merge($auth, [ 'deltaLinks' => [ $this->getNewPersonalDeltaLink($auth) ] ]);
            } catch (\Exception $e) {
                dd('Error in callback microsoft driver', $e);
            }


            // And we request the Graph to get the Sharepoint URL
            $graph = new Graph();
            $graph->setAccessToken($auth["access_token"]);

            $success = true;

            try {
                $sitesResponse = $graph->createRequest('GET', "/sites/root")->execute();
                logs()->info("Sharepoint Sites request status: " . $sitesResponse->getStatus());
                $orgResponse = $graph->createRequest('GET', "/organization")->execute();
                logs()->info("Organization request status: " . $sitesResponse->getStatus());
                $drivesResponse = $graph->createRequest('GET', "/me/drives")->execute();
                logs()->info("Drives request status: " . $drivesResponse->getStatus());
            } catch (\Exception $ex) {
                // We can't get the Sharepoint URL, we skip
                logs()->info("Failed to get Sharepoint");
                $success = false;
            }

            logs()->info("Fetch success: " . $success);

            if ($success && $sitesResponse->getStatus() == 200 && $orgResponse->getStatus() == 200 && $drivesResponse->getStatus() == 200) {
                $sites = $sitesResponse->getBody();
                logs()->info("Sites response: " . json_encode($sites));
                $orgs = $orgResponse->getBody();
                $oneDrives = $drivesResponse->getBody();

                $tenantId = Arr::get($orgs, "value.0.id");
                Session::put($state . "_tenant_id", $tenantId);

                $tenantUrl = Arr::get($sites, "webUrl");
                Session::put($state . "_tenant_url", $tenantUrl);

                $drives = [];
                $personalUrl = null;
                foreach (Arr::get($oneDrives, "value", []) as $drive) {
                    $url = $drive["webUrl"];
                    $urlinfo = parse_url($url);
                    $pathinfo = pathinfo($urlinfo["path"]);
                    $drives[] = [
                        "id" => $drive["id"],
                        "name" => $drive["name"],
                        "url" => $url,
                        "domain" => $urlinfo["host"],
                        "path" => $pathinfo["dirname"],
                    ];
                    $personalUrl ="https://" . $urlinfo["host"];
                }
                Session::put($state . "_drives", $drives);

                $oauthClient = $this->getSharepointOAuthClient($tenantId, $tenantUrl);

                $authUrl = $oauthClient->getAuthorizationUrl();

                $newState = $oauthClient->getState();
                logs()->info("New state: " . $newState);

                Session::put($newState, Session::get($state));
                Session::put($newState . "_tenant_id", Session::get($state . "_tenant_id"));
                Session::put($newState . "_tenant_url", Session::get($state . "_tenant_url"));
                Session::put($newState . "_personal_url", $personalUrl);
                Session::put($newState . "_drives", Session::get($state . "_drives"));
                Session::put($newState . "_step", 2);
//                Session::put($newState . "_scopes", implode(" ", $scopes));
                Session::put($newState . "_auth", $auth);

                logs()->info("Redirecting to: " . $authUrl);

                return redirect()->away($authUrl);
            }

        } else if ($step === 2) {
            logs()->info("Step 2");
            $options = Session::get("{$state}_auth");

            if (Session::has($state . "_tenant_id") && Session::has($state . "_tenant_url")) {
                $tenantId = Session::get($state . "_tenant_id");
                $tenantUrl = Session::get($state . "_tenant_url");
                $personalUrl = Session::get($state . "_personal_url");
                $oauthClient = $this->getSharepointOAuthClient($tenantId, $tenantUrl);

                try {
                    $authCode = $request->get('code');
                    abort_unless($authCode, 400, 'No code');

                    $auth = $oauthClient->getAccessToken('authorization_code', [ 'code' => $authCode ])
                                        ->jsonSerialize();
                    $auth = array_merge($auth, [ 'deltaLinks' => [ $this->getNewPersonalDeltaLink($options) ] ]);
                    $options["sharepoint"] = $auth;

                    $options["sharepoint"]["tenant_url"] = $tenantUrl;
                    logs()->info("Retrieved tenant URL from session: " . $tenantUrl);

                    $options["sharepoint"]["tenant_id"] = $tenantId;
                    logs()->info("Retrieved tenant ID from session: " . $tenantId);

                    $options["sharepoint"]["drives"] = Session::get($state . "_drives");

                    if ($personalUrl) {
                        // We have to go to step 3 for Personal SharePoint
                        logs()->info("Need to authenticate on: " . $personalUrl);
                        $oauthClient = $this->getSharepointOAuthClient($tenantId, $personalUrl);

                        $authUrl = $oauthClient->getAuthorizationUrl();

                        $newState = $oauthClient->getState();
                        logs()->info("New state: " . $newState);

                        Session::put($newState, Session::get($state));
                        Session::put($newState . "_tenant_id", Session::get($state . "_tenant_id"));
                        Session::put($newState . "_tenant_url", Session::get($state . "_tenant_url"));
                        Session::put($newState . "_personal_url", $personalUrl);
                        Session::put($newState . "_drives", Session::get($state . "_drives"));
                        Session::put($newState . "_step", 3);
//                Session::put($newState . "_scopes", implode(" ", $scopes));
                        Session::put($newState . "_auth", $options);

                        logs()->info("Redirecting to: " . $authUrl);

                        return redirect()->away($authUrl);
                    }

                    Session::forget($state . "_tenant_id");
                    Session::forget($state . "_drives");
                    Session::forget($state . "_tenant_url");
                } catch (\Exception $e) {
                    dd('Error in callback microsoft driver', $e);
                }
            }

        } else if ($step === 3) {
            // Third step : We authenticate using "SharePoint Personal" scopes (-my)

            $options = Session::get("{$state}_auth");
            $authCode = $request->get('code');
            abort_unless($authCode, 400, 'No code');

            $tenantId = Session::get($state . "_tenant_id");
            $personalUrl = Session::get($state . "_personal_url");
            $oauthClient = $this->getSharepointOAuthClient($tenantId, $personalUrl);

            $auth = $oauthClient->getAccessToken('authorization_code', [ 'code' => $authCode ])
                                ->jsonSerialize();
            $auth = array_merge($auth, [ 'deltaLinks' => [ $this->getNewPersonalDeltaLink($options) ] ]);
            $options["PersonalSharePoint"] = $auth;
        }


        $authCode = $request->get('code');
        abort_unless($authCode, 400, 'No code');

        $oauthClient = $this->getOAuthClient();

        try {
            $options = $oauthClient->getAccessToken('authorization_code', [ 'code' => $authCode ])->jsonSerialize();
            $options = array_merge($options, [ 'deltaLinks' => [ $this->getNewPersonalDeltaLink($options) ] ]);
        } catch (\Exception $e) {
            dd('Error in callback microsoft driver', $e);
        }

        Session::forget($state);
        Session::forget($state . "_step");
        $data = $this->processOptions($options);
        $dataStr = json_encode($data);

        return redirect()->away($redirectUrl . "&data=$dataStr");
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
            if ($site_id) {
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

        } catch (\Exception $e) {
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

        } catch (\Exception $e) {
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

        if (empty($deltaLinks)) {
            return false;
        }

        $files_changes = collect();

        foreach ($deltaLinks as $delta) {
            try {
                $resp = $graph->createRequest('GET', $delta)->execute()->getBody();
                $files_changes = $files_changes->merge(collect($resp['value'])->pluck('id'));
            } catch (\Exception $e) {
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
        if (empty($site_id)) {
            $endpoint = "/me/$endpoint";
        } else {
            $endpoint = "/sites/$site_id/$endpoint";
        }

        return $graph->createRequest('GET', $endpoint)->execute()->getRawBody();
    }

    public function refreshToken($config = null) : array
    {
        $config = $config ?? $this->default_config;

        $oauthClient = $this->getOAuthClient();

        return array_merge($config, $oauthClient->getAccessToken('refresh_token', [ 'refresh_token' => Arr::get($config, 'refresh_token') ])
                                                ->jsonSerialize());
    }
}
