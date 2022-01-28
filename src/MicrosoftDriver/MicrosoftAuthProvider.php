<?php


namespace DynamicScreen\Microsoft\MicrosoftDriver;

use DynamicScreen\SdkPhp\Interfaces\IModule;
use Illuminate\Support\Arr;
use Unsplash\HttpClient as UnsplashClient;
use Unsplash\Photo;
use DynamicScreen\SdkPhp\Handlers\OAuthProviderHandler;

class MicrosoftAuthProvider extends OAuthProviderHandler
{
    protected static string $provider = 'unsplash';

    public function __construct(IModule $module, $config = null)
    {
        parent::__construct($module, $config);
    }

    public function signin($callbackUrl = null)
    {
        $callbackUrl = $callbackUrl ?? route('api.oauth.callback');
        $scopes = ['public', 'read_user'];
        $config = ['callback_url' => $callbackUrl];

        $this->initConnection($config);

        return UnsplashClient::$connection->getConnectionUrl($scopes);
    }

    public function callback($request, $redirectUrl = null)
    {
        $code = $request->input('code');

        $this->initConnection(['callback_url' => route('api.oauth.callback')]);
        $accessToken = UnsplashClient::$connection->generateToken($code);

        $data = $this->processOptions($accessToken->jsonSerialize());

        $dataStr = json_encode($data);

        return redirect()->away($redirectUrl ."&data=$dataStr");
    }

    public function getUserInfos($config = null)
    {
        $config = $config ?? $this->default_config;

        $this->initConnection($config);
        $infos = UnsplashClient::$connection->getResourceOwner()->toArray();
        if(isset($infos['errors'])){
            throw new \Exception($infos['error'][0]);
        }
        return $infos;
    }

    public function testConnection($config = null)
    {
        $config = $config ?? $this->default_config;

        return $this->getUserInfos($config);
    }

    public function initConnection($config = null)
    {
        $provider = $this->getProviderIdentifier();
        $config = $config ?? $this->default_config;

        if (isset($config['callback_url'])) {
            UnsplashClient::init([
                'applicationId' => config("services.$provider.client_id"),
                'secret' => config("services.$provider.client_secret"),
                'callbackUrl' => $config['callback_url'],
                'utmSource' => config("services.$provider.app_name")
            ]);
        }
        else if (isset($config['access_token'])) {
            UnsplashClient::init([
                'applicationId' => config("services.$provider.client_id"),
                'secret' => config("services.$provider.client_secret"),
                'utmSource' => config("services.$provider.app_name")
            ],[
                'access_token' => $config['access_token'],
                'expires_in' => 30000
            ]);
        }
    }

//    Used for poc
    public function provideData($settings = [])
    {
        $this->addData('me', fn () => $this->getUserInfos());
    }

    public function getRandomPhoto($options, $config = null)
    {
        $config = $config ?? $this->default_config;

        $category = Arr::get($options, 'category');

        $this->initConnection($config);
        $randPhotos = Photo::random(['query' => $category, 'orientation' => 'landscape']);
        return ['photographer' => $randPhotos->user['name'], 'url' => $randPhotos->urls['raw']];
    }
}
