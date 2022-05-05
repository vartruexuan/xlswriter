<?php

namespace Vartruexuan\Xlswriter\Common\Utils;
use GuzzleHttp\Client;
use GuzzleHttp\ClientInterface;
use Vartruexuan\Xlswriter\Common\Library\Singleton;

class HttpClient
{

    use Singleton;

    /**
     * @var Client
     */
    protected $httpClient = null;

    /**
     * 并发请求
     *
     * @param $requestParam
     *          [
     *              ["url"=>"",'method'=>'','option'=>[]],
     *          ]
     *
     * @return array
     * @throws \Throwable
     */
    public function multiRequest($requestParam)
    {
        foreach ($requestParam as $key => $param) {
            $promises[$key] = $this->getHttpClient()->requestAsync($param["method"], $param['url'], $param['option']);
        }
        return \GuzzleHttp\Promise\Utils::unwrap($promises);
    }

    public function getHttpClient()
    {
        if (!($this->httpClient instanceof ClientInterface)) {
            $this->httpClient = new Client();
        }

        return $this->httpClient;
    }

}