<?php

namespace Vartruexuan\Xlswriter;

class BaseExcel
{

    protected $config = [];

    protected function getConfig()
    {
        return $this->config;
    }
    protected function setConfig($config)
    {
        $this->config = array_merge($this->config, $config);
    }

}