<?php
/*
 * This file is part of the vartruexuan/xlswriter.
 *
 * (c) vartruexuan <guozhaoxuanx@163.com>
 *
 * This source file is subject to the MIT license that is bundled
 * with this source code in the file LICENSE.
 */
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