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

/**
 * 图表
 */
class Chart extends BaseExcel
{
    /**
     * @var \Vtiful\Kernel\Chart
     */
    protected $chart=null;
    public function __construct($config, $fileHandle,$type)
    {
        $this->setConfig($config);
        $this->chart = new \Vtiful\Kernel\Chart($fileHandle,$type);
    }
    public function toResource()
    {
        return $this->chart->toResource();
    }

}