<?php

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