<?php

namespace Vartruexuan\Xlswriter;

class Sheet
{
    /**
     * @var \Vtiful\Kernel\Excel
     */
    private $excel = null;

    private $config = null;

    public function __construct($sheetConfig, $excel)
    {
        $this->config = $sheetConfig;
        $this->excel = $excel;
        $this->excel->checkoutSheet($sheetConfig['sheetName']);
    }

    /**
     * 工作表缩放
     *   范围：10 <= $scale <= 400
     *   默认值： 100
     *   缩放比例不影响打印比例
     *
     * @param $zoom
     * @return \Vtiful\Kernel\Excel
     */
    public function zoom($zoom)
    {
        $this->excel->zoom($zoom);
        return $this;
    }

    public function hide($isHide = true)
    {
        if ($isHide) {
            $this->excel->setCurrentSheetHide();
        }
        return $this;
    }
    public function gridline($gridline = \Vtiful\Kernel\Excel::GRIDLINES_HIDE_ALL)
    {
         $this->excel->gridline($gridline);
         return $this;
    }
    public function defaultFormat($style)
    {

        $format=new StyleFormat($style,$this->excel->getHandle());
        $this->excel->defaultFormat($format->toResource());
        return $this;
    }
}