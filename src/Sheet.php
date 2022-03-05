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

class Sheet extends BaseExcel
{
    /**
     * @var \Vtiful\Kernel\Excel
     */
    protected $excel = null;

    public function __construct($sheetConfig, $excel)
    {
        $this->setConfig($sheetConfig);
        $this->excel = $excel;
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

    public function gridline($gridline = Constant::GRIDLINES_HIDE_ALL)
    {
        $this->excel->gridline($gridline);
        return $this;
    }

    public function defaultFormat($style)
    {
        $format = new StyleFormat($style, $this->excel->getHandle());
        $this->excel->defaultFormat($format->toResource());
        return $this;
    }

    public function protection($password)
    {
        $this->excel->protection($password);
        return $this;
    }
}