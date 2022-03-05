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

use Vtiful\Kernel\Excel;

class Constant
{
    // 数据类型


    // 样式
    const FORMAT_BORDER_THIN=\Vtiful\Kernel\Format::BORDER_THIN;


    // 颜色
     const COLOR_BLACK = \Vtiful\Kernel\Format::COLOR_BLACK;
     const COLOR_BLUE = \Vtiful\Kernel\Format::COLOR_BLUE;
     const COLOR_BROWN = \Vtiful\Kernel\Format::COLOR_BROWN;
     const COLOR_CYAN = \Vtiful\Kernel\Format::COLOR_CYAN;
     const COLOR_GRAY =\Vtiful\Kernel\Format::COLOR_GRAY;
     const COLOR_GREEN = \Vtiful\Kernel\Format::COLOR_GREEN;
     const COLOR_LIME = \Vtiful\Kernel\Format::COLOR_LIME;
     const COLOR_MAGENTA =\Vtiful\Kernel\Format::COLOR_MAGENTA;
     const COLOR_NAVY = \Vtiful\Kernel\Format::COLOR_NAVY;
     const COLOR_ORANGE = \Vtiful\Kernel\Format::COLOR_ORANGE;
     const COLOR_PINK = \Vtiful\Kernel\Format::COLOR_PINK;
     const COLOR_PURPLE = \Vtiful\Kernel\Format::COLOR_PURPLE;
     const COLOR_RED = \Vtiful\Kernel\Format::COLOR_RED;
     const COLOR_SILVER = \Vtiful\Kernel\Format::COLOR_SILVER;
     const COLOR_WHITE = \Vtiful\Kernel\Format::COLOR_WHITE;
     const COLOR_YELLOW = \Vtiful\Kernel\Format::COLOR_YELLOW;

    // 网格线
    const GRIDLINES_HIDE_ALL    = Excel::GRIDLINES_HIDE_ALL; // 隐藏 屏幕网格线 和 打印网格线
    const GRIDLINES_SHOW_SCREEN = Excel::GRIDLINES_SHOW_SCREEN; // 显示屏幕网格线
    const GRIDLINES_SHOW_PRINT  = Excel::GRIDLINES_SHOW_PRINT; // 显示打印网格线
    const GRIDLINES_SHOW_ALL    = Excel::GRIDLINES_SHOW_ALL; // 显示 屏幕网格线 和 打印网格线

}