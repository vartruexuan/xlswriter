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


    const SKIP_NONE = 0x00;         // 不忽略任何单元格、行
    const SKIP_EMPTY_ROW = 0x01;    // 忽略空行
    const SKIP_EMPTY_CELLS = 0x02;  // 忽略空单元格（肉眼观察单元格内无数据，并不代表单元格未定义、未使用）
    const SKIP_EMPTY_VALUE = 0X100; // 忽略单元格空数据


    const TYPE_STRING = 0x01;    // 字符串
    const TYPE_INT = 0x02;       // 整型
    const TYPE_DOUBLE = 0x04;    // 浮点型
    const TYPE_TIMESTAMP = 0x08; // 时间戳，可以将 xlsx 文件中的格式化时间字符转为时间戳

}