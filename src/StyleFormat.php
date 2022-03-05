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

use Vtiful\Kernel\Format;

class StyleFormat extends BaseExcel
{
    /**
     * @var \Vtiful\Kernel\Format
     */
    protected $format = null;

    public function __construct($config, $fileHandle)
    {
        $this->setConfig($config);
        $this->format = new Format($fileHandle);
        $this->initFormat();
    }

    protected function initFormat()
    {
        foreach ($this->getConfig() as $method => $param) {
            if (method_exists($this, $method)) {
                call_user_func_array([$this, $method], ['param' => $param]);
            }
        }
    }

    public function toResource()
    {
        return $this->format->toResource();
    }

    protected function bold($param)
    {
        if ($param) {
            $this->format->bold();
        }
        return $this;
    }

    protected function wrap($param)
    {
        if ($param) {
            $this->format->wrap();
        }
        return $this;
    }

    protected function fontColor($fontColor)
    {
        $this->format->fontColor($fontColor);
        return $this;
    }

    protected function fontSize($fontSize)
    {
        $this->format->fontSize($fontSize);
        return $this;
    }

    protected function border($style)
    {
        $this->format->border($style);
        return $this;
    }

    protected function background($param)
    {
        $this->format->background($param['color'], $param['pattern']);
        return $this;
    }

    protected function font($fontName)
    {
        $this->format->font($fontName);
        return $this;
    }

    protected function number($format)
    {
        $this->format->number($format);
        return $this;
    }

    protected function align($param)
    {
        $this->format->align(...$param);
        return $this;
    }

    protected function underline($style)
    {
        $this->format->underline($style);
        return $this;
    }

    protected function unlocked()
    {
        $this->format->unlocked();
        return $this;
    }

}