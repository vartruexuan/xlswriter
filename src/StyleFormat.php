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

class StyleFormat
{
    private $config = [];

    /**
     * @var \Vtiful\Kernel\Format
     */
    private $format = null;

    public function __construct($config, $fileHandle)
    {
        $this->config = $config;
        $this->format = new Format($fileHandle);
        $this->initFormat();
    }

    protected function initFormat()
    {
        foreach ($this->config as $k => $v) {
            if (method_exists($this, $k)) {
                call_user_func_array([$this, $k], ['param' => $v]);
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

    protected function border($param)
    {
        $this->format->border($param);
        return $this;
    }

    protected function background($param)
    {
        $this->format->background($param['color'],$param['pattern']);
        return $this;
    }

    protected function font($param)
    {
        $this->format->font($param);
        return $this;
    }

    protected function number($param)
    {
        $this->format->number($param);
        return $this;
    }

    protected function align($param)
    {
        $this->format->align(...$param);
        return $this;
    }

    protected function underline($param)
    {
        $this->format->underline($param);
        return $this;
    }

}