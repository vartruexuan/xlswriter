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

use \Vtiful\Kernel\Excel;
use \Vtiful\Kernel\Format;

class XlsWriter
{
    /**
     * @var \Vtiful\Kernel\Excel
     */
    private $excel = null;
    private $config = [];

    public function __construct($config = [])
    {
        $this->setConfig($config);
        $this->initExcel();
    }

    /**
     * @param array $sheetsConfig
     * @param $fileName
     * @param $isConstMemory
     * @return string
     */
    public function export(array $sheetsConfig, $fileName = 'demo.xlsx', $isConstMemory = false)
    {
        // 固定内存模式
        if ($isConstMemory) {
            $this->excel = $this->excel->constMemory($fileName);
        } else {
            $this->excel = $this->excel->fileName($fileName);
        }
        foreach ($sheetsConfig as $sheetConfig) {
            $this->exportSheet($sheetConfig);
        }
        $filePath = $this->excel->output();
        $this->closeExcel();
        return $filePath;
    }

    public function exportSheet($sheetConfig)
    {
        // 设置header
        $this->setHeader($this->calculationColspan($sheetConfig['header']));
        // 导出数据

        // 处理回调

    }

    public function getConfig()
    {
        return $this->config;
    }

    private function setConfig($config)
    {
        $this->config = array_merge($this->config, $config);
    }

    /**
     * @param array $headers
     * @param $maxRow
     * @param $dataHeaders
     * @param $rowIndex
     * @param $endColIndex
     * @return bool
     */
    private function setHeader(array $headers, &$maxRow = 1, &$dataHeaders = [], $rowIndex = 1, &$endColIndex = -1)
    {
        foreach ($headers as $head) {
            $head = Config::getHeaderConfig($head);
            if ($head['key']) {
                $dataHeaders[] = [
                    'key' => $head['key'],
                    'type' => $head['type'],
                    'dataFormat' => $head['dataFormat'] ?? null,
                ];
            }
            $startColIndex = $endColIndex + 1;
            $endColIndex = $startColIndex + $head['colspan'] - 1;

            $startCol = self::stringFromColumnIndex($startColIndex);
            $endCol = self::stringFromColumnIndex($endColIndex);

            if ($head['width']) {
                $this->excel->setColumn("{$startCol}:{$endCol}", $head['width']);
            }

            $startRow = $rowIndex;
            $endRow = $startRow + $head['rowspan'] - 1;
            if ($endRow > $maxRow) {
                $maxRow = $endRow;
            }

            // 默认样式
            $format = new StyleFormat([
                "align"=>[Format::FORMAT_ALIGN_CENTER, Format::FORMAT_ALIGN_VERTICAL_CENTER],
                "bold"=>true,
            ],$this->excel->getHandle());

            // 合并单元格 [A1:B3]
            $this->excel->mergeCells("{$startCol}{$startRow}:{$endCol}{$endRow}", $head['title'],  $format->toResource());

            // 子集操作
            if (isset($head['children']) && $head['children']) {
                $endColIndex = $startColIndex - 1;
                $this->setHeader($head['children'], $maxRow, $dataHeaders, $rowIndex + $head['rowspan'], $endColIndex);
            }
        }
        return true;
    }

    /**
     * 计算colspan(多级)
     *
     * @param     $header
     * @param int $level
     *
     * @return mixed
     */
    protected function calculationColspan($header, $level = 1)
    {
        // 子集colspan之和
        foreach ($header as &$head) {
            $children = $head['children'] ?? [];
            if ($children) {
                $head['children'] = $this->calculationColspan($children, $level + 1);
                $head['colspan'] = array_sum(array_column($head['children'], 'colspan'));
            } else {
                $head['colspan'] = 1;
            }
        }
        return $header;
    }

    private function initExcel()
    {
        if (!$this->excel instanceof \Vtiful\Kernel\Excel) {
            $this->excel = new Excel($this->config);
        }
        return $this->excel;
    }

    private function closeExcel()
    {
        return $this->excel->close();
    }

    public function stringFromColumnIndex($index)
    {
        return Excel::stringFromColumnIndex($index);
    }

    public function setSheetZoom($zoom)
    {
        return $this->excel->zoom($zoom);
    }

    public function setSheetHide()
    {
        return $this->setCurrentSheetHide();
    }

    public function setSheetGridline($gridline = \Vtiful\Kernel\Excel::GRIDLINES_HIDE_ALL)
    {
        return $this->excel->gridline($gridline);
    }
}