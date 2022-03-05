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
     * @param string $fileName
     * @param bool $isConstMemory
     * @return string
     */
    public function export($sheetsConfig, $fileName = 'demo.xlsx', $isConstMemory = false)
    {
        foreach ($sheetsConfig as $k => $sheetConfig) {
            // 固定内存模式
            if (!$k) {
                if ($isConstMemory) {
                    $this->excel = $this->excel->constMemory($fileName, $sheetConfig['sheetName'], false); // wps
                } else {
                    $this->excel = $this->excel->fileName($fileName, $sheetConfig['sheetName']);
                }
            }
            $this->exportSheet($sheetConfig, $k);
        }
        $filePath = $this->excel->output();
        $this->closeExcel();
        return $filePath;
    }

    protected function exportSheet($sheetConfig, $isAdd = false)
    {
        if ($isAdd) {
            $this->excel->addSheet($sheetConfig['sheetName']);
        }
        // 设置sheet
        $this->setSheet($sheetConfig);

        // 设置header
        $dataHeaders = [];
        $endColIndex = -1;
        $rowIndex = 1;
        $this->setHeader($this->calculationColspan($sheetConfig['header']), $maxRow, $dataHeaders, $rowIndex, $endColIndex);
        // 导出数据
        $this->exportData($dataHeaders, $sheetConfig, $maxRow);

    }

    protected function setSheet($sheetConfig)
    {
        $sheet=new Sheet($sheetConfig,$this->excel);
        foreach ($sheetConfig as $method => $param) {
            if(method_exists($sheet,$method)){
                call_user_func_array([$sheet,$method],['param'=>$param]);
            }
        }
    }
    protected function setHeader(array $headers, &$maxRow = 1, &$dataHeaders = [], $rowIndex = 1, &$endColIndex = -1)
    {
        foreach ($headers as $head) {
            $head = DefaultConfig::getHeaderConfig($head);
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
                "align" => [Format::FORMAT_ALIGN_CENTER, Format::FORMAT_ALIGN_VERTICAL_CENTER],
                "bold" => true,
            ], $this->excel->getHandle());

            // 合并单元格 [A1:B3]
            $this->excel->mergeCells("{$startCol}{$startRow}:{$endCol}{$endRow}", $head['title'], $format->toResource());
            // 子集操作
            if (isset($head['children']) && $head['children']) {
                $endColIndex = $startColIndex - 1;
                $this->setHeader($head['children'], $maxRow, $dataHeaders, $rowIndex + $head['rowspan'], $endColIndex);
            }
        }
        return true;
    }

    protected function exportData($dataHeaders, $sheetConfig, $maxRow)
    {
        $i = 1;
        $keysIndex = array_flip(array_column($dataHeaders, 'key'));
        $this->excel->setCurrentLine($maxRow);
        do {
            $data = $sheetConfig['data'];
            $rs = true;
            if (is_callable($sheetConfig['data'])) {
                $rs = $sheetConfig['data']($i, $data);
            }
            if ($i == 1) {
                $pageSize = count($data);
            }
            // 格式化数据
            foreach ($data as $k => $v) {
                foreach ($dataHeaders as $colIndex => $head) {
                    // 格式化
                    if (is_callable($head['dataFormat'])) {
                        $newVal[$head['key']] = call_user_func_array($head['dataFormat'], [
                            'row' => $v,
                            'rowIndex' => $pageSize * ($i - 1) + $k,
                            'colIndex' => $keysIndex[$head['key']]
                        ]);
                    } else {
                        $newVal[$head['key']] = $v[$head['key']] ?? '';
                    }
                    // 样式



                }
                // 写入数据
                $this->excel->data([array_values($newVal)]);
            }
            $i++;
        } while (!$rs);


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

    public static function stringFromColumnIndex($index)
    {
        return Excel::stringFromColumnIndex($index);
    }

    public static function columnIndexFromString($index)
    {

        return Excel::columnIndexFromString($index);
    }

}