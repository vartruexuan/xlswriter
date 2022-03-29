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

class XlsWriter extends BaseExcel
{
    /**
     * @var \Vtiful\Kernel\Excel
     */
    public $excel = null;

    protected $config = [
        "path" => "./", // 导出地址
    ];

    public function __construct($config = [])
    {
        $this->setConfig($config);
        $this->initExcel();
    }

    /**
     * @param array  $sheetsConfig
     * @param string $fileName
     * @param bool   $isConstMemory
     *
     * @return string
     */
    public function export($sheetsConfig, $fileName = 'demo.xlsx', $isConstMemory = false)
    {
        set_time_limit(0);
        foreach ($sheetsConfig as $k => $sheetConfig) {

            if (!$k) {
                // 固定内存模式
                if ($isConstMemory) {
                    $this->excel = $this->excel->constMemory($fileName, $sheetConfig['sheetName'], false); // wps
                } else {
                    $this->excel = $this->excel->fileName($fileName, $sheetConfig['sheetName']);
                }
            }
            $this->exportSheet($sheetConfig, $k, $isConstMemory);
        }
        $filePath = $this->excel->output();
        $this->closeExcel();
        return $filePath;
    }

    public function import($config, $fileName = 'demo.xlsx')
    {
        $this->excel = $this->excel->openFile($fileName);
        $sheetList = $this->excel->sheetList();

        $list = [];
        foreach ($sheetList as $sheet) {
            $data = $this->excel->openSheet($sheet)->getSheetData();
            $list[$sheet] = $data;
        }
        return $list;
    }

    protected function exportSheet($sheetConfig, $isAdd = false, $isConstMemory = false)
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
        if ($isConstMemory) {
            // $this->excel->header(array_column($sheetConfig['header'],'title'));
        }

        $sheetType = $sheetConfig['sheetType'] ?? 'default';
        if ($sheetType == 'default') {
            $this->setHeader($this->calculationColspan($sheetConfig['header']), $maxRow, $dataHeaders, $rowIndex, $endColIndex);
            // 导出数据
            $this->exportData($dataHeaders, $sheetConfig, $maxRow);
        } else {
            // 图表操作
            $this->exportChart($sheetConfig);

        }

    }

    protected function setSheet($sheetConfig)
    {
        $sheet = new Sheet($sheetConfig, $this->excel);
        foreach ($sheetConfig as $method => $param) {
            if (method_exists($sheet, $method)) {
                call_user_func_array([$sheet, $method], ['param' => $param]);
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
                    'field' => $head['field'] ?? $head['key'],
                    'dataFormat' => $head['dataFormat'] ?? null,
                ];
            }
            $startColIndex = $endColIndex + 1;
            $endColIndex = $startColIndex + $head['colspan'] - 1;

            $startCol = self::stringFromColumnIndex($startColIndex);
            $endCol = self::stringFromColumnIndex($endColIndex);

            $head['width'] = $head['width'] > 0 ? $head['width'] : strlen($head['title']) + 5;
            $this->excel->setColumn("{$startCol}:{$endCol}", $head['width']);

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

        $this->excel->setCurrentLine($maxRow);
        do {
            $data = $sheetConfig['data'];
            $isWhile = false;
            if (is_callable($sheetConfig['data'])) {
                $data = $sheetConfig['data']($this, $i, $dataHeaders, $isWhile,$maxRow);
            }
            if ($i == 1) {
                $pageSize = count($data);
            }
            $startRowIndex = $maxRow + $pageSize * ($i - 1);
            // 格式化数据
            $this->writerData($data, $dataHeaders, $sheetConfig, $startRowIndex);
            $i++;
        } while ($isWhile);

    }

    protected function exportChart($sheetConfig)
    {
        return false;
    }

    /**
     * 写入数据
     *
     * @param $data
     * @param $dataHeaders
     *
     * @return void
     */
    public function writerData($data, $dataHeaders, $sheetConfig, $startRowIndex = 0)
    {
        // $startRowIndex=$this->excel->getCurrentLine();
        $keysIndex = array_flip(array_column($dataHeaders, 'key'));
        // 格式化数据
        $newData = [];
        foreach ($data as $k => $v) {
            $rowIndex = $startRowIndex + $k;
            // 行处理
            if(isset($sheetConfig['rowFormat']['merge'])){
                $mergeList = array_merge($mergeList ?? [], $sheetConfig['rowFormat']['merge']($data[$k - 1] ?? null, $v, $data[$k + 1] ?? null, $rowIndex + 1, $keysIndex));
            }
            foreach ($dataHeaders as $colIndex => $head) {

                // 格式化
                if (is_callable($head['dataFormat'])) {
                    $newVal[$colIndex] = call_user_func_array($head['dataFormat'], [
                        'row' => $v,
                        'rowIndex' => $rowIndex,
                        'colIndex' => $keysIndex[$head['key']],
                        "keysIndex" => $keysIndex
                    ]);
                } else {
                    $newVal[$colIndex] = $v[$head['field'] ?? $head['key']] ?? '';
                }

                $dataType = $head['type'];
                $dataTypeParam = [];
                if (is_array($dataType)) {
                    $dataType = $head['type'][0];
                    $dataTypeParam = $head['type'][1] ?? [];
                }
                // 数据类型
                $this->insertCell($dataType ?? 'text', $rowIndex, $keysIndex[$head['key']], $newVal[$colIndex], $dataTypeParam);
                // 列样式
            }
            $newData[$k] = array_values($newVal ?? []);

        }
        // 写入数据
        //$this->excel->data(array_values($newData));
        $this->rowFormat($mergeList ?? []);

    }

    protected function rowFormat($mergeList)
    {
        // 计算最终合并

        // 合并行单元格
        foreach ($mergeList as $merge) {
            // 默认样式
            $format = new StyleFormat([
                "align" => [Format::FORMAT_ALIGN_CENTER, Format::FORMAT_ALIGN_VERTICAL_CENTER],
            ], $this->excel->getHandle());
            $startCol = self::stringFromColumnIndex($merge['col_start']);
            $endCol = self::stringFromColumnIndex($merge['col_end']);
            $startRow = $merge['row_start'];
            $endRow = $merge['row_end'];
            // 合并单元格 [A1:B3]
            $this->excel->mergeCells("{$startCol}{$startRow}:{$endCol}{$endRow}", $merge['col_value'], $format->toResource());
        }
    }


    /**
     * 插入单元格
     *
     * @param $dataType 数据类型
     * @param $rowIndex 行下标
     * @param $colIndex 列下标
     * @param $param    附加参数
     *
     * @return void
     */
    public function insertCell($dataType, $rowIndex, $colIndex, $value, $param = [])
    {
        // 数据类型
        $dataTypeList = [
            // 文本
            'text' => [
                'format' => null,
                'formatHandle' => null,
            ],
            // 链接
            'url' => [
                'text' => null,// 链接文字
                'tooltip' => null,// 链接提示
                'formatHandle' => null,
            ],
            // 公式
            'formula' => [
                'formatHandle' => null,
            ],
            // 时间
            'date' => [
                'dateFormat' => 'yyyy-mm-dd hh:mm:ss',// 时间格式
                'formatHandle' => null,
            ],
            // 图片
            'image' => [
                'widthScale' => 1, // 宽度缩放比例
                'heightScale' => 1,// 高度缩放比例
            ],
        ];
        $dataTypes = array_keys($dataTypeList);
        $dataType = in_array($dataType, $dataTypes) ? $dataType : $dataTypes[0];
        // 排除非附加属性字段
        $param = array_intersect_key($param, $dataTypeList[$dataType]);
        $param = array_merge([
            'row' => $rowIndex,
            'column' => $colIndex,
            'value' => $value,
        ], $dataTypeList[$dataType], $param);
        $dataType = ucfirst($dataType);
        call_user_func_array([$this->excel, "insert{$dataType}"], $param);
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

    protected function initExcel()
    {
        if (!$this->excel instanceof \Vtiful\Kernel\Excel) {
            $this->excel = new Excel($this->getConfig());
        }
        return $this->excel;
    }

    protected function closeExcel()
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