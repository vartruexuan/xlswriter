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

use Intervention\Image\ImageManagerStatic as Image;
use mysql_xdevapi\Exception;
use Vartruexuan\Xlswriter\common\exception\XlswriterException;
use Vartruexuan\Xlswriter\common\utils\Arr;
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

    /**
     * 导入
     *
     * @param $config
     * @param $fileName
     *
     * @return array
     */
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

    /**
     * 导出sheet
     *
     * @param $sheetConfig
     * @param $isAdd
     * @param $isConstMemory
     *
     * @return void
     */
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
        }

    }

    /**
     * 设置sheet
     *
     * @param $sheetConfig
     *
     * @return void
     */
    protected function setSheet($sheetConfig)
    {
        $sheet = new Sheet($sheetConfig, $this->excel);
        foreach ($sheetConfig as $method => $param) {
            if (method_exists($sheet, $method)) {
                call_user_func_array([$sheet, $method], ['param' => $param]);
            }
        }
    }

    /**
     * 设置头信息
     *
     * @param array $headers
     * @param       $maxRow
     * @param       $dataHeaders
     * @param       $rowIndex
     * @param       $endColIndex
     *
     * @return bool
     */
    protected function setHeader(array $headers, &$maxRow = 1, &$dataHeaders = [], $rowIndex = 1, &$endColIndex = -1)
    {
        foreach ($headers as $head) {
            $head = DefaultConfig::getHeaderConfig($head);
            if ($head['field']) {
                $dataHeaders[] = [
                    'key' => $head['key'] ?? $head['field'],
                    'type' => $head['type'],
                    'field' => $head['field'],
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
            $format = $this->getStyleFormat([
                "align" => [Format::FORMAT_ALIGN_CENTER, Format::FORMAT_ALIGN_VERTICAL_CENTER],
                "bold" => true,
            ]);
            // 合并单元格 [A1:B3]
            $this->excel->mergeCells("{$startCol}{$startRow}:{$endCol}{$endRow}", $head['title'], $format);
            // 子集操作
            if (isset($head['children']) && $head['children']) {
                $endColIndex = $startColIndex - 1;
                $this->setHeader($head['children'], $maxRow, $dataHeaders, $rowIndex + $head['rowspan'], $endColIndex);
            }
        }
        return true;
    }

    /**
     * 导出数据
     *
     * @param $dataHeaders
     * @param $sheetConfig
     * @param $maxRow
     *
     * @return void
     */
    protected function exportData($dataHeaders, $sheetConfig, $maxRow)
    {
        $i = 1;
        $this->excel->setCurrentLine($maxRow - 1);
        do {
            $data = $sheetConfig['data'];
            $isWhile = false;
            if (is_callable($sheetConfig['data'])) {
                $data = $sheetConfig['data']($this, $i, $dataHeaders, $isWhile, $maxRow);
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
        $startRowIndex = $this->excel->getCurrentLine();
        $keysIndex = array_flip(array_column($dataHeaders, 'key'));
        // 格式化数据
        foreach ($data as $k => $v) {
            $rowIndex = $startRowIndex + $k + 1;
            // 行处理
            if (isset($sheetConfig['rowFormat']['merge'])) {
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
                    $newVal[$colIndex] = Arr::get($v, $head['field'], '');
                }

                $dataType = $head['type'];
                $dataTypeParam = [];
                if (is_array($dataType)) {
                    $dataType = $head['type'][0];
                    $dataTypeParam = $head['type'][1] ?? [];
                }
                $this->insertCell($dataType ?? 'text', $rowIndex, $keysIndex[$head['key']], $newVal[$colIndex], $dataTypeParam);
            }
            // $newData[$k] = array_values($newVal ?? []);
        }
        // 写入数据
        //$this->excel->data(array_values($newData));
        $this->rowFormat($mergeList ?? []);

    }

    /**
     * 行格式化
     *
     * @param $mergeList
     *
     * @return void
     */
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
     * @param $option    附加参数
     *
     * @return void
     */
    public function insertCell($dataType, $rowIndex, $colIndex, $value, $option = [])
    {
        $callFunName = "insert{$dataType}";
        if (!method_exists($this, $callFunName)) {
            throw  new XlswriterException("{$dataType} type not exists");
        }
        call_user_func_array([$this, $callFunName], [
            'rowIndex' => $rowIndex,
            'colIndex' => $colIndex,
            'value' => $value,
            'option' => $option,
        ]);
        return $this;
    }

    /**
     * 插入文本
     *
     * @param $rowIndex
     * @param $colIndex
     * @param $value
     * @param $option
     *
     * @return Excel
     */
    public function insertText($rowIndex, $colIndex, $value, $option = [])
    {
        $option = array_merge([
            'format' => null,
            'formatHandler' => null,
        ], $option);
        $option['formatHandler'] = $this->getStyleFormat($option['formatHandler']);
        $this->excel->insertText($rowIndex, $colIndex, $value, $option['format'], $option['formatHandler']);
        return $this;
    }

    /**
     * 插入图片
     *
     * @param $rowIndex
     * @param $colIndex
     * @param $value
     * @param $param
     *
     * @return Excel
     */
    public function insertImage($rowIndex, $colIndex, $value, $option = [])
    {
        $option = array_merge([
            'widthScale' => 1, // 宽度缩放比例
            'heightScale' => 1,// 高度缩放比例
        ], $option);
        // 图片特殊处理
        if (!file_exists($value)) {
            // 下载图片
            return $this->insertText($rowIndex, $colIndex, $value, $option);
        }
        $image = Image::make($value);
        $height = $image->height();
        $width = $image->width();
        // 设置行高
        if ($height) {
            $height = $height * ($option['heightScale'] ?? 1); // 比例计算
            $this->excel->setRow("A:" . ($rowIndex + 1), $height);
        }
        // 设置列宽
        if ($width) {
            $width = $width * ($option['widthScale'] ?? 1); // 比例计算
            $colIndexStr = self::stringFromColumnIndex($colIndex);
            $this->excel->setColumn("{$colIndexStr}:{$colIndexStr}", ceil($width / 8) + 5);
        }
        return $this->excel->insertImage($rowIndex, $colIndex, $value, $option['widthScale'], $option['heightScale']);
    }

    /**
     * 插入链接
     *
     * @param $rowIndex
     * @param $colIndex
     * @param $value
     * @param $option
     *
     * @return $this
     */
    public function insertUrl($rowIndex, $colIndex, $value, $option = [])
    {
        $option = array_merge([
            'text' => null,// 链接文字
            'tooltip' => null,// 链接提示
            'formatHandler' => null,
        ], $option);
        $this->excel->insertUrl($rowIndex, $colIndex, $value, $option['text'], $option['tooltip'], $this->getStyleFormat($option['formatHandler']));
        return $this;
    }

    /**
     * 插入公式
     *
     * @param $rowIndex
     * @param $colIndex
     * @param $value
     * @param $param
     *
     * @return $this
     */
    public function insertFormula($rowIndex, $colIndex, $value, $param = [])
    {
        $option = array_merge([
            'formatHandler' => null,
        ], $param);
        $this->excel->insertFormula($rowIndex, $colIndex, $value, $this->getStyleFormat($option['formatHandler']));
        return $this;
    }

    /**
     * 插入时间
     *
     * @param $rowIndex
     * @param $colIndex
     * @param $value
     * @param $param
     *
     * @return $this
     */
    public function insertDate($rowIndex, $colIndex, $value, $param = [])
    {
        $option = array_merge([
            'dateFormat' => null,
            'formatHandler' => null,
        ], $param);
        $this->excel->insertDate($rowIndex, $colIndex, $value, $option['dateFormat'], $this->getStyleFormat($option['formatHandler']));
        return $this;
    }

    /**
     * 获取样式资源
     *
     * @param $style
     *
     * @return resource
     */
    public function getStyleFormat($style)
    {
        return $style ? (new StyleFormat($style, $this->excel->getHandle()))->toResource() : null;
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
        static $fields = [];
        // 子集colspan之和
        foreach ($header as &$head) {
            $head = array_merge($head, [
                'children' => $head['children'] ?? [],
                'colspan' => 1,
            ]);
            $field = Arr::get($head, 'field');
            if ($field) {
                // 生成key 标识
                $fields[$field] = ($fields[$field] ?? 0) + 1;
                $head['key'] = $fields[$field] > 0 ? ($field . '-' . $fields[$field]) : $field;
            }
            if ($head['children']) {
                $head['children'] = $this->calculationColspan($head['children'], $level + 1);
                $head['colspan'] = array_sum(array_column($head['children'], 'colspan'));
            }
        }
        return $header;
    }

    /**
     * 初始化excel
     *
     * @return Excel
     * @throws \Exception
     */
    protected function initExcel()
    {
        if (!$this->excel instanceof \Vtiful\Kernel\Excel) {
            $this->excel = new Excel($this->getConfig());
        }
        return $this->excel;
    }

    /**
     * 关闭excel
     *
     * @return mixed
     */
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