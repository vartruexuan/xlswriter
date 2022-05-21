<?php

use Vartruexuan\Xlswriter\Constant;
use Vartruexuan\Xlswriter\XlsWriter;

class ExcelTest extends \PHPUnit\Framework\TestCase
{

    public function testExport()
    {
        // 基础数据导出
        $data= [
            [
                'name' => '测试1',
                'url' => 'https://baidu.com',
                'image' => 'http://www.ykzx.cn/file/upload/1_181030100319_1.jpg',
                'date' => time(),
                'arr'=>[
                    'arr1'=>'测试下'
                ]
            ],
            [
                'name' => '测试2',
                'url' => 'https://baidu.com',
                'image' => 'http://www.ykzx.cn/file/upload/1-1305131621534Z.jpg',
                'date' => time(),
            ],
            [
                'name' => '测试2',
                'url' => 'https://baidu.com',
                'image' => __DIR__ . '/image/3.jpeg',
                'date' => time(),
            ],
            [
                'name' => '测试2',
                'url' => 'https://baidu.com',
                'image' => __DIR__ . '/image/1.jpeg',
                'date' => time(),
            ],
        ];

        $xls = new XlsWriter([
            'path' => __DIR__."/demo/"
        ]);
        $param = [
            [
                "sheetName" => 'sheet1',
                "sheetType" => "default",
                // 头信息
                "header" => [
                    [
                        "title" => "姓名",
                        "field" => "name",
                        "type" => ["text", [
                            'formatHandler' => [
                                'fontColor' => Constant::COLOR_RED,
                            ],
                        ]],
                    ],
                    [
                        "title" => "测试数组key",
                        "field" => "arr.arr1",
                        "type" => ["text", [
                            'formatHandler' => [
                                'fontColor' => Constant::COLOR_RED,
                            ],
                        ]],
                    ],
                    [
                        "title" => "链接",
                        "field" => "url",
                        "type" => ["url", [
                            "text" => "我是个链接",
                            "tooltip" => "提示文字",
                            'formatHandler' => [
                                'fontColor' => Constant::COLOR_RED,
                            ],
                        ]],

                    ],
                    [
                        "title" => "图片",
                        'field'=>'image',
                        "type" => ["image", [
                            'widthScale' => 0.1, // 宽度缩放比例
                            'heightScale' => 0.1,// 高度缩放比例
                        ]],
                    ],
                    [
                        "title" => "时间",
                        "field" => "date",
                        "type" => ["date", [
                            "dateFormat" => "yyyy-mm-dd hh:mm"
                        ]],
                    ],
                    [
                        "title" => "公式",
                        "field" => "formula",
                        "type" => ["formula", [
                        ]],
                        "dataFormat" => function ($row, $rowIndex, $colIndex, $keysIndex) {
                            $colStr = XlsWriter::stringFromColumnIndex($keysIndex['index']);
                            $colStrEnd = XlsWriter::stringFromColumnIndex($keysIndex['index']);
                            return "=SUM({$colStr}{$rowIndex}:{$colStrEnd}{$rowIndex})";
                        }
                    ],
                ],
                // 数据
                "data" => $data,
            ]
        ];
        $filePath = $xls->export($param);
        $this->assertEquals(file_exists($filePath),true,'文件地址');
    }
    public function testImageMoreExport()
    {
        // 基础数据导出
        $data= [
            [
                'name' => '测试1',
                'url' => 'https://baidu.com',
                'image' =>['http://www.ykzx.cn/file/upload/1_181030100319_1.jpg','http://www.ykzx.cn/file/upload/1-1305131621534Z.jpg',__DIR__ . '/image/3.jpeg'],
                'date' => time(),
                'arr'=>[
                    'arr1'=>'测试下'
                ]
            ],

        ];

        $xls = new XlsWriter([
            'path' => __DIR__."/demo/"
        ]);
        $param = [
            [
                "sheetName" => 'sheet1',
                "sheetType" => "default",
                // 头信息
                "header" => [
                    [
                        "title" => "姓名",
                        "field" => "name",
                        "type" => ["text", [
                            'formatHandler' => [
                                'fontColor' => Constant::COLOR_RED,
                            ],
                        ]],
                    ],
                    [
                        "title" => "测试数组key",
                        "field" => "arr.arr1",
                        "type" => ["text", [
                            'formatHandler' => [
                                'fontColor' => Constant::COLOR_RED,
                            ],
                        ]],
                    ],
                    [
                        "title" => "链接",
                        "field" => "url",
                        "type" => ["url", [
                            "text" => "我是个链接",
                            "tooltip" => "提示文字",
                            'formatHandler' => [
                                'fontColor' => Constant::COLOR_RED,
                            ],
                        ]],

                    ],
                    [
                        "title" => "图片",
                        'field'=>'image',
                        "type" => ["image", [
                            'widthScale' => 0.1, // 宽度缩放比例
                            'heightScale' => 0.1,// 高度缩放比例
                        ]],
                    ],
                    [
                        "title" => "时间",
                        "field" => "date",
                        "type" => ["date", [
                            "dateFormat" => "yyyy-mm-dd hh:mm"
                        ]],
                    ],
                ],
                // 数据
                "data" => $data,
            ]
        ];
        $filePath = $xls->export($param,'moreImage.xlsx');
        $this->assertEquals(file_exists($filePath),true,'文件地址');
    }


}