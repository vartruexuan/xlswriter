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

class DefaultConfig
{

    public static function getSheetConfig($config){

        return array_merge([

            "sheetName" =>"",
            "sheetType" =>"",
            "data" =>[
                [
                    "name"=>"测试",
                ]
            ],
            "zoom" =>"",
            "gridline" =>"",
            "isHide" =>false,
            "rowStyle" =>[],
            "protection" =>[],
            "validation" => [
                "type" =>"Validation::TYPE_LIST",
                "config" =>""
            ]

        ],$config);
    }
    public static function getHeaderConfig($config=[])
    {
        return array_merge([
            "title" => "",
            "type" => "string",
            "key" => "",
            "style" => "",
            "dataFormat" =>null,
            "rowspan" => 1,
            "colspan" => 1,
            "children" => [],
            "width" => 0
        ], $config);
    }

}