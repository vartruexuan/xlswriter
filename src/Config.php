<?php

namespace Vartruexuan\Xlswriter;

class Config
{



    public static function getCommonConfig($config=[])
    {
        return array_merge([
            "path"=>"./", // 导出地址
        ],$config);

    }


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
            "title" => "名称",
            "type" => "string",
            "key" => "name",
            "style" => "",
            "dataFormat" => function () {
            },
            "rowspan" => 1,
            "colspan" => 1,
            "children" => [],
            "width" => 0
        ], $config);
    }

}