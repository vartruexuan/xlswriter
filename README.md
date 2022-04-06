<h1 align="center"> xlswriter </h1>

<p align="center"> .</p>


## 安装

```shell
$ composer require vartruexuan/xlswriter -vvv
```

## 使用
```php
 $excel = new XlsWriter([
    'path'=>'./excel', // 导出文件存放目录
]);
$excel->export([
   	// 设置页码
    [
        "sheetName" =>"sheet1",
        "sheetType" =>"default",
        "header" =>[
            [
                "title" =>"name",
                "type" =>"text",
                "key" =>"name",
                "children" =>[],//可设置多级表头
            ],
            [
                "title" =>"年龄",
                "type" =>"text", // 数据类型: text,url,formula,date,image
                "key" =>"age",  // 数据key
            ]
        ],
        "data" =>[

        		[
        			"name"=>"小黄",
        			"age"=>11,
        		],
				[
        			"name"=>"小红",
        			"age"=>11,
        		],

        ],
            
    ]  
]);
```
## 配置
### 导出配置
```php
[
    // 页码配置(可多个)
    [
        "sheetName" =>"sheet1", // 页码名称（必须唯一）
        "sheetType" =>"default", // 页码类型
        // 设置表头
        "header" =>[
            [
                "title" =>"name", // 展示列名
                "type" =>"text", // 数据类型: text,url,formula,date,image
                "key" =>"name",// 数据key值
                "style"=>[ // 列样式
                            "italic" =>false,
                            "align" =>[
                                1,
                                2
                            ],
                            "strikeout" =>false,
                            "underline" =>1,
                            "wrap" =>false,
                            "fontColor" =>"0xFF69B4",
                            "fontSize" =>1.2,
                            "bold" =>false,
                            "border" =>1,
                            "background" => [
                                "color" =>1,
                                "pattern" =>""
                            ],
                            "font" =>"微软雅黑",
                            "number" =>"'#,##0'"
                ],
                "children" =>[],//可设置多级表头
            ],
            [
                "title" =>"年龄",
                "type" =>"text", // 数据类型: text,url,formula,date,image
                "key" =>"age",  
                "children" =>[],
            ]
        ],
        
        // 数据 array | callback(XlsWriter $xlsWriter, $i, $dataHeaders, &$isWhile,$maxRow
        "data" =>[
        		[
        			"name"=>"小黄",
        			"age"=>11,
        		],
				[
        			"name"=>"小红",
        			"age"=>11,
        		],

        ],
        "zoom" =>"20", // 缩放比例:10~400, 默认100
        "gridline" =>"1", // 网格线 
        "isHide" =>false, // 是否隐藏
        "protection" =>"123",// 设置操作密码
    ]
]


```

TODO
## License

MIT