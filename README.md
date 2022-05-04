# 概述


- 基于组件 [viest/php-ext-xlswriter](https://github.com/viest/php-ext-xlswriter) 封装
- 支持无限极表头
- 多页码配置
- 多数据类型配置
  
# 安装
```shell
composer require vartruexuan/xlswriter -vvv
```
# demo
```php
$excel = new XlsWriter([
    'path' => './excel', // 导出文件存放目录
]);
$excel->export([
    // 设置页码
    [
       // 页码名
        "sheetName" => "sheet1",
       // 头信息
        "header" => [
            [
                "title" => "name", // 列展示标题
                "type" => "text",  // 数据类型: text,url,formula,date,image
                "key" => "name",   // 列标识(必须唯一)
                "field"=>"name",   // 数据字段
            ],
            [
                "title" => "年龄",
                "type" => "text",
                "key" => "age", 
            ]
        ],
        "data" => [

            [
                "name" => "小黄",
                "age" => 11,
            ],
            [
                "name" => "小红",
                "age" => 11,
            ],

        ],

    ]
]);
```
# 配置
## sheet
```php
[
  [
    "sheetName"=>"sheet1", // 页码名（必须唯一）
    "header"=>[], // 表头： 参考header
    "hide"=>false,// 是否隐藏表
    "zoom"=>100, // 比例缩放: 默认值: 10010 <= 值 <= 400/
    /*
      网格线：
          0: 隐藏 屏幕网格线 和 打印网格线
          1: 显示屏幕网格线
          2: 显示打印网格线
          3: 显示 屏幕网格线 和 打印网格线
    */
    "gridline"=>1, 
    "protection"=>"123",// 设置密码
  ]

]
```
## header
```php
[
    [
        "title" =>"姓名", // 列标题
        // 数据字段: 多级表头时，非底级表头不建议设置
        "field" =>"name", 
        // 子头信息
        "children" =>[],  
        // 数据类型
        "type" =>[
            "text",
            [
                "format" =>null, 
                "formatHandler" => []
            ]
        ]
    ],
    [
        "type" =>[
            "date",
            [
                "dateFormat" =>"yyyy-mm-dd hh:mm:ss",
                "formatHandler" => []
            ]
        ]
    ],
    [
        "type" =>[
            "url",
            [
                "text" =>"我是个链接",
                "tooltip" =>"提示我是个链接",
                "formatHandler" =>null
            ]
        ]
    ],
    [
        "type" =>[
            "image",
            [
                "widthScale" =>1,
                "heightScale" =>1
            ]
        ]
    ],
    [
        "type" =>[
            "formula",
            [
                "formatHandler" => []
            ]
        ]
    ]
]
```
### type-数据类型
#### text

文本类型，自动识别数据类型【默认】
**附加参数**：

- format    数据格式化

![1651636334(1).jpg](https://cdn.nlark.com/yuque/0/2022/jpeg/395716/1651636381204-e4e995b9-b761-4414-967f-f6466b2d3517.jpeg#clientId=uf20f5e67-7549-4&crop=0&crop=0&crop=1&crop=1&from=paste&height=404&id=xZVXk&margin=%5Bobject%20Object%5D&name=1651636334%281%29.jpg&originHeight=790&originWidth=1093&originalType=binary&ratio=1&rotation=0&showTitle=false&size=77050&status=done&style=none&taskId=u4c8bc6bc-ca09-4ac0-94e8-232ebe16abe&title=&width=559)

- formatHandler   参考style
  
#### url

url地址（限制数量65530）
**附加参数**：

      -  text  链接文字 
      -  tooltip  链接提示 
      -  formatHandler  样式参考 style




#### formula

公式
**附加参数**:

      -  formatHandler  样式参考 style

#### date

时间（数据必须是时间戳）

**附加参数**


      -   dateFormat 时间格式 （yyyy-mm-dd hh:mm:ss）

  ![image.png](https://cdn.nlark.com/yuque/0/2022/png/395716/1651638896655-42e7c633-6424-4596-b198-9473c334e270.png#clientId=uf20f5e67-7549-4&crop=0&crop=0&crop=1&crop=1&from=paste&height=411&id=u34ab3d00&margin=%5Bobject%20Object%5D&name=image.png&originHeight=612&originWidth=728&originalType=binary&ratio=1&rotation=0&showTitle=false&size=34670&status=done&style=none&taskId=u277bb537-34cf-49f6-94d3-124b6b4ae13&title=&width=489)

      -   formatHandler    样式参考 style


#### image

图片

**附加参数**

-  widthScale  宽度缩放比例
-  heightScale  高度缩放比例

## style
```php
[
    "italic" =>false,// 是否斜体
   /*
         文本对齐： 
             1 水平左对齐 
             2 水平剧中对齐 
             3 水平右对齐 
             4 水平填充对齐 
             5 水平两端对齐 
             6 横向中心对齐 
             7 分散对齐 
             8 顶部垂直对齐 
             9 底部垂直对齐 
             10 垂直剧中对齐 
             11 垂直两端对齐 
             12 垂直分散对齐
   */
    "align" =>[ 
        1,
        2
    ],
    "strikeout" =>false, // 文本删除（文本中间划线）
   /*
        1 单下划线 
        2 双下划线 
        3 会计用单下划线 
        4 会计用双下划线
   */
    "underline" =>1, 
    "wrap" =>false,// 文本换行：如果单元格内文本包含 \n ，将处理换行样式
    "fontColor" =>16711935,// 字体颜色: 16进制 0xffff
    "fontSize" =>1.2,// 字体大小 
    "bold" =>false,// 是否加粗
   /*
          1 薄边框风格 
          2 中等边框风格 
          3 虚线边框风格 
          4 虚线边框样式 
          5 厚边框风格 
          6 双边风格 
          7 头发边框样式 
          8 中等虚线边框样式 
          9 短划线边框样式 
          10 中等点划线边框样式 
          11 Dash-dot-dot边框样式 
          12 中等点划线边框样式 
          13 倾斜的点划线边框风格
   */
    "border" =>1,// 边框
   /*
      背景颜色（int）： color 【16进制】 
      背景图案样式(int)：pattern 
              0        无 , 
              1        实体（solid）【默认】 , 
              2        中灰色(MEDIUM_GRAY) , 
              3        深灰色（DARK_GRAY） , 
              4        浅灰色（LIGHT_GRAY） , 
              5        黑色水平（DARK_HORIZONTAL） , 
              6        黑色垂直 （DARK_VERTICAL） , 
              7        黑色下（DARK_DOWN） , 
              8        黑色上（DARK_UP） , 
              9        黑色网格（DARK_GRID） , 
              10        黑色格子（DARK_TRELLIS） , 
              11        LIGHT_HORIZONTAL , 
              12        LIGHT_VERTICAL , 
              13        LIGHT_DOWN , 
              14        LIGHT_UP , 
              15        LIGHT_GRID , 
              16        LIGHT_TRELLIS , 
              17        GRAY_125 , 
              18        GRAY_0625 ,
   */
    "background" => [ // 背景
        "color" =>16711935, // 背景颜色
        "pattern" =>1 // 背景图案样式
    ],
    "font" =>"微软雅黑", // 字体
    "number" =>"#,##0" // 数字格式化
]
```
## 




