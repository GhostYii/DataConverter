游戏中的导表工具是指的将Excel表导入到游戏中的工具，全程为DataConverter，以下简称DCT。

# DCT是什么?

DCT是使用CSharp语言编写的一套可用于游戏中的数据表导入工具，其中主要的工作流为Excel->Json->C# Class/Struct。
其中Excel的解析是基于[SpreadsheetLight库](https://spreadsheetlight.com/)实现的，而Json的部分则是基于[Newtonsoft.Json](https://www.newtonsoft.com/json)来搞的，Bson的部分是基于[MongoDB.bson](https://github.com/mongodb/mongo-csharp-driver/tree/master/src/MongoDB.Bson)搞的，我只做了一点微小的工作。

# 如何使用？

DCT的核心内容也编译为了DLL库（DataConverter.Core），可以通过将该文件导入项目中来调用。
同时，该项目也支持了一个CLI工具，任何人都可以下载编译好的CLI工具来使用。

# 有什么用？

DCT的主要工作流程是将Excel表导出为json/bson，并且可以自动生成对应的C#结构体，从而实现数据加载。

# 有何规范？

DCT规定了一些excel的规范以用于解析，其主要内容如下（以下内容如非特殊提醒，“表”均指代“Excel数据表”，“数据”均指代“json数据”，“结构体”指代C#中的结构体）:

- 数据的最小单位为一张表，一张表指excel中一个sheet
- 表中的有效行指的是存在数据的非注释行，有效列指的是存在数据的非注释列
- 以‘#’开头的行/列为注释行/列，注释行/列不会参与解析
- 表中可以存在任意多个注释行/列
- 表中的第一个有效行中的数据必须为类型定义字符串，其格式如下：
  - `{ "format": "array" }`：该表格按行导出为数组
  - `{ "format": "map", "key": row_name }`：该表格按照字典导出，键为某一行的名字，当指定导出格式为字典时，必须同时指定合法的键
- 表中的第二个有效行必须为类型定义行，支持的类型如下：
  - `bool`：布尔值，也可写作`boolean`
  - `int`：整数
  - `float`：浮点数/小数
  - `string`：字符串，也可写作`str`
  - `array:type`：类型为type的数组，也可写作`arr`或者`list`
  - `map:value_type`：键类型为string（键类型不可更改），值类型为value_type的字典，也可以写作`dict`、`pairs`、`hash`
  - `object:object_type`：类型为object_type的对象，此对象类型必须被DCT所识别，也可写作`obj`
  - 类型支持嵌套，即可存在多维数组如`arr:arr:float`、`map:map:int`这样的类型定义
  - 由于C#是强类型语言，当类型为数组、字典或对象这些引用类型时，必须指定其类型
- 表中的第三个有效行必须为变量名
  - 变量名要求和C#中起名要求一致，当名字与关键字冲突时，DCT会在生成的代码中自动添加'@'符号
- 变量名存在特殊的前后缀
  - `#`前缀：表示此数据列注释，该列所有数据均不会被DCT识别并导出
  - `*`后缀：表示此列不得为空值，当出现空值时DCT将会抛出异常
- 数据将从第四个有效行开始被导出，其中数据填写规范如下：
  - `bool`：填写`true(不区分大小写)`或者`1`会被识别为`true`，其他所有识别为`false`
  - `int`：填写整数，如`1`
  - `float`：填写单精度小数，如`1.0`
  - `string`：填写任意长度字符串，如`hello, world!`
  - `array`：以`[]`包裹的对应类型的数据，如`[1, 2, 3]`、`[ [1,2,3], [4,5,6] ]`
  - `map`：以`{"key":value, }`包裹的对应类型的数据，如`{ "1": 1, "2":2 }`
  - `object`：填写可被反序列化为对应对象的json字符串，如`{"num":145647163,"percent":0.588,"text":"0.54244547387638576486438546"}`

# 示例代码

以下代码演示了如何将数据表导出为json数据

```csharp
using DataConvert.Core;

void Example()
{
    ExcelConverter convert = new ExcelConverter();
    string json = convert.ToJson("example.xlsx", "sheet1");
    // string json = convert.ToJson("example.xlsx", 2);
    Console.Print(json);
}
```

以下代码演示了如何将数据表格式导出为CSharp结构体

```csharp
using DataConvert.Core;

void Example()
{
    ExcelConvert convert = new ExcelConvert();
    var sheetNames = ExcelHelper.GetWorksheetNames("example.xlsx");

    int id = 0;
    foreach (var name in sheetNames)
    {
        string csContext = convert.ToCSharp("example.xlsx", id++, name);
        if (string.IsNullOrEmpty(csContext))
            continue;

        Console.Print(csContext);        
    }
}
```

更多的示例代码可以查看CLI源码中的函数。
