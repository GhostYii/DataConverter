﻿using DataConverter;
using DataConverter.Core;
using System.Collections.Generic;
using System.IO;

namespace DataConverter
{
    class CustomCommands
    {
        //[CMD("get_fmt", "获取数据表格式")]
        //private static void ParseFormat(string path)
        //{
        //    var fmt = ExcelHelper.GetDataFormat(path);
        //    Commands.Terminal?.Append($"格式为{fmt.format}, key:{fmt.key}, type:{fmt.type}");
        //}

        //[CMD("types", "获取数据表内的类型")]
        //private static void PrintTypes(int index)
        //{
        //    var tick = System.DateTime.Now;
        //    var types = ExcelHelper.GetTypes("测试表格.xlsx", index);
        //    var span = System.DateTime.Now - tick;
        //    Console.Terminal?.Append($"解析类型消耗时间:{span.Milliseconds}ms");

        //    if (types == null)
        //        return;
        //    foreach (var type in types)
        //    {
        //        Commands.Terminal?.Append($"{type.Key}:{type.Value.type} ({type.Value.ToType()?.Name})");
        //    }
        //}

        //[CMD("names", "获取数据表内的所有字段名称")]
        //private static void PrintNames(int index)
        //{
        //    var names = ExcelHelper.GetNames("测试表格.xlsx", index);
        //    if (names == null)
        //        return;

        //    foreach (var name in names)
        //    {
        //        Commands.Terminal?.Append($"{name.Key}:{name.Value.name}({name.Value.fieldName})");
        //    }
        //}

        [CMD("array")]
        private static void ArrayTest()
        {
            ExcelConverter ec = new ExcelConverter();
            string json = ec.ToJson("测试表格.xlsx", "数组测试");
            Console.Print(json);
        }

        [CMD("map")]
        private static void MapTest()
        {
            ExcelConverter ec = new ExcelConverter();
            string json = ec.ToJson("测试表格.xlsx", "字典测试");
            Console.Print(json);
        }

        [CMD("tocs", "转为CS文件")]
        private static void ToCSFiles()
        {
            Console.Print("");
            ExcelConverter ec = new ExcelConverter();
            var names = ExcelHelper.GetWorksheetNames("测试表格.xlsx");

            int id = 0;
            foreach (var name in names)
            {
                string cs = ec.ToCSharp("测试表格.xlsx", id++, name, string.Empty);
                if (string.IsNullOrEmpty(cs))
                    continue;

                Console.Print(cs);
                Console.Print("===========");
            }
        }

        [CMD("sds")]
        private static void SeAndDeSe(string filename, int sheetIndex)
        {
            ExcelConverter ec = new ExcelConverter();
            string json = ec.ToJson(filename, sheetIndex);

            var obj = ec.FromData<List<GoodArray>>(json);
            
            Console.Print($"{obj}");
        }

    }
}
