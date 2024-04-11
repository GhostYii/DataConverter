﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace DataConverter.Core
{
    public class ExcelConverter : ConverterBase
    {
        private static string[] _supportExtensions = new string[]
        {
            ".xlsx",
            ".xls"
        };

        public override bool CheckConvert(string extension)
        {
            return _supportExtensions.Contains(extension);
        }

        public override T FromData<T>(string filename)
        {
            return new T();
        }

        public override string ToCSharp(string filename, int sheetIndex, string typename)
        {
            // support keyword name           
            if (!Utils.IsValidIdentifier(typename))
                typename = $"@{typename}";

            ExcelData data = ExcelHelper.GetExcelData(filename, sheetIndex);
            if (data == null)
                return string.Empty;

            CodeWriter writer = new CodeWriter();
            writer.WriteLine("using System;");
            writer.WriteLine("using System.Collections.Generic;");

            writer.WriteLine();

            writer.WriteLine($"public struct {typename}");
            writer.BeginBlock();

            foreach (var pair in data.Names)
            {
                string columnName = pair.Key;
                var cellName = pair.Value;

                if (cellName.settings.isIgnore)
                    continue;

                writer.WriteLine($"public {data.Types[columnName].ClassTypeName} {cellName.fieldName} {{ get; set; }}");
            }

            writer.EndBlock();

            return writer.ToString();
        }

        public override string ToJson(string filename, int sheetIndex)
        {
            string name = ExcelHelper.GetSheetNameByIndex(filename, sheetIndex);
            return ToJson(filename, name);
        }

        public string ToJson(string filename, string sheetName)
        {
            ExcelData excelData = ExcelHelper.GetExcelData(filename, sheetName);

            if (excelData == null)
                return string.Empty;

            switch (excelData.Format.format)
            {
                case FormatType.Array:
                    JArray array = new JArray();
                    foreach (var dataPair in excelData.Datas)
                    {
                        array.Add(ToJsonObject(excelData, dataPair.Key));
                    }
                    return JsonConvert.SerializeObject(array);
                case FormatType.KeyValuePair:
                    JObject mapObj = new JObject();
                    foreach (var dataPair in excelData.Datas)
                    {
                        var item = ToJsonObject(excelData, dataPair.Key);
                        var keyToken = item[excelData.Format.key].ToString();
                        if (mapObj.ContainsKey(keyToken))
                        {
                            Console.PrintError($"数据表{Path.GetFileName(filename)}表{sheetName}中" +
                                $"包含重复key（{excelData.Format.key}）值{keyToken}");
                            return string.Empty;
                        }
                        mapObj[keyToken] = item;
                    }
                    return JsonConvert.SerializeObject(mapObj);
            }

            return string.Empty;
        }

        private JObject ToJsonObject(ExcelData data, int rowNumber)
        {
            JObject jsonData = new JObject();
            if (!data.Datas.ContainsKey(rowNumber))
                return null;

            var rowData = data.Datas[rowNumber];

            foreach (var cellPair in rowData)
            {
                var columnName = cellPair.Key;
                var cellData = cellPair.Value;

                var name = data.Names[columnName];
                if (name.settings.isIgnore)
                    continue;

                if (name.settings.cantEmpty && cellData == null)
                {
                    Console.PrintError($"数据表{Path.GetFileName(data.Filename)}表{data.SheetName}中字段{name.name}" +
                        $"（位置{columnName}{rowNumber}）不得为空");
                    return null;
                }

                var type = data.Types[columnName].ToType();
                jsonData[data.Names[columnName].name] = new JValue(To(cellData, type));
            }

            return jsonData;
        }

        private object To(object value, Type type)
        {
            if (type == null)
                return value;

            if (value == null)
                return type.IsValueType ? Activator.CreateInstance(type) : null;

            if (type == typeof(int))
                return Convert.ToInt32(value);

            if (type == typeof(float))
                return float.Parse(value.ToString());

            if (type == typeof(string))
                return value.ToString();

            if (type == typeof(object))
                return JsonConvert.DeserializeObject(value.ToString(), type);

            return value;
        }


    }
}
