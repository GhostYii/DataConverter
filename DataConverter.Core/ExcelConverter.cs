﻿using DocumentFormat.OpenXml.Spreadsheet;
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

        public override bool CheckConvert(string extension) => _supportExtensions.Contains(extension);

        public bool CheckToJson(string filename, int sheetIndex)
        {
            if (!CheckConvert(Path.GetExtension(filename)))
                return false;

            ExcelData excelData = ExcelHelper.GetExcelData(filename, sheetIndex);

            // template sheet couldnt convert to json            
            if (excelData == null || excelData.Config.isTemplate)
                return false;

            return true;
        }

        public bool CheckToJson(string filename, string sheetName) => CheckToJson(filename, ExcelHelper.GetSheetIndexByName(filename, sheetName));

        public override T FromData<T>(string json)
        {
            try
            {
                return JsonConvert.DeserializeObject<T>(json)!;
            }
            catch (Exception e)
            {
                Console.PrintError(e.Message);
                return default;
            }
        }

        public override string ToCSharp(string filename, int sheetIndex, string typename, string nameSpace)
        {
            if (!CheckConvert(Path.GetExtension(filename)))
            {
                Console.PrintError($"数据表'{Path.GetFileName(filename)}'不支持的格式");
                return string.Empty;
            }

            ExcelData data = ExcelHelper.GetExcelData(filename, sheetIndex);
            if (data == null || !string.IsNullOrEmpty(data.Config.templateName) || data.Config.objectType == ObjectType.None)
                return string.Empty;

            string sheetName = ExcelHelper.GetSheetNameByIndex(filename, sheetIndex);

            if (!data.Config.isTemplate && data.Config.format == FormatType.None)
            {
                Console.PrintError($"数据表'{Path.GetFileName(filename)}'中子表'{sheetName}'不支持的转换格式");
                return string.Empty;
            }

            typename = string.IsNullOrEmpty(data.Config.objectName) ? typename : data.Config.objectName;

            if (string.IsNullOrEmpty(typename) || !Utils.IsValidTypeName(typename))
            {
                Console.PrintError($"数据表'{Path.GetFileName(filename)}'中子表'{sheetName}'转CS文件typename({typename})无法作为变量名");
                return string.Empty;
            }

            // number typename add '_' prefix
            if (int.TryParse(typename.Substring(0, 1), out int _))
                typename = $"_{typename}";

            // support keyword name           
            if (!Utils.IsValidIdentifier(typename))
                typename = $"@{typename}";

            CodeWriter writer = new CodeWriter();

            writer.WriteLine("// This code was auto-generated by ExcelConverter tool");
            //The source file of this code is table b in a.xlsx
            writer.WriteLine($"// The source file of this code is sheet {sheetName} in {Path.GetFileName(filename)}");
            if (data.Config.isTemplate)
                writer.WriteLine("// Template sheet to define types");

            // tip sheet format first
            switch (data.Config.format)
            {
                case FormatType.Array:
                    writer.WriteLine($"// File format should be List<{typename}>");
                    break;
                case FormatType.KeyValuePair:
                    writer.WriteLine($"// File format should be Dictionary<{Utils.GetKeyType(data)}, {typename}>");
                    break;
                default:
                    writer.WriteLine("// Unknown file format");
                    break;
            }
            writer.WriteLine();

            writer.WriteLine("using System;");
            writer.WriteLine("using System.Collections.Generic;");

            if (!string.IsNullOrEmpty(nameSpace))
            {
                writer.WriteLine();

                // number typename add '_' prefix
                if (int.TryParse(nameSpace.Substring(0, 1), out int _))
                    nameSpace = $"_{nameSpace}";

                // support keyword name           
                if (!Utils.IsValidIdentifier(nameSpace))
                    nameSpace = $"@{nameSpace}";

                writer.WriteLine($"namespace {nameSpace}");
                writer.BeginBlock();
            }
            else
            {
                // define sheet type first
                switch (data.Config.format)
                {
                    case FormatType.Array:
                        writer.WriteLine($"using {typename}Data = System.Collections.Generic.List<{typename}>;");
                        break;
                    case FormatType.KeyValuePair:
                        writer.WriteLine($"using {typename}Data = System.Collections.Generic.Dictionary<{Utils.GetKeyType(data)}, {typename}>;");
                        break;
                    default:
                        break;
                }
            }

            // write enum types
            if (data.Config.genEnumType)
            {
                foreach (var (enumName, enumValues) in data.Enums)
                {
                    writer.WriteLine($"public enum {enumName}");
                    writer.BeginBlock();

                    for (int i = 0; i < enumValues.Count; ++i)
                    {
                        writer.WriteIndent();
                        writer.WriteFormat
                        (
                            "{0}{1}{2}",
                            enumValues[i],
                            i == 0 ? " = 0" : string.Empty,
                            i == enumValues.Count - 1 ? string.Empty : ","
                        );
                        writer.WriteLine();
                    }

                    writer.EndBlock();
                    writer.WriteLine();
                }
            }

            string codeType = data.Config.objectType.ToString().ToLower();
            // default use struct to avoid gc
            if (string.IsNullOrEmpty(codeType) || codeType == "none")
                codeType = "struct";

            writer.WriteLine($"public {codeType} {typename}");
            writer.BeginBlock();

            foreach (var (columnName, cellName) in data.Names)
            {
                if (!data.Types.ContainsKey(columnName) || cellName.settings.isIgnore)
                    continue;

                writer.WriteLine($"public {data.Types[columnName].TypeName} {cellName.name};");
                //writer.WriteLine($"public {data.Types[columnName].TypeName} {cellName.fieldName} {{ get; set; }}");
            }

            writer.EndBlock();

            if (!string.IsNullOrEmpty(nameSpace))
                writer.EndBlock();

            return writer.ToString();
        }

        public override string ToJson(string filename, int sheetIndex)
        {
            if (!CheckConvert(Path.GetExtension(filename)))
            {
                Console.PrintError($"数据表'{Path.GetFileName(filename)}'不支持的格式");
                return string.Empty;
            }

            string name = ExcelHelper.GetSheetNameByIndex(filename, sheetIndex);
            return ToJson(filename, name);
        }

        public string ToJson(string filename, string sheetName)
        {
            if (!CheckConvert(Path.GetExtension(filename)))
            {
                Console.PrintError($"数据表'{Path.GetFileName(filename)}'不支持的格式");
                return string.Empty;
            }

            ExcelData excelData = ExcelHelper.GetExcelData(filename, sheetName);

            // template sheet dont convert to json
            if (excelData == null || excelData.Config.isTemplate)
                return string.Empty;

            switch (excelData.Config.format)
            {
                case FormatType.Array:
                    JArray array = new JArray();
                    foreach (var (row, _) in excelData.Datas)
                    {
                        array.Add(ToJsonObject(excelData, row));
                    }
                    return JsonConvert.SerializeObject(array);
                case FormatType.KeyValuePair:
                    JObject mapObj = new JObject();
                    foreach (var (row, _) in excelData.Datas)
                    {
                        var item = ToJsonObject(excelData, row);
                        var keyToken = item[excelData.Config.key]!.ToString();
                        if (mapObj.ContainsKey(keyToken))
                        {
                            Console.PrintError($"数据表'{Path.GetFileName(filename)}'表'{sheetName}'中" +
                                $"包含重复key（{excelData.Config.key}）值{keyToken}");
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

            foreach (var (columnName, cellData) in rowData)
            {
                var name = data.SelfNames[columnName];
                if (name.settings.isIgnore)
                    continue;

                if (name.settings.cantEmpty && cellData == null)
                {
                    Console.PrintError($"数据表'{Path.GetFileName(data.Filename)}'表'{data.SheetName}'中字段{name.name}" +
                        $"（位置{columnName}{rowNumber}）不得为空");
                    return null;
                }

                var type = data.SelfTypes[columnName].Type;
                string cellName = data.SelfNames[columnName].name;

                if (cellData == null)
                {
                    jsonData[cellName] = data.SelfTypes[columnName].DefaultJsonValue;
                    continue;
                }
                else if (type == null)
                {
                    jsonData[cellName] = JsonConvert.DeserializeObject(cellData?.ToString(), data.SelfTypes[columnName].JsonType) as JToken;
                    //jsonData[cellName] = JsonConvert.DeserializeObject<JObject>(cellData?.ToString());                    
                    continue;
                }

                // special for boolean
                if (type.Equals(typeof(bool)))
                    jsonData[data.SelfNames[columnName].name] = new JValue(Utils.ParseToBool(cellData));
                else if (type.IsValueType || type.Equals(typeof(string)))
                    jsonData[data.SelfNames[columnName].name] = new JValue(cellData);
                else if (type.IsList())
                    jsonData[cellName] = JsonConvert.DeserializeObject<JArray>(cellData.ToString());
                else
                    jsonData[cellName] = JsonConvert.DeserializeObject<JObject>(cellData.ToString());
            }

            return jsonData;
        }
    }
}
