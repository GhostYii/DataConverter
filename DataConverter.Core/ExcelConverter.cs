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

        public override T FromData<T>(string json)
        {
            try
            {
                return JsonConvert.DeserializeObject<T>(json);
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

            if (string.IsNullOrEmpty(typename) || !Utils.IsValidTypeName(typename))
            {
                Console.PrintError($"数据表'{Path.GetFileName(filename)}'表{sheetIndex}转CS文件typename({typename})无法作为变量名");
                return string.Empty;
            }

            // number typename add '_' prefix
            if (int.TryParse(typename.Substring(0, 1), out int _))
                typename = $"_{typename}";

            // support keyword name           
            if (!Utils.IsValidIdentifier(typename))
                typename = $"@{typename}";

            ExcelData data = ExcelHelper.GetExcelData(filename, sheetIndex);
            if (data == null)
                return string.Empty;

            CodeWriter writer = new CodeWriter();
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

                // define sheet type first
                switch (data.Format.format)
                {
                    case FormatType.Array:
                        writer.WriteLine($"using {typename}Data = List<{typename}>;");
                        break;
                    case FormatType.KeyValuePair:
                        writer.WriteLine($"using {typename}Data = Dictionary<{Utils.GetKeyType(data)}, {typename}>;");
                        break;
                    default:
                        break;
                }
            }
            else
            {
                // define sheet type first
                switch (data.Format.format)
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

            writer.WriteLine();

            writer.WriteLine($"public struct {typename}");
            writer.BeginBlock();

            foreach (var pair in data.Names)
            {
                string columnName = pair.Key;
                var cellName = pair.Value;

                if (cellName.settings.isIgnore)
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
                            Console.PrintError($"数据表'{Path.GetFileName(filename)}'表'{sheetName}'中" +
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
                    Console.PrintError($"数据表'{Path.GetFileName(data.Filename)}'表'{data.SheetName}'中字段{name.name}" +
                        $"（位置{columnName}{rowNumber}）不得为空");
                    return null;
                }

                var type = data.Types[columnName].Type;
                string cellName = data.Names[columnName].name;

                if (cellData == null)
                {
                    jsonData[cellName] = data.Types[columnName].DefaultJsonValue;
                    continue;
                }
                else if (type == null)
                {
                    jsonData[cellName] = JsonConvert.DeserializeObject(cellData?.ToString(), data.Types[columnName].JsonType) as JToken;
                    //jsonData[cellName] = JsonConvert.DeserializeObject<JObject>(cellData?.ToString());
                    //Console.Print($"GTMD----> {cellData?.ToString()}");
                    continue;
                }

                // special for boolean
                if (type.Equals(typeof(bool)))
                    jsonData[data.Names[columnName].name] = new JValue(Utils.ParseToBool(cellData));                
                else if (type.IsValueType || type.Equals(typeof(string)))
                    jsonData[data.Names[columnName].name] = new JValue(cellData);
                else if (type.IsList())
                    jsonData[cellName] = JsonConvert.DeserializeObject<JArray>(cellData.ToString());
                else
                    jsonData[cellName] = JsonConvert.DeserializeObject<JObject>(cellData.ToString());
            }

            return jsonData;
        }

        
    }
}
