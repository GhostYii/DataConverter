using Newtonsoft.Json;
using SpreadsheetLight;

namespace DataConverter.Core
{
    using DataDict = Dictionary<int, Dictionary<string, object>>;
    using DataNameDict = Dictionary<string, CellName>;
    using DataTypeDict = Dictionary<string, CellType>;
    using ExcelDocument = SLDocument;
    using Row = Dictionary<int, SLCell>;
    using Rows = Dictionary<int, Dictionary<int, SLCell>>;

    public static class ExcelHelper
    {
        public static int ToColumnIndex(string columnName)
        {
            return SLConvert.ToColumnIndex(columnName);
        }
        public static string ToColumnName(int columnIndex)
        {
            return SLConvert.ToColumnName(columnIndex);
        }

        public static bool CheckValid(string filename, int sheetIndex)
        {
            if (!File.Exists(filename))
            {
                Console.PrintError($"不存在数据表文件{filename}");
                return false;
            }

            try
            {
                ExcelDocument file = new ExcelDocument(filename);
                if (sheetIndex >= file.GetWorksheetNames().Count)
                {
                    Console.PrintError($"数据表'{Path.GetFileName(filename)}'不存在第{sheetIndex}张表");
                    return false;
                }

                return true;
            }
            catch (Exception e)
            {
                Console.PrintError(e.Message);
                return false;
            }

        }
        public static bool CheckValid(string filename, string sheetName)
        {
            if (!File.Exists(filename))
            {
                Console.PrintError($"不存在数据表文件:{filename}");
                return false;
            }

            try
            {
                ExcelDocument file = new ExcelDocument(filename, sheetName);
                if (file == null)
                {
                    Console.PrintError($"数据表'{Path.GetFileName(filename)}'不存在名为'{sheetName}'的数据表");
                    return false;
                }

                return true;
            }
            catch (Exception e)
            {
                Console.PrintError(e.Message);
                return false;
            }
        }

        // 获取数据表表格数量
        public static int GetWorksheetCount(string filename)
        {
            if (!File.Exists(filename))
            {
                Console.PrintError($"文件{filename}不存在");
                return 0;
            }

            using (ExcelDocument file = new ExcelDocument(filename))
            {
                return file.GetSheetNames().Count;
            }
        }

        // 获取数据表所有表格名称
        public static string[] GetWorksheetNames(string filename)
        {
            try
            {
                if (!File.Exists(filename))
                {
                    Console.PrintError($"文件{filename}不存在");
                    return null;
                }

                using (ExcelDocument file = new ExcelDocument(filename))
                {
                    return file.GetSheetNames().ToArray();
                }
            }
            catch (Exception e)
            {
                Console.PrintError(e.Message);
                return null;
            }
        }

        public static string GetSheetNameByIndex(string filename, int sheetIndex)
        {
            if (!File.Exists(filename))
            {
                Console.PrintError($"文件{filename}不存在");
                return null;
            }

            using (ExcelDocument file = new ExcelDocument(filename))
            {
                var names = file.GetSheetNames();
                if (sheetIndex >= names.Count)
                {
                    Console.PrintError($"文件{filename}不存在索引为{sheetIndex}的表");
                    return null;
                }

                return names[sheetIndex];
            }

        }
        public static int GetSheetIndexByName(string filename, string sheetName)
        {
            if (!File.Exists(filename))
            {
                Console.PrintError($"文件{filename}不存在");
                return 0;
            }

            using (ExcelDocument file = new ExcelDocument(filename))
            {
                int index = file.GetSheetNames().IndexOf(sheetName);
                if (index == -1)
                {
                    Console.PrintError($"文件{filename}不存在名称为{sheetName}的表");
                    return 0;
                }

                return index;
            }
        }

        #region Internal

        // 获取表格数据配置
        internal static DataConfig GetDataConfig(string filename, int sheetIndex = 0)
        {
            DataConfig result = new DataConfig() { format = FormatType.None };
            if (!CheckValid(filename, sheetIndex))
                return result;

            Rows rows = GetValidRows(filename, sheetIndex);
            if (rows == null)
            {
                Console.PrintError($"不存在的表格{Path.GetFileName(filename)}");
                return result;
            }

            var fmt = GetDataConfig(rows);
            if (fmt.HasValue && fmt.Value.format == FormatType.None)
                Console.PrintError($"数据表'{Path.GetFileName(filename)}'第{sheetIndex}张表配置了不支持的格式");
            else
                Console.PrintError($"数据表'{Path.GetFileName(filename)}'第{Const.ROW_LINE_NUM_CONFIG}个有效行不是配置控制字段");

            return fmt ?? result;
        }
        internal static DataConfig GetDataFormat(string filename, string sheetName)
        {
            DataConfig defaultFmt = new DataConfig() { format = FormatType.None };
            if (!CheckValid(filename, sheetName))
                return defaultFmt;

            if (!File.Exists(filename))
                return defaultFmt;

            return GetDataConfig(filename, GetSheetIndexByName(filename, sheetName));
        }

        // 获取表格数据名称
        internal static DataNameDict GetNames(string filename, int sheetIndex = 0)
        {
            if (!CheckValid(filename, sheetIndex))
                return null;

            Rows rows = GetValidRows(filename, sheetIndex);
            if (rows.Count < Const.ROW_LINE_NUM_NAME)
            {
                Console.PrintError($"数据表'{Path.GetFileName(filename)}'表{sheetIndex}第{Const.ROW_LINE_NUM_NAME}个有效行不是字段名称行");
                return null;
            }

            return GetNames(rows, filename, sheetIndex);
        }
        internal static DataNameDict GetNames(string filename, string sheetName)
        {
            if (!CheckValid(filename, sheetName))
                return null;

            return GetNames(filename, GetSheetIndexByName(filename, sheetName));
        }

        // 获取表格数据类型
        internal static DataTypeDict GetTypes(string filename, int sheetIndex = 0)
        {
            if (!CheckValid(filename, sheetIndex))
                return null;

            Rows rows = GetValidRows(filename, sheetIndex);
            if (rows.Count < Const.ROW_LINE_NUM_TYPE)
            {
                Console.PrintError($"数据表'{Path.GetFileName(filename)}'表{sheetIndex}第{Const.ROW_LINE_NUM_TYPE}个有效行不是字段类型行");
                return null;
            }

            return GetTypes(rows, filename, sheetIndex);
        }
        internal static DataTypeDict GetTypes(string filename, string sheetName)
        {
            if (!CheckValid(filename, sheetName))
                return null;

            return GetTypes(filename, GetSheetIndexByName(filename, sheetName));
        }

        // 获取表格数据（除格式、类型、名称外）
        internal static DataDict GetTableData(string filename, int sheetIndex = 0)
        {
            if (!CheckValid(filename, sheetIndex))
                return null;

            return GetTableData(GetValidRows(filename, sheetIndex));
        }
        internal static DataDict GetTableData(string filename, string sheetName)
        {
            if (!CheckValid(filename, sheetName))
                return null;

            return GetTableData(filename, GetSheetIndexByName(filename, sheetName));
        }

        internal static ExcelData GetExcelData(string filename, int sheetIndex = 0)
        {
            if (!CheckValid(filename, sheetIndex))
                return null;

            Rows rows = GetValidRows(filename, sheetIndex);

            var cfg = GetDataConfig(rows);
            if (cfg.HasValue)
            {
                ExcelData data = new ExcelData();
                data.Filename = filename;
                data.SheetIndex = sheetIndex;
                data.SheetName = GetSheetNameByIndex(filename, sheetIndex);
                data.Config = cfg.Value;
                data.Types = GetTypes(rows, filename, sheetIndex);
                data.Names = GetNames(rows, filename, sheetIndex);
                data.Datas = GetTableData(rows, data.Names.Keys);
                data.DataBeginRowNumber = rows.Count <= Const.ROW_LINE_NUM_DATA ? 0 :
                    rows.ElementAt(Const.ROW_LINE_NUM_DATA).Key;

                if (data.Config.format == FormatType.KeyValuePair)
                {
                    string pos = data.Names.First(pair => pair.Value.name == data.Config.key).Key;
                    if (!data.Types[pos].type.IsValueType())
                        Console.PrintWarning($"数据表'{Path.GetFileName(filename)}'表{sheetIndex}字典键非值类型");
                }

                return data;
            }
            else
            {
                return null;
            }
        }
        internal static ExcelData GetExcelData(string filename, string sheetName)
        {
            if (!CheckValid(filename, sheetName))
                return null;

            return GetExcelData(filename, GetSheetIndexByName(filename, sheetName));

        }

        private static ExcelDocument GetWorksheet(string filename, int sheetIndex = 0)
        {
            try
            {
                if (!CheckValid(filename, sheetIndex))
                    return null;

                return new ExcelDocument(filename, GetSheetNameByIndex(filename, sheetIndex));
            }
            catch (Exception e)
            {
                Console.PrintError($"加载数据表'{Path.GetFileName(filename)}'失败，{e.Message}");
                return null;
            }
        }

        private static ExcelDocument GetWorksheet(string filename, string sheetName = "Sheet1")
        {
            try
            {
                if (!CheckValid(filename, sheetName))
                    return null;

                return new ExcelDocument(filename, sheetName);
            }
            catch (Exception e)
            {
                Console.PrintError($"加载数据表'{Path.GetFileName(filename)}'失败，{e.Message}");
                return null;
            }
        }

        // 获取所有有效行
        private static Rows GetValidRows(string filename, int sheetIndex = 0/*, bool includeNote = false*/)
        {
            if (!File.Exists(filename))
            {
                Console.PrintError($"不存在的表格\"{Path.GetFileName(filename)}\"");
                return null;
            }

            var sheet = GetWorksheet(filename, sheetIndex);
            if (sheet == null)
                return null;

            var sstr = sheet.GetSharedStrings();
            sstr[0].GetText();

            //if (includeNote)
            //    return sheet.GetCells();

            Rows rows = new Rows();
            foreach (var (rowNumber, data) in sheet.GetCells())
            {
                string firstCellValue = sheet.GetCellValueAsString($"A{rowNumber}");
                bool isNote = firstCellValue.Trim().StartsWith(Const.NOTE_PREFIX);

                if (isNote)
                    continue;

                Row row = new Row();
                foreach (var (columnIndex, cell) in data)
                {
                    cell.CellText = sheet.GetCellValueAsString(rowNumber, columnIndex);
                    row[columnIndex] = cell;
                }

                rows[rowNumber] = row;
            }

            return rows;
        }

        private static DataConfig? GetDataConfig(Rows rows)
        {
            var row = rows.ElementAt(Const.ROW_LINE_NUM_CONFIG).Value;
            string fmtStr = row.Values.First().CellText.Trim();
            try
            {
                var result = JsonConvert.DeserializeObject<DataConfig>(fmtStr);
                return result;
            }
            catch
            {
                return null;
            }
        }

        private static DataNameDict GetNames(Rows rows, string filename, int sheetIndex)
        {
            DataNameDict result = new DataNameDict();
            HashSet<string> names = new HashSet<string>();

            var target = rows.ElementAt(Const.ROW_LINE_NUM_NAME);
            int rowIndex = target.Key;
            Row row = target.Value;

            foreach (var (columnIndex, cell) in row)
            {
                string columnName = SLConvert.ToColumnName(columnIndex);
                string cellStr = cell.CellText.Trim();

                if (string.IsNullOrEmpty(cellStr))
                    continue;

                string fieldName = Utils.ToFieldName(cellStr);

                if (string.IsNullOrEmpty(fieldName))
                {
                    Console.PrintError($"数据表'{Path.GetFileName(filename)}'表{sheetIndex}位置为" +
                        $"{columnName}{rowIndex}的数据名称非法，非法名称将被忽略");
                    continue;
                }

                if (names.Contains(fieldName))
                {
                    Console.PrintWarning($"数据表'{Path.GetFileName(filename)}'表{sheetIndex}位置为" +
                        $"{columnName}{rowIndex}的数据名称重复，重复数据将被忽略");
                    continue;
                }

                names.Add(fieldName);

                ConverterSettings cs = new ConverterSettings()
                {
                    isIgnore = cellStr.StartsWith(Const.NOTE_PREFIX),
                    cantEmpty = cellStr.EndsWith(Const.NON_EMPTY_SUFFIX)
                };

                result[columnName] = new CellName()
                {
                    name = cellStr.Trim(Const.NOTE_PREFIX, Const.NON_EMPTY_SUFFIX),
                    fieldName = fieldName,
                    settings = cs
                };
            }

            return result;
        }
        private static DataTypeDict GetTypes(Rows rows, string filename, int sheetIndex)
        {
            var target = rows.ElementAt(Const.ROW_LINE_NUM_TYPE);
            int rowIndex = target.Key;
            Row row = target.Value;

            DataTypeDict result = new DataTypeDict();
            foreach (var (columnIndex, cell) in row)
            {
                string columnName = SLConvert.ToColumnName(columnIndex);

                string cellStr = cell.CellText.Trim();
                if (string.IsNullOrEmpty(cellStr))
                    continue;
                var type = TypeParser.Parse(cellStr);
                if (type == null)
                {
                    Console.PrintError($"不支持的数据类型\'{TypeParser.SplitType(cellStr)[0]}\'，位于数据表'{Path.GetFileName(filename)}'表{sheetIndex}的" +
                                        $"{SLConvert.ToCellReference(rowIndex, columnIndex)}项，该项数据将被忽略");
                    continue;
                }
                else
                {
                    result[columnName] = type;
                }
            }
            return result;
        }
        private static DataDict GetTableData(Rows rows, IEnumerable<string> columnNames = null)
        {
            DataDict data = new DataDict();
            for (int i = Const.ROW_LINE_NUM_DATA; i < rows.Count; i++)
            {
                Dictionary<string, object> rowData = new Dictionary<string, object>();

                var rowInfo = rows.ElementAt(i);
                int rowIndex = rowInfo.Key;
                var row = rowInfo.Value;

                if (columnNames == null)
                {
                    foreach (var (columnIndex, cell) in row)
                    {
                        string columnName = SLConvert.ToColumnName(columnIndex);
                        rowData[columnName] = cell.CellText;
                    }
                }
                else
                {
                    foreach (var name in columnNames)
                    {
                        int columnIndex = SLConvert.ToColumnIndex(name);
                        rowData[name] = row.ContainsKey(columnIndex) ? row[columnIndex].CellText : null;
                    }
                }

                data[rowIndex] = rowData;
            }

            return data;
        }

        #endregion

    }
}
