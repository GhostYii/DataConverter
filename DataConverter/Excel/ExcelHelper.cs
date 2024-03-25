using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using FastExcel;
using Newtonsoft.Json;
using ExcelFile = FastExcel.FastExcel;

namespace DataConverter
{
    using DataNameDict = Dictionary<string, CellName>;
    using DataTypeDict = Dictionary<string, CellType>;
    using DataDict = Dictionary<int, Dictionary<string, object>>;

    public static class ExcelHelper
    {         
        public static bool CheckValid(string filename, int sheetIndex)
        {            
            if (!File.Exists(filename))
            {
                Console.PrintError($"不存在数据表文件{filename}");
                return false;
            }

            try
            {
                ExcelFile file = new ExcelFile(new FileInfo(filename), true);
                if (!file.Read(sheetIndex).Exists)
                {
                    Console.PrintError($"数据{Path.GetFileName(filename)}不存在第{sheetIndex}张表");
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
                Console.PrintError($"不存在数据表文件{filename}");
                return false;
            }

            try
            {
                ExcelFile file = new ExcelFile(new FileInfo(filename), true);
                if (!file.Read(sheetName).Exists)
                {
                    Console.PrintError($"数据{Path.GetFileName(filename)}不存在名为{sheetName}的数据表");
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
            using (ExcelFile file = new ExcelFile(new FileInfo(filename), true))
            {
                if (!file.ExcelFile.Exists)
                {
                    Console.PrintError($"文件{filename}不存在");
                    return 0;
                }
                return file.Worksheets.Length;
            }
        }

        // 获取数据表所有表格名称
        public static string[] GetWorkshetNames(string filename)
        {
            try
            {
                using (ExcelFile file = new ExcelFile(new FileInfo(filename), true))
                {
                    if (!file.ExcelFile.Exists)
                    {
                        Console.PrintError($"文件{filename}不存在");
                        return null;
                    }

                    string[] names = new string[file.Worksheets.Length];
                    for (int i = 0; i < file.Worksheets.Length; i++)
                    {
                        names[i] = file.Worksheets[i].Name;
                    }
                    return names;
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
            using (ExcelFile file = new ExcelFile(new FileInfo(filename), true))
            {
                if (!file.ExcelFile.Exists)
                {
                    Console.PrintError($"文件{filename}不存在");
                    return null;
                }
                var sheet = file.Read(sheetIndex);
                if (!sheet.Exists)
                {
                    Console.PrintError($"文件{filename}不存在索引为{sheetIndex}的表");
                    return null;
                }

                return sheet.Name;
            }

        }
        public static int GetSheetIndexByName(string filename, string sheetName)
        {
            using (ExcelFile file = new ExcelFile(new FileInfo(filename), true))
            {
                if (!file.ExcelFile.Exists)
                {
                    Console.PrintError($"文件{filename}不存在");
                    return 0;
                }
                var sheet = file.Read(sheetName);
                if (!sheet.Exists)
                {
                    Console.PrintError($"文件{filename}不存在名称为{sheetName}的表");
                    return 0;
                }

                return sheet.Index;
            }
        }

        // 获取表格数据格式
        public static DataFormat GetDataFormat(string filename, int sheetIndex = 1)
        {
            DataFormat result = new DataFormat() { format = FormatType.None };
            if (!CheckValid(filename, sheetIndex))
                return result;

            Row[] rows = GetValidRows(filename, sheetIndex);
            if (rows == null)
            {
                Console.PrintError($"不存在的表格{Path.GetFileName(filename)}");
                return result;
            }

            var fmt = GetDataFormat(rows);
            if (fmt.HasValue && fmt.Value.format == FormatType.None)
                Console.PrintError($"数据表{Path.GetFileName(filename)}第{sheetIndex}张表配置了不支持的格式");
            else
                Console.PrintError($"数据表{Path.GetFileName(filename)}第{Const.ROW_LINE_NUM_FORMAT}个有效行不是格式控制字段");

            return fmt ?? result;
        }
        public static DataFormat GetDataFormat(string filename, string sheetName)
        {
            DataFormat defaultFmt = new DataFormat() { format = FormatType.None };
            if (!CheckValid(filename, sheetName))
                return defaultFmt;

            if (!File.Exists(filename))
                return defaultFmt;

            using (ExcelFile excel = new ExcelFile(new FileInfo(filename), true))
            {
                int index = excel.GetWorksheetIndexFromName(sheetName);
                return GetDataFormat(filename, index);
            }
        }

        // 获取表格数据名称
        public static DataNameDict GetNames(string filename, int sheetIndex = 1)
        {
            if (!CheckValid(filename, sheetIndex))
                return null;

            Row[] rows = GetValidRows(filename, sheetIndex);
            if (rows.Length < Const.ROW_LINE_NUM_NAME)
            {
                Console.PrintError($"数据表{Path.GetFileName(filename)}表{sheetIndex}第{Const.ROW_LINE_NUM_NAME}个有效行不是字段名称行");
                return null;
            }

            return GetNames(rows, filename, sheetIndex);
        }
        public static DataNameDict GetNames(string filename, string sheetName)
        {
            if (!CheckValid(filename, sheetName))
                return null;

            using (ExcelFile excel = new ExcelFile(new FileInfo(filename), true))
            {
                int index = excel.GetWorksheetIndexFromName(sheetName);
                if (index == 0)
                {
                    Console.PrintError($"数据表{Path.GetFileName(filename)}不存在名为{sheetName}的表格");
                }
                return GetNames(filename, index);
            }
        }

        // 获取表格数据类型
        public static DataTypeDict GetTypes(string filename, int sheetIndex = 1)
        {
            if (!CheckValid(filename, sheetIndex))
                return null;

            Row[] rows = GetValidRows(filename, sheetIndex);
            if (rows.Length < Const.ROW_LINE_NUM_TYPE)
            {
                Console.PrintError($"数据表{Path.GetFileName(filename)}表{sheetIndex}第{Const.ROW_LINE_NUM_TYPE}个有效行不是字段类型行");
                return null;
            }

            return GetTypes(rows, filename, sheetIndex);
        }
        public static DataTypeDict GetTypes(string filename, string sheetName)
        {
            if (!CheckValid(filename, sheetName))
                return null;

            using (ExcelFile excel = new ExcelFile(new FileInfo(filename), true))
            {
                int index = excel.GetWorksheetIndexFromName(sheetName);
                if (index == 0)
                {
                    Console.PrintError($"数据表{Path.GetFileName(filename)}不存在名为{sheetName}的表格");
                }
                return GetTypes(filename, index);
            }
        }

        // 获取表格数据（除格式、类型、名称外）
        public static DataDict GetTableData(string filename, int sheetIndex = 1)
        {
            if (!CheckValid(filename, sheetIndex))
                return null;

            return GetTableData(GetValidRows(filename, sheetIndex));
        }
        public static DataDict GetTableData(string filename, string sheetName)
        {
            if (!CheckValid(filename, sheetName))
                return null;

            using (ExcelFile excel = new ExcelFile(new FileInfo(filename), true))
            {
                int index = excel.GetWorksheetIndexFromName(sheetName);
                if (index == 0)
                {
                    Console.PrintError($"数据表{Path.GetFileName(filename)}不存在名为{sheetName}的表格");
                }
                return GetTableData(filename, index);
            }
        }

        public static ExcelData GetExcelData(string filename, int sheetIndex = 1)
        {
            if (!CheckValid(filename, sheetIndex))
                return null;

            Row[] rows = GetValidRows(filename, sheetIndex);

            var fmt = GetDataFormat(rows);
            if (fmt.HasValue)
            {
                ExcelData data = new ExcelData();
                data.Filename = filename;
                data.SheetIndex = sheetIndex;
                data.SheetName = GetSheetNameByIndex(filename, sheetIndex);
                data.Format = fmt.Value;
                data.Types = GetTypes(rows, filename, sheetIndex);
                data.Names = GetNames(rows, filename, sheetIndex);
                data.Datas = GetTableData(rows, data.Names.Keys);
                data.DataBeginRowNumber = rows.Length <= Const.ROW_LINE_NUM_DATA ? 0 :
                    rows[Const.ROW_LINE_NUM_DATA].RowNumber;
                return data;
            }
            else
            {
                return null;
            }
        }
        public static ExcelData GetExcelData(string filename, string sheetName)
        {
            if (!CheckValid(filename, sheetName))
                return null;

            using (ExcelFile excel = new ExcelFile(new FileInfo(filename), true))
            {
                int index = excel.GetWorksheetIndexFromName(sheetName);
                if (index == 0)
                {
                    Console.PrintError($"数据表{Path.GetFileName(filename)}不存在名为{sheetName}的表格");
                }
                return GetExcelData(filename, index);
            }
        }

        public static void Test(string filename, int idx)
        {
            var sheet = GetWorksheet(filename, idx);
            //Console.Print($"{sheet.ExistingHeadingRows}");
            foreach (var row in sheet.Rows)
            {
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                foreach (var cell in row.Cells)
                {
                    sb.AppendFormat("{0}:{1} ", cell.CellName, cell.Value);
                }
                Console.Print(sb.ToString());
            }
        }

        #region Internal

        private static Worksheet GetWorksheet(string filename, int sheetIndex = 1)
        {
            if (!File.Exists(filename))
                return null;

            try
            {
                ExcelFile excel = new ExcelFile(new FileInfo(filename), true);
                if (excel.Worksheets.Length <= sheetIndex - 1)
                {
                    excel.Dispose();
                    return null;
                }

                return excel.Read(sheetIndex);
            }
            catch (Exception e)
            {
                Console.PrintError($"加载数据表{Path.GetFileName(filename)}失败，{e.Message}");
                return null;
            }
        }

        // 获取所有有效行，不过滤注释列
        private static Row[] GetValidRows(string filename, int sheetIndex = 1)
        {
            if (!File.Exists(filename))
            {
                Console.PrintError($"不存在的表格\"{Path.GetFileName(filename)}\"");
                return null;
            }

            Worksheet sheet = GetWorksheet(filename, sheetIndex);
            if (sheet == null)
                return null;

            if (!sheet.Exists)
            {
                Console.PrintError($"不存在的表格\"{Path.GetFileName(filename)}\"");
                return null;
            }

            List<Row> rows = new List<Row>();

            foreach (var row in sheet.Rows)
            {
                bool isNote = row.Cells.ElementAt(0).Value.ToString().Trim().StartsWith(Const.NOTE_PREFIX);

                //Console.Print($"表格{filename}第{row.RowNumber}行{row.Cells.ElementAt(0).CellName}是不是注释？{isNote}");

                if (isNote)
                    continue;

                rows.Add(row);
            }

            return rows.ToArray();
        }

        private static DataFormat? GetDataFormat(Row[] rows)
        {
            string fmtStr = rows[Const.ROW_LINE_NUM_FORMAT].Cells.ToArray()[0].Value.ToString();
            try
            {
                var result = JsonConvert.DeserializeObject<DataFormat>(fmtStr);
                return result;
            }
            catch
            {
                return null;
            }
        }
        private static DataNameDict GetNames(Row[] rows, string filename, int sheetIndex)
        {
            DataNameDict result = new DataNameDict();
            HashSet<string> names = new HashSet<string>();
            foreach (var cell in rows[Const.ROW_LINE_NUM_NAME].Cells)
            {
                string cellStr = cell.Value.ToString().Trim();
                string fieldName = Utils.ToFieldName(cellStr);
                if (string.IsNullOrEmpty(fieldName))
                {
                    Console.PrintError($"数据表{Path.GetFileName(filename)}表{sheetIndex}位置为" +
                        $"{cell.ColumnName}{rows[Const.ROW_LINE_NUM_NAME].RowNumber}的数据名称非法，非法名称将被忽略");
                    continue;
                }

                if (names.Contains(fieldName))
                {
                    Console.PrintWarning($"数据表{Path.GetFileName(filename)}表{sheetIndex}位置为" +
                        $"{cell.ColumnName}{rows[Const.ROW_LINE_NUM_NAME].RowNumber}的数据名称重复，重复数据将被忽略");
                    continue;
                }

                names.Add(fieldName);

                ConverterSettings cs = new ConverterSettings()
                {
                    isIgnore = cellStr.StartsWith(Const.NOTE_PREFIX),
                    cantEmpty = cellStr.EndsWith(Const.NON_EMPTY_SUFFIX)
                };

                result[cell.ColumnName] = new CellName()
                {
                    name = cellStr.Trim(Const.NOTE_PREFIX, Const.NON_EMPTY_SUFFIX),
                    fieldName = fieldName,
                    settings = cs
                };
            }

            return result;
        }
        private static DataTypeDict GetTypes(Row[] rows, string filename, int sheetIndex)
        {
            DataTypeDict result = new DataTypeDict();            
            foreach (var cell in rows[Const.ROW_LINE_NUM_TYPE].Cells)
            {
                string cellStr = cell.Value.ToString().Trim();                
                var type = TypeParser.Parse(cellStr);
                if (type == null)
                {
                    Console.PrintError($"不支持的数据类型\'{TypeParser.SplitType(cellStr)[0]}\'，位于数据表{Path.GetFileName(filename)}表{sheetIndex}的" +
                                        $"{cell.CellName}项，该项数据将被忽略");
                    continue;
                }
                else
                {
                    result[cell.ColumnName] = type;
                }
            }

            return result;
        }
        private static DataDict GetTableData(Row[] rows, IEnumerable<string> columnNames = null)
        {
            DataDict data = new DataDict();
            for (int i = Const.ROW_LINE_NUM_DATA; i < rows.Length; i++)
            {
                Dictionary<string, object> rowData = new Dictionary<string, object>();
                if (columnNames == null)
                {
                    foreach (var cell in rows[i].Cells)
                    {
                        rowData[cell.ColumnName] = cell.Value;
                    }
                }
                else
                {
                    foreach (var name in columnNames)
                    {
                        rowData[name] = rows[i].GetCellByColumnName(name)?.Value;
                    }
                }

                data[rows[i].RowNumber] = rowData;
            }

            return data;
        }

        #endregion


    }
}
