namespace DataConverter.Core
{
    internal class ExcelData
    {
        public int DataBeginRowNumber { get; internal set; } = 0;
        public string Filename { get; internal set; } = string.Empty;
        public int SheetIndex { get; internal set; } = 0;
        public string SheetName { get; internal set; } = string.Empty;

        public DataConfig Config { get; internal set; } = new();
        public Dictionary<string, CellName> Names { get; internal set; } = new();
        public Dictionary<string, CellType> Types { get; internal set; } = new();

        // { rowNumber : { columnName : value } }
        public Dictionary<int, Dictionary<string, string>> Datas { get; internal set; } = new();

        public Dictionary<string, List<string>> Enums { get; internal set; } = new();

        public void UpdateDataByEnums()
        {
            if (Enums.Count == 0)
                return;

            var newData = new List<ValueTuple<int, string, string>>();

            foreach (var (colName, cellType) in Types)
            {
                if (cellType.type != CellValueType.Enum)
                    continue;

                foreach (var (rowNum, data) in Datas)
                {
                    var val = data[colName];
                    if (string.IsNullOrEmpty(val))
                    {
                        newData.Add((rowNum, colName, "0"));
                        continue;
                    }

                    if (int.TryParse(val, out int index))
                    {
                        if (index >= Enums[cellType.objName].Count)
                            newData.Add((rowNum, colName, "0"));
                        else
                            newData.Add((rowNum, colName, val));
                        continue;
                    }

                    string fieldName = Utils.ToFieldName(val);
                    if (string.IsNullOrEmpty(fieldName))
                    {
                        newData.Add((rowNum, colName, "0"));
                        continue;
                    }

                    newData.Add((rowNum, colName, Enums[cellType.objName].IndexOf(fieldName).ToString()));
                }
            }

            foreach (var (rowNum, colName, val) in newData)
            {
                Datas[rowNum][colName] = val;
            }
        }
    }
}
