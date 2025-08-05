namespace DataConverter.Core
{
    internal class ExcelData
    {
        public int DataBeginRowNumber { get; internal set; } = 0;
        public string Filename { get; internal set; } = string.Empty;
        public int SheetIndex { get; internal set; } = 0;
        public string SheetName { get; internal set; } = string.Empty;
        public DataConfig Config { get; internal set; } = new();

        // { rowNumber : { columnName : value } }        
        public Dictionary<int, Dictionary<string, string>> Datas { get; internal set; } = new();

        private Dictionary<string, CellName> _names = new();
        public Dictionary<string, CellName> SelfNames => _names;
        public Dictionary<string, CellName> Names { get => Template == null ? _names : Template.Names; internal set => _names = value; }

        private Dictionary<string, CellType> _types = new();
        public Dictionary<string, CellType> SelfTypes => _types;
        public Dictionary<string, CellType> Types { get => Template == null ? _types : Template.Types; internal set => _types = value; }

        private Dictionary<string, List<string>> _enums = new();
        public Dictionary<string, List<string>> Enums { get => Template == null ? _enums : Template.Enums; internal set => _enums = value; }

        internal ExcelData Template { get; set; } = null;

        public void UpdateConfigsByTemplate()
        {
            if (Template == null)
                return;

            DataConfig config = Template.Config;
            config.isTemplate = false;
            config.templateName = Config.templateName;
            config.genEnumType = false;

            Config = config;
        }

        public void UpdateDataByEnums()
        {
            if (Enums.Count == 0)
                return;

            // TODO: support more depth of enum

            var newData = new List<ValueTuple<int, string, string>>();

            foreach (var (colName, cellType) in SelfTypes)
            {
                if (cellType.type == CellValueType.Enum)
                    UpdateEnumValue(colName, cellType, newData);
                else if (cellType.type == CellValueType.Array && cellType.subType.type == CellValueType.Enum)
                    UpdateArrayEnumValue(colName, cellType, newData);
                else if (cellType.type == CellValueType.Map && cellType.subType.type == CellValueType.Enum)
                    UpdateMapEnumValue(colName, cellType, newData);

                // TODO: support object enum
            }

            foreach (var (rowNum, colName, val) in newData)
            {
                Datas[rowNum][colName] = val;
            }
        }

        private void UpdateEnumValue(string colName, CellType cellType, List<ValueTuple<int, string, string>> newData)
        {
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

        private void UpdateArrayEnumValue(string colName, CellType cellType, List<ValueTuple<int, string, string>> newData)
        {
            foreach (var (rowNum, data) in Datas)
            {
                var val = data[colName];
                List<int> enumVals = new();
                if (string.IsNullOrEmpty(val))
                    continue;

                val = val.TrimStart('[').TrimEnd(']');
                foreach (var singleData in val.Split(',', StringSplitOptions.RemoveEmptyEntries))
                {
                    if (int.TryParse(singleData, out var _))
                        continue;

                    var enumVal = singleData.TrimStart('"').TrimEnd('"').Trim();
                    enumVals.Add(Enums[cellType.subType.objName].IndexOf(Utils.ToFieldName(enumVal)));
                }

                string newVal = "[";
                foreach (var item in enumVals)
                {
                    newVal += $"{item},";
                }

                newData.Add((rowNum, colName, newVal.TrimEnd(',') + ']'));
            }
        }

        private void UpdateMapEnumValue(string colName, CellType cellType, List<(int, string, string)> newData)
        {
            foreach (var (rowNum, data) in Datas)
            {
                var val = data[colName];
                if (string.IsNullOrEmpty(val))
                    continue;

                List<string> pairs = new();
                val = val.TrimStart('{').TrimEnd('}');
                foreach (var pair in val.Split(',', StringSplitOptions.RemoveEmptyEntries))
                {
                    var pairData = pair.Split(':');
                    if (pairData.Length != 2)
                        continue;

                    var singleData = pairData[1];
                    if (int.TryParse(singleData, out var _))
                        continue;

                    var enumVal = singleData.TrimStart('"').TrimEnd('"').Trim();
                    pairs.Add($"{pairData[0]}:{Enums[cellType.subType.objName].IndexOf(Utils.ToFieldName(enumVal))}");
                }

                string newVal = "{";
                foreach (var pair in pairs)
                {
                    newVal += $"{pair},";
                }

                newData.Add((rowNum, colName, newVal.TrimEnd(',') + '}'));
            }
        }
    }
}
