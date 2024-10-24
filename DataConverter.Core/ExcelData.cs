﻿namespace DataConverter.Core
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

            var newData = new List<ValueTuple<int, string, string>>();

            foreach (var (colName, cellType) in SelfTypes)
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
