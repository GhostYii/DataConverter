namespace DataConverter.Core
{
    internal class ExcelData
    {
        public int DataBeginRowNumber { get; internal set; }
        public string Filename { get; internal set; }
        public int SheetIndex { get; internal set; }
        public string SheetName { get; internal set; }

        public DataFormat Format { get; internal set; }
        public Dictionary<string, CellName> Names { get; internal set; }
        public Dictionary<string, CellType> Types { get; internal set; }
        public Dictionary<int, Dictionary<string, object>> Datas { get; internal set; }        
    }
}
