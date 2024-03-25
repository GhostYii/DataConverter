using System.Collections.Generic;
using Newtonsoft.Json;

namespace DataConverter
{
    public enum FormatType
    {
        None,
        Array,
        KeyValuePair,
        Object
    }

    public enum ValueType
    {
        //Auto,
        Int,
        Float,
        String,
        Object,
        Ref
    }

    [JsonConverter(typeof(DataFormatConverter))]
    public struct DataFormat
    {
        public FormatType format;
        public string key;
        public string type;
    }

    public struct CellType
    {
        public ValueType type;
        public string objName;
        public string refPath;
    }

    public struct ConverterSettings
    {
        // 忽略数据
        public bool isIgnore;
        // 不得为空
        public bool cantEmpty;
    }

    public struct CellName
    {
        public string name;
        public string fieldName;
        public ConverterSettings settings;
    }

}
