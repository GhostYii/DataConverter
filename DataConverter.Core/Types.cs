using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace DataConverter.Core
{
    internal enum FormatType
    {
        None = 0,
        Array,
        KeyValuePair
    }

    internal enum ObjectType
    {
        None = 0,
        Struct,
        Class
    }

    internal enum CellValueType
    {
        Null = 0,
        Bool,
        Int,
        Float,
        String,
        Object,
        Array,
        Map
    }

    [JsonConverter(typeof(DataConfigConverter))]
    internal struct DataConfig
    {
        public FormatType format;
        public string key;
        public string type;
        public ObjectType objectType;
        public string objectName;
    }


    internal class CellType
    {
        public CellValueType type = CellValueType.Null;
        public CellType subType = null;
        public string objName = string.Empty;

        public static CellType Default => new CellType() { type = CellValueType.Null };

        public bool IsValueType { get => type.IsValueType(); }

        public Type Type
        {
            get
            {
                switch (type)
                {
                    case CellValueType.Bool:
                        return typeof(bool);
                    case CellValueType.Int:
                        return typeof(int);
                    case CellValueType.Float:
                        return typeof(float);
                    case CellValueType.String:
                        return typeof(string);
                    case CellValueType.Object:
                        return Type.GetType(objName) ?? DC.GetType(objName);
                    case CellValueType.Array:
                        return Type.GetType(FullTypeName);
                    case CellValueType.Map:
                        return Type.GetType(FullTypeName);
                }
                return typeof(object);
            }
        }

        public Type JsonType
        {
            get
            {
                switch (type)
                {
                    case CellValueType.Null:
                        return typeof(JObject);
                    case CellValueType.Bool:
                        return typeof(JValue);
                    case CellValueType.Int:
                        return typeof(JValue);
                    case CellValueType.Float:
                        return typeof(JValue);
                    case CellValueType.String:
                        return typeof(JValue);
                    case CellValueType.Object:
                        return typeof(JObject);
                    case CellValueType.Array:
                        return typeof(JArray);
                    case CellValueType.Map:
                        return typeof(JObject);
                }
                return typeof(JToken);
            }
        }

        // dictionary key's type can only be string
        public string FullTypeName
        {
            get
            {
                switch (type)
                {
                    case CellValueType.Null:
                        return null;     // Error Case
                    case CellValueType.Bool:
                        return typeof(bool).FullName;
                    case CellValueType.Int:
                        return typeof(int).FullName;
                    case CellValueType.Float:
                        return typeof(float).FullName;
                    case CellValueType.String:
                        return typeof(string).FullName;
                    case CellValueType.Object:
                        return objName;
                    case CellValueType.Array:
                        return $"System.Collections.Generic.List`1[{subType.FullTypeName}]";
                    case CellValueType.Map:
                        return $"System.Collections.Generic.Dictionary`2[{typeof(string).FullName},{subType.FullTypeName}]";
                }
                return typeof(object).FullName;
            }
        }

        public string TypeName
        {
            get
            {
                switch (type)
                {
                    case CellValueType.Null:
                        return string.Empty;
                    case CellValueType.Bool:
                        return "bool";
                    case CellValueType.Int:
                        return "int";
                    case CellValueType.Float:
                        return "float";
                    case CellValueType.String:
                        return "string";
                    case CellValueType.Object:
                        return objName;
                    case CellValueType.Array:
                        return $"List<{subType.TypeName}>";
                    case CellValueType.Map:
                        return $"Dictionary<string, {subType.TypeName}>";
                }
                return string.Empty;
            }
        }

        public JToken DefaultJsonValue
        {
            get
            {
                switch (type)
                {
                    case CellValueType.Bool:
                        return new JValue(false);
                    case CellValueType.Int:
                        return new JValue(0);
                    case CellValueType.Float:
                        return new JValue(0f);
                    case CellValueType.String:
                        return new JValue(string.Empty);
                    case CellValueType.Object:
                        return new JObject();
                    case CellValueType.Array:
                        return new JArray();
                    case CellValueType.Map:
                        return new JObject();
                }
                return new JObject();
            }
        }
    }

    internal struct ConverterSettings
    {
        // 忽略数据
        public bool isIgnore;
        // 不得为空
        public bool cantEmpty;
    }

    internal struct CellName
    {
        public string name;
        public string fieldName;
        public ConverterSettings settings;
    }

}
