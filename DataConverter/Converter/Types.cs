﻿using System;
using System.Collections;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace DataConverter
{
    public enum FormatType
    {
        None = 0,
        Array,
        KeyValuePair,
        Object
    }

    public enum ValueType
    {
        Null = 0,
        Int,
        Float,
        String,
        Object,
        Array,
        Map,
    }

    [JsonConverter(typeof(DataFormatConverter))]
    public struct DataFormat
    {
        public FormatType format;
        public string key;
        public string type;
    }


    public class CellType
    {
        public ValueType type = ValueType.Null;
        public CellType subType = null;
        public string objName = string.Empty;

        public static CellType Default => new CellType() { type = ValueType.Null };

        public Type ToType()
        {
            switch (type)
            {
                case ValueType.Int:
                    return typeof(int);
                case ValueType.Float:
                    return typeof(float);
                case ValueType.String:
                    return typeof(string);
                case ValueType.Object:
                    return Type.GetType(objName);
                case ValueType.Array:
                    return Type.GetType(TypeName);
                case ValueType.Map:
                    return Type.GetType(TypeName);
            }
            return typeof(object);
        }

        // TODO: array & dictionary not implement
        public string TypeName
        {
            get
            {
                switch (type)
                {
                    case ValueType.Null:
                        return null;     // Error Case
                    case ValueType.Int:
                        return typeof(float).FullName;
                    case ValueType.Float:
                        return typeof(float).FullName;
                    case ValueType.String:
                        return typeof(string).FullName;
                    case ValueType.Object:
                        return objName;
                    case ValueType.Array:
                        return $"System.Collections.Generic.List`1[{subType.TypeName}]";
                    case ValueType.Map:
                        return $"System.Collections.Generic.Dictionary`2[{typeof(string).FullName},{subType.TypeName}]";
                    default:
                        return typeof(object).FullName;
                }
            }
        }
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
