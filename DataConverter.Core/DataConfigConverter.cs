using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace DataConverter.Core
{
    internal class DataConfigConverter : JsonConverter
    {
        public override bool CanWrite => false;

        public override bool CanConvert(Type objectType)
        {
            return objectType == typeof(DataConfig);
        }

        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            DataConfig result = new DataConfig()
            {
                format = FormatType.None,
                key = string.Empty,
                type = string.Empty,
                objectType = ObjectType.None,
                objectName = string.Empty
            };

            JObject jsonObj = serializer.Deserialize<JObject>(reader);
            switch (jsonObj.Value<string>("format").ToLower())
            {
                case "arr":
                    result.format = FormatType.Array;
                    break;
                case "array":
                    result.format = FormatType.Array;
                    break;
                case "map":
                    result.format = FormatType.KeyValuePair;
                    result.key = jsonObj.Value<string>("key");
                    break;
                default:
                    result.format = FormatType.None;
                    break;
            }

            // default generate struct
            string objTypeStr = jsonObj.Value<string>("obj_type");
            if (!string.IsNullOrEmpty(objTypeStr))
            {
                objTypeStr = objTypeStr.Trim().ToLower();
                objTypeStr = $"{char.ToUpper(objTypeStr[0])}{objTypeStr.Substring(1, objTypeStr.Length - 1)}";
            }
            if (!Enum.TryParse(objTypeStr, out result.objectType))
                result.objectType = ObjectType.Struct;

            result.objectName = jsonObj.Value<string>("obj_name") ?? string.Empty;

            return result;
        }

        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            return;
        }
    }
}
