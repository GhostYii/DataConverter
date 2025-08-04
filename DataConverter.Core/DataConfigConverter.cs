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
                isTemplate = false,
                templateName = string.Empty,
                format = FormatType.None,
                key = string.Empty,
                type = string.Empty,
                objectType = ObjectType.None,
                objectName = string.Empty,
                genEnumType = true
            };

            JObject jsonObj = serializer.Deserialize<JObject>(reader)!;
            result.isTemplate = jsonObj.ContainsKey("is_template") ? jsonObj.Value<bool>("is_template") : false;
            result.templateName = jsonObj.ContainsKey("template_name") ? jsonObj.Value<string>("template_name") ?? string.Empty : string.Empty;

            if (jsonObj.ContainsKey("format") || jsonObj.ContainsKey("fmt"))
            {
                string format = jsonObj.ContainsKey("format") ? jsonObj.Value<string>("format")!.ToLower() : jsonObj.Value<string>("fmt")!.ToLower();

                switch (format)
                {
                    case "arr":
                        result.format = FormatType.Array;
                        break;
                    case "array":
                        result.format = FormatType.Array;
                        break;
                    case "map":
                        result.format = FormatType.KeyValuePair;
                        result.key = jsonObj.Value<string>("key")!;
                        break;
                    default:
                        result.format = FormatType.None;
                        break;
                }
            }

            // default generate struct
            string objTypeStr = jsonObj.Value<string>("obj_type")!;
            if (!string.IsNullOrEmpty(objTypeStr))
            {
                objTypeStr = objTypeStr.Trim().ToLower();
                objTypeStr = $"{char.ToUpper(objTypeStr[0])}{objTypeStr.Substring(1, objTypeStr.Length - 1)}";
            }
            if (!Enum.TryParse(objTypeStr, out result.objectType))
                result.objectType = ObjectType.Struct;

            result.objectName = jsonObj.Value<string>("obj_name") ?? string.Empty;
            result.genEnumType = jsonObj.ContainsKey("gen_enum") ? jsonObj.Value<bool>("gen_enum") : true;

            return result;
        }

        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            return;
        }
    }
}
