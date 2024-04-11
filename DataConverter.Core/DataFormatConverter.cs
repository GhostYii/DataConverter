using System;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace DataConverter.Core
{
    internal class DataFormatConverter : JsonConverter
    {
        public override bool CanWrite => false;

        public override bool CanConvert(Type objectType)
        {
            return objectType == typeof(DataFormat);
        }

        public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
        {
            DataFormat result = new DataFormat()
            {
                format = FormatType.None,
                key = string.Empty,
                type = string.Empty
            };

            JObject jsonObj = serializer.Deserialize<JObject>(reader);
            switch (jsonObj.Value<string>("format").ToLower())
            {
                case "array":
                    result.format = FormatType.Array;
                    break;
                case "map":
                    result.format = FormatType.KeyValuePair;
                    result.key = jsonObj.Value<string>("key");
                    break;
                case "object":
                    result.format = FormatType.Object;
                    result.type = jsonObj.Value<string>("type");
                    break;
                case "obj":
                    result.format = FormatType.Object;
                    result.type = jsonObj.Value<string>("type");
                    break;
                default:
                    result.format = FormatType.None;
                    break;
            }            

            return result;
        }

        public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
        {
            return;                
        }
    }
}
