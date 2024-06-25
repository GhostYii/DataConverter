using MongoDB.Bson;
using MongoDB.Bson.Serialization;

namespace DataConverter.Core
{
    public static class BsonConverter
    {
        public static byte[] ToBson<T>(T obj)
        {
            return obj.ToBson();
        }

        public static byte[] ExcelToBson(string filename, string sheetName)
        {
            ExcelConverter converter = new ExcelConverter();
            string json = converter.ToJson(filename, sheetName);
            return BsonDocument.Parse(json).ToBson();
        }

        public static T FromBson<T>(byte[] bson)
        {
            return BsonSerializer.Deserialize<T>(bson);
        }

        public static T FromBson<T>(string filename)
        {
            return FromBson<T>(File.ReadAllBytes(filename));
        }

    }
}
