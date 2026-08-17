using System.Collections;
using MongoDB.Bson;
using MongoDB.Bson.Serialization;
using MongoDB.Bson.Serialization.Options;
using MongoDB.Bson.Serialization.Serializers;

namespace DataConverter.Core
{
    public static class BsonConverter
    {
        // MongoDB.Bson 默认拒绝可能丢失精度的 double -> float 转换。
        // 数据表中的 float 在 JSON/BSON 中会被表示为 double，因此允许截断转换。
        static BsonConverter()
        {
            RepresentationConverter converter = new RepresentationConverter(false, true);
            BsonSerializer.RegisterSerializer(typeof(float), new SingleSerializer(BsonType.Double, converter));
        }

        // BSON 的根节点必须是文档，数组表统一包装在该字段下。
        public const string ArrayItemsElementName = "items";

        // MongoDB.Bson 默认会把名为 id/Id 的成员映射为 _id，
        // ExcelToBson 时同步该字段名，保证 FromBson<T> 能正确读回。
        private const string ID_ELEMENT_NAME = "_id";

        public static byte[] ToBson<T>(T obj)
        {
            if (obj == null)
                throw new ArgumentNullException(nameof(obj));

            if (TryGetListElementType(typeof(T), out _))
            {
                BsonArray array = new BsonArray();
                foreach (object? item in (IEnumerable)obj)
                    array.Add(ToBsonValue(item));

                return new BsonDocument(ArrayItemsElementName, array).ToBson();
            }

            return obj.ToBson();
        }

        public static byte[] ExcelToBson(string filename, string sheetName)
        {
            ExcelConverter converter = new ExcelConverter();
            if (!converter.CheckConvert(Path.GetExtension(filename)))
            {
                Console.PrintError($"数据表'{Path.GetFileName(filename)}'不支持BSON转换");
                return Array.Empty<byte>();
            }

            return ConvertJsonToBson(converter.ToJson(filename, sheetName), filename, sheetName);
        }

        public static byte[] ExcelToBson(string filename, int sheetIndex)
        {
            ExcelConverter converter = new ExcelConverter();
            if (!converter.CheckConvert(Path.GetExtension(filename)))
            {
                Console.PrintError($"数据表'{Path.GetFileName(filename)}'不支持BSON转换");
                return Array.Empty<byte>();
            }

            return ConvertJsonToBson(converter.ToJson(filename, sheetIndex), filename, sheetIndex.ToString());
        }

        public static T FromBson<T>(byte[] bson)
        {
            if (bson == null)
                throw new ArgumentNullException(nameof(bson));
            if (bson.Length == 0)
                throw new ArgumentException("BSON数据不能为空", nameof(bson));

            Type type = typeof(T);
            if (TryGetListElementType(type, out Type elementType))
                return DeserializeList<T>(bson, type, elementType);

            return BsonSerializer.Deserialize<T>(bson);
        }

        public static T FromBson<T>(string filename)
        {
            using (FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (MemoryStream ms = new MemoryStream())
            {
                fs.CopyTo(ms);
                return FromBson<T>(ms.ToArray());
            }
        }

        private static byte[] ConvertJsonToBson(string json, string filename, string sheetName)
        {
            if (string.IsNullOrEmpty(json))
            {
                Console.PrintError($"数据表'{Path.GetFileName(filename)}'表'{sheetName}'没有可转换为BSON的数据");
                return Array.Empty<byte>();
            }

            BsonValue root = NormalizeIdElement(BsonSerializer.Deserialize<BsonValue>(json));
            if (root is BsonArray array)
                return new BsonDocument(ArrayItemsElementName, array).ToBson();

            return root.AsBsonDocument.ToBson();
        }

        private static BsonValue NormalizeIdElement(BsonValue value)
        {
            switch (value.BsonType)
            {
                case BsonType.Document:
                {
                    BsonDocument document = value.AsBsonDocument;
                    BsonDocument result = new BsonDocument();
                    foreach (BsonElement element in document)
                    {
                        string name = string.Equals(element.Name, "id", StringComparison.OrdinalIgnoreCase)
                            ? ID_ELEMENT_NAME
                            : element.Name;

                        if (result.Contains(name))
                        {
                            Console.PrintWarning($"BSON字段'{element.Name}'与已有字段'{name}'冲突，重复字段将被忽略");
                            continue;
                        }

                        result.Add(name, NormalizeIdElement(element.Value));
                    }

                    return result;
                }
                case BsonType.Array:
                {
                    BsonArray array = new BsonArray();
                    foreach (BsonValue item in value.AsBsonArray)
                        array.Add(NormalizeIdElement(item));
                    return array;
                }
                default:
                    return value;
            }
        }

        private static T DeserializeList<T>(byte[] bson, Type targetType, Type elementType)
        {
            BsonDocument document = BsonSerializer.Deserialize<BsonDocument>(bson);
            if (!document.TryGetElement(ArrayItemsElementName, out BsonElement element) ||
                element.Value is not BsonArray array)
            {
                throw new FormatException($"BSON数据缺少数组字段'{ArrayItemsElementName}'");
            }

            Type listType = typeof(List<>).MakeGenericType(elementType);
            IList items = (IList)Activator.CreateInstance(listType)!;
            foreach (BsonValue item in array)
            {
                if (item is not BsonDocument itemDocument)
                    throw new FormatException($"BSON数组'{ArrayItemsElementName}'中包含非文档元素，无法转换为'{targetType}'");

                items.Add(BsonSerializer.Deserialize(itemDocument, elementType)!);
            }

            if (targetType.IsAssignableFrom(items.GetType()))
                return (T)items;

            if (targetType.IsArray)
            {
                Array result = Array.CreateInstance(elementType, items.Count);
                items.CopyTo(result, 0);
                return (T)(object)result;
            }

            throw new FormatException($"不支持将BSON数组转换为类型'{targetType}'");
        }

        private static bool TryGetListElementType(Type type, out Type elementType)
        {
            elementType = typeof(object);
            if (type.IsArray)
            {
                elementType = type.GetElementType()!;
                return true;
            }

            if (!type.IsGenericType)
                return false;

            Type definition = type.GetGenericTypeDefinition();
            if (definition == typeof(List<>) ||
                definition == typeof(IList<>) ||
                definition == typeof(ICollection<>) ||
                definition == typeof(IEnumerable<>) ||
                definition == typeof(IReadOnlyList<>) ||
                definition == typeof(IReadOnlyCollection<>))
            {
                elementType = type.GetGenericArguments()[0];
                return true;
            }

            return false;
        }

        private static BsonValue ToBsonValue(object? item)
        {
            if (item == null)
                return BsonNull.Value;

            if (item is BsonValue value)
                return value;

            return BsonDocumentWrapper.Create(item.GetType(), item);
        }
    }
}
