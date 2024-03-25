namespace DataConverter
{
    class CustomCommands
    {
        [CMD("get_fmt", "获取数据表格式")]
        private static void ParseFormat(string path)
        {
            var fmt = ExcelHelper.GetDataFormat(path);
            Commands.Terminal?.Append($"格式为{fmt.format}, key:{fmt.key}, type:{fmt.type}");
        }

        [CMD("types", "获取数据表内的类型")]
        private static void PrintTypes(int index)
        {
            var types = ExcelHelper.GetTypes("测试表格.xlsx", index);
            if (types == null)
                return;
            foreach (var type in types)
            {
                Commands.Terminal?.Append($"{type.Key}:{type.Value.type}");
            }
        }

        [CMD("names", "获取数据表内的所有字段名称")]
        private static void PrintNames(int index)
        {
            var names = ExcelHelper.GetNames("测试表格.xlsx", index);
            if (names == null)
                return;

            foreach (var name in names)
            {
                Commands.Terminal?.Append($"{name.Key}:{name.Value.name}({name.Value.fieldName})");
            }
        }

        [CMD("array")]
        private static void ArrayTest()
        {
            Converter.ExcelConverter ec = new Converter.ExcelConverter();
            string json = ec.ToJson("测试表格.xlsx", "数组测试");
            Console.Print(json);
        }

        [CMD("map")]
        private static void MapTest()
        {
            Converter.ExcelConverter ec = new Converter.ExcelConverter();
            string json = ec.ToJson("测试表格.xlsx", "字典测试");
            Console.Print(json);
        }
    }

    [System.Serializable]
    public class TestObject
    {
        public int num;
        public float percent;
        public string text;

        //public System.Collections.Generic.List<int> list;
    }
}
