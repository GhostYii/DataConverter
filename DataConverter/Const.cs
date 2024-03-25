namespace DataConverter
{
    internal static class Const
    {
        public const char NOTE_PREFIX = '#';    // 注释前缀        
        public const char NON_EMPTY_SUFFIX = '*';   // 非空后缀
        public const char TYPE_SPLIT_CHAR = ':';    // 类型分隔符

        public const int ROW_LINE_NUM_FORMAT = 0;   // 格式所在有效行
        public const int ROW_LINE_NUM_TYPE = 1;     // 类型所在有效行
        public const int ROW_LINE_NUM_NAME = 2;     // 名称所在有效行
        public const int ROW_LINE_NUM_DATA = 3;     // 数据起始有效行

        // 用于分割类型的分隔符
        //public static readonly char[] TYPE_SPLIT_CHARS = new char[] { ':' };
    }

}
