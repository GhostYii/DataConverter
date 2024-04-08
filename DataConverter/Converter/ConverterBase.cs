namespace DataConverter.Converter
{
    internal abstract class ConverterBase
    {        
        public abstract bool CheckConvert(string extension);
        public abstract string ToJson(string filename, int sheetIndex);
        public abstract string ToCSharpClass(string filename, int sheetIndex, string className);
        public abstract T FromData<T>(string filename) where T : new();
    }
}
