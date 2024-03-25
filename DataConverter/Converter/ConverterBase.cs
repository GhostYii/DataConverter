namespace DataConverter.Converter
{
    internal abstract class ConverterBase
    {        
        public abstract bool CheckConvert(string extension);
        public abstract string ToJson(string filename);
        public abstract string ToCSharpClass(string filename);
        public abstract T FromData<T>(string filename) where T : new();
    }
}
