namespace DataConverter.Core
{
    public abstract class ConverterBase
    {        
        public abstract bool CheckConvert(string extension);
        public abstract string ToJson(string filename, int sheetIndex);        
        public abstract string ToCSharp(string filename, int sheetIndex, string typename);        
        public abstract T FromData<T>(string filename) where T : new();        
    }
}
