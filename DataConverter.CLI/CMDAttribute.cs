using System;

namespace DataConverter.CLI
{
    [AttributeUsage(AttributeTargets.Method, AllowMultiple = true)]
    internal class CMDAttribute : Attribute
    {
        public string Name { get; private set; }
        public string Desc { get; private set; }

        public CMDAttribute()
        {
        }

        public CMDAttribute(string name, string desc = "")
        {
            Name = name.Replace(' ', '_');
            Desc = desc;
        }
    }
}
