using System;

namespace NPOI.Objects.Attributes
{
    public class NPOIColumnAttribute : Attribute
    {
        public int Index { get; set; }

        public string Name { get; set; }

        public NPOIColumnAttribute()
        {
            Index = -1;
        }

        public NPOIColumnAttribute(int index)
        {
            Index = index;
        }

        public NPOIColumnAttribute(string name)
        {
            Name = name;
            Index = -1;
        }

        public NPOIColumnAttribute(int index, string name)
        {
            Index = index;
            Name = name;
        }
    }
}
