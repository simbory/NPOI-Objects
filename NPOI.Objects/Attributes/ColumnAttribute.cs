using System;

namespace NPOI.Objects.Attributes
{
    public class ColumnAttribute : Attribute
    {
        public int Index { get; set; }

        public string Name { get; set; }

        public ColumnAttribute()
        {
            Index = -1;
        }

        public ColumnAttribute(int index)
        {
            Index = index;
        }

        public ColumnAttribute(string name)
        {
            Name = name;
            Index = -1;
        }

        public ColumnAttribute(int index, string name)
        {
            Index = index;
            Name = name;
        }
    }
}
