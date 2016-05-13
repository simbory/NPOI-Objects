using System;

namespace NPOI.Objects
{
    /// <summary>
    /// NPOIColumnAttribute indicate that this property will be mapped to a featured excel column
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    [Serializable]
    public class NPOIColumnAttribute : Attribute
    {
        /// <summary>
        /// the column index
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        /// the column name
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// the constructor
        /// </summary>
        public NPOIColumnAttribute()
        {
            Index = -1;
        }

        /// <summary>
        /// the constructor
        /// </summary>
        /// <param name="index"></param>
        public NPOIColumnAttribute(int index)
        {
            Index = index;
        }

        /// <summary>
        /// the constructor
        /// </summary>
        /// <param name="name">the name of the column</param>
        public NPOIColumnAttribute(string name)
        {
            Name = name;
            Index = -1;
        }

        /// <summary>
        /// the constructor
        /// </summary>
        /// <param name="index">the column index</param>
        /// <param name="name">the column name</param>
        public NPOIColumnAttribute(int index, string name)
        {
            Index = index;
            Name = name;
        }
    }
}