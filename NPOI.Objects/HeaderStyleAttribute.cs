using System;

namespace NPOI.Objects
{
    /// <summary>
    /// the HeaderStyleAttribute class is used to set the header style while drawing the excel file
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    [Serializable]
    public class HeaderStyleAttribute : StyleAttribute
    {
        /// <summary>
        /// the column width, the default value is 6
        /// </summary>
        public ushort ColumnWidth { get; set; }

        /// <summary>
        /// the constructor
        /// </summary>
        public HeaderStyleAttribute()
        {
            ColumnWidth = 6;
        }
    }
}