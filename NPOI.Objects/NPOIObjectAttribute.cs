using System;

namespace NPOI.Objects
{
    /// <summary>
    /// NPOIObjectAttribute indicate that this model will be mapped to excel
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class NPOIObjectAttribute : Attribute
    {
        /// <summary>
        /// the header row index
        /// </summary>
        public int HeaderRowIndex { get; set; }

        /// <summary>
        /// the first data row index
        /// </summary>
        public int StartIndex { get; set; }

        /// <summary>
        /// the last data row index
        /// </summary>
        public int EndIndex { get; set; }

        /// <summary>
        /// the constructor
        /// </summary>
        /// <param name="headerRow">the header row index</param>
        /// <param name="startIndex">the first data row</param>
        /// <param name="endIndex">the last data row</param>
        public NPOIObjectAttribute(int headerRow = 0, int startIndex = 1, int endIndex = -1)
        {
            HeaderRowIndex = headerRow;
            StartIndex = startIndex;
            EndIndex = endIndex;
        }
    }
}