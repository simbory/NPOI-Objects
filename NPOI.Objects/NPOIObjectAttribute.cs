using System;

namespace NPOI.Objects
{
    public class NPOIObjectAttribute : Attribute
    {
        public int HeaderRowIndex { get; set; }

        public int StartIndex { get; set; }

        public int EndIndex { get; set; }

        public NPOIObjectAttribute(int headerRow = 0, int startIndex = 1, int endIndex = -1)
        {
            HeaderRowIndex = headerRow;
            StartIndex = startIndex;
            EndIndex = endIndex;
        }
    }
}