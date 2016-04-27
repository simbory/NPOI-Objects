using System;
using NPOI.SS.UserModel;

namespace NPOI.Objects
{
    public abstract class StyleAttribute : Attribute
    {
        public int Height { get; set; }

        public string TextColor { get; set; }

        public string BackgroundColor { get; set; }

        public string ForegroundColor { get; set; }

        public HorizontalAlignment TextAlign { get; set; }

        public VerticalAlignment VerticalAlign { get; set; }

        public FillPattern FillPattern { get; set; }

        public short FontWeight { get; set; }

        public string FontFamily { get; set; }

        public short FontSize { get; set; }

        public bool IsItalic { get; set; }

        protected StyleAttribute()
        {
            TextAlign = HorizontalAlignment.General;
            VerticalAlign = VerticalAlignment.Top;
            FontSize = -1;
            FillPattern = FillPattern.SolidForeground;
        }
    }
}