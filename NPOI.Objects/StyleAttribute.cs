using System;
using NPOI.SS.UserModel;

namespace NPOI.Objects
{
    /// <summary>
    /// the StyleAttribute is used to the the style of the excel cell (header cell or the data cell)
    /// </summary>
    public abstract class StyleAttribute : Attribute
    {
        /// <summary>
        /// the height
        /// </summary>
        public int Height { get; set; }

        /// <summary>
        /// the text color
        /// </summary>
        public string TextColor { get; set; }

        /// <summary>
        /// the background colur
        /// </summary>
        public string BackgroundColor { get; set; }

        /// <summary>
        /// the foreground color
        /// </summary>
        public string ForegroundColor { get; set; }

        /// <summary>
        /// text aling
        /// </summary>
        public HorizontalAlignment TextAlign { get; set; }

        /// <summary>
        /// vertical aling
        /// </summary>
        public VerticalAlignment VerticalAlign { get; set; }

        /// <summary>
        /// fill pattern
        /// </summary>
        public FillPattern FillPattern { get; set; }

        /// <summary>
        /// the font weight
        /// </summary>
        public short FontWeight { get; set; }

        /// <summary>
        /// the font family
        /// </summary>
        public string FontFamily { get; set; }

        /// <summary>
        /// the font size
        /// </summary>
        public short FontSize { get; set; }

        /// <summary>
        /// the font is italic or not
        /// </summary>
        public bool IsItalic { get; set; }

        /// <summary>
        /// the constructor
        /// </summary>
        protected StyleAttribute()
        {
            TextAlign = HorizontalAlignment.General;
            VerticalAlign = VerticalAlignment.Top;
            FontSize = -1;
            FillPattern = FillPattern.SolidForeground;
        }
    }
}