using System;
using System.Text;
using NPOI.SS.UserModel;

namespace NPOI.Objects
{
    /// <summary>
    /// the default cell value converters
    /// </summary>
    public static class CellValueConverters
    {
        /// <summary>
        /// convert the cell value to rich text string(HTML)
        /// </summary>
        /// <param name="cell">the excel cell</param>
        /// <returns>the rich text string (HTML)</returns>
        public static string RichTextConverter(ICell cell)
        {
            var stringBuilder = new StringBuilder();
            stringBuilder.Append("<p>");
            var styleChars = new RichStyleString();
            var richText = cell.RichStringCellValue;
            for (int i = 0; i < richText.Length; i++)
            {
                var chara = richText.String[i];
                if (chara.Equals('\r'))
                    continue;
                if (chara.Equals('\n'))
                {
                    stringBuilder.Append("<br/>");
                    continue;
                }
                IFont font = null;
                try
                {
                    var fontIndex = richText.GetFontAtIndex(i);
                    if (fontIndex >= 0)
                        font = cell.Sheet.Workbook.GetFontAt(fontIndex);
                }
                catch (Exception)
                {
                    font = null;
                }
                if (font == null)
                {
                    stringBuilder.Append(chara);
                    continue;
                }
                if (styleChars.IsCurrentStyle(font))
                {
                    styleChars.CharList.Add(chara);
                }
                else
                {
                    stringBuilder.Append(styleChars.ToHtml());
                    styleChars.CurrentFont = font;
                    styleChars.CharList.Clear();
                    styleChars.CharList.Add(chara);
                }
            }
            if (styleChars.CharList.Count > 0)
            {
                stringBuilder.Append(styleChars.ToHtml());
            }
            stringBuilder.Append("</p>");
            return stringBuilder.ToString();
        }

        /// <summary>
        /// the default converter that be used to convert the cell value to boolean 
        /// </summary>
        /// <param name="cell">the excel cell</param>
        /// <returns>the boolean value</returns>
        public static bool BooleanConverter(ICell cell)
        {
            try
            {
                return cell.BooleanCellValue;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// the default converter used to convert the cell value to number
        /// </summary>
        /// <param name="cell">the excel cell</param>
        /// <returns>the number</returns>
        public static double NumericConverter(ICell cell)
        {
            try
            {
                return cell.NumericCellValue;
            }
            catch (Exception)
            {
                return 0;
            }
        }

        /// <summary>
        /// the default converter used to convert the cell value to DateTime
        /// </summary>
        /// <param name="cell">the excel cell</param>
        /// <returns>the datetime value</returns>
        public static DateTime DateTimeConverter(ICell cell)
        {
            try
            {
                return cell.DateCellValue;
            }
            catch (Exception)
            {
                return DateTime.MinValue;
            }
        }

        /// <summary>
        /// the default converter used to convert the cell value to byte
        /// </summary>
        /// <param name="cell">the excel cell</param>
        /// <returns>the byte value</returns>
        public static byte ByteConverter(ICell cell)
        {
            try
            {
                return cell.ErrorCellValue;
            }
            catch (Exception)
            {
                return 0;
            }
        }

        /// <summary>
        /// the default converter used to convert the cell value to char
        /// </summary>
        /// <param name="cell">the excel value</param>
        /// <returns>the char value</returns>
        public static char CharConverter(ICell cell)
        {
            try
            {
                if (cell.CellType == CellType.String)
                {
                    var str = cell.StringCellValue;
                    return string.IsNullOrEmpty(str) ? (char)0 : str[0];
                }
                return (char) 0;
            }
            catch (Exception)
            {
                return (char) 0;
            }
        }

        /// <summary>
        /// the default converter used to convert the cell value to Guid
        /// </summary>
        /// <param name="cell">the excel cell</param>
        /// <returns>the Guid value</returns>
        public static Guid GuidConverter(ICell cell)
        {
            try
            {
                if (cell.CellType == CellType.String)
                {
                    var strValue = cell.StringCellValue;
                    if (!string.IsNullOrEmpty(strValue))
                    {
                        Guid guid;
                        return Guid.TryParse(strValue, out guid) ? guid : Guid.Empty;
                    }
                }
                return Guid.Empty;
            }
            catch (Exception)
            {
                return Guid.Empty;
            }
        }

        /// <summary>
        /// the default converter used to convert the cell value to any other object
        /// </summary>
        /// <param name="cell">the excel cell</param>
        /// <param name="type">the type of the object</param>
        /// <returns>the value of the object</returns>
        public static object UnknownTypeConverter(ICell cell, Type type)
        {
            return null;
        }
    }
}