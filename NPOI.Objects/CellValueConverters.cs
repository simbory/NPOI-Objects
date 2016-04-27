using System;
using System.Text;
using NPOI.SS.UserModel;

namespace NPOI.Objects
{
    public static class CellValueConverters
    {
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

        public static object UnknowTypeConverter(ICell cell, Type type)
        {
            return null;
        }
    }
}