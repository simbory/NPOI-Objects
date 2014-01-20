using System;
using System.Linq;
using System.Reflection;
using System.Text;
using NPOI.SS.UserModel;

namespace NPOI.Objects
{
    public static class NPOIExtension
    {
        public static string ToHtml(this ICell cell)
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
                catch(Exception ex)
                {
                    Console.Write(ex);
                    font = null;
                }
                if (font == null )
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

        public static T GetCustomAttribute<T>(this PropertyInfo property) where T: class
        {
            var attrs = property.GetCustomAttributes(typeof (T), false);
            if (attrs.Length < 1)
            {
                return null;
            }
            return (T) attrs.First();
        }

        public static T GetCustomAttribute<T>(this Type type) where T : class
        {
            var attrs = type.GetCustomAttributes(typeof(T), false);
            if (attrs.Length < 1)
            {
                return null;
            }
            return (T)attrs.First();
        }
    }
}
