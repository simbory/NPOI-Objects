using System.Collections.Generic;
using System.Text;
using NPOI.SS.UserModel;

namespace NPOI.Objects
{
    class RichStyleString
    {
        public IFont CurrentFont { get; set; }

        public List<char> CharList { get; private set; }

        public RichStyleString()
        {
            CharList = new List<char>();
        }

        public bool IsCurrentStyle(IFont font)
        {
            if (CurrentFont == null)
                return false;
            return font.Boldweight == CurrentFont.Boldweight
                   && font.IsItalic == CurrentFont.IsItalic
                   && font.FontName == CurrentFont.FontName;
        }

        public string ToHtml()
        {
            if (CharList == null || CharList.Count < 1)
                return "";
            var builder = new StringBuilder();
            builder.Append("<span");
            if (CurrentFont != null)
            {
                builder.AppendFormat(@" style=""font-weight:{0};font-style:{1};font-family:'{2}'""",
                    CurrentFont.Boldweight,
                    CurrentFont.IsItalic ? "italic": "normal",
                    CurrentFont.FontName);
            }
            builder.Append(">");
            builder.Append(CharList.ToArray());
            builder.Append("</span>");
            return builder.ToString();
        }
    }
}