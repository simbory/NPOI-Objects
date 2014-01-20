using System;
using System.Drawing;

namespace NPOI.Objects
{
    public static class ColorConvert
    {
        public static Color? ToColor(this string hex)
        {
            if (string.IsNullOrEmpty(hex))
                return null;
            try
            {
                var colorConverter = new ColorConverter();
                var color = colorConverter.ConvertFromString(hex);
                if (color != null)
                    return (Color) color;
            }
            catch (Exception ex)
            {
                return null;
            }
            return null;
        }
    }
}
