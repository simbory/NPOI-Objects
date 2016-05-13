using System;
using System.Drawing;

namespace NPOI.Objects
{
    /// <summary>
    /// convert the color string to color object
    /// </summary>
    public static class ColorConvert
    {
        /// <summary>
        /// convert the color string (#FF0000 or red) to color object
        /// </summary>
        /// <param name="hex">the string of the color, for example #FF0000, red</param>
        /// <returns>the color object</returns>
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
            catch (Exception)
            {
                return null;
            }
            return null;
        }
    }
}