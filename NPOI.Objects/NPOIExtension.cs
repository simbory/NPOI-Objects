using System;
using System.Linq;
using System.Reflection;
using System.Text;
using NPOI.SS.UserModel;

namespace NPOI.Objects
{
    internal static class NPOIExtension
    {
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