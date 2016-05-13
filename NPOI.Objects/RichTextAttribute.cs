using System;

namespace NPOI.Objects
{
    /// <summary>
    /// indicate that this property is the rich text string
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    [Serializable]
    public class RichTextAttribute : Attribute
    {
    }
}
