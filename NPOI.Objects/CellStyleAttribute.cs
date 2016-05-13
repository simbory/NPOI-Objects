using System;

namespace NPOI.Objects
{
    /// <summary>
    /// CellStyleAttribute attribute class is used to set the cell style when convert the object to excel file
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    [Serializable]
    public class CellStyleAttribute : StyleAttribute
    {
    }
}