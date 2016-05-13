using System.Reflection;
using NPOI.SS.UserModel;

namespace NPOI.Objects
{
    /// <summary>
    /// the ColumnDrawing class
    /// </summary>
    public class ColumnDrawing
    {
        /// <summary>
        /// the name of the column
        /// </summary>
        public string ColumnName { get; set; }

        /// <summary>
        /// the column inde
        /// </summary>
        public int ColumnIndex { get; set; }

        /// <summary>
        /// the column width
        /// </summary>
        public int ColumnWidth { get; set; }

        /// <summary>
        /// indicate that if the DrawingFactory will set alternate style to the alternate excel row
        /// </summary>
        public bool HasAlternate { get; set; }

        /// <summary>
        /// the Property info of the model
        /// </summary>
        public PropertyInfo Property { get; set; }

        /// <summary>
        /// the style of the cell
        /// </summary>
        public ICellStyle CellStyle { get; set; }

        /// <summary>
        /// the font of the cell
        /// </summary>
        public IFont CellFont { get; set; }

        /// <summary>
        /// the style of the alternate cell
        /// </summary>
        public ICellStyle AlternateCellStyle { get; set; }

        /// <summary>
        /// the font of the alternate cell
        /// </summary>
        public IFont AlternateCellFont { get; set; }

        /// <summary>
        /// the style of the header cell
        /// </summary>
        public ICellStyle HeaderStyle { get; set; }
    }
}