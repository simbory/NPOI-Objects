using System.Reflection;
using NPOI.SS.UserModel;

namespace NPOI.Objects
{
    public class ColumnDrawing
    {
        public string ColumnName { get; set; }

        public int ColumnIndex { get; set; }

        public int ColumnWidth { get; set; }

        public bool HasAlternate { get; set; }

        public PropertyInfo Property { get; set; }

        public ICellStyle CellStyle { get; set; }
        public IFont CellFont { get; set; }

        public ICellStyle AlternateCellStyle { get; set; }
        public IFont AlternateCellFont { get; set; }

        public ICellStyle HeaderStyle { get; set; }
    }
}
