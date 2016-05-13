using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NPOI.Objects
{
    /// <summary>
    /// DrawingFactory is used to create an excel file/stream and write the value of the model to the excel file/stream
    /// </summary>
    public class DrawingFactory : IDisposable
    {
        /// <summary>
        /// Excel workbook object
        /// </summary>
        protected IWorkbook Workbook;

        /// <summary>
        /// excel output stream
        /// </summary>
        protected readonly Stream ExcelStream;

        /// <summary>
        /// is output stream or not
        /// </summary>
        protected readonly bool IsOutStream;

        /// <summary>
        /// use excel template or not
        /// </summary>
        protected bool UseTemplate;
        
        /// <summary>
        /// the file path of the excel file
        /// </summary>
        public string ExcelPath { get; protected set; }

        /// <summary>
        /// the type of the workbook, Excel2003(.xls) or Excel2007(.xlsx)
        /// </summary>
        public ExcelType WorkbookType { get; protected set; }

        /// <summary>
        /// the constructor
        /// </summary>
        /// <param name="path">the file path of the excel. the file extension must be .xls or .xlsx</param>
        public DrawingFactory(string path)
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentNullException("path");
            ExcelPath = path;
            var ext = Path.GetExtension(path);
            if (string.IsNullOrEmpty(ext))
            {
                throw new FileLoadException("File extension cannot be empty", path);
            }
            if (ext.Equals(".xls", StringComparison.InvariantCultureIgnoreCase))
            {
                var dir = Path.GetDirectoryName(ExcelPath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);
                ExcelStream = new FileStream(path, File.Exists(ExcelPath) ? FileMode.Open : FileMode.CreateNew, FileAccess.Write);
                Workbook = new HSSFWorkbook();
                WorkbookType = ExcelType.Excel2003;
                IsOutStream = false;
            }
            else if (ext.Equals(".xlsx", StringComparison.InvariantCultureIgnoreCase))
            {
                var dir = Path.GetDirectoryName(ExcelPath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);
                ExcelStream = new FileStream(path, File.Exists(ExcelPath) ? FileMode.Open : FileMode.CreateNew, FileAccess.Write);
                Workbook = new XSSFWorkbook();
                WorkbookType = ExcelType.Excel2007;
                IsOutStream = false;
            }
        }

        /// <summary>
        /// the constructor
        /// </summary>
        /// <param name="stream">the excel output stream</param>
        /// <param name="workbookType">the excel type</param>
        public DrawingFactory(Stream stream, ExcelType workbookType)
        {
            if (stream == null)
                throw new ArgumentNullException("stream");
            if (!stream.CanWrite)
                throw new IOException("The stream is no writable");
            ExcelStream = stream;
            WorkbookType = workbookType;
            if (WorkbookType == ExcelType.Excel2003)
            {
                Workbook = new HSSFWorkbook();
            }
            else
            {
                Workbook = new XSSFWorkbook();
            }
            IsOutStream = true;
        }

        /// <summary>
        /// the constructor
        /// </summary>
        /// <param name="stream">the excel output stream</param>
        public DrawingFactory(Stream stream): this(stream, ExcelType.Excel2003)
        {
        }
        
        /// <summary>
        /// convert the StyleAttribute to the CellStyle object
        /// </summary>
        /// <param name="attr">the object of the StyleAttribute</param>
        /// <returns>the CellStyle</returns>
        protected virtual ICellStyle FillCellStyle(StyleAttribute attr)
        {
            if (attr == null)
            {
                return null;
            }
            ICellStyle style;
            if (WorkbookType == ExcelType.Excel2003)
            {
                style = FillCellStyle2003(attr);
            }
            else
            {
                style = FillCellStyle2007(attr);
            }
            style.Alignment = attr.TextAlign;
            style.VerticalAlignment = attr.VerticalAlign;
            var font = FillFont(attr);
            if (font != null)
            {
                style.SetFont(font);
            }
            return style;
        }

        private HSSFCellStyle FillCellStyle2003(StyleAttribute attr)
        {
            var style = (HSSFCellStyle) ((HSSFWorkbook)Workbook).CreateCellStyle();
            if (!string.IsNullOrEmpty(attr.BackgroundColor))
            {
                var color = attr.BackgroundColor.ToColor();
                if (color.HasValue)
                {
                    style.FillBackgroundColor = GetColor2003(color.Value);
                    style.FillPattern = attr.FillPattern;
                }
            }
            if (!string.IsNullOrEmpty(attr.ForegroundColor))
            {
                var color = attr.ForegroundColor.ToColor();
                if (color.HasValue)
                {
                    style.FillForegroundColor = GetColor2003(color.Value);
                    style.FillPattern = attr.FillPattern;
                }
            }
            return style;
        }

        private XSSFCellStyle FillCellStyle2007(StyleAttribute attr)
        {
            var style = (XSSFCellStyle) ((XSSFWorkbook) Workbook).CreateCellStyle();
            if (!string.IsNullOrEmpty(attr.BackgroundColor))
            {
                var color = attr.BackgroundColor.ToColor();
                if (color.HasValue)
                {
                    style.SetFillBackgroundColor(GetColor2007(color.Value));
                    style.FillPattern = attr.FillPattern;
                }
            }
            if (!string.IsNullOrEmpty(attr.ForegroundColor))
            {
                var color = attr.ForegroundColor.ToColor();
                if (color.HasValue)
                {
                    style.SetFillForegroundColor(GetColor2007(color.Value));
                    style.FillPattern = attr.FillPattern;
                }
            }
            return style;
        }

        /// <summary>
        /// convert StyleAttribute to the excel font 
        /// </summary>
        /// <param name="attr">the value of the StyleAttribute</param>
        /// <returns>the font</returns>
        protected virtual IFont FillFont(StyleAttribute attr)
        {
            if (attr == null)
            {
                return null;
            }
            if (attr.FontWeight > 0
                || !string.IsNullOrEmpty(attr.FontFamily)
                || attr.FontSize > 0
                || attr.IsItalic
                || !string.IsNullOrEmpty(attr.TextColor))
            {
                IFont font;
                if (WorkbookType == ExcelType.Excel2003)
                {
                    font = FillFont2003(attr);
                }
                else
                {
                    font = FillFont2007(attr);
                }
                if (attr.FontWeight > 0)
                    font.Boldweight = attr.FontWeight;
                if (!string.IsNullOrEmpty(attr.FontFamily))
                    font.FontName = attr.FontFamily;
                if (attr.FontSize > 0)
                    font.FontHeightInPoints = attr.FontSize;
                if (attr.IsItalic)
                    font.IsItalic = true;
                return font;
            }
            return null;
        }

        private HSSFFont FillFont2003(StyleAttribute attr)
        {
            var font = (HSSFFont)((HSSFWorkbook) Workbook).CreateFont();
            if (!string.IsNullOrEmpty(attr.TextColor))
            {
                var color = attr.TextColor.ToColor();
                if (color.HasValue)
                {
                    font.Color = GetColor2003(color.Value);
                }
            }
            return font;
        }

        private XSSFFont FillFont2007(StyleAttribute attr)
        {
            var font = (XSSFFont)((XSSFWorkbook)Workbook).CreateFont();
            if (!string.IsNullOrEmpty(attr.TextColor))
            {
                var color = attr.TextColor.ToColor();
                if (color.HasValue)
                {
                    font.SetColor(GetColor2007(color.Value));
                }
            }
            return font;
        }

        private short GetColor2003(Color color)
        {
            var hssfWorkbook = Workbook as HSSFWorkbook;
            if (hssfWorkbook != null)
            {
                var palette = hssfWorkbook.GetCustomPalette();
                var workbookColor = palette.FindColor(color.R, color.G, color.B);
                if (workbookColor != null)
                    return workbookColor.Indexed;
                try
                {
                    workbookColor = palette.AddColor(color.R, color.G, color.B);
                    return workbookColor.Indexed;
                }
                catch (Exception)
                {
                    return palette.FindSimilarColor(color.R, color.G, color.B).Indexed;
                }
            }
            return HSSFColor.COLOR_NORMAL;
        }

        private XSSFColor GetColor2007(Color color)
        {
            return new XSSFColor(color);
        }

        /// <summary>
        /// get the ColumnDrawing info from the model type
        /// </summary>
        /// <param name="classType">the type of the model class</param>
        /// <returns>the ColumnDrawing array</returns>
        protected virtual ColumnDrawing[] GetColumnDrawings(Type classType)
        {
            var cellList = new List<ColumnDrawing>();
            var classProperties = classType.GetProperties(BindingFlags.Instance | BindingFlags.Public | BindingFlags.GetProperty);
            if (classProperties.Length <= 0)
                return new ColumnDrawing[0];
            var columnIndexes = new List<int>();
            foreach (var property in classProperties)
            {
                var ignoreAttr = property.GetCustomAttribute<DrawingIgnoreAttribute>();
                var propAttr = property.GetCustomAttribute<NPOIColumnAttribute>();
                if (ignoreAttr != null || propAttr == null)
                    continue;
                var cellInfo = new ColumnDrawing();
                // Column basic information
                if (propAttr.Index < 0)
                {
                    throw new Exception("Column Index is out of range.\r\nThe value must not be smaller than 0.");
                }
                if (columnIndexes.Contains(propAttr.Index))
                    throw new Exception("Duplicate column index " + propAttr.Index);
                cellInfo.ColumnIndex = propAttr.Index;
                columnIndexes.Add(propAttr.Index);

                if (!string.IsNullOrEmpty(propAttr.Name))
                    cellInfo.ColumnName = propAttr.Name;
                if (string.IsNullOrEmpty(cellInfo.ColumnName))
                    cellInfo.ColumnName = property.Name;
                // Column style information
                var headerStyleAttr = property.GetCustomAttribute<HeaderStyleAttribute>();
                var headerStyle = FillCellStyle(headerStyleAttr);
                if (headerStyle != null)
                    cellInfo.HeaderStyle = headerStyle;
                if (headerStyleAttr != null)
                    cellInfo.ColumnWidth = headerStyleAttr.ColumnWidth;

                var cellStyleAttr = property.GetCustomAttribute<CellStyleAttribute>();
                if (cellStyleAttr != null)
                {
                    var cellStyle = FillCellStyle(cellStyleAttr);
                    if (cellStyle != null)
                        cellInfo.CellStyle = cellStyle;
                    var cellFont = FillFont(cellStyleAttr);
                    if (cellFont != null)
                        cellInfo.CellFont = cellFont;
                }
                var alternateCellStyleAttr = property.GetCustomAttribute<AlternateCellStyleAttribute>();
                if (alternateCellStyleAttr != null)
                {
                    cellInfo.HasAlternate = true;
                    var cellStyle = FillCellStyle(alternateCellStyleAttr);
                    if (cellStyle != null)
                        cellInfo.AlternateCellStyle = cellStyle;
                    var cellFont = FillFont(alternateCellStyleAttr);
                    if (cellFont != null)
                        cellInfo.AlternateCellFont = cellFont;
                }
                else
                {
                    cellInfo.HasAlternate = false;
                }
                cellInfo.Property = property;
                cellList.Add(cellInfo);
            }

            return cellList.OrderBy(x => x.ColumnIndex).ToArray();
        }

        /// <summary>
        /// draw the header of the excel
        /// </summary>
        /// <param name="drawings">the ColumnDrawing array</param>
        /// <param name="sheet">the excel sheet</param>
        /// <param name="headerRowIndex">the header row index</param>
        protected virtual void DrawHeader(IEnumerable<ColumnDrawing> drawings, ISheet sheet, int headerRowIndex)
        {
            var headerRow = sheet.CreateRow(headerRowIndex);
            foreach (var drawing in drawings)
            {
                var headerCell = headerRow.CreateCell(drawing.ColumnIndex);
                DrawHeaderFontAndStyle(headerCell, drawing);
                if (drawing.ColumnWidth > 255)
                {
                    drawing.ColumnWidth = 255;
                }
                sheet.SetColumnWidth(drawing.ColumnIndex, drawing.ColumnWidth * 256);
            }
        }

        /// <summary>
        /// draw the cell value of the excel
        /// </summary>
        /// <param name="cell">the excel cell</param>
        /// <param name="value">the excel cell value</param>
        protected virtual void DrawCellValue(ICell cell, object value)
        {
            if (value == null)
            {
                cell.SetCellType(CellType.Blank);
                return;
            }
            var strValue = value as string;
            if (strValue != null)
            {
                cell.SetCellType(CellType.String);
                cell.SetCellValue(strValue);
                return;
            }
            var charValue = value as char[];
            if (charValue != null)
            {
                cell.SetCellType(CellType.String);
                cell.SetCellValue(new string(charValue));
                return;
            }
            if (value is bool)
            {
                cell.SetCellType(CellType.Boolean);
                cell.SetCellValue((bool) value);
                return;
            }
            if (value is int
                || value is uint
                || value is long
                || value is ulong
                || value is short
                || value is ushort
                || value is float
                || value is double)
            {
                cell.SetCellType(CellType.Numeric);
                cell.SetCellValue(Convert.ToDouble(value));
                return;
            }
            if (value is byte)
            {
                cell.SetCellType(CellType.Error);
                cell.SetCellValue(Convert.ToDouble(value));
                return;
            }
            if (value is DateTime)
            {
                cell.SetCellType(CellType.Formula);
                cell.SetCellValue((DateTime) value);
                return;
            }
            cell.SetCellType(CellType.String);
            cell.SetCellValue(value.ToString());
        }

        /// <summary>
        /// draw the cell font and style
        /// </summary>
        /// <param name="cell">the excel cell</param>
        /// <param name="drawing">the ColumnDrawing object</param>
        /// <param name="alternate">is alternate excel cell</param>
        protected virtual void DrawCellFontAndStyle(ICell cell, ColumnDrawing drawing, bool alternate)
        {
            if (UseTemplate)
                return;
            ICellStyle style;
            IFont font;
            if (alternate)
            {
                style = drawing.AlternateCellStyle;
                font = drawing.AlternateCellFont;
            }
            else
            {
                style = drawing.CellStyle;
                font = drawing.CellFont;
            }
            if (style == null)
                style = cell.Sheet.Workbook.GetCellStyleAt(0);
            cell.CellStyle = style;
            if (font != null)
                cell.CellStyle.SetFont(font);
        }

        /// <summary>
        /// draw the header cell font and style
        /// </summary>
        /// <param name="cell">the excel cell</param>
        /// <param name="drawing">the ColumnDrawing object</param>
        protected virtual void DrawHeaderFontAndStyle(ICell cell, ColumnDrawing drawing)
        {
            cell.SetCellType(CellType.String);
            cell.SetCellValue(drawing.ColumnName);
            if (UseTemplate)
                return;
            cell.CellStyle = drawing.HeaderStyle;
        }

        /// <summary>
        /// draw the whole row
        /// </summary>
        /// <param name="drawings">the array of the ColumnDrawing</param>
        /// <param name="sheet">the excel sheet</param>
        /// <param name="rowIndex">the row index</param>
        /// <param name="obj">the value object of the whole row</param>
        protected virtual void DrawRow(IEnumerable<ColumnDrawing> drawings, ISheet sheet, int rowIndex, object obj)
        {
            var row = sheet.CreateRow(rowIndex);
            foreach (var drawing in drawings)
            {
                var cell = row.CreateCell(drawing.ColumnIndex);
                var value = drawing.Property.GetValue(obj, null);
                DrawCellValue(cell, value);
                DrawCellFontAndStyle(cell, drawing, rowIndex%2 == 1 && drawing.HasAlternate);
            }
        }

        /// <summary>
        /// write objects to worksheet
        /// </summary>
        /// <typeparam name="T">any type of model</typeparam>
        /// <param name="sheetIndex">the worksheet index</param>
        /// <param name="sheetName">the worksheet name</param>
        /// <param name="objects">the arry of object</param>
        public void Draw<T>(int sheetIndex, string sheetName, params T[] objects)
        {
            if (objects == null)
                throw new ArgumentNullException("objects");
            var type = typeof (T);
            var objAttr = type.GetCustomAttribute<NPOIObjectAttribute>() ?? new NPOIObjectAttribute();
            var drawings = GetColumnDrawings(type);
            var sheet = Workbook.GetSheet(sheetName);
            if (sheet == null)
            {
                sheet = Workbook.CreateSheet(sheetName);
                DrawHeader(drawings, sheet, objAttr.HeaderRowIndex);
            }
            for (int i = 0; i < objects.Length; i++)
            {
                var obj = objects[i];
                DrawRow(drawings, sheet, objAttr.StartIndex + i, obj);
            }
        }

        /// <summary>
        /// Dispose
        /// </summary>
        public void Dispose()
        {
            if (Workbook != null)
            {
                Workbook.Write(ExcelStream);
            }
            if (!IsOutStream && ExcelStream != null)
                ExcelStream.Close();
        }

        /// <summary>
        /// set the excel template to the worksheet
        /// </summary>
        /// <param name="path">the path of the excel template</param>
        public void SetTemplate(string path)
        {
            if (!File.Exists(path))
                throw new FileNotFoundException("The excel template file cannot be found.");
            var ext = Path.GetExtension(path);
            if (string.IsNullOrEmpty(ext) || (!ext.ToLower().Equals(".xls", StringComparison.InvariantCultureIgnoreCase) && !ext.Equals(".xlsx", StringComparison.InvariantCultureIgnoreCase)))
                throw new FileLoadException("Invalid file extension. The file extension must be .xls or .xlsx", path);
            Workbook = new HSSFWorkbook(new FileStream(path, FileMode.Open, FileAccess.ReadWrite));
            UseTemplate = true;
        }
    }
}