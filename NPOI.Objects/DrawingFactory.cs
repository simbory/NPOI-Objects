using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using NPOI.HSSF.UserModel;
using NPOI.Objects.Attributes;
using NPOI.SS.UserModel;

namespace NPOI.Objects
{
    public class DrawingFactory : IDisposable
    {
        private HSSFWorkbook _workbook;

        private readonly Stream _excelStream;

        private readonly bool _isOutStream;

        private bool _useTemplate;
        
        public string ExcelPath { get; private set; }

        public DrawingFactory(string path)
        {
            if (string.IsNullOrEmpty(path))
                throw new ArgumentNullException("path");
            ExcelPath = path;
            var ext = Path.GetExtension(path);
            if (string.IsNullOrEmpty(ext) || ext.ToLower() != ".xls")
                throw new FileLoadException("File extension is invalid. The file extension must be .xls", path);
            var dir = Path.GetDirectoryName(ExcelPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            _excelStream = new FileStream(path, File.Exists(ExcelPath)? FileMode.Open : FileMode.CreateNew, FileAccess.Write);
            _workbook = new HSSFWorkbook();
            _isOutStream = false;
        }

        public DrawingFactory(Stream stream)
        {
            if (stream == null)
                throw new ArgumentNullException("stream");
            if (!stream.CanWrite)
                throw new IOException("The stream is no writable");
            _excelStream = stream;
            _workbook = new HSSFWorkbook();
            _isOutStream = true;
        }

        private short GetColor(Color color)
        {
            var workbookColor = _workbook.GetCustomPalette().FindColor(color.R, color.G, color.B);
            if (workbookColor == null)
            {
                try
                {
                    workbookColor = _workbook.GetCustomPalette().AddColor(color.R, color.G, color.B);
                    return workbookColor.GetIndex();
                }
                catch (Exception)
                {
                    return _workbook.GetCustomPalette().FindSimilarColor(color.R, color.G, color.B).GetIndex();
                }
            }
            return workbookColor.GetIndex();
        }

        private ICellStyle FillStyle(StyleAttribute attr)
        {
            if (attr == null)
            {
                return null;
            }
            var style = _workbook.CreateCellStyle();
            style.Alignment = attr.TextAlign;
            style.VerticalAlignment = attr.VerticalAlign;
            if (!string.IsNullOrEmpty(attr.BackgroundColor))
            {
                var color = attr.BackgroundColor.ToColor();
                if (color.HasValue)
                {
                    style.FillForegroundColor = GetColor(color.Value);
                    style.FillPattern = FillPattern.SolidForeground;
                }
            }
            return style;
        }

        private IFont FillFont(StyleAttribute attr)
        {
            if (attr.FontWeight > 0
                || !string.IsNullOrEmpty(attr.FontFamily)
                || attr.FontSize > 0
                || attr.IsItalic
                || !string.IsNullOrEmpty(attr.TextColor))
            {
                var font = _workbook.CreateFont();
                if (attr.FontWeight > 0)
                    font.Boldweight = attr.FontWeight;
                if (!string.IsNullOrEmpty(attr.FontFamily))
                    font.FontName = attr.FontFamily;
                if (attr.FontSize > 0)
                    font.FontHeightInPoints = attr.FontSize;
                if (!string.IsNullOrEmpty(attr.TextColor))
                {
                    var color = attr.TextColor.ToColor();
                    if (color.HasValue)
                    {
                        font.Color = GetColor(color.Value);
                    }
                }
                if (attr.IsItalic)
                    font.IsItalic = true;
                return font;
            }
            return null;
        }

        private ColumnDrawing[] GetColumnDrawings(Type classType)
        {
            var cellList = new List<ColumnDrawing>();
            var classProperties = classType.GetProperties(BindingFlags.Instance | BindingFlags.Public | BindingFlags.GetProperty);
            if (classProperties.Length <= 0)
                return new ColumnDrawing[0];
            foreach (var property in classProperties)
            {
                var ignoreAttr = property.GetCustomAttribute<DrawingIgnoreAttribute>();
                var propAttr = property.GetCustomAttribute<NPOIColumnAttribute>();
                if (ignoreAttr != null || propAttr == null)
                    continue;
                var cellInfo = new ColumnDrawing();
                // Column basic information
                var columnIndexes = new List<int>();
                if (propAttr.Index >= 0)
                {
                    if (columnIndexes.Contains(propAttr.Index))
                        throw new Exception("Duplicate column index " + propAttr.Index);
                    cellInfo.ColumnIndex = propAttr.Index;
                    columnIndexes.Add(propAttr.Index);
                }
                else
                {
                    throw new Exception("Column Index is out of range.\r\nThe value must not be smaller than 0.");
                }
                if (!string.IsNullOrEmpty(propAttr.Name))
                    cellInfo.ColumnName = propAttr.Name;
                if (string.IsNullOrEmpty(cellInfo.ColumnName))
                    cellInfo.ColumnName = property.Name;
                // Column style information
                var headerStyleAttr = property.GetCustomAttribute<HeaderStyleAttribute>();
                var headerStyle = FillStyle(headerStyleAttr);
                if (headerStyle != null)
                    cellInfo.HeaderStyle = headerStyle;
                var headerFont = FillFont(headerStyleAttr);
                if (headerFont != null)
                    cellInfo.HeaderFont = headerFont;
                cellInfo.ColumnWidth = headerStyleAttr.ColumnWidth;

                var cellStyleAttr = property.GetCustomAttribute<CellStyleAttribute>();
                if (cellStyleAttr != null)
                {
                    var cellStyle = FillStyle(cellStyleAttr);
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
                    var cellStyle = FillStyle(alternateCellStyleAttr);
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

        private void DrawHeader(IEnumerable<ColumnDrawing> drawings, ISheet sheet, int headerRowIndex)
        {
            var headerRow = sheet.CreateRow(headerRowIndex);
            foreach (var drawing in drawings)
            {
                var headerCell = headerRow.CreateCell(drawing.ColumnIndex);
                headerCell.SetCellType(CellType.String);
                headerCell.SetCellValue(drawing.ColumnName);
                DrawHeaderFontAndStyle(headerCell, drawing);
                if (drawing.ColumnWidth > 255)
                    throw new ArgumentException("The maximum column width for an individual cell is 255 characters.");
                sheet.SetColumnWidth(drawing.ColumnIndex, drawing.ColumnWidth * 256);
            }
        }

        private void DrawCellValue(ICell cell, object value)
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

        private void DrawCellFontAndStyle(ICell cell, ColumnDrawing drawing, bool alternate)
        {
            if (_useTemplate)
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

        private void DrawHeaderFontAndStyle(ICell cell, ColumnDrawing drawing)
        {
            if (_useTemplate)
                return;
            cell.CellStyle = drawing.HeaderStyle ?? cell.Sheet.Workbook.GetCellStyleAt(0);
            if (drawing.HeaderFont != null)
                cell.CellStyle.SetFont(drawing.HeaderFont);
        }

        private void DrawRow(IEnumerable<ColumnDrawing> drawings, ISheet sheet, int rowIndex, object obj)
        {
            if (rowIndex == 3)
            {
                Console.Write(rowIndex);
            }
            var row = sheet.CreateRow(rowIndex);
            foreach (var drawing in drawings)
            {
                var cell = row.CreateCell(drawing.ColumnIndex);
                var value = drawing.Property.GetValue(obj, null);
                DrawCellValue(cell, value);
                DrawCellFontAndStyle(cell, drawing, rowIndex%2 == 1 && drawing.HasAlternate);
            }
        }

        public void Draw<T>(int sheetIndex, string sheetName, params T[] objects)
        {
            if (objects == null)
                throw new ArgumentNullException("objects");
            var type = typeof (T);
            var objAttr = type.GetCustomAttribute<NPOIObjectAttribute>() ?? new NPOIObjectAttribute();
            var sheet = _workbook.GetSheet(sheetName);
            if (sheet == null)
            {
                sheet = _workbook.CreateSheet(sheetName);
            }
            else
            {
                _workbook.Remove(sheet);
                sheet = _workbook.CreateSheet(sheetName);
            }
            var drawings = GetColumnDrawings(type);
            DrawHeader(drawings, sheet, objAttr.HeaderRowIndex);
            for (int i = 0; i < objects.Length; i++)
            {
                var obj = objects[i];
                DrawRow(drawings, sheet, objAttr.StartIndex + i, obj);
            }
        }

        public void Dispose()
        {
            if (_workbook != null)
            {
                _workbook.Write(_excelStream);
            }
            if (!_isOutStream && _excelStream != null)
                _excelStream.Close();
        }

        public void SetTemplate(string path)
        {
            if (!File.Exists(path))
                throw new FileNotFoundException("The excel template file cannot be found.");
            var ext = Path.GetExtension(path);
            if (string.IsNullOrEmpty(ext) || ext.ToLower() != ".xls")
                throw new FileLoadException("Invalid file extension. The file extension must be .xls", path);
            _workbook = new HSSFWorkbook(new FileStream(path, FileMode.Open, FileAccess.ReadWrite));
            _useTemplate = true;
        }
    }
}