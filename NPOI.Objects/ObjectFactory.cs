using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NPOI.Objects
{
    public sealed class ObjectFactory : IDisposable
    {
        private readonly IWorkbook _workbook;

        private readonly Stream _excelStream;

        private readonly bool _needClose;

        public string ExcelPath { get; private set; }

        public ExcelType ExcelType { get; private set; }

        public ObjectFactory(string path)
        {
            ExcelPath = path;
            var ext = Path.GetExtension(path);
            if (string.IsNullOrEmpty(ext) || (ext.ToLower() != ".xls" && ext.ToLower() != ".xlsx"))
                throw new FileLoadException("File extension is invalid", path);
            ExcelType = ext == ".xls" ? ExcelType.Excel2003 : ExcelType.Excel2007;
            _excelStream = new FileStream(path, FileMode.Open, FileAccess.Read);
            _workbook = ExcelType == ExcelType.Excel2003
                    ? (IWorkbook)new HSSFWorkbook(_excelStream)
                    : new XSSFWorkbook(_excelStream);
            _needClose = true;
        }

        public ObjectFactory(Stream stream, ExcelType excelType)
        {
            ExcelType = excelType;
            if (stream == null)
                throw new ArgumentNullException("stream");
            if (!stream.CanRead)
                throw new IOException("The file stream is not readable.");
            _workbook = ExcelType == ExcelType.Excel2003
                ? (IWorkbook)new HSSFWorkbook(stream)
                : new XSSFWorkbook(stream);
            _excelStream = stream;
            _needClose = false;
        }



        public T[] SheetToObjects<T>(int sheetIndex = 0) where T : class
        {
            AssertType(typeof(T));
            if (_workbook == null)
                return new T[0];
            var sheet = _workbook.GetSheetAt(sheetIndex);
            return ConvertSheetToObjects<T>(sheet);
        }

        public T[] SheetToObjects<T>(string sheetName) where T : class
        {
            AssertType(typeof(T));
            if (_workbook == null)
                return new T[0];
            var sheet = _workbook.GetSheet(sheetName);
            return ConvertSheetToObjects<T>(sheet);
        }

        private Dictionary<PropertyInfo, int> GetClassProperties(Type classType, ISheet sheet)
        {
            if (sheet == null)
                return null;
            var props = new Dictionary<PropertyInfo, int>();
            var attr = AssertType(classType);
            var headerRowDic = new Dictionary<string, int>();
            var headerRow = sheet.GetRow(attr.HeaderRowIndex);
            var headerLastColumn = headerRow.LastCellNum;
            for (int i = 0; i < headerLastColumn; i++)
            {
                var cell = headerRow.GetCell(i);
                if (cell == null)
                    continue;
                var cellValue = cell.StringCellValue;
                if (string.IsNullOrEmpty(cellValue))
                    continue;
                cellValue = cellValue.Trim().Replace("\t", " ").Replace("\r", "").Replace("\n", " ").ToLower();
                if (string.IsNullOrEmpty(cellValue))
                    continue;
                if (headerRowDic.ContainsKey(cellValue))
                    throw new DuplicateColumnException(cellValue, headerRowDic[cellValue], i, attr.HeaderRowIndex);
                headerRowDic.Add(cellValue, i);
            }
            var classProperties = classType.GetProperties(BindingFlags.Instance | BindingFlags.Public | BindingFlags.SetProperty);
            if (classProperties.Length > 0)
            {
                foreach (var property in classProperties)
                {
                    var propAttr = property.GetCustomAttribute<NPOIColumnAttribute>();
                    if (propAttr == null)
                        continue;
                    if (propAttr.Index >= 0 && headerRowDic.ContainsValue(propAttr.Index))
                    {
                        props.Add(property, propAttr.Index);
                        continue;
                    }
                    if (!string.IsNullOrEmpty(propAttr.Name) && headerRowDic.ContainsKey(propAttr.Name.ToLower()))
                    {
                        props.Add(property, headerRowDic[propAttr.Name.ToLower()]);
                        continue;
                    }
                    if (string.IsNullOrEmpty(propAttr.Name) && propAttr.Index < 0)
                    {
                        if (headerRowDic.ContainsKey(property.Name.ToLower()))
                        {
                            props.Add(property, headerRowDic[property.Name.ToLower()]);
                            continue;
                        }
                    }
                    throw new InvalidColumnException(string.IsNullOrEmpty(propAttr.Name) ? property.Name : propAttr.Name);
                }
            }
            return props;
        }

        private T[] ConvertSheetToObjects<T>(ISheet sheet) where T : class
        {
            if (sheet == null)
                return new T[0];
            var objectList = new List<T>();
            var objInfo = AssertType(typeof(T));
            var properties = GetClassProperties(typeof(T), sheet);
            var end = objInfo.EndIndex;
            if (end < 1 || end > sheet.LastRowNum)
            {
                end = sheet.LastRowNum;
            }
            for (int i = objInfo.StartIndex; i <= end; i++)
            {
                objectList.Add(RowToObject<T>(properties, sheet.GetRow(i)));
            }
            return objectList.ToArray();
        }

        private T RowToObject<T>(Dictionary<PropertyInfo, int> props, IRow row)
        {
            var obj = Activator.CreateInstance<T>();
            if (row != null)
            {
                foreach (var propPair in props)
                {
                    var prop = propPair.Key;
                    var cell = row.GetCell(propPair.Value);
                    if (cell == null)
                        continue;
                    var value = CellToObject(prop, cell);
                    if (value != null)
                    {
                        prop.SetValue(obj, value, null);
                    }
                }
            }
            return obj;
        }

        private object CellToObject(PropertyInfo prop, ICell cell)
        {
            var propType = prop.PropertyType;
            object value = null;
            if (propType == typeof(int))
            {
                value = (int)cell.NumericCellValue;
            }
            else if (propType == typeof(uint))
            {
                value = (uint)cell.NumericCellValue;
            }
            else if (propType == typeof(long))
            {
                value = (long)cell.NumericCellValue;
            }
            else if (propType == typeof(ulong))
            {
                value = (ulong)cell.NumericCellValue;
            }
            else if (propType == typeof(short))
            {
                value = (short)cell.NumericCellValue;
            }
            else if (propType == typeof(ushort))
            {
                value = (ushort)cell.NumericCellValue;
            }
            else if (propType == typeof(float))
            {
                value = (float)cell.NumericCellValue;
            }
            else if (propType == typeof(double))
            {
                value = cell.NumericCellValue;
            }
            else if (propType == typeof(bool))
            {
                value = cell.BooleanCellValue;
            }
            else if (propType == typeof(DateTime))
            {
                value = cell.DateCellValue;
            }
            else if (propType == typeof(byte))
            {
                value = cell.ErrorCellValue;
            }
            else if (propType == typeof(char))
            {
                var strValue = cell.StringCellValue;
                if (!string.IsNullOrEmpty(strValue))
                {
                    value = strValue[0];
                }
            }
            else if (propType == typeof(char[]))
            {
                var strValue = cell.StringCellValue;
                if (!string.IsNullOrEmpty(strValue))
                {
                    value = strValue.ToArray();
                }
            }
            else if (propType == typeof(Guid))
            {
                var strValue = cell.StringCellValue;
                if (!string.IsNullOrEmpty(strValue))
                {
                    Guid guid;
                    value = Guid.TryParse(strValue, out guid) ? guid : Guid.Empty;
                }
            }
            else if (propType == typeof(Uri))
            {
                try
                {
                    var strValue = cell.StringCellValue;
                    if (!string.IsNullOrEmpty(strValue))
                    {
                        value = new Uri(strValue);
                    }
                }
                catch (Exception)
                {
                    return null;
                }
            }
            else if (propType == typeof(string))
            {
                try
                {
                    var richTextAttrs = prop.GetCustomAttributes(typeof(RichTextAttribute), false);
                    value = richTextAttrs.Length < 1 ? cell.StringCellValue : cell.ToHtml();
                }
                catch (Exception)
                {
                    return null;
                }
            }
            return value;
        }

        private NPOIObjectAttribute AssertType(Type type)
        {
            var attr = type.GetCustomAttribute<NPOIObjectAttribute>();
            if (attr == null)
            {
                throw new CustomAttributeFormatException("Invalid class type.\r\nThe class " + type + " must has " + typeof(NPOIObjectAttribute) + " attribute");
            }
            return attr;
        }

        public void Dispose()
        {
            if (_excelStream == null || !_needClose)
                return;
            try
            {
                _excelStream.Close();
            }
            catch (Exception)
            {
                // ignored
            }
        }
    }
}