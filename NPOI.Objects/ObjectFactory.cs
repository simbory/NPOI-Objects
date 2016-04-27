using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NPOI.Objects
{
    public class ObjectFactory : IDisposable
    {
        protected readonly IWorkbook Workbook;

        protected readonly Stream ExcelStream;

        protected readonly bool NeedClose;

        public string ExcelPath { get; protected set; }

        public ExcelType ExcelType { get; protected set; }

        public Func<ICell, string> RichTextConverter { get; set; }

        public Func<ICell, bool> BooleanConverter { get; set; }

        public Func<ICell, double> NumericConverter { get; set; }

        public Func<ICell, DateTime> DateTimeConverter { get; set; }

        public Func<ICell, byte> ByteConverter { get; set; }

        public Func<ICell, char> CharConverter { get; set; }

        public Func<ICell, Guid> GuidConverter { get; set; }

        public Func<ICell, Type, object> UnknowTypeConverter { get; set; }

        public ObjectFactory(string path)
        {
            ExcelPath = path;
            var ext = Path.GetExtension(path);
            if (string.IsNullOrEmpty(ext) || (ext.ToLower() != ".xls" && ext.ToLower() != ".xlsx"))
                throw new FileLoadException("File extension is invalid", path);
            ExcelType = ext == ".xls" ? ExcelType.Excel2003 : ExcelType.Excel2007;
            ExcelStream = new FileStream(path, FileMode.Open, FileAccess.Read);
            Workbook = ExcelType == ExcelType.Excel2003
                    ? (IWorkbook)new HSSFWorkbook(ExcelStream)
                    : new XSSFWorkbook(ExcelStream);
            NeedClose = true;
            RichTextConverter = CellValueConverters.RichTextConverter;
            BooleanConverter = CellValueConverters.BooleanConverter;
            NumericConverter = CellValueConverters.NumericConverter;
            DateTimeConverter = CellValueConverters.DateTimeConverter;
            ByteConverter = CellValueConverters.ByteConverter;
            CharConverter = CellValueConverters.CharConverter;
            GuidConverter = CellValueConverters.GuidConverter;
            UnknowTypeConverter = CellValueConverters.UnknowTypeConverter;
        }

        public ObjectFactory(Stream stream, ExcelType excelType)
        {
            ExcelType = excelType;
            if (stream == null)
                throw new ArgumentNullException("stream");
            if (!stream.CanRead)
                throw new IOException("The file stream is not readable.");
            Workbook = ExcelType == ExcelType.Excel2003
                ? (IWorkbook)new HSSFWorkbook(stream)
                : new XSSFWorkbook(stream);
            ExcelStream = stream;
            NeedClose = false;
        }
        
        public T[] SheetToObjects<T>(int sheetIndex = 0) where T : class
        {
            AssertType(typeof(T));
            if (Workbook == null)
                return new T[0];
            var sheet = Workbook.GetSheetAt(sheetIndex);
            return ConvertSheetToObjects<T>(sheet);
        }

        public T[] SheetToObjects<T>(string sheetName) where T : class
        {
            AssertType(typeof(T));
            if (Workbook == null)
                return new T[0];
            var sheet = Workbook.GetSheet(sheetName);
            return ConvertSheetToObjects<T>(sheet);
        }

        protected Dictionary<PropertyInfo, int> GetClassProperties(Type classType, ISheet sheet)
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

        protected T[] ConvertSheetToObjects<T>(ISheet sheet) where T : class
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

        protected T RowToObject<T>(Dictionary<PropertyInfo, int> props, IRow row)
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
                    var value = GetCellValue(prop, cell);
                    if (value != null)
                    {
                        prop.SetValue(obj, value, null);
                    }
                }
            }
            return obj;
        }

        protected virtual object GetCellValue(PropertyInfo prop, ICell cell)
        {
            var propType = prop.PropertyType;

            if (propType == typeof(string))
            {
                var isRichText = prop.GetCustomAttributes(typeof(RichTextAttribute), false).Length > 0;
                if (isRichText)
                {
                    var richText = cell.RichStringCellValue;
                    return richText != null ? RichTextConverter(cell) : null;
                }
                try
                {
                    return cell.StringCellValue;
                }
                catch (Exception)
                {
                    var rich = cell.RichStringCellValue;
                    return rich != null ? rich.String : null;
                }
            }
            if (propType == typeof(bool))
            {
                return BooleanConverter(cell);
            }
            if (propType == typeof(int))
            {
                return (int) NumericConverter(cell);
            }
            if (propType == typeof(uint))
            {
                return (uint)NumericConverter(cell);
            }
            if (propType == typeof(long))
            {
                return (long)NumericConverter(cell);
            }
            if (propType == typeof(ulong))
            {
                return (ulong)NumericConverter(cell);
            }
            if (propType == typeof(short))
            {
                return (short)NumericConverter(cell);
            }
            if (propType == typeof(ushort))
            {
                return (ushort)NumericConverter(cell);
            }
            if (propType == typeof(float))
            {
                return (float)NumericConverter(cell);
            }
            if (propType == typeof(double))
            {
                return NumericConverter(cell);
            }
            if (propType == typeof(DateTime))
            {
                return DateTimeConverter(cell);
            }
            if (propType == typeof(byte))
            {
                return ByteConverter(cell);
            }
            if (propType == typeof(char))
            {
                return CharConverter(cell);
            }
            if (propType == typeof(Guid))
            {
                return GuidConverter(cell);
            }
            return UnknowTypeConverter(cell, propType);
        }

        protected NPOIObjectAttribute AssertType(Type type)
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
            if (ExcelStream == null || !NeedClose)
                return;
            try
            {
                ExcelStream.Close();
            }
            catch (Exception)
            {
                // ignored
            }
        }
    }
}