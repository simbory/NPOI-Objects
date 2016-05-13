using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace NPOI.Objects
{
    /// <summary>
    /// the ObjectFactory class is used to read the excel file/stream and convert the value in excel to the model
    /// </summary>
    public class ObjectFactory : IDisposable
    {
        /// <summary>
        /// the excel workbook
        /// </summary>
        protected IWorkbook Workbook;

        /// <summary>
        /// the excel stream
        /// </summary>
        protected readonly Stream ExcelStream;

        /// <summary>
        /// indicate that the ObjectFactory need auto-close the stream or not after reading the excel file/stream 
        /// </summary>
        protected bool NeedClose;

        /// <summary>
        /// the file path of the excel
        /// </summary>
        public string ExcelPath { get; protected set; }

        /// <summary>
        /// the excel type
        /// </summary>
        public ExcelType ExcelType { get; protected set; }

        /// <summary>
        /// rich text converter
        /// </summary>
        public Func<ICell, string> RichTextConverter { get; set; }

        /// <summary>
        /// boolean converter
        /// </summary>
        public Func<ICell, bool> BooleanConverter { get; set; }

        /// <summary>
        /// number converter
        /// </summary>
        public Func<ICell, double> NumericConverter { get; set; }

        /// <summary>
        /// DateTime convertor
        /// </summary>
        public Func<ICell, DateTime> DateTimeConverter { get; set; }

        /// <summary>
        /// byte converter
        /// </summary>
        public Func<ICell, byte> ByteConverter { get; set; }

        /// <summary>
        /// char converter
        /// </summary>
        public Func<ICell, char> CharConverter { get; set; }

        /// <summary>
        /// Guid Converter
        /// </summary>
        public Func<ICell, Guid> GuidConverter { get; set; }

        /// <summary>
        /// Unknown type converter
        /// </summary>
        public Func<ICell, Type, object> UnknownTypeConverter { get; set; }

        /// <summary>
        /// constructor
        /// </summary>
        /// <param name="path">the excel path</param>
        public ObjectFactory(string path)
        {
            ExcelPath = path;
            var ext = Path.GetExtension(path);
            if (string.IsNullOrEmpty(ext) || (ext.ToLower() != ".xls" && ext.ToLower() != ".xlsx"))
                throw new FileLoadException("File extension is invalid", path);
            ExcelType = ext == ".xls" ? ExcelType.Excel2003 : ExcelType.Excel2007;
            ExcelStream = new FileStream(path, FileMode.Open, FileAccess.Read);
            NeedClose = true;
            Init();
        }

        /// <summary>
        /// constuctor
        /// </summary>
        /// <param name="stream">the excel stream</param>
        /// <param name="excelType">the excel type</param>
        public ObjectFactory(Stream stream, ExcelType excelType)
        {
            ExcelType = excelType;
            if (stream == null)
                throw new ArgumentNullException("stream");
            if (!stream.CanRead)
                throw new IOException("The file stream is not readable.");
            ExcelStream = stream;
            NeedClose = false;
            Init();
        }

        private void Init()
        {
            Workbook = ExcelType == ExcelType.Excel2003
                    ? (IWorkbook)new HSSFWorkbook(ExcelStream)
                    : new XSSFWorkbook(ExcelStream);
            RichTextConverter = CellValueConverters.RichTextConverter;
            BooleanConverter = CellValueConverters.BooleanConverter;
            NumericConverter = CellValueConverters.NumericConverter;
            DateTimeConverter = CellValueConverters.DateTimeConverter;
            ByteConverter = CellValueConverters.ByteConverter;
            CharConverter = CellValueConverters.CharConverter;
            GuidConverter = CellValueConverters.GuidConverter;
            UnknownTypeConverter = CellValueConverters.UnknownTypeConverter;
        }
        
        /// <summary>
        /// convert the excel worksheet to model array
        /// </summary>
        /// <typeparam name="T">the type of the model</typeparam>
        /// <param name="sheetIndex">the sheeet index</param>
        /// <returns>the model array</returns>
        public T[] SheetToObjects<T>(int sheetIndex = 0) where T : class
        {
            AssertType(typeof(T));
            if (Workbook == null)
                return new T[0];
            var sheet = Workbook.GetSheetAt(sheetIndex);
            return ConvertSheetToObjects<T>(sheet);
        }

        /// <summary>
        /// convert the excel worksheet to model array
        /// </summary>
        /// <typeparam name="T">the type of the model</typeparam>
        /// <param name="sheetName">the sheet name</param>
        /// <returns>the model array</returns>
        public T[] SheetToObjects<T>(string sheetName) where T : class
        {
            AssertType(typeof(T));
            if (Workbook == null)
                return new T[0];
            var sheet = Workbook.GetSheet(sheetName);
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
                    var value = GetCellValue(prop, cell);
                    if (value != null)
                    {
                        prop.SetValue(obj, value, null);
                    }
                }
            }
            return obj;
        }

        private object GetCellValue(PropertyInfo prop, ICell cell)
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
            return UnknownTypeConverter(cell, propType);
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

        /// <summary>
        /// Dispose
        /// </summary>
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