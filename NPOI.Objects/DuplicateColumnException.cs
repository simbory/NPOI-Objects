using System;

namespace NPOI.Objects
{
    public class DuplicateColumnException : Exception
    {
        private readonly string _message;
        public override string Message
        {
            get
            {
                return _message;
            }
        }

        public DuplicateColumnException(string columnName, int oldColumn, int newColumn, int rowIndex)
        {
            _message = string.Format(@"Duplicate column name ""{3}"" at column {0} and column {1} in row {2}.",
                oldColumn + 1, newColumn + 1, rowIndex + 1, columnName);
        }
    }
}