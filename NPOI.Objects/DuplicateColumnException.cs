using System;

namespace NPOI.Objects
{
    /// <summary>
    /// this exception indicate that the model has set the duplicate column index
    /// </summary>
    public class DuplicateColumnException : Exception
    {
        private readonly string _message;
        
        /// <summary>
        /// the messagge text of the exceltion
        /// </summary>
        public override string Message
        {
            get
            {
                return _message;
            }
        }

        /// <summary>
        /// constructor of the exception
        /// </summary>
        /// <param name="columnName">the column name</param>
        /// <param name="oldColumn">the old column index</param>
        /// <param name="newColumn">the new column index</param>
        /// <param name="rowIndex">the row index</param>
        public DuplicateColumnException(string columnName, int oldColumn, int newColumn, int rowIndex)
        {
            _message = string.Format(@"Duplicate column name ""{3}"" at column {0} and column {1} in row {2}.",
                oldColumn + 1, newColumn + 1, rowIndex + 1, columnName);
        }
    }
}