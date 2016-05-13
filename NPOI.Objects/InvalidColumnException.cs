using System;

namespace NPOI.Objects
{
    /// <summary>
    /// indicate the column set is invalid or canot be found
    /// </summary>
    public class InvalidColumnException : Exception
    {
        private readonly string _message;
        /// <summary>
        /// the message of the exception
        /// </summary>
        public override string Message
        {
            get
            {
                return _message;
            }
        }

        /// <summary>
        /// the constructor
        /// </summary>
        /// <param name="columnName">the column name</param>
        public InvalidColumnException(string columnName)
        {
            _message = string.Format(@"Cannot find the field ""{0}"" in the excel table.", columnName);
        }
    }
}
