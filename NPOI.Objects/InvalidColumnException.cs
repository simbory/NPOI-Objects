using System;

namespace NPOI.Objects
{
    public class InvalidColumnException : Exception
    {
        private readonly string _message;
        public override string Message
        {
            get
            {
                return _message;
            }
        }

        public InvalidColumnException(string columnName)
        {
            _message = string.Format(@"Cannot find the field ""{0}"" in the excel table.", columnName);
        }
    }
}
