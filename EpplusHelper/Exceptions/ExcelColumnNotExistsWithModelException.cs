using System;

namespace EpplusExtensions.Exceptions
{
    /// <summary>
    /// excel列不在Model中异常
    /// </summary>
    public class ExcelColumnNotExistsWithModelException : System.Exception
    {
        public ExcelColumnNotExistsWithModelException() : base() { }
        public ExcelColumnNotExistsWithModelException(string message) : base(message) { }
    }
}
