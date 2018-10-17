using System;

namespace EpplusExtensions
{
    /// <summary>
    /// excel列不在Model中异常
    /// </summary>
    public class ExcelColumnNotExistsWithModelException : Exception
    {
        public ExcelColumnNotExistsWithModelException() : base() { }
        public ExcelColumnNotExistsWithModelException(string message) : base(message) { }
    }
}
