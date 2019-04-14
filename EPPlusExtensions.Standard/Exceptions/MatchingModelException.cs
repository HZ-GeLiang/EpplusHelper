using System;
using System.Collections.Generic;

namespace EPPlusExtensions.Exceptions
{

    internal class MatchingModelException : System.Exception
    {
        public MatchingModelException() : base() { }
        public MatchingModelException(string message) : base(message) { }
        public List<ExcelCellInfoAndModelType> ListExcelCellInfoAndModelType { get; set; }
        internal MatchingModel MatchingModel { get; set; }
    }


    public class ExcelCellInfoAndModelType
    {
        public ExcelCellInfo ExcelCellInfo { get; set; }
        public Type ModelType { get; set; }
    }
}
