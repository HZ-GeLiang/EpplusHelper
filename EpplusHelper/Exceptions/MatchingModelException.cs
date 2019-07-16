using System;
using System.Collections.Generic;

namespace EpplusExtensions.Exceptions
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
        public List<ExcelCellInfo> ExcelCellInfoList { get; set; }
        public Type ModelType { get; set; }
    }
}
