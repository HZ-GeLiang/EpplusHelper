using System;
using System.Data;

namespace SampleApp.Core
{
    class Program
    {
        static void Main(string[] args)
        {
            DataTable dt = new DataTable();
            dt.AsEnumerable();
            Console.WriteLine("Hello World!");
        }
    }
}
