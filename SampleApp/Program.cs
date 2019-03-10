﻿using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EpplusExtensions;
using OfficeOpenXml;

namespace SampleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            new Sample04_3().Run();
            OpenDirectoryHelp.OpenFilePath(System.IO.Path.Combine(OpenDirectoryHelp.GetSaveFilePath(), @"Debug\模版\"));
        }

    }
}
