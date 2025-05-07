using EPPlusExtensions.ExtensionMethods;
using OfficeOpenXml;
using System.Text;

namespace EPPlusExtensions
{
    public sealed class ExcelPackageHelper
    {
        /// <summary>
        /// 将workSheetIndex转换为代码中确切的值
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="workSheetIndex">从1开始</param>
        /// <returns></returns>
        internal static int ConvertWorkSheetIndex(ExcelPackage excelPackage, int workSheetIndex)
        {
            //if (!excelPackage.Compatibility.IsWorksheets1Based)
            //{
            //    workSheetIndex -= 1; //从0开始的, 需要自己 -1;
            //}
            //return workSheetIndex;

            //var offset = excelPackage.Compatibility.IsWorksheets1Based ? 0 : -1;
            //return workSheetIndex + offset;

            return workSheetIndex + (excelPackage.Compatibility.IsWorksheets1Based ? 0 : -1);
        }


        /// <summary>
        /// 保存文件
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <param name="filePath">文件路径</param>
        public static void Save(ExcelPackage excelPackage, string filePath)
        {
            //File.Delete(savePath); //删除文件。如果文件不存在,也不报错

            //FileMode.Create
            //若指定路径下的文件不存在，系统会创建一个新文件。
            //若指定路径下的文件已经存在，系统会将该文件截断为零字节，也就是清除文件原有的内容。

            var dirPath = Path.GetDirectoryName(filePath);
            if (Directory.Exists(dirPath) == false)
            {
                Directory.CreateDirectory(dirPath);
            }

            //FileMode.Create
            //若指定路径下的文件不存在，系统会创建一个新文件。
            //若指定路径下的文件已经存在，系统会将该文件截断为零字节，也就是清除文件原有的内容。


            using var memoryStream = ExcelPackageHelper.GetMemoryStream(excelPackage);
            using var file = new FileStream(filePath, FileMode.Create, FileAccess.Write);

            //byte[] bytes = new byte[memoryStream.Length];
            //_ = memoryStream.Read(bytes, 0, (int)memoryStream.Length);
            byte[] bytes = memoryStream.ToArray();
            file.Write(bytes, 0, bytes.Length);
        }


        /// <summary>
        /// 获得内存流
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <returns></returns>
        public static MemoryStream GetMemoryStream(ExcelPackage excelPackage)
        {
            var ms = new MemoryStream();
            excelPackage.SaveAs(ms);
            if (ms.CanSeek && ms.Position != 0)
            {
                ms.Position = 0;
                //ms.Seek(0, SeekOrigin.Begin);
            }
            return ms;
        }

    }
}