using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusExtensions.Helpers
{

    public sealed class FileHelper
    {

        /// <summary>
        /// 读取一个文件,获得一个文件流
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="mode"></param>
        /// <param name="access"></param>
        /// <param name="share"></param>
        /// <returns></returns>
        public static FileStream GetFileStream(string filePath,
            FileMode mode = FileMode.Open,
            FileAccess access = FileAccess.Read,
            FileShare share = FileShare.ReadWrite)
        {
            return new FileStream(filePath, mode, access, share);
        }
    }
}
