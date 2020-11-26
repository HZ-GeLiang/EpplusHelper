using System.IO;

namespace EPPlusTool.MethodExtension
{
    internal static class MemoryStreamExtensions
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="memoryStream">内存流</param>
        /// <param name="savePath">保存的路径</param>
        /// <param name="memoryStreamPosition">设置流中的当前位置,默认0</param>
        public static void Save(this MemoryStream memoryStream, string savePath, int memoryStreamPosition = 0)
        {
            memoryStream.Position = memoryStreamPosition;
            using (var file = new FileStream(savePath, FileMode.Create, System.IO.FileAccess.Write))
            {
                byte[] bytes = new byte[memoryStream.Length];
                memoryStream.Read(bytes, 0, (int)memoryStream.Length);
                file.Write(bytes, 0, bytes.Length);
            }
        }
    }
}
