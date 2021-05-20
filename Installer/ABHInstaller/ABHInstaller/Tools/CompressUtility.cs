using ICSharpCode.SharpZipLib.Zip;
using System;
using System.IO;

namespace ABHInstaller.Tools
{
    public class CompressUtility
    {
        /// <summary>
        /// 解压并保存 Zip 格式的文件流
        /// </summary>
        /// <param name="resStream">文件流</param>
        /// <param name="rootDir">解压根路径</param>
        /// <returns></returns>
        public static bool UnZipFromStream(Stream resStream, string rootDir)
        {
            ZipInputStream zippedStream = null;
            string filePath;
            try
            {
                zippedStream = new ZipInputStream(resStream);
                ZipEntry ent;
                while ((ent = zippedStream.GetNextEntry()) != null)
                {
                    if (!string.IsNullOrEmpty(ent.Name))
                    {
                        filePath = Path.Combine(rootDir, ent.Name);
                        filePath = filePath.Replace('/', '\\'); // 斜杠换反斜杠
                        Console.WriteLine("|| ZipEntry：{0}", filePath);

                        // 纯文件夹，创建文件夹即可
                        if (filePath.EndsWith("\\"))
                        {
                            Directory.CreateDirectory(filePath);
                            continue;
                        }
                        else
                        {
                            Console.WriteLine("|| 需要文件夹：{0}", Path.GetDirectoryName(filePath));
                            Directory.CreateDirectory(Path.GetDirectoryName(filePath));
                            using (FileStream fs = File.Create(filePath))   // 创建或覆盖文件
                            {
                                int size = 2048;
                                byte[] data = new byte[size];
                                while (true)
                                {
                                    size = zippedStream.Read(data, 0, data.Length);
                                    if (size > 0)
                                        fs.Write(data, 0, size);
                                    else
                                        break;
                                }
                            }
                        }
                    }
                }
                Console.WriteLine("|| 解压完成。");
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine("|| 解压失败");
                return false;
            }
            finally
            {
                if (zippedStream != null)
                    zippedStream.Close();
            }
        }
    }
}
