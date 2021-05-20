using ABHInstaller.Tools;
using ICSharpCode.SharpZipLib.Zip;
using System;
using System.IO;
using System.Windows;
using System.Windows.Resources;

namespace ABHInstaller
{
    public static class InstallUtility
    {
        /// <summary>
        /// 复制资源文件到目标文件夹
        /// </summary>
        /// <param name="rl">ResourceLoader 实例</param>
        /// <param name="targetDir">目标文件夹路径</param>
        /// <returns>结果</returns>
        public static bool CopyResourcesToDir(ResourcesLoader rl, string targetDir)
        {
            // 方法一：使用资源字典resx，不便于组织大量的文件
            //File.Copy(@"Resources\Alertiy'sBirthdayHat.gms", targetDir);
            // var file = Properties.Resources.Albertiy_sBirthdayHat;
            // /InstallResources/X7/Albertiy'sBirthdayHat.gms
            // /InstallResources/X7/ABH-Resource/type-1-blue.bmp

            foreach (string key in rl.resourceList.Keys)
            {
                Uri uri = new Uri(key, UriKind.Relative);
                try
                {
                    // 方法二：使用文件属性为Resource的文件路径，读取为流
                    StreamResourceInfo info = Application.GetResourceStream(uri);

                    ResourceKeeper rk = rl.resourceList[key];

                    if (!rk.isCompressed)   // 不是压缩文件，正常写入文件流
                    {
                        CopyNormalStream(info.Stream, rk, targetDir);
                    }
                    else
                    {
                        CopyCompressedStream(info.Stream, rk, targetDir);
                    }
                }
                catch (IOException)
                {
                    Console.WriteLine("|| 资源 " + uri + " 不存在");
                    continue;
                }
                catch (Exception e)
                {
                    Console.WriteLine("|| 出现意料之外的错误，跳过复制");
                }
            }
            return true;
        }

        /// <summary>
        /// 复制普通的流到文件
        /// </summary>
        /// <param name="resStream">来自资源文件的流</param>
        /// <param name="rk">文件信息</param>
        /// <param name="targetDir">目标文件夹</param>
        public static void CopyNormalStream(Stream resStream, ResourceKeeper rk, string targetDir)
        {
            // 目标路径
            string targetUrl = Path.Combine(targetDir, rk.name);

            // Console.WriteLine("|| 创建文件夹：" + Path.GetDirectoryName(targetUrl));
            // File.Create 文件创建需要文件夹必须存在；
            // Directory.CreateDirectory 可以创建多级文件夹
            Directory.CreateDirectory(Path.GetDirectoryName(targetUrl));

            using (FileStream fileStream = File.Create(targetUrl))
            {
                resStream.Seek(0, SeekOrigin.Begin); // 从第一个字节开始复制
                resStream.CopyTo(fileStream);
                Console.WriteLine("|| Write to " + targetUrl + " Done!");
            }
        }

        public static void CopyCompressedStream(Stream resStream, ResourceKeeper rk, string targetDir)
        {
            Console.WriteLine("|| 需要解压的文件：" + rk.name);
            // 1. 创建目标文件夹
            string targetUncompressDir = Path.GetDirectoryName(Path.Combine(targetDir, rk.name));
            Directory.CreateDirectory(targetUncompressDir);
            Console.WriteLine("|| 解压到目录：{0}", targetUncompressDir);
            // 当做 Zip 文件解压
            CompressUtility.UnZipFromStream(resStream, targetUncompressDir);
        }

    }
}
