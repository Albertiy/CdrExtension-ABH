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
        public static bool CopyResourceToDir(ResourcesLoader rl, string targetDir)
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

                    // 目标路径
                    string targetUrl = Path.Combine(targetDir, rl.resourceList[key]);

                    // Console.WriteLine("|| 创建文件夹：" + Path.GetDirectoryName(targetUrl));
                    // File.Create 文件创建需要文件夹必须存在；
                    // Directory.CreateDirectory 可以创建多级文件夹
                    Directory.CreateDirectory(Path.GetDirectoryName(targetUrl));

                    using (FileStream fileStream = File.Create(targetUrl))
                    {
                        info.Stream.Seek(0, SeekOrigin.Begin); // 从第一个字节开始复制
                        info.Stream.CopyTo(fileStream);
                        Console.WriteLine("|| Write to " + targetUrl + " Done!");
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
    }
}
