using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Resources;
using WinForm = System.Windows.Forms;

namespace ABHInstaller
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public ViewModel viewModel = new ViewModel();

        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = viewModel; // 全局界面数据上下文绑定
            RefreshSoftwareList();

        }

        private void InstalledListComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Console.WriteLine(" 你选择了：  " + InstalledListComboBox.SelectedItem.ToString());
            string oldStr = this.InstallPathTextBox.Text;
            string newStr = this.viewModel.InstalledList[InstalledListComboBox.SelectedItem.ToString()];
            if (!newStr.Equals(""))
                this.viewModel.TextPath = newStr;
            Console.WriteLine("new Text: " + this.viewModel.TextPath);
        }

        private void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            this.RefreshSoftwareList();
        }

        private void ChooseDirButton_Click(object sender, RoutedEventArgs e)
        {
            this.InstalledListComboBox.SelectedItem = "自定义";
            //WinForm.FolderBrowserDialog dialog = new WinForm.FolderBrowserDialog();
            //WinForm.DialogResult res = dialog.ShowDialog();
            //Console.WriteLine(res);

            CommonOpenFileDialog dialog = new CommonOpenFileDialog
            {
                IsFolderPicker = true,//设置为选择文件夹
                Title = "选择 CorelDRAW 套装所在文件夹（包含 Draw 目录）",
                InitialDirectory = @"C:\",
                Multiselect = false
            };
            String oldPath = this.viewModel.TextPath;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                this.viewModel.TextPath = dialog.FileName + @"\";
                Console.WriteLine("|| The Folder you choosed： " + this.viewModel.TextPath);
            }
            else
            {
                Console.WriteLine("|| 未选择文件夹。");
            }

            //if (dialog.ShowDialog(this) == false) return;
            //_fileName = dialog.FileName;
            ////初始化图片
            //BitmapImage tempImage = new BitmapImage();
            //tempImage.BeginInit();
            //tempImage.UriSource = new Uri(_fileName, UriKind.RelativeOrAbsolute);
            //tempImage.EndInit();

            ////imageCavas为窗口中设置的Image控件
            //imageCavas.Source = tempImage;
        }

        private void InstallButton_Click(object sender, RoutedEventArgs e)
        {
            string folderPath = this.viewModel.TextPath;

            string errorReason = "";
            Console.WriteLine("|| 检测目录 “" + folderPath + "” 是否存在；");
            if (Directory.Exists(folderPath))
            {
                string copyToPath = Path.Combine(folderPath, @"Draw\GMS\");
                Console.WriteLine("|| 检测目录 “" + copyToPath + "” 是否存在；");
                if (Directory.Exists(copyToPath))
                {
                    Console.WriteLine("|| 是有效的路径，开始复制：");
                    ResourcesLoader rl = new ResourcesLoader();
                    if(this.CopyResourceToDir(rl, copyToPath))
                    {
                        MessageBox.Show("安装完成", "提示", MessageBoxButton.OK);
                        System.Windows.Application.Current.Shutdown();
                    } else
                    {
                        MessageBox.Show("安装失败", "提示", MessageBoxButton.OK);
                    }
                    return;
                }
                else errorReason = "未找到有效的CorelDRAW安装路径！";
            }
            else errorReason = "文件夹不存在！";
            Console.WriteLine("|| " + errorReason);
            MessageBoxResult result = MessageBox.Show(errorReason, "提示", MessageBoxButton.OK);
        }

        public void RefreshSoftwareList()
        {
            viewModel.InstalledList = SearchInstalledDirection();

        }
        /// <summary>
        /// 从注册表查找本机CorelDraw安装位置列表
        /// </summary>
        protected Dictionary<string, string> SearchInstalledDirection()
        {
            // 可能有也可能没有；可能有多个子项，也可能没有子项。ShellExt 是无效的，CorelDRAW Graphics Suite 17 是有效的；Destination 是需要的。
            string rootItem = @"SOFTWARE\Corel\Setup";
            ArrayList items = this.GetRegistryItems(rootItem);
            Dictionary<string, string> itemAndParm = new Dictionary<string, string>();
            foreach (string i in items)
            {
                string p = GetRegistryValue(rootItem + @"\" + i, "Destination");
                if (!p.Equals(string.Empty))
                {
                    Console.WriteLine("|| 有效项：{ " + i + " , " + p + " }");
                    // 直接覆盖，这样即使是没有继承通知接口的集合对象，也可以通知页面刷新
                    itemAndParm.Add(i, p);
                }
            }
            return itemAndParm;
        }
        //1.向注册表中写信息
        //protected void SetRegistryValue()
        //{
        //    using (RegistryKey key = Registry.LocalMachine.OpenSubKey(@"", true))
        //    {
        //        if (key == null)
        //        {
        //            using (RegistryKey myKey = Registry.LocalMachine.CreateSubKey(@"SOFTWARE\EricSun\MyTestKey"))
        //            {
        //                myKey.SetValue("MyKeyName", "Hello EricSun." + DateTime.Now.ToString());
        //            }
        //        }
        //        else
        //        {
        //            key.SetValue("MyKeyName", "Hello EricSun." + DateTime.Now.ToString());
        //        }
        //    }

        //}

        /// <summary>
        /// 读取注册表项上的属性参数(IIS也可以正常读取)
        /// </summary>
        /// <param name="path">项路径</param>
        /// <param name="paramName">属性名</param>
        /// <example>"SOFTWARE\\EricSun\\MyTestKey", "MyKeyName"</example>
        /// <returns></returns>
        protected string GetRegistryValue(string path, string paramName)
        {
            Console.WriteLine("|| GetRegistryValue: " + path + " param: " + paramName);
            string value = string.Empty;
            RegistryKey root = Registry.LocalMachine;
            RegistryKey rk = root.OpenSubKey(path);
            if (rk != null)
            {
                value = (string)rk.GetValue(paramName, null);
            }
            Console.WriteLine("|| 找到 值：" + value);
            return value;
        }

        protected ArrayList GetRegistryItems(string path)
        {
            Console.WriteLine("|| GetRegistryItems: " + path);
            ArrayList items = new ArrayList();

            RegistryKey root = Registry.LocalMachine;
            RegistryKey rk = root.OpenSubKey(path);
            if (rk != null)
            {
                string[] keys = rk.GetSubKeyNames();
                foreach (var keyName in keys)
                {
                    Console.WriteLine("|| 找到 项: " + keyName);
                    if (!keyName.Equals("ShellExt"))
                        items.Add(keyName);
                }
            }
            return items;
        }

        protected bool CopyResourceToDir(ResourcesLoader rl, string targetDir)
        {
            //File.Copy(@"Resources\Alertiy'sBirthdayHat.gms", targetDir);
            // var file = Properties.Resources.Albertiy_sBirthdayHat;
            // /Resources/GMS/Albertiy'sBirthdayHat.gms
            // /Resources/GMS/ABH-Resource/type-1-blue.bmp
            foreach (string key in rl.resourceList.Keys)
            {
                Uri uri = new Uri(key, UriKind.Relative);
                try
                {
                    StreamResourceInfo info = Application.GetResourceStream(uri);

                    // 目标路径
                    string targetUrl = Path.Combine(targetDir, rl.resourceList[key]);

                    Console.WriteLine("|| 创建文件夹：" + Path.GetDirectoryName(targetUrl));
                    
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
                catch (System.IO.IOException)
                {
                    Console.WriteLine("|| 资源 " + uri + " 不存在");
                    continue;
                }
                catch(Exception e)
                {
                    Console.WriteLine("|| 出现意料之外的错误，跳过复制");
                }
            }
            return true;
        }
    }
}
