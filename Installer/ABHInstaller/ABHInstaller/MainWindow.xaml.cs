using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
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
            WinForm.FolderBrowserDialog dialog = new WinForm.FolderBrowserDialog();
            dialog.ShowDialog();

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
    }
}
