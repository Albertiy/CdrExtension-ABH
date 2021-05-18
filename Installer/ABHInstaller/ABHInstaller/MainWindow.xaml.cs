using ABHInstaller.Tools;
using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace ABHInstaller
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        /// <summary>
        /// 共用的资源文件列表
        /// </summary>
        private string[] commonFileList = {
            "ABH-Resource/type-1-blue.bmp",
            "ABH-Resource/type-1-pink.bmp",
            "ABH-Resource/type-1.cdr",
            "ABH-Resource/type-1.cdt",
            "ABH-Resource/type-2.bmp",
            "ABH-Resource/type-2.cdr",
            "ABH-Resource/type-2.cdt",
            "ABH-Resource/type-3-blue.bmp",
            "ABH-Resource/type-3-pink.bmp",
            "ABH-Resource/type-3.cdr",
            "ABH-Resource/type-3.cdt",
            "ABH-Resource/type-4-blue.bmp",
            "ABH-Resource/type-4.cdr",
            "ABH-Resource/type-4.cdt",
            "ABH-Resource/type-5-pink.bmp",
            "ABH-Resource/type-5.cdr",
            "ABH-Resource/type-5.cdt",
            "ABH-Resource/type-6-classicred.bmp",
            "ABH-Resource/type-6.cdr",
            "ABH-Resource/type-6.cdt",
            "ABH-Resource/type.bmp"
        };
        /// <summary>
        /// X4版本独有的资源文件
        /// </summary>
        private string[] X4FileList = {
            "ABHX4.gms",
            "ABH-Resource/Toolbar_X4.xslt",
        };
        /// <summary>
        /// X7版本独有的资源文件
        /// </summary>
        private string[] X7FileList = {
            "Albertiy'sBirthdayHat.gms",
            "ABH-Resource/ABH_Toolbar.cdws",
        };

        /// <summary>
        /// 动态绑定的视图模型
        /// </summary>
        protected ViewModel viewModel = new ViewModel();
        /// <summary>
        /// 构造函数
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
            this.DataContext = viewModel; // 全局界面数据上下文绑定
            RefreshSoftwareList();
        }
        /// <summary>
        /// 复选框选择改变事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void InstalledListComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Console.WriteLine("|| 复选框选择改变：" + InstalledListComboBox.SelectedItem.ToString());
            string oldStr = this.InstallPathTextBox.Text;
            string newStr = this.viewModel.InstalledList[InstalledListComboBox.SelectedItem.ToString()];
            if (!newStr.Equals(""))
            {
                this.viewModel.TextPath = newStr;
                Console.WriteLine("|| 文本框内容改变: " + this.viewModel.TextPath);
            }
        }
        /// <summary>
        /// 刷新复选框按钮点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            this.RefreshSoftwareList();
        }
        /// <summary>
        /// 手动选择目录按钮点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ChooseDirButton_Click(object sender, RoutedEventArgs e)
        {
            InstalledListComboBox.SelectedItem = "自定义";    // 修改下拉列表项，方便用户重新选择

            CommonOpenFileDialog dialog = new CommonOpenFileDialog
            {
                IsFolderPicker = true,  //设置为选择文件夹
                Title = "选择 CorelDRAW 套装所在文件夹（包含 Draw 目录）",
                InitialDirectory = Directory.Exists(viewModel.TextPath.Trim()) ? viewModel.TextPath.Trim() : @"C:\",  // 初始位置，若当前文本框文件夹存在，则显示当前文件夹
                Multiselect = false     // 是否多选
            };
            string oldPath = viewModel.TextPath;
            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
            {
                string newPath = dialog.FileName + @"\";
                if (oldPath != newPath)
                {
                    viewModel.TextPath = newPath;
                    Console.WriteLine("|| The Folder you choosed： " + newPath);
                }
                else
                    Console.WriteLine("|| 选择的文件夹未改变");
            }
            else
            {
                Console.WriteLine("|| 未选择文件夹");
            }
        }
        /// <summary>
        /// 安装按钮点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void InstallButton_Click(object sender, RoutedEventArgs e)
        {
            string errorReason = InstallResources(viewModel.TextPath);
            if (errorReason != string.Empty)
            {
                Console.WriteLine("|| " + errorReason);
                MessageBox.Show(errorReason, "提示", MessageBoxButton.OK);
            }
        }
        /// <summary>
        /// 刷新软件路径列表
        /// </summary>
        public void RefreshSoftwareList()
        {
            // 直接覆盖，这样即使是没有继承通知接口的集合对象，也可以通知页面刷新
            viewModel.InstalledList = SearchInstalledDirection();
        }
        /// <summary>
        /// 查找本机CorelDraw安装位置列表，本项目中是从注册表参数查找
        /// </summary>
        protected Dictionary<string, string> SearchInstalledDirection()
        {
            // 可能有也可能没有；可能有多个子项，也可能没有子项。ShellExt 是无效的，CorelDRAW Graphics Suite 17 是有效的；Destination 是需要的。
            string rootItem = @"SOFTWARE\Corel\Setup";  // X7 注册表项的位置
            string rootItem2 = @"SOFTWARE\WOW6432Node\Corel\Setup"; // X4 或 32位注册表地址的位置

            ArrayList items = RegistryUtility.GetRegistryItems(Registry.LocalMachine, rootItem, new string[] { "ShellExt" });
            ArrayList items2 = RegistryUtility.GetRegistryItems(Registry.LocalMachine, rootItem2, new string[] { "ShellExt" });

            Dictionary<string, string> itemAndParm = new Dictionary<string, string>();

            foreach (string i in items)
            {
                string p = RegistryUtility.GetRegistryValue(Registry.LocalMachine, rootItem + @"\" + i, "Destination");
                if (!p.Equals(string.Empty))
                {
                    Console.WriteLine("|| 下拉列表有效项：{ " + i + " , " + p + " }");
                    itemAndParm.Add(i, p);
                }
            }
            foreach (string i in items2)
            {
                string p = RegistryUtility.GetRegistryValue(Registry.LocalMachine, rootItem2 + @"\" + i, "Destination");
                if (!p.Equals(string.Empty))
                {
                    Console.WriteLine("|| 下拉列表有效项：{ " + i + " , " + p + " }");
                    itemAndParm.Add(i, p);
                }
            }
            return itemAndParm;
        }

        public string InstallResources(string folderPath)
        {
            Console.WriteLine("|| 目标安装路径{0}", folderPath);
            string errorReason = string.Empty;
            if (Directory.Exists(folderPath))
            {
                string copyToPath = Path.Combine(folderPath, @"Draw\GMS\"); // 子路径
                Console.WriteLine("|| 检测目录 “" + copyToPath + "” 是否存在；");
                if (Directory.Exists(copyToPath))
                {
                    Console.WriteLine("|| 是有效的路径，开始复制：");
                    ResourcesLoader rl;
                    List<string> tempList = new List<string>();
                    tempList.AddRange(commonFileList);
                    if (folderPath.Contains("X4") || InstalledListComboBox.SelectedItem.ToString().Contains("X4"))  // 复制 X4 版本的资源文件
                    {
                        tempList.AddRange(X4FileList);
                        rl = new ResourcesLoader(@"/InstallResources/X4/", tempList.ToArray());
                    }
                    else // 复制 X7 版本的资源文件
                    {
                        tempList.AddRange(X7FileList);
                        rl = new ResourcesLoader(@"/InstallResources/X7/", tempList.ToArray());
                    }
                    if (InstallUtility.CopyResourceToDir(rl, copyToPath))
                    {
                        MessageBox.Show("安装完成", "提示", MessageBoxButton.OK);
                        Application.Current.Shutdown();  // 关闭程序
                    }
                    else
                    {
                        MessageBox.Show("安装失败", "提示", MessageBoxButton.OK);
                        errorReason = "Bad Copy Result!";
                    }
                }
                else errorReason = "未找到有效的CorelDRAW安装路径！";
            }
            else
                errorReason = "文件夹不存在！";
            return errorReason;
        }
    }
}
