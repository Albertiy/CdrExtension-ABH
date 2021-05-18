using System.Collections.Generic;
using System.ComponentModel;

namespace ABHInstaller
{
    public class ViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        // 手动告知属性变化，只需用在除string以外的引用类型上，如对象或集合
        public void RaisePropertyChange(string propertyName)
        {
            if (this.PropertyChanged != null)
            {
                this.PropertyChanged.Invoke(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        private Dictionary<string, string> installedList;

        public Dictionary<string, string> InstalledList
        {
            get => installedList;
            set
            {
                value.Add("自定义", ""); // 添加一项
                installedList = value;
            }
        }

        public string TextPath
        {
            get => textPath;
            set
            {
                textPath = value;
                this.RaisePropertyChange("TextPath");   // 必须放在值修改后
            }
        }

        private string textPath;

        public ViewModel()
        {
            InstalledList = new Dictionary<string, string>();
        }

    }
}
