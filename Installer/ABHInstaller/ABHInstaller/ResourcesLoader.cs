using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ABHInstaller
{
    public class ResourcesLoader
    {
        private string resourceRoot = "/Resources/GMS/";

        private string[] fileList = {
            "/Resources/GMS/Albertiy'sBirthdayHat.gms",
            "/Resources/GMS/ABH-Resource/ABH_Toolbar.cdws",
            "/Resources/GMS/ABH-Resource/type-1-blue.bmp",
            "/Resources/GMS/ABH-Resource/type-1-pink.bmp",
            "/Resources/GMS/ABH-Resource/type-1.cdr",
            "/Resources/GMS/ABH-Resource/type-1.cdt",
            "/Resources/GMS/ABH-Resource/type-2.bmp",
            "/Resources/GMS/ABH-Resource/type-2.cdr",
            "/Resources/GMS/ABH-Resource/type-2.cdt",
            "/Resources/GMS/ABH-Resource/type-3-blue.bmp",
            "/Resources/GMS/ABH-Resource/type-3-pink.bmp",
            "/Resources/GMS/ABH-Resource/type-3.cdr",
            "/Resources/GMS/ABH-Resource/type-3.cdt",
            "/Resources/GMS/ABH-Resource/type-4-blue.bmp",
            "/Resources/GMS/ABH-Resource/type-4.cdr",
            "/Resources/GMS/ABH-Resource/type-4.cdt",
            "/Resources/GMS/ABH-Resource/type-5-pink.bmp",
            "/Resources/GMS/ABH-Resource/type-5.cdr",
            "/Resources/GMS/ABH-Resource/type-5.cdt",
            "/Resources/GMS/ABH-Resource/type-6-classicred.bmp",
            "/Resources/GMS/ABH-Resource/type-6.cdr",
            "/Resources/GMS/ABH-Resource/type-6.cdt",
            "/Resources/GMS/ABH-Resource/type.bmp"
        };

        // key：项目Resource文件路径，value：目标相对路径
        public Dictionary<string, string> resourceList = new Dictionary<string, string>();
        public ResourcesLoader()
        {
            foreach (string s in this.fileList)
            {
                this.resourceList.Add(s, s.Replace(this.resourceRoot,""));
            }
        }
    }
}
