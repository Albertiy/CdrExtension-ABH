using System.Collections.Generic;

namespace ABHInstaller
{
    public class ResourcesLoader
    {
        /// <summary>
        /// 安装文件在本项目中的资源根路径
        /// </summary>
        private readonly string resourceRoot;
        /// <summary>
        /// 默认文件列表
        /// </summary>
        private string[] defaultFileList = { };
        /// <summary>
        /// 按照安装的目录结构组织的文件列表
        /// </summary>
        private string[] fileList;
        /// <summary>
        /// key：项目Resource文件路径，value：目标相对路径
        /// </summary>
        public Dictionary<string, string> resourceList = new Dictionary<string, string>();
        /// <summary>
        /// 构造函数
        /// <param name="resourceRoot">项目资源根路径</param>
        /// </summary>
        public ResourcesLoader(string resourceRoot = "")
        {
            this.resourceRoot = resourceRoot;
            fileList = defaultFileList;
            InitializeResourceList();
        }
        /// <summary>
        /// 构造函数重载，带有文件列表参数
        /// </summary>
        /// <param name="resourceRoot">项目资源根路径</param>
        /// <param name="fileList">文件列表</param>
        public ResourcesLoader(string resourceRoot, string[] fileList)
        {
            this.resourceRoot = resourceRoot;
            this.fileList = fileList;
            InitializeResourceList();
        }
        /// <summary>
        /// 初始化 ResourceList 字典
        /// </summary>
        private void InitializeResourceList()
        {
            foreach (string s in fileList)
            {
                // s.Replace(this.resourceRoot, "")
                this.resourceList.Add(resourceRoot + s, s);
            }
        }
    }
}
