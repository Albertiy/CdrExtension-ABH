using ABHInstaller.Tools;
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
        /// 默认文件列表，按照安装的目录结构组织的文件列表
        /// </summary>
        private readonly string[] defaultFileList = { };

        /// <summary>
        /// key：项目Resource文件路径，value：ResourceKeeper对象，保存文件名等必要信息
        /// </summary>
        public Dictionary<string, ResourceKeeper> resourceList = new Dictionary<string, ResourceKeeper>();

        /// <summary>
        /// 构造函数
        /// <param name="resourceRoot">项目资源根路径</param>
        /// </summary>
        public ResourcesLoader(string resourceRoot = "")
        {
            this.resourceRoot = resourceRoot;
            PushToResourceList(defaultFileList);
        }

        /// <summary>
        /// 构造函数，带有文件列表参数
        /// </summary>
        /// <param name="resourceRoot">项目资源根路径</param>
        /// <param name="fileList">文件列表</param>
        public ResourcesLoader(string resourceRoot, string[] fileList)
        {
            this.resourceRoot = resourceRoot;
            PushToResourceList(fileList);
        }

        /// <summary>
        /// 构造函数，带有文件名键值对集合
        /// </summary>
        /// <param name="resourceRoot"></param>
        /// <param name="filePairs"></param>
        public ResourcesLoader(string resourceRoot, Dictionary<string, bool> filePairs)
        {
            this.resourceRoot = resourceRoot;
            PushToResourceList(filePairs);
        }

        /// <summary>
        /// 使用 字符串数组 初始化 ResourceList 字典
        /// </summary>
        /// <param name="fileList">文件名数组</param>
        /// <param name="isCompressed">批量设置是否已压缩</param>
        public void PushToResourceList(string[] fileList, bool isCompressed = false)
        {
            foreach (string f in fileList)
            {
                // s.Replace(this.resourceRoot, "")
                this.resourceList.Add(resourceRoot + f, new ResourceKeeper(f, isCompressed));
            }
        }

        /// <summary>
        /// 使用 键值对集合 初始化 ResourceList 字典
        /// </summary>
        /// <param name="filePairs">文件名和 “是否压缩” 键值对集合</param>
        public void PushToResourceList(Dictionary<string, bool> filePairs)
        {
            foreach (KeyValuePair<string, bool> p in filePairs)
            {
                this.resourceList.Add(resourceRoot + p.Key, new ResourceKeeper(p.Key, p.Value));
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fileName"></param>
        public void PushToResourceList(string fileName)
        {
            this.resourceList.Add(resourceRoot + fileName, new ResourceKeeper(fileName));
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePair"></param>
        public void PushToResourceList(KeyValuePair<string, bool> filePair)
        {
            this.resourceList.Add(resourceRoot + filePair.Key, new ResourceKeeper(filePair.Key, filePair.Value));
        }
    }
}
