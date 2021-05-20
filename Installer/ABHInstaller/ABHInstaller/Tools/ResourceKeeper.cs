namespace ABHInstaller.Tools
{
    public class ResourceKeeper
    {
        public string name;         // 文件名（或相对路径）
        public bool isCompressed;   // 是否是压缩文件（用于自动解压缩）

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="name">文件名</param>
        /// <param name="isCompressed">是否已压缩</param>
        public ResourceKeeper(string name, bool isCompressed = false)
        {
            this.name = name;
            this.isCompressed = isCompressed;
        }
    }
}
