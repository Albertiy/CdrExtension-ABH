using Microsoft.Win32;
using System;
using System.Collections;

namespace ABHInstaller.Tools
{
    public static class RegistryUtility
    {
        //1.
        /// <summary>
        /// 向注册表项中设置（或创建）参数
        /// </summary>
        /// <param name="root">注册表顶级节点</param>
        /// <param name="path">注册表项路径</param>
        /// <param name="paramKey">参数名</param>
        /// <param name="paramValue">参数值</param>
        /// <param name="isExist">注册表项不存在时是否创建项</param>
        /// <returns>创建是否成功</returns>
        public static bool SetRegistryParam(RegistryKey root, string path, string paramKey, string paramValue, bool isExist)
        {
            try
            {
                using (RegistryKey key = root.OpenSubKey(path, true))
                {
                    if (key != null)
                    {
                        key.SetValue(paramKey, paramValue);
                        return true;
                    }
                    else if (key == null && !isExist)
                    {
                        using (RegistryKey myKey = root.CreateSubKey(path))
                        {
                            myKey.SetValue(paramKey, paramValue);
                            return true;
                        }
                    }
                    else
                    {
                        return false;
                    }

                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 获取注册表项的子项列表
        /// </summary>
        /// <param name="root">注册表顶级节点</param>
        /// <param name="path">注册表项路径</param>
        /// <param name="excludeList">排除的子项</param>
        /// <returns>返回ArrayList对象</returns>
        public static ArrayList GetRegistryItems(RegistryKey root, string path, string[] excludeList = null)
        {
            Console.WriteLine("|| GetRegistryItems: " + path);
            ArrayList items = new ArrayList();
            RegistryKey rk = root.OpenSubKey(path);
            if (rk != null)
            {
                string[] subKeys = rk.GetSubKeyNames();
                foreach (var key in subKeys)
                {
                    if (excludeList == null || Array.IndexOf(excludeList, key) == -1)
                        items.Add(key);
                }
                Console.WriteLine("|| 注册表项 {0}\\{1} 有以下子项：{2}", root, path, string.Join(", ", (string[])items.ToArray(typeof(string))));
            }
            else
                Console.WriteLine("|| 无效的注册表路径：{0}\\{1}", root, path);
            return items;
        }

        /// <summary>
        /// 读取注册表项上的参数的值 (IIS也可以正常读取)
        /// </summary>
        /// <param name="root">注册表顶级节点</param>
        /// <param name="path">注册表项路径</param>
        /// <param name="paramKey">注册表项参数名</param>
        /// <returns>如果存在，返回注册表项的参数值；否则返回 string.Empty</returns>
        public static string GetRegistryValue(RegistryKey root, string path, string paramKey)
        {
            Console.WriteLine("|| GetRegistryValue: {0}\\{1} param: {2}", root.ToString(), path, paramKey);
            string paramValue = string.Empty;
            RegistryKey rk = root.OpenSubKey(path);
            if (rk != null)
            {
                paramValue = (string)rk.GetValue(paramKey, string.Empty);
                if (paramValue != string.Empty)
                    Console.WriteLine("|| 找到注册表项 {0}\\{1} 的参数 {2} 的值：{3}", root.ToString(), path, paramKey, paramValue);
                else
                    Console.WriteLine("|| 未找到注册表项 {0}\\{1} 的参数 {2} 的值", root.ToString(), path, paramKey);
            }
            else
            {
                Console.WriteLine("|| 未找到注册表项 {0}\\{1}", root.ToString(), path);
            }
            return paramValue;
        }
    }
}
