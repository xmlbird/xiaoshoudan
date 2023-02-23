/*
 * 由SharpDevelop创建。
 * 用户： e
 * 日期: 2023/2/23
 * 时间: 21:30
 * 
 * 要改变这种模板请点击 工具|选项|代码编写|编辑标准头文件
 */
using System;
using System .IO;

namespace xiaoshoudan
{
	/// <summary>
	/// Description of Class1.
	/// </summary>
	public static class AboutFile
	{
		public static bool isFileLocked(string pathName)
        {
            try
            {
                if (!File.Exists(pathName))
                {
                    return false;
                }
                using (var fs = new FileStream(pathName, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    fs.Close();
                }
            }
            catch
            {
                return true;
            }
            return false;
        }
	}
}
