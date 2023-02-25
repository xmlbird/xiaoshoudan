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
		
		//判断一个文件是否被打开
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
	
	
	
	//判断是否为空文件夹
	
	public static bool isEmptyFolder(string str1)
	{
	    if (Directory.GetDirectories(str1 + "\\").Length > 0 || Directory.GetFiles(str1+ "\\").Length > 0)
			{
	
			return false ;
			}
	    return true;
	}

    //创建文件并写入内容
	public static void WriteFiles(string str,string str1)
    {
		
       if (!File.Exists(str))
            {
               //没有则创建这个文件
                FileStream fs1 = new FileStream(str, FileMode.Create, FileAccess.Write);//创建写入文件 
               //设置文件属性为隐藏
                //File.SetAttributes(@"c:\\users\\administrator\\desktop\\webapplication1\\webapplication1\\testtxt.txt",  FileAttributes.System );    
                StreamWriter sw = new StreamWriter(fs1);
                sw.Write(str1);//开始写入值
                sw.Close();
                fs1.Close();
               
            }

      }
	
	
	
	
	
	}


}
