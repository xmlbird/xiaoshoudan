/*
 * 由SharpDevelop创建。
 * 用户： e
 * 日期: 2023/2/16
 * 时间: 18:59
 * 
 * 要改变这种模板请点击 工具|选项|代码编写|编辑标准头文件
 */

using System;
using System .Text ;
using System .Data.OleDb ;
using System .Data;
using System .IO ;
using System.Security.Cryptography;  
using System .Collections; 
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace xiaoshoudan
{
	/// <summary>
	/// Description of Class1.
	/// </summary>

	
//对填写内容进行校验类
public static class myvalidate{
	
		//1.判断是否为空
		 public static bool ifblank(string str){
			if (str.Trim() == null || str.Trim().Length == 0)    //验证这个参数是否为空
			{return true;}                           //是，就返回False
             else
             {return false ;}
		}
		
		//2.判断是否都是数字
		public static bool myvalidateNumric(string str){
     		 ASCIIEncoding ascii = new ASCIIEncoding();//new ASCIIEncoding 的实例
             byte[] bytestr = ascii.GetBytes(str);         //把string类型的参数保存到数组里
             foreach (byte c in bytestr)                   //遍历这个数组里的内容
            {
                if (c < 48 || c > 57)                          //判断是否为数字
                {
                    return false;                              //不是，就返回False
                                    }
            }
            return true;                                        //是，就返回True
        }
		
		
		//3.判断是否超过某个 位数
		 public static bool ifchaoguoweishu(string str , int n){
			if (str.Trim().Length > n)    //验证这个参数是否大于某个位数
                return false;                           //是，就返回False
            else
             	return true;    
		}
		
		//4.判断是否低于某个位数
		 public static bool ifdiyuweishu(string str , int n){
			if (str.Trim().Length < n)    //验证这个参数是否大于某个位数
                return false;                           //是，就返回False
			else
             	return true;
            }
				
		//判断是否是两位小数或整数
		public static bool ifshuozi(string str)
		{Regex reg =new Regex("^[0-9]*(\\.[0-9]{1,2})?$");
		
		if (reg.IsMatch (str))
		{
			return true ;
		}
		else 
		{
			return false;
		}
		
		
		}
		
		}

//读文件类
public static class readtxt{

	public static string IP="";
		
	public static string getIP(){
		foreach (string line in File.ReadLines(@"IPAddress.txt"))
                    {   
		      //if ((line.ToString().Trim ().Substring(0,2))=="IP")
		        //    {         
		      	IP =(line .ToString ().Trim ());
		    	//  }
		   
		}
	   return IP;
	
	}
}

//数据库类
public  class DataHelper{
	//"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dataAddress;    
	// @"data source="+chuandi .IPa +";initial catalog=zhanghu;user id=sa;pwd=111111";
	
	//创建数据库连接
     
     public OleDbConnection  link { get; set; }
     public DataHelper ()
     {
     }
     public DataHelper (string str)
     {
     link = new OleDbConnection( "provider= Microsoft.Jet.OLEDB.4.0;extended properties=Excel 8.0;data source="+ str);
     
     }
	 
		   
     //1.打开数据库
	   public void open(){
            if (link.State==0)
    	    link.Open();
    		}
	  
        //2.读取数据库
       public OleDbDataReader dataread(string selectstring ) {
    	OleDbCommand  mycommand=link .CreateCommand ();
        mycommand.CommandText =selectstring ;
        OleDbDataReader myReader=mycommand.ExecuteReader ();
        return myReader;
      }
    
        //3.执行数据库命令
        public void dataexecute(string exestr){
          OleDbCommand  mycommand=link.CreateCommand();
          mycommand.CommandText =exestr;
          mycommand.ExecuteNonQuery ();
       }
    
       //4.关闭数据库
        public void close(){
    	   if (link.State!=0)
    	   link.Close ();
           link =null ;
        }

       //5.填充DATAset
        public  void FullDataSet(string sqlstr,DataSet mydataset,string TableName){
    	OleDbDataAdapter  myDataAdapter=new OleDbDataAdapter  (sqlstr ,link);
    	myDataAdapter.Fill (mydataset ,TableName );
        }
  
}


//传递值类
public class chuandi{
public static int zhi1;//级别
public static object zhi2;//登陆窗体
public static string zhi3;  //用户名
public static string IPa;  
public static bool issev; //是否是服务器
}


//询问类

static class Ask
{
	static public bool YesOrNot(string str)
	{
		DialogResult a = MessageBox .Show(str,"注意",MessageBoxButtons .YesNo ,MessageBoxIcon .Question ,MessageBoxDefaultButton .Button1 );
	if (a == DialogResult.No )
	{return false;}
	else
	{return true;}
	}
}


//添加数据库类

static class  DataExec
{
	static public int DE(string str){
		DataHelper helper1=new DataHelper ();
	   try {
			helper1 .open();
			helper1 .dataexecute(str);
			return 1 ;
			}
			catch{
			MessageBox .Show("错误");
			return 2;
		      }
			finally {
			helper1.close ();
			}
	}
}

}













///取汉字首字母类
 static class GetFirstLetter
 {
/// <summary>
 
/// 在指定的字符串列表CnStr中检索符合拼音索引字符串
 
///</summary>
 
/// <param name="CnStr">汉字字符串</param>
 
/// <returns>相对应的汉语拼音首字母串</returns>
 
public static string GetSpellCode(string CnStr) {
 
　  string strTemp="";
 
　　int iLen=CnStr.Length;
 
　　int i=0;
 
　　for (i=0;i<=iLen-1;i++) {
 
　　　　 strTemp+=GetCharSpellCode(CnStr.Substring(i,1));
 
　　　　 
　　}
 
　　return strTemp;
 
}
 
/// <summary>
 
/// 得到一个汉字的拼音第一个字母，如果是一个英文字母则直接返回大写字母
 
/// </summary>
 
/// <param name="CnChar">单个汉字</param>
 
/// <returns>单个大写字母</returns>
 
private static string GetCharSpellCode(string CnChar) {
 
　　long iCnChar;
 
　　byte[] ZW = System.Text.Encoding.Default.GetBytes(CnChar);
 
　　//如果是字母，则直接返回
 
　　if (ZW.Length==1) {
 
　　　　 return CnChar.ToUpper();
 
　　}
 
　　else {
 
　　　　 // get the array of byte from the single char
 
　　　　int i1 = (short)(ZW[0]);
 
　　　　int i2 = (short)(ZW[1]);
 
　　　　iCnChar = i1*256+i2;
 
　　　　}　　
 
// iCnChar match the constant
 
　　if ((iCnChar>=45217) && (iCnChar<=45252)) {
 
　　　　 return "A";
 
　　}
 
　　else if ((iCnChar>=45253) && (iCnChar<=45760)) {
 
　　　　return "B";
 
　　} else if ((iCnChar>=45761) && (iCnChar<=46317)) {
 
　　　　return "C";
 
　　} else if ((iCnChar>=46318) && (iCnChar<=46825)) {
 
　　　　return "D";
 
　　} else if ((iCnChar>=46826) && (iCnChar<=47009)) {
 
　　　　return "E";
 
　　} else if ((iCnChar>=47010) && (iCnChar<=47296)) {
 
　　　　return "F";
 
　　} else if ((iCnChar>=47297) && (iCnChar<=47613)) {
 
　　　　return "G";
 
　　} else if ((iCnChar>=47614) && (iCnChar<=48118)) {
 
　　　　return "H";
 
　　} else if ((iCnChar>=48119) && (iCnChar<=49061)) {
 
　　　　return "J";
 
　　} else if ((iCnChar>=49062) && (iCnChar<=49323)) {
 
　　　　return "K";
 
　　} else if ((iCnChar>=49324) && (iCnChar<=49895)) {
 
　　　　return "L";
 
　　} else if ((iCnChar>=49896) && (iCnChar<=50370)) {
 
　　　　return "M";
 
　　}else if ((iCnChar>=50371) && (iCnChar<=50613)) {
 
　　　　return "N";
 
　　} else if ((iCnChar>=50614) && (iCnChar<=50621)) {
 
　　　　return "O";
 
　　} else if ((iCnChar>=50622) && (iCnChar<=50905)) {
 
　　　　return "P";
 
　　} else if ((iCnChar>=50906) && (iCnChar<=51386)) {
 
　　　　return "Q";
 
　　} else if ((iCnChar>=51387) && (iCnChar<=51445)) {
 
　　　　return "R";
 
　　} else if ((iCnChar>=51446) && (iCnChar<=52217)) {
 
　　　　return "S";
 
　　} else if ((iCnChar>=52218) && (iCnChar<=52697)) {
 
　　　　return "T";
 
　　} else if ((iCnChar>=52698) && (iCnChar<=52979)) {
 
　　　　return "W";
 
　　} else if ((iCnChar>=52980) && (iCnChar<=53640)) {
 
　　　　return "X";
 
　　} else if ((iCnChar>=53689) && (iCnChar<=54480)) {
 
　　　　return "Y";
 
　　} else if ((iCnChar>=54481) && (iCnChar<=55289)) {
 
　　　　return "Z";
 
　　} else
 
　　return ("?");
 
}
















































}


