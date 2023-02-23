/*
 * 由SharpDevelop创建。
 * 用户： e
 * 日期: 2023/2/22
 * 时间: 18:16
 * 
 * 要改变这种模板请点击 工具|选项|代码编写|编辑标准头文件
 */
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System .Data .OleDb ;
using System .IO ;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.SS.Util;
using NPOI.SS.UserModel;

namespace xiaoshoudan
{
	/// <summary>
	/// Description of MainForm.
	/// </summary>
	public partial class MainForm : Form
	{
		public MainForm()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
		}
		void MainFormLoad(object sender, EventArgs e)
		{
			     
	            
		         
		}
		void Button1Click(object sender, EventArgs e)
		{
	             OpenFileDialog fileDialog = new OpenFileDialog();
	             fileDialog.Multiselect = false; 
	             fileDialog.Title = "请选择文件"; 
	             fileDialog.Filter="Excel工作簿(97-2003)|*.xls|Excel工作簿(2007以上）|*.xlsx";
	             fileDialog.InitialDirectory = @"C:\";
	              
	             if (fileDialog .ShowDialog () == DialogResult.OK )
	                 {
	             	
	             	label1 .Text =fileDialog .FileName ;
	             		
	             }
	             if (AboutFile.isFileLocked (label1 .Text .Trim ()))
	             {
	             	MessageBox .Show ("请先关闭该文件");
	             	label1 .Text ="";
	             	return ;
	             }
		
		}
		
		void Button2Click(object sender, EventArgs e)
		{
			 
			if (label1 .Text .Trim () == "")
			{
				MessageBox .Show ("请先选择文件");
				return ;
			
			}
			
			
			string str1 = "";
			FolderBrowserDialog folder = new FolderBrowserDialog ();
			folder .Description ="选择要保存的文件夹";
			if (folder .ShowDialog() == DialogResult.OK)
			{
				str1 =folder.SelectedPath ;
			}
			
			if (Directory.GetDirectories(str1 + "\\").Length > 0 || Directory.GetFiles(str1+ "\\").Length > 0)
			{
				MessageBox .Show ("文件夹不为空，请选择一个空文件夹");
				return;
			}
					
			DataHelper helper =new DataHelper (label1 .Text.Trim () );
	       helper.open ();
	       string str2 = "select distinct name from [sheet1$]";
	       OleDbDataReader reader = helper .dataread (str2);
	       while(reader .Read ())
	       {      	
	       	DataHelper helper1 = new DataHelper(label1 .Text .Trim());
	       	helper1.open ();
	       	string str3 ="select * from [sheet1$] where name = '"+ reader[0].ToString()+"' order by date ";
	       
	       	OleDbDataReader reader1= helper1.dataread (str3);
	        	
	       	HSSFWorkbook workbook =new HSSFWorkbook();
	       	ISheet sheet = workbook.CreateSheet(reader[0].ToString ().Trim () );
	      	
	       sheet.SetColumnWidth(0,  (int)((5.83+0.78) *256));
           sheet.SetColumnWidth(1,  (int)((55.75+0.78) *256));
           sheet.SetColumnWidth(2,  (int)((7.15+0.68) *256));
           sheet.SetColumnWidth(3,  (int)((13.83+0.78) *256));
           sheet.SetColumnWidth(4,  (int)((15.63+0.78) *256));
           sheet.SetColumnWidth(5,  (int)((13.88+0.78) *256));
           sheet.SetColumnWidth(6,  (int)((16.13+0.78) *256));
           sheet.SetColumnWidth(7,  (int)((13.38+0.78) *256));   
	       	
	       IRow row = sheet.CreateRow(0);
           ICell cell=row.CreateCell(0);	
            row.Height =41*20;
            cell.SetCellValue("销售单");
           sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, 7));
           ICellStyle style=workbook .CreateCellStyle ();
           style.VerticalAlignment =VerticalAlignment .CENTER ;
           style.Alignment = NPOI.SS.UserModel.HorizontalAlignment .CENTER ;
           IFont font= workbook .CreateFont();
           font.FontName ="宋体";
           font.Boldweight =short.MaxValue ;
           font.FontHeightInPoints =16;
           style.SetFont (font);
           row.GetCell(0).CellStyle =style ;
           
           
                     
           ICellStyle style1=workbook.CreateCellStyle();
           style1.VerticalAlignment =VerticalAlignment .CENTER ;
           style1.Alignment = NPOI.SS.UserModel.HorizontalAlignment .CENTER ;
           IFont font1=workbook .CreateFont ();
           font1.Boldweight =short.MaxValue;
           font1.FontHeightInPoints =12;
           font1.FontName ="宋体";
           style1.BorderBottom = NPOI.SS.UserModel.BorderStyle.THIN;
           style1.BorderLeft =NPOI.SS.UserModel.BorderStyle.THIN;
           style1.BorderRight =NPOI.SS.UserModel.BorderStyle.THIN;
           style1.BorderTop =NPOI.SS.UserModel.BorderStyle.THIN;
           style1 .SetFont (font1);
           
           ICellStyle style2=workbook.CreateCellStyle();
           style2.VerticalAlignment =VerticalAlignment .CENTER ;
           style2.Alignment = NPOI.SS.UserModel.HorizontalAlignment .CENTER ;
           IFont font2=workbook .CreateFont ();
           font2.FontHeightInPoints =12;
           font2.FontName ="宋体";
           style2.BorderBottom = NPOI.SS.UserModel.BorderStyle.THIN;
           style2.BorderLeft =NPOI.SS.UserModel.BorderStyle.THIN;
           style2.BorderRight =NPOI.SS.UserModel.BorderStyle.THIN;
           style2.BorderTop =NPOI.SS.UserModel.BorderStyle.THIN;
           style2 .SetFont (font2);
           
           
           ICellStyle style3=workbook.CreateCellStyle();
           style3.VerticalAlignment =VerticalAlignment.CENTER ;
           style3.Alignment = NPOI.SS.UserModel.HorizontalAlignment.LEFT ;
           IFont font3=workbook .CreateFont ();
           font3.FontHeightInPoints =13;
           font3.FontName ="黑体";
           style3 .SetFont (font3);
           
           IRow kehu =sheet.CreateRow (1);
           kehu.Height =23*20;
           sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, 7));
           ICell cellname = kehu.CreateCell (0);
           cellname .SetCellValue(" 客户姓名："+ reader[0]);
           cellname .CellStyle =style3 ;
           
           
           IRow row1 = sheet.CreateRow(2);
           row1.Height =23*20;
           string[] strarr={"序号","内容","数量","单价","金额","是否付款","日期","备注"};
           for(int i=0;i<=7;i++){
           ICell cell1=row1.CreateCell(i);
           cell1.SetCellValue(strarr[i]);
           cell1.CellStyle =style1 ;      
                                 }
           
           
           
           int n=0; int sl = 0; decimal hj=0; decimal jizhang=0; decimal yifukuan= 0;
            while(reader1.Read ()){
           	IRow row2=sheet.CreateRow(n + 3);
           	row2.Height =21*20 ;
           	int a1 = 0; decimal  a2 = 0; decimal  a3 = 0;
           	
           	if (reader1[2].ToString ().Trim ()=="")
           	  {a1=0;}
           	else
           	   {a1=Convert.ToInt32(reader1[2].ToString ().Trim ());}
           	
           	if (reader1[4].ToString ().Trim ()=="")
           	    {a2=0;a3=0;}
           	else 
           	{
           	  if(reader1 [5].ToString ().Trim ()=="记账")
           	     { a2=Convert.ToDecimal(reader1[4].ToString ().Trim ());}
           	  else
                {a3=Convert.ToDecimal (reader1[4].ToString ().Trim ());}
           	}
           		
           	
           	sl =sl + a1;
           	jizhang  = jizhang  + a2;
           	yifukuan =yifukuan + a3;
           	
           	//将一条记录内容放入数组
                       string[] strarray=new string[8];
                       strarray[0]=(n+1).ToString ().Trim ();
                       strarray [1]=reader1[1].ToString ().Trim ();
                       strarray [2]=reader1[2].ToString ().Trim ();
                       strarray [3]=string.Format ("{0:#0.00}",reader1[3]);
                       strarray [4]=string.Format ("{0:#0.00}",reader1[4]);
                       strarray [5]=reader1[5].ToString ().Trim ();
                       //string[] c=reader1[6].ToString ().Split (' ');
                       strarray [6]=String.Format("{0:yyyy-MM-dd}",reader1[6]);;
                        strarray [7]=reader1[7].ToString ().Trim ();
                  //在每一行建立7个单元格，赋值，并指定样式
                  for (int a=0;a<=7;a++){
                     	ICell cell5=row2.CreateCell(a);
                     	cell5.SetCellValue(strarray [a]);
                     	cell5.CellStyle =style2;
                      	
                  }
                  
                  n++;
                 
                  }
                
                 hj= jizhang + yifukuan ;
                IRow row7=sheet.CreateRow(n +3);
           	    row7.Height =21*20 ;
           	    ICell cell10 = row7.CreateCell (0); cell10.CellStyle =style2;
           	    ICell cell11 = row7.CreateCell (1);cell11.SetCellValue("合计");cell11.CellStyle =style1;
           	    ICell cell12 = row7.CreateCell (2);cell12.SetCellValue(sl.ToString ()); cell12.CellStyle =style1;
           	    ICell cell13 = row7.CreateCell (3);cell13.CellStyle =style2;
           	    ICell cell14 = row7.CreateCell (4);cell14.SetCellValue(string.Format ("{0:#0.00}",hj));cell14.CellStyle =style1;
           	    ICell cell15 = row7.CreateCell (5); cell15.CellStyle =style2;
           	    ICell cell16 = row7.CreateCell (6); cell16.CellStyle =style2;
           	    ICell cell17 = row7.CreateCell (7); cell17.CellStyle =style2;
           	    
           	    n++;
           	              	    
           	    IRow row8=sheet.CreateRow(n +3);
           	    row8.Height =21*20 ;
           	    ICell cell20 = row8.CreateCell (0); cell20.CellStyle =style2;
           	    ICell cell21 = row8.CreateCell (1); cell21.SetCellValue("其中:记  账");cell21.CellStyle =style2;
           	    ICell cell22 = row8.CreateCell (2); cell22.CellStyle =style2;
           	    ICell cell23 = row8.CreateCell (3); cell23.CellStyle =style2;
           	    ICell cell24 = row8.CreateCell (4); cell24.CellStyle =style2;
           	    if (jizhang  != 0)
           	    {cell24.SetCellValue(string.Format ("{0:#0.00}",jizhang ));}
           	    else
           	    {cell24.SetCellValue("");}
           	    
           	    ICell cell25 = row8.CreateCell (5); cell25.CellStyle =style2;
           	    ICell cell26 = row8.CreateCell (6); cell26.CellStyle =style2;
           	    ICell cell27 = row8.CreateCell (7); cell27.CellStyle =style2;
           	    
           	    
                n++;
           	    IRow row9=sheet.CreateRow(n +3);
           	    row9.Height =21*20 ;
           	    ICell cell30 = row9.CreateCell (0); cell30.CellStyle =style2;
           	    ICell cell31 = row9.CreateCell (1); cell31.SetCellValue("      已付款");cell31.CellStyle =style2;
           	    ICell cell32 = row9.CreateCell (2); cell32.CellStyle =style2;
           	    ICell cell33 = row9.CreateCell (3); cell33.CellStyle =style2;
           	    ICell cell34 = row9.CreateCell (4); cell34.CellStyle =style2;
           	    if (yifukuan != 0)
           	    {cell34.SetCellValue(string.Format ("{0:#0.00}",yifukuan));}
           	    else
           	    {cell34.SetCellValue("");}
           	    ICell cell35 = row9.CreateCell (5); cell35.CellStyle =style2;
           	    ICell cell36 = row9.CreateCell (6); cell36.CellStyle =style2;
           	    ICell cell37 = row9.CreateCell (7); cell37.CellStyle =style2;
           	    
           	    FileStream fs = File.OpenWrite(str1+ "\\" + reader[0].ToString ().Trim () + ".xls");
		             workbook .Write(fs);
		                fs.Close ();
           
		                helper1.close ();
           
                           	
	       }
	       
	       helper .close ();
	       	
	       	MessageBox .Show("已成功生成完毕,文件位于文件夹\n" + str1 + "\\" );
	    }
		
		
		
		
		
		
	}
	
	
	
	
	
}
