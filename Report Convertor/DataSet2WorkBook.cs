/*
 * Created by SharpDevelop.
 * User: yuqingzh
 * Date: 2011-6-17
 * Time: 22:21
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Data;
using System.Windows.Forms;
using System.Data.Sql;
using System.Data.OleDb;
//using Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace Report_Convertor
{
 /// <summary>
 /// 将数据集转换成excel工作簿
 /// </summary>
 
 public class DataSet2WorkBook
 {

  private DataSet mDs = new DataSet(); //存放数据源
  private string mFilePath = System.Environment.CurrentDirectory + "\\Data\\Output.xls"; //excel文件名，保存的路径

  public DataSet2WorkBook(ref DataSet ds , string filePath )
  {
   //
   // TODO: 在此处添加构造函数逻辑
            //
   this.mDs = ds ;
   this.mFilePath = filePath ;
  }
  

  /// <summary>
  /// 将数据表转换成excel工作簿中的sheet
  /// </summary>
  /// <param name="tb">要转换的数据表（引用类型）</param>
  /// <param name="xSheet">目标sheet</param>
  /// <param name="SheetName">sheet名字</param>
  /// <returns></returns>
  private bool DataTable2Sheet( ref System.Data.DataTable tb ,ref Excel._Worksheet xSheet ,string SheetName )
  {
  	try
	{
		int rowIndex=2;
		int colIndex=0;
		if(SheetName == "")
		{
			xSheet.Name = tb.TableName ;
		}
		else
		{
			xSheet.Name = SheetName ;
		}

		foreach(DataColumn tempCol in tb.Columns )
		{
		  
		     xSheet.Cells[1,colIndex+1]=tempCol.ColumnName;
	     
		     rowIndex = 2 ;
		     //if (tempCol.DataType == System.Type.GetType("System.DateTime"))
		     if (tempCol.ColumnName.ToLower().Contains("date"))
		     {
				foreach(DataRow tempRow in tb.Rows )
				{
					if (tempRow[colIndex].ToString().Length != 0)
					{
			//	              	xSheet.Cells[rowIndex, colIndex+1] = (Convert.ToDateTime(tempRow[colIndex].ToString())).ToString("yyyy-MM-dd");
					
			//	  				xSheet.Cells[rowIndex, colIndex+1] = (Convert.ToDateTime(tempRow[colIndex].ToString()).ToOADate());

			//	              	string str, subStr;
			//	              	str = (Convert.ToDateTime(tempRow[colIndex].ToString()).GetDateTimeFormats('r')[0]).ToString();
			//					subStr = str.Substring(5, str.IndexOf(" 00:00:00 GMT") - 5);
			//					xSheet.Cells[rowIndex, colIndex+1] = subStr;
			//			string[] dt = Convert.ToDateTime(tempRow[colIndex].ToString()).GetDateTimeFormats('d');	
						try
						{
							xSheet.Cells[rowIndex, colIndex+1] = Convert.ToDateTime(tempRow[colIndex].ToString()).GetDateTimeFormats('d')[3];				
						}
						catch
						{
							xSheet.Cells[rowIndex ,colIndex+1] = "";
						}
					}
					else
					{
						xSheet.Cells[rowIndex ,colIndex+1] = tempRow[colIndex].ToString();
					}
				
//		     		xSheet.get_Range(xSheet.Cells[rowIndex, colIndex+1], xSheet.Cells[rowIndex, colIndex]).HorizontalAlignment = XlHAlign.xlHAlignCenter; //XlVAlign.xlVAlignLeft;//设置居中对齐 
					rowIndex++ ;
				}
		     }
//		     else if (tempCol.DataType == System.Type.GetType("System.String"))
//		     {
//				foreach(DataRow tempRow in tb.Rows )
//				{
//					xSheet.Cells[rowIndex ,colIndex+1] = tempRow[colIndex].ToString();
////			      	xSheet.get_Range(xSheet.Cells[rowIndex, colIndex+1], xSheet.Cells[rowIndex, colIndex]).HorizontalAlignment = XlHAlign.xlHAlignLeft;  
//					rowIndex++ ;
//				}
//		     }
		     else
		     {
				foreach(DataRow tempRow in tb.Rows )
				{
					xSheet.Cells[rowIndex ,colIndex+1] = tempRow[colIndex].ToString();
//					xSheet.get_Range(xSheet.Cells[rowIndex, colIndex+1], xSheet.Cells[rowIndex, colIndex]).HorizontalAlignment = XlHAlign.xlHAlignLeft;  
					rowIndex++ ;
				}
		     }
			 
		     colIndex++;
	    }
	    
	    return true ;
	}
	catch
	{
		MessageBox.Show("Error Occurred when exporting the sheet " + SheetName + "!");
		return false ;
	}
}


   		
  /// <summary>
  /// 将指定数据集里的表转换成工作簿里sheet
  /// </summary>
  /// <param name="starPos">数据表开始位置从0开始计数</param>
  /// <param name="Count">要转换数据表的数目</param>
  /// <returns>成功返回true</returns>
  public bool ConvertSheet(int starPos ,int Count)
  {
    System.Data .DataTable tempTable ; //创建临时表
    Excel.Application xApp= new Excel.Application();
    xApp.Visible = false ;
    object objOpt = System.Reflection.Missing.Value;
      
    Excel.Workbook xBook = xApp.Workbooks.Add(true) ;//添加新工作簿
    Excel.Sheets xSheets = xBook.Sheets ;
    Excel._Worksheet xSheet = null ;
    
    bool ret = true;
    
   try
   {
    //
    //转换从指定起始位置以后一定数目的数据集
    //
    for(int i = starPos , iCount = 1 ; iCount <= Count && i< this.mDs.Tables.Count ; i++ ,iCount++ )
    {
     tempTable = this.mDs.Tables[i] ;
     //
     //创建空的sheet
     //
     xSheet = (Excel._Worksheet)(xBook.Sheets.Add(objOpt,objOpt,objOpt,objOpt)) ;

     ret = DataTable2Sheet(ref tempTable  ,ref  xSheet ,"") ;
     
     if (ret == false)
     {
     	break;
     }
     
    }

    //
    //获取默认生成的sheet并将其删除
    //
    //Excel._Worksheet tempXSheet = (Excel._Worksheet) (xSheets.get_Item(1)) ;
    //
    Excel._Worksheet tempXSheet = (Excel._Worksheet) (xBook.Worksheets[Count+1]) ;
    tempXSheet.Delete() ;
    System.Runtime.InteropServices.Marshal.ReleaseComObject(tempXSheet) ;
    tempXSheet=null ;
    //
    //保存
    //
    xBook.Saved = true ;
    xBook.SaveCopyAs(this.mFilePath ) ;
    
   }
   catch
   {
    ret = false ;
   }
   finally
   {
   	//
    //释放资源
    //
    System.Runtime.InteropServices.Marshal.ReleaseComObject(xSheet) ;
    xSheet=null ;
    System.Runtime.InteropServices.Marshal.ReleaseComObject(xSheets) ;
    xSheets=null ;
    
    System.Runtime.InteropServices.Marshal.ReleaseComObject(xBook) ;
    xBook=null ;
    xApp.Quit();
    System.Runtime.InteropServices.Marshal.ReleaseComObject(xApp);
    xApp = null ;
    GC.Collect();//强行销毁
   }
   
   return ret;
  }
/// <summary>
/// 重载convert，将数据集里所有的表转换工作簿的sheet
/// </summary>
/// <returns></returns>
  
  public bool ConvertAll()
  {
   return this.ConvertSheet( 0 ,this.mDs.Tables.Count ) ;
  }
  
 }

/// <summary>
 /// WorkBook2DataSet 的摘要说明。将工作簿转换成dataset
 /// </summary>
 public class WorkBook2DataSet
 {
  private string mFilePath = "" ;
  private DataSet mDs = new DataSet() ;
  

  public WorkBook2DataSet(string path , ref DataSet ds)
  {
   //
   // TODO: 在此处添加构造函数逻辑
   //
   this.mDs = ds ;
   this.mFilePath = path ;

  }

/// <summary>
/// 将工作簿中指定的sheet转换成dataset中的表
/// </summary>
/// <param name="pos">sheet在工作簿中的位置</param>
/// <returns>成功返回true</returns>
  public bool Convert(int pos)
  {
   bool r = false ;
   string strSql = "" ;
   string sheetName = "" ;
   System.Data.DataTable tTable;
   OleDbDataAdapter objDa ;
   //
   //创建excel进程
   //
   object obj = System.Reflection.Missing.Value;
   Excel.ApplicationClass xxApp= new Excel.ApplicationClass() ;//.Application();
   Excel.Workbook  xxBook =null ;
   Excel._Worksheet xxSheet =null ;
   
   try
   {

    //
    //打开excel文件，并获取指定sheet的名字
    //
    xxBook =  xxApp.Workbooks.Open(this.mFilePath ,obj ,obj,obj,obj,obj,obj,obj,obj,obj,obj,obj,obj,obj,obj) ;//添加新工作簿
    xxSheet = (Excel._Worksheet) (xxBook.Worksheets[pos]) ;
    sheetName =xxSheet.Name.ToString() ;
    //
    //释放excel资源
    //
    System.Runtime.InteropServices.Marshal.ReleaseComObject(xxSheet) ;
    xxSheet=null ;
    GC.Collect() ;
    xxBook.Close(false,obj,obj) ;
    System.Runtime.InteropServices.Marshal.ReleaseComObject(xxBook) ;
    xxBook=null ;
    xxApp.Quit();
    System.Runtime.InteropServices.Marshal.ReleaseComObject(xxApp);
    xxApp = null ;
    //
    //创建数据连接
    //
//    OleDbConnection objConn = new OleDbConnection(
//     "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="+ this.mFilePath +";Extended Properties=Excel 8.0;");
	OleDbConnection objConn = new OleDbConnection(
     "Provider=Microsoft.Jet.OLEDB.12.0;Data Source="+ this.mFilePath +";Extended Properties=Excel 12.0;");
      
    //
    //获取工作簿中的表
    //
    strSql = "select * from [" + sheetName +"$]" ;
    tTable = new System.Data.DataTable( sheetName ) ;
    //
    //将sheet填入table中
    //
    objDa = new OleDbDataAdapter(strSql ,objConn) ;
    objDa.Fill(tTable) ;
    this.mDs.Tables.Add(tTable) ;
    //
    //摧毁连接
    //
    objConn.Dispose() ;
    r = true ;
    
   }
   catch
   {
    r = false ;
   }

   GC.Collect() ;
   return r ;
  }
/// <summary>
/// 转换工作簿中所有的sheet到dataset
/// </summary>
/// <returns></returns>
  public bool Convert()
  { 
   bool r = false ; //返回值
   //
   //创建excel进程
   //
   object obj = System.Reflection.Missing.Value;
   Excel.Application xApp= new Excel.Application();
   xApp.Visible = false ;
   Excel.Workbook xBook = xApp.Workbooks.Open(this.mFilePath ,false ,false,obj,obj,obj,obj,obj,obj,obj,obj,obj,obj,obj,obj) ;//

   int count = xBook.Sheets.Count ;
   //
   //释放资源
   //
   xBook.Close(false , this.mFilePath ,obj) ;
   System.Runtime.InteropServices.Marshal.ReleaseComObject(xBook) ;
   xBook=null ;
   xApp.Quit() ;
   System.Runtime.InteropServices.Marshal.ReleaseComObject(xApp);
   xApp = null ;
   GC.Collect() ;
   for(int i = 1 ; i <= count ; i++)
   {
    r = Convert(i) ;
   }
    
   return r ;
   //return this.Convert(1,count) ;
   }
}
 
}