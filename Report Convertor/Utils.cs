/*
 * Created by SharpDevelop.
 * User: zhangyuqing3
 * Date: 2017-2-28
 * Time: 9:29
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Drawing;
using System.Windows.Forms;
using System.Data;
using System.Data.Sql;
using System.Data.OleDb;
using System.IO;
using Microsoft.Win32;

namespace Report_Convertor
{
	/// <summary>
	/// Description of Class1.
	/// </summary>
	public class Utils
	{
		public Utils()
		{
			
		}
					/// <summary>
		/// //////////////////////////////////////////////////////////////////////////////////////////
		/// </summary>
		/// <param name="inputPath"></param>
		/// <returns></returns>
		
		public static string SetInputPath(string inputPath)
		{
			string ret = "";
			if (File.Exists(inputPath + ".xls"))
			{
				ret = inputPath + ".xls";
			}
			else
			{
				ret = inputPath + ".xlsx";
			}
			
			return ret;
		}
		
		public static bool Office2003Exists()
        {
            bool exist = false;
            RegistryKey rk = Registry.LocalMachine;
            RegistryKey akey = rk.OpenSubKey(@"SOFTWARE\\Microsoft\\Office\\11.0\\Word\\InstallRoot\\");
            RegistryKey akeytwo = rk.OpenSubKey(@"SOFTWARE\\Microsoft\\Office\\12.0\\Word\\InstallRoot\\");
            //检查本机是否安装Office2003
            if (akey != null)
            {
                string file03 = akey.GetValue("Path").ToString();
                if (File.Exists(file03 + "Excel.exe"))
                {
                    exist = true;
                }
            }
            
            return exist;
		}
		
		public static bool Office2007Exists()
        {
            bool exist = false;
            RegistryKey rk = Registry.LocalMachine;
            RegistryKey akey = rk.OpenSubKey(@"SOFTWARE\\Microsoft\\Office\\12.0\\Word\\InstallRoot\\");
            //检查本机是否安装Office2007
            if (akey != null)
            {
                string file07 = akey.GetValue("Path").ToString();
                if (File.Exists(file07 + "Excel.exe"))
                {
                    exist = true;
                }
            }
            
            return exist;
        }
		
		public static bool Office2010Exists()
        {
            bool exist = false;
            RegistryKey rk = Registry.LocalMachine;
            RegistryKey akey = rk.OpenSubKey(@"SOFTWARE\\Microsoft\\Office\\14.0\\Word\\InstallRoot\\");
            //检查本机是否安装Office2007
            if (akey != null)
            {
                string file10 = akey.GetValue("Path").ToString();
                if (File.Exists(file10 + "Excel.exe"))
                {
                    exist = true;
                }
            }
            
            return exist;
        }
		public static bool Office2016Exists()
        {
            bool exist = false;
            RegistryKey rk = Registry.LocalMachine;
            RegistryKey akey = rk.OpenSubKey(@"SOFTWARE\\Microsoft\\Office\\15.0\\Word\\InstallRoot\\");
            //检查本机是否安装Office2007
            if (akey != null)
            {
                string file16 = akey.GetValue("Path").ToString();
                if (File.Exists(file16 + "Excel.exe"))
                {
                    exist = true;
                }
            }
            
            return exist;
        }
		
		/// <summary>
		/// 修改注册表TypeGuessRows的值
		/// 
		/// ADO.NET读取Excel表格时，OLEDB（Excel 2000-2003一般是是Jet 4.0，Excel 2007是ACE 12.0，
		/// 即Access Connectivity Engine，ACE也可以用来访问Excel 2000-2003）。会默认扫面Sheet中的
		/// 前几行来决定数据类型，这个行数是由注册表中
		/// Excel 2000-2003 : HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Jet\4.0\Engines\Excel
		/// Excel 2007 : HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\12.0\Access Connectivity Engine\Engines\Excel
		/// 中的TypeGuessRows值来控制，默认是8。 
		/// 在执行Excel读取之前将TypeGuessRows值设为0，那样Jet就会扫描最多16384行。
		/// 当然，如果文件太大的话，这里就有效率问题了。
		/// 采用这个方案一般还要在Excel文件连接字符串中的Extended Properties加入IMEX=1
		/// 当 IMEX=0 时为“汇出模式”，这个模式开启的 Excel 档案只能用来做“写入”用途。
		/// 当 IMEX=1 时为“汇入模式”，这个模式开启的 Excel 档案只能用来做“读取”用途。
		/// 当 IMEX=2 时为“连結模式”，这个模式开启的 Excel 档案可同时支援“读取”与“写入”用途。
		/// </summary>
		/// <param name="setDefault">为true表示设为默认值8，为flase表示设为0</param>
		/// <returns></returns>
		public static bool ModifyTypeGuessRows(bool setDefault)
		{
		    int toSetValue = 0;
			string JetKeyRoot = "SOFTWARE\\Microsoft\\Jet\\4.0\\Engines\\Excel";
			string ACEKeyRoot = "SOFTWARE\\Microsoft\\Office\\12.0\\Access Connectivity Engine\\Engines\\Excel";
			string ACEKeyRoot10 = "SOFTWARE\\Microsoft\\Office\\14.0\\Access Connectivity Engine\\Engines\\Excel";
			string ACEKeyRoot16 = "SOFTWARE\\Microsoft\\Office\\15.0\\Access Connectivity Engine\\Engines\\Excel";
	  		
			if(setDefault)
		    {
		        toSetValue = 8;
		    }
		  
		    try
		    {
		    	if ( Office2003Exists() )
		    	{
		        	RegistryKey JetRegKey = Registry.LocalMachine.OpenSubKey(JetKeyRoot, true);
		        	JetRegKey.SetValue("TypeGuessRows", Convert.ToString(toSetValue, 16), RegistryValueKind.DWord);
		        	JetRegKey.Close();
		    	}
		    	if ( Office2007Exists() )
		    	{
		       		RegistryKey ACERegKey = Registry.LocalMachine.OpenSubKey(ACEKeyRoot, true);
		        	ACERegKey.SetValue("TypeGuessRows", Convert.ToString(toSetValue, 16), RegistryValueKind.DWord);
		        	ACERegKey.Close();
		    	}		   
		    	if ( Office2010Exists() )
		    	{
		       		RegistryKey ACERegKey = Registry.LocalMachine.OpenSubKey(ACEKeyRoot10, true);
		        	ACERegKey.SetValue("TypeGuessRows", Convert.ToString(toSetValue, 16), RegistryValueKind.DWord);
		        	ACERegKey.Close();
		    	}	
		    	if ( Office2016Exists() )
		    	{
		       		RegistryKey ACERegKey = Registry.LocalMachine.OpenSubKey(ACEKeyRoot16, true);
		        	ACERegKey.SetValue("TypeGuessRows", Convert.ToString(toSetValue, 16), RegistryValueKind.DWord);
		        	ACERegKey.Close();
		    	}	
		    }
		    catch(Exception ex)
		    {
//		        MessageBox.Show("Registry update failed："+ ex.Message);
		        return false;
		    }
		    
		    return true;
		}
	}
}
