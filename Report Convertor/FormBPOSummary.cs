/*
 * Created by SharpDevelop.
 * User: zhangyuqing3
 * Date: 2017-2-22
 * Time: 18:21
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
	/// Description of Form5.
	/// </summary>
	public partial class frmSummary : Form
	{
		public static OleDbConnection conn;
		public static string strCon;
		public static string inputPath;			
		public static DataSet ds = new DataSet();
		public static string sql;
		
		public frmSummary(Report_Convertor.frmContainer parent)
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
			this.MdiParent = parent;
			Utils.ModifyTypeGuessRows( false );			
			
			inputPath = Utils.SetInputPath(System.Environment.CurrentDirectory + "\\Data\\BPO\\BPO");
			
			//if ( Office2010Exists() && inputPath.Contains("xlsx") )
			if ( Utils.Office2010Exists() && inputPath.Contains("xlsx") )
			{
				strCon = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source=" + 
										inputPath + ";Extended Properties=\"Excel 12.0; HDR=YES; IMEX=10\"";
			}
			else if ( Utils.Office2007Exists() && inputPath.Contains("xlsx") )
			{
				strCon = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source=" + 
										inputPath + ";Extended Properties=\"Excel 12.0; HDR=YES; IMEX=10\"";
			}
			else
			{
//				strCon = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source =" +
//										inputPath + ";Extended Properties=\"Excel 8.0; HDR=Yes; IMEX=10\"";
				strCon = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source=" + 
										inputPath + ";Extended Properties=\"Excel 12.0; HDR=YES; IMEX=10\"";
			}	
			
			/// 
			/// </summary>
			/// <returns></returns>
			//   	
      	
        	UpdateDataSet(); 
			CalcSummary();        	
		}
		
		public bool OpenDB()
		{	
			conn = new OleDbConnection(strCon);
        	try
        	{
        		if (conn.State == ConnectionState.Closed)
        		{
        			conn.Open();
        			return true;
        		}
        	}        	
        	catch//(OleDbException E)
        	{
        		string text = "Error occurred when reading input file. Please check the sheet and attribute name.\n\n" +
        					  "Input file: " + inputPath + "\n";
        		MessageBox.Show(text);
        		//throw new Exception(E.Message);
        		return false;
        	}	
        	
        	return false;
		}
			
		public void CloseDB()
		{
			if (conn.State == ConnectionState.Open)
			{
 				conn.Close();
			}
		}
		
		
		public void UpdateDataSet()
		{
			ds.Clear();
        	
			if ( File.Exists(inputPath))
            {
        		if ( OpenDB() )
				{
        			sql = "select * from [Summary$]";
        			
					OleDbDataAdapter da = new OleDbDataAdapter(sql, strCon);        			
        			da.Fill(ds, "Summary");  
        		
        			this.dgSummary.DataMember = "Summary";
        			this.dgSummary.DataSource = ds;
        			
        			sql = "select [Submit month],[Invoice date],[Invoice Number],[Invoice Value],[Currency],[Supplier]," +
        			"[BPO Category],[BPO Number],[BPO Category Details] from [Invoice$]";
        			
					OleDbDataAdapter da1 = new OleDbDataAdapter(sql, strCon);        			
        			da1.Fill(ds, "Invoice");  
        		
        			sql = String.Format("select [Supplier],[BPO Category],[BPO Number],[Amount],[BPO date],[BPO start date]," +
				                    "[BPO end date] from [BPO$]");
        			OleDbDataAdapter da2 = new OleDbDataAdapter(sql, strCon);     
        			da2.Fill(ds, "BPO");
  			
        			CloseDB();
        		}   
            }
		}
		
		public void CalcSummary()
		{	
			foreach (DataRow drBPO in ds.Tables["BPO"].Rows)
			{
				DataRow drSummary = ds.Tables["Summary"].NewRow();

				drSummary["Supplier"] = drBPO["Supplier"];
				drSummary["BPO Number"] = drBPO["BPO Number"];
				drSummary["BPO Category"] = drBPO["BPO Category"];
				drSummary["BPO start date"] = drBPO["BPO start date"];
				drSummary["BPO end date"] = drBPO["BPO end date"];
				drSummary["Amount"] = drBPO["Amount"];
				
				int invoicedValue = 0;
				int balance = 0;
				foreach (DataRow drInvoice in ds.Tables["Invoice"].Rows)
				{					
					if (drBPO["BPO Number"].ToString() == drInvoice["BPO Number"].ToString() &&
					   drBPO["Supplier"].ToString() == drInvoice["Supplier"].ToString() &&
					   drBPO["BPO Category"].ToString() == drInvoice["BPO Category"].ToString())
					{
						invoicedValue += int.Parse(drInvoice["Invoice Value"].ToString());
					}
				}
				
				drSummary["Invoiced Value"] = invoicedValue.ToString();
				balance = int.Parse(drBPO["Amount"].ToString()) - invoicedValue;
				drSummary["Balance"] = balance.ToString();
				
				ds.Tables["Summary"].Rows.Add(drSummary);	
			}
		}
	}
}
