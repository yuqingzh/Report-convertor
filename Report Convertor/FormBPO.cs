/*
 * Created by SharpDevelop.
 * User: zhangyuqing3
 * Date: 2017-2-22
 * Time: 15:44
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
	public partial class frmBPO : Form
	{
		public static OleDbConnection conn;
		public static string strCon;
		public static string inputPath;			
		public static DataSet ds = new DataSet();
		public static string sql;
		
		public frmBPO(Report_Convertor.frmContainer parent)
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
				strCon = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source =" +
										inputPath + ";Extended Properties=\"Excel 8.0; HDR=Yes; IMEX=10\"";
			}	
			
			
			//
			// Read BPO Table
			//   	
      	
        	UpdateDataSet();
			cbSupplier.SelectedIndex = 0;
			cbBpoCategory.SelectedIndex = 0;	
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
        	//ds.Tables.Clear();         	
        	
			if ( File.Exists(inputPath))
            {
        		sql = "select [Supplier],[BPO Category],[BPO Number],[Amount],[BPO date],[BPO start date],[BPO end date] from [BPO$]";
        		if ( OpenDB() )
				{
        			OleDbDataAdapter da = new OleDbDataAdapter(sql, strCon);     
        			da.Fill(ds, "BPO");  			
        			CloseDB();
        		}   
        		
        		this.dgBPO.DataMember = "BPO";
        		this.dgBPO.DataSource = ds;
            }
		}
		
		//insert
		void BtInsertClick(object sender, System.EventArgs e)
		{	
			string supplier = this.cbSupplier.Text;
			string bpoCategory = this.cbBpoCategory.Text;
			string bpoNumber = tbBpoNumber.Text;
			string amount = this.tbAmount.Text;
			string bpoDate = this.dtBpoDate.Value.Date.ToShortDateString();
			string bpoStartDate = this.dtBpoPeriodStart.Value.Date.ToShortDateString();
			string bpoEndDate = this.dtBpoPeriodEnd.Value.Date.ToShortDateString();

			if (supplier == "" || bpoCategory == "" || bpoNumber == "" || amount == "" ||
			    bpoDate == "" || bpoStartDate == "" || bpoEndDate == "")
			{
				MessageBox.Show("Please fill in data. Database not updated.");
				return;
			}
			
			if (KeyExist() == true)
			{
				MessageBox.Show("BPO has been exist in the database!");
				return;	
			}
			
			sql = String.Format("insert into [BPO$]([Supplier],[BPO Category],[BPO Number],[Amount],[BPO date],[BPO start date]," +
			                    "[BPO end date]) values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}')",
			                     supplier, bpoCategory, bpoNumber, amount, bpoDate, bpoStartDate,bpoEndDate);
			if ( OpenDB() )
			{
				OleDbCommand cmd = new OleDbCommand(sql, conn);
				cmd.ExecuteNonQuery();
        		CloseDB();
        		UpdateDataSet();
        		MessageBox.Show("Insert success!");
        	}        
		}
		
		void TbBpoNumberLeave(object sender, EventArgs e)
		{
			if (tbBpoNumber.Text == "")
			{
				return;
			}
			
			bool foundFlag = false;
			foreach (DataRow dr in ds.Tables["BPO"].Rows)
			{
				if (tbBpoNumber.Text == dr["BPO Number"].ToString()) 
				{
					this.cbSupplier.Text = dr["Supplier"].ToString();
					this.cbBpoCategory.Text = dr["BPO Category"].ToString();
					this.tbAmount.Text = dr["Amount"].ToString();
					this.dtBpoDate.Value = Convert.ToDateTime(dr["BPO date"].ToString());
					this.dtBpoPeriodStart.Value = Convert.ToDateTime(dr["BPO start date"].ToString());
					this.dtBpoPeriodEnd.Value = Convert.ToDateTime(dr["BPO end date"].ToString());

					foundFlag = true;					
				}
				
			}      	
        	
			if ( foundFlag == true && File.Exists(inputPath))
            {
				ds.Clear();
				
				string bpoNumber = tbBpoNumber.Text;
        		sql = String.Format("select [Supplier],[BPO Category],[BPO Number],[Amount],[BPO date],[BPO start date]," +
				                    "[BPO end date] from [BPO$] where [BPO Number]='{0}'", bpoNumber);
        		if ( OpenDB() )
				{
        			OleDbDataAdapter da = new OleDbDataAdapter(sql, strCon);     
        			da.Fill(ds, "BPO");
  			
        			CloseDB();
        		}   
        		
        		this.dgBPO.DataMember = "BPO";
        		this.dgBPO.DataSource = ds;
            }
		}
					
		void BtUpdateClick(object sender, EventArgs e)
		{
			string supplier = this.cbSupplier.Text;
			string bpoCategory = this.cbBpoCategory.Text;
			string bpoNumber = tbBpoNumber.Text;
			string amount = this.tbAmount.Text;
			string bpoDate = this.dtBpoDate.Value.Date.ToShortDateString();
			string bpoStartDate = this.dtBpoPeriodStart.Value.Date.ToShortDateString();
			string bpoEndDate = this.dtBpoPeriodEnd.Value.Date.ToShortDateString();
			
			if (supplier == "" || bpoCategory == "" || bpoNumber == "" || amount == "" ||
			    bpoDate == "" || bpoStartDate == "" || bpoEndDate == "")
			{
				MessageBox.Show("Please fill in data. Database not updated.");
				return;
			}
			
			if (KeyExist() == false)
			{
				MessageBox.Show("BPO is not found in the database!");
				return;
			}

			sql = String.Format("update [BPO$] set [Supplier]='{0}',[BPO Category]='{1}',[BPO Number]='{2}',[Amount]='{3}'," +
			                    "[BPO date]='{4}',[BPO start date]='{5}',[BPO end date]='{6}' where [BPO Number]='{7}'",
			                    supplier, bpoCategory, bpoNumber, amount, bpoDate, bpoStartDate,bpoEndDate, bpoNumber);
			if ( OpenDB() )
			{
				OleDbCommand cmd = new OleDbCommand(sql, conn);
				cmd.ExecuteNonQuery();
        		CloseDB();
        		UpdateDataSet();
        		MessageBox.Show("Update success!");
        	}        
			
		}
		
		
		void BtDeleteClick(object sender, EventArgs e)
		{
//			if (tbBpoNumber.Text == "")
//			{
//				MessageBox.Show("Please fill in BPO number. Database not updated.");
//				return;
//			}
//			string bpoNumber = tbBpoNumber.Text;
//        	sql = String.Format("delete * from [BPO$] where [BPO Number]='{0}'", bpoNumber);
//			if ( OpenDB() )
//			{
//				OleDbCommand cmd = new OleDbCommand(sql, conn);
//				cmd.ExecuteNonQuery();
//        		CloseDB();
//        		UpdateDataSet();
//        		MessageBox.Show("Update success!");
//        	}        
//			
		}
		
		bool KeyExist()
		{ 
			foreach (DataRow dr in ds.Tables["BPO"].Rows)
			{
				if (tbBpoNumber.Text == dr["BPO Number"].ToString()) 
				{					
					return true;
				}				
			}      	
			return false;
		}
		
		void EnableButtons()
		{
			if (KeyExist() == true)
			{
				btInsert.Enabled = false;
				btUpdate.Enabled = true;
			}
			else
			{
				btInsert.Enabled = true;
				btUpdate.Enabled = false;
			}
		}
//		void BtInsertEnter(object sender, EventArgs e)
//		{
//			EnableButtons();
//		}
//		
//		void BtUpdateEnter(object sender, EventArgs e)
//		{
//			EnableButtons();
//		}

	}
}
