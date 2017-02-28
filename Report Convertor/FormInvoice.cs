/*
 * Created by SharpDevelop.
 * User: zhangyuqing3
 * Date: 2017-2-22
 * Time: 17:58
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
	public partial class frmInvoice : Form
	{
		public static OleDbConnection conn;
		public static string strCon;
		public static string inputPath;			
		public static DataSet ds = new DataSet();
		public static string sql;
		
		public frmInvoice(Report_Convertor.frmContainer parent)
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
			// Read Invoice Table
			//   	
      	
        	UpdateDataSet();       

			foreach (DataRow dr in ds.Tables["BPO"].Rows)
			{
        		cbBpoNumber.Items.Add(dr["BPO Number"].ToString());
			}
			
			cbSubmitMonth.SelectedIndex = 0;
			cbCurrency.SelectedIndex = 0;
			cbSupplier.SelectedIndex = 0;
			cbBpoCategory.SelectedIndex = 0;
			cbBpoNumber.SelectedIndex = 0;
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
        		if ( OpenDB() )
				{
        			sql = "select [Submit month],[Invoice date],[Invoice Number],[Invoice Value],[Currency],[Supplier]," +
        			"[BPO Category],[BPO Number],[BPO Category Details] from [Invoice$]";
        			
					OleDbDataAdapter da = new OleDbDataAdapter(sql, strCon);        			
        			da.Fill(ds, "Invoice");  
        		
        			this.dgInvoice.DataMember = "Invoice";
        			this.dgInvoice.DataSource = ds;
        		
        			sql = String.Format("select [Supplier],[BPO Category],[BPO Number],[Amount],[BPO date],[BPO start date]," +
				                    "[BPO end date] from [BPO$]");
        			OleDbDataAdapter da1 = new OleDbDataAdapter(sql, strCon);     
        			da1.Fill(ds, "BPO");
        			
        			foreach (DataRow dr in ds.Tables["BPO"].Rows)
					{
        				cbBpoNumber.Items.Add(dr["BPO Number"].ToString());
					}
  			
        			CloseDB();
        		}   
            }
		}
		
		//insert
		void BtInsertClick(object sender, System.EventArgs e)
		{	
			
			
			string submitMonth = this.cbSubmitMonth.Text;
			string invoiceDate = this.dtInvoiceDate.Value.Date.ToShortDateString();
			string invoiceNumber = tbInvoiceNumber.Text;
			string invoiceAmount = tbInvoiceAmount.Text;
			string currency = cbCurrency.Text;
			string supplier = this.cbSupplier.Text;
			string bpoCategory = this.cbBpoCategory.Text;
			string bpoNumber = cbBpoNumber.Text;
			string bpoCategoryDetails = tbBpoCategoryDetails.Text;
			
			
			if (submitMonth == "" || invoiceDate == "" || invoiceNumber == "" || invoiceAmount == "" ||
			    currency == "" || supplier == "" || bpoCategory == "" || bpoNumber == "")
			{
				MessageBox.Show("Please fill in data. Database not updated.");
				return;
			}
			if (bpoCategoryDetails == "")
			{
				bpoCategoryDetails = "NA";
			}
			
			if (KeyExist() == true)
			{
				MessageBox.Show("Invoice has been exist in the database!");
				return;	
			}
		
			sql = String.Format("insert into [Invoice$]([Submit month],[Invoice date],[Invoice Number],[Invoice Value],[Currency],[Supplier]," +
			                    "[BPO Category],[BPO Number],[BPO Category Details]) values ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}')",
			                     submitMonth, invoiceDate, invoiceNumber, invoiceAmount, currency, supplier, bpoCategory, bpoNumber, bpoCategoryDetails);
			if ( OpenDB() )
			{
				OleDbCommand cmd = new OleDbCommand(sql, conn);
				cmd.ExecuteNonQuery();
        		CloseDB();
        		UpdateDataSet();
        		MessageBox.Show("Insert success!");
        	}        
		}
		
		void BtUpdateClick(object sender, EventArgs e)
		{
			string submitMonth = this.cbSubmitMonth.Text;
			string invoiceDate = this.dtInvoiceDate.Value.Date.ToShortDateString();
			string invoiceNumber = tbInvoiceNumber.Text;
			string invoiceAmount = tbInvoiceAmount.Text;
			string currency = cbCurrency.Text;
			string supplier = this.cbSupplier.Text;
			string bpoCategory = this.cbBpoCategory.Text;
			string bpoNumber = cbBpoNumber.Text;	
			string bpoCategoryDetails = tbBpoCategoryDetails.Text;	

			
			if (KeyExist() == false)
			{
				MessageBox.Show("Invoice is not found in the database!");
				return;
			}
			
			if (submitMonth == "" || invoiceDate == "" || invoiceNumber == "" || invoiceAmount == "" ||
			    currency == "" || supplier == "" || bpoCategory == "" || bpoNumber == "")
			{
				MessageBox.Show("Please fill in data. Database not updated.");
				return;
			}
			if (bpoCategoryDetails == "")
			{
				bpoCategoryDetails = "NA";
			}

			sql = String.Format("update [Invoice$] set [Submit month]='{0}',[Invoice date]='{1}',[Invoice Number]='{2}'," +
			                    "[Invoice Value]='{3}',[Currency]='{4}',[Supplier]='{5}',[BPO Category]='{6}',[BPO Number]='{7}'," +
			                    "[BPO Category Details]='{8}' where [Invoice Number]='{9}'",
			                    submitMonth, invoiceDate, invoiceNumber, invoiceAmount, currency, supplier, bpoCategory, bpoNumber, bpoCategoryDetails, invoiceNumber);
			if ( OpenDB() )
			{
				OleDbCommand cmd = new OleDbCommand(sql, conn);
				cmd.ExecuteNonQuery();
        		CloseDB();
        		UpdateDataSet();
        		MessageBox.Show("Update success!");
        	}        
			
		}		
		
		void TbInvoiceNumberLeave(object sender, EventArgs e)
		{
			if (tbInvoiceNumber.Text == "")
			{
				return;
			}	

        	bool foundFlag = false;
			foreach (DataRow dr in ds.Tables["Invoice"].Rows)
			{
				if (tbInvoiceNumber.Text == dr["Invoice Number"].ToString()) 
				{
					this.cbSubmitMonth.Text = dr["Submit month"].ToString();
					this.dtInvoiceDate.Value = Convert.ToDateTime(dr["Invoice date"].ToString());
					tbInvoiceNumber.Text = dr["Invoice Number"].ToString();
					tbInvoiceAmount.Text = dr["Invoice Value"].ToString();
					cbCurrency.Text = dr["Currency"].ToString();
					this.cbSupplier.Text = dr["Supplier"].ToString();
					this.cbBpoCategory.Text = dr["BPO Category"].ToString();
					cbBpoNumber.Text = dr["BPO Number"].ToString();	
					tbBpoCategoryDetails.Text = dr["BPO Category Details"].ToString();		

					foundFlag = true;										
				}				
			} 
        	
			if ( foundFlag == true && File.Exists(inputPath))
            {
				ds.Clear();

				string invoiceNumber = tbInvoiceNumber.Text;
        		sql = String.Format("select [Submit month],[Invoice date],[Invoice Number],[Invoice Value],[Currency],[Supplier]," +
        				"[BPO Category],[BPO Number],[BPO Category Details] from [Invoice$] where [Invoice Number]='{0}'", invoiceNumber);
        		if ( OpenDB() )
				{
        			OleDbDataAdapter da = new OleDbDataAdapter(sql, strCon);     
        			da.Fill(ds, "Invoice");
  			
        			CloseDB();
        		}   
        		
        		this.dgInvoice.DataMember = "Invoice";
        		this.dgInvoice.DataSource = ds;
            }
		}
		
		void cbBpoNumberLeave(object sender, EventArgs e)
		{
			if (cbBpoNumber.Text == "")
			{
				return;
			}
			
			bool foundFlag = false;
			foreach (DataRow dr in ds.Tables["BPO"].Rows)
			{
				if (cbBpoNumber.Text == dr["BPO Number"].ToString()) 
				{
					this.cbSupplier.Text = dr["Supplier"].ToString();
					this.cbBpoCategory.Text = dr["BPO Category"].ToString();	
			
					foundFlag = true;					
				}
			}

			if (foundFlag == false)
			{
				MessageBox.Show("BPO not found!");
			}
		}
		
		bool KeyExist()
		{ 
			foreach (DataRow dr in ds.Tables["Invoice"].Rows)
			{
				if (tbInvoiceNumber.Text == dr["Invoice Number"].ToString()) 
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
