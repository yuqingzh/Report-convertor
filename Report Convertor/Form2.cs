/*
 * Created by SharpDevelop.
 * User: yuqingzh
 * Date: 2011-6-16
 * Time: 16:18
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

namespace Report_Convertor
{
	/// <summary>
	/// Description of Form2.
	/// </summary>
	public partial class frmOutput : Form
	{
		public static DataSet ds = new DataSet();
		public static OleDbDataAdapter da;
		public static OleDbCommandBuilder cb;
		public static string OutputPath;
		public static string strCon;		
		public static OleDbConnection conn;
		
		public static bool OutputTemplateExist = false;	
		
		public frmOutput(Report_Convertor.frmContainer parent)
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
			this.MdiParent = parent;
			ReadAllOutput();
		}
		
		void ReadOutput(string outputPath, string sql, string tableName)
		{			
			strCon = " Provider = Microsoft.Jet.OLEDB.4.0 ; Data Source =" +
										outputPath + ";Extended Properties=Excel 8.0";
			conn = new OleDbConnection(strCon);
        	try
        	{
        		if (conn.State == ConnectionState.Closed)
        		{
        			conn.Open();
        		}
        		da = new OleDbDataAdapter(sql, strCon);
        		cb = new OleDbCommandBuilder(da);
        		da.Fill(ds, tableName);    
        	}        	
        	catch//(OleDbException E)
        	{
        		string text = "Error occurred when reading output template file. Please check the sheet and attribute name.\n\n" +
        					  "Output file: " + outputPath + "\n\n" +
        					  "SQL: " + sql + "\n";
        		MessageBox.Show(text);
        		//throw new Exception(E.Message);
        	}
			finally
			{
				if (conn.State == ConnectionState.Open)
				{
 					conn.Close();
				}
			} 	
		}
		
		void ReadAllOutput()
		{			
        	string sql;
        	
        	// OpenLog
        	OutputPath = System.Environment.CurrentDirectory + "\\Data\\Output_Template.xls";
        	
        	if (!(File.Exists(OutputPath)))
            { 
        		OutputTemplateExist = false;
        		MessageBox.Show( "Output_Template.xls file does not exist!" );
        		return;
        	}
        	else
        	{
        		OutputTemplateExist = true;
        	}
        	
        	// 
        	sql = "select * from [OTD$]";
        	this.ReadOutput(OutputPath, sql, "OTD");	
        	dgOTD.DataMember = "OTD";
        	dgOTD.DataSource = ds;    
        			
        	// Onelog
        	sql = "select * from [Onelog$]";
        	this.ReadOutput(OutputPath, sql, "Onelog");	
        	dgOpenLog.DataMember = "Onelog";
        	dgOpenLog.DataSource = ds;     	

			// NonOneLog
        	sql = "select * from [Non-Onelog$]";
        	this.ReadOutput(OutputPath, sql, "NonOneLog");	
        	dgNonOnelog.DataMember = "NonOneLog";
        	dgNonOnelog.DataSource = ds;  

			// NonDesp
        	sql = "select * from [Non-Desp$]";
        	this.ReadOutput(OutputPath, sql, "NonDesp");	
        	dgNonDesp.DataMember = "NonDesp";
        	dgNonDesp.DataSource = ds;  	
        	
			// OpenOrderOneLog
        	sql = "select * from [OpenOrder-Onelog$]";
        	this.ReadOutput(OutputPath, sql, "OpenOrderOnelog");	
        	dgOpenOrderOnelog.DataMember = "OpenOrderOnelog";
        	dgOpenOrderOnelog.DataSource = ds;          	
        	
			// OpenOrderNonOneLog
        	sql = "select * from [OpenOrder-NonOnelog$]";
        	this.ReadOutput(OutputPath, sql, "OpenOrderNonOnelog");	
        	dgOpenOrderNonOnelog.DataMember = "OpenOrderNonOnelog";
        	dgOpenOrderNonOnelog.DataSource = ds; 
        	
        	// OpenOrdereSparesNew
        	sql = "select * from [OpenOrder-eSparesNew$]";
        	this.ReadOutput(OutputPath, sql, "OpenOrdereSparesNew");	
        	dgOpenOrdereSparesNew.DataMember = "OpenOrdereSparesNew";
        	dgOpenOrdereSparesNew.DataSource = ds; 
        	
        	//MonthlyOTDAEMetrics
        	sql = "select * from [AE Metrics$]";
        	this.ReadOutput(OutputPath, sql, "MonthlyOTDAEMetrics");	
        	dgMonthlyOTDAEMetrics.DataMember = "MonthlyOTDAEMetrics";
        	dgMonthlyOTDAEMetrics.DataSource = ds;
        	
        	//MonthlyOTDR4SMetrics
        	sql = "select * from [R4S Metrics$]";
        	this.ReadOutput(OutputPath, sql, "MonthlyOTDR4SMetrics");	
        	dgMonthlyOTDR4SMetrics.DataMember = "MonthlyOTDR4SMetrics";
        	dgMonthlyOTDR4SMetrics.DataSource = ds; 
        	
        	//MonthlyOTDTWTranslation
        	sql = "select * from [MonthlyOTD TW Translation$]";
        	this.ReadOutput(OutputPath, sql, "MonthlyOTDTWTranslation");	
        	dgMonthlyOTDTWTranslation.DataMember = "MonthlyOTDTWTranslation";
        	dgMonthlyOTDTWTranslation.DataSource = ds; 
        	
        	//WeeklyOTDAnalysis
        	sql = "select * from [Weekly OTD Analysis$]";
        	this.ReadOutput(OutputPath, sql, "WeeklyOTDAnalysis");	
        	dgWeeklyOTDAnalysis.DataMember = "WeeklyOTDAnalysis";
        	dgWeeklyOTDAnalysis.DataSource = ds; 
        	
        	//MonthlyActuallyClosedRMA
        	sql = "select * from [Closed RMA-NZ$]";
        	this.ReadOutput(OutputPath, sql, "Closed RMA-NZ");
        	dgClosedRMANZ.DataMember = "Closed RMA-NZ";
        	dgClosedRMANZ.DataSource = ds; 
        	
        	sql = "select * from [Closed RMA-ASB$]";
        	this.ReadOutput(OutputPath, sql, "Closed RMA-ASB");
        	dgClosedRMAASB.DataMember = "Closed RMA-ASB";
        	dgClosedRMAASB.DataSource = ds; 
        	
        	sql = "select * from [Closed RMA-Citadel$]";
        	this.ReadOutput(OutputPath, sql, "Closed RMA-Citadel");
        	dgClosedRMACitadel.DataMember = "Closed RMA-Citadel";
        	dgClosedRMACitadel.DataSource = ds; 
        	
        	sql = "select * from [Closed RMA-eSpares TW$]";
        	this.ReadOutput(OutputPath, sql, "Closed RMA-eSpares TW");
        	dgClosedRMATW.DataMember = "Closed RMA-eSpares TW";
        	dgClosedRMATW.DataSource = ds; 
        	
        	sql = "select * from [Closed RMA-Ormes$]";
        	this.ReadOutput(OutputPath, sql, "Closed RMA-Ormes");
        	dgClosedRMAOrmes.DataMember = "Closed RMA-Ormes";
        	dgClosedRMAOrmes.DataSource = ds; 
        	
        	sql = "select * from [Closed RMA-eSpares New$]";
        	this.ReadOutput(OutputPath, sql, "Closed RMA-eSpares New");
        	dgeSparesNew.DataMember = "Closed RMA-eSpares New";
        	dgeSparesNew.DataSource = ds; 
        	
        	sql = "select * from [Closed RMA$]";
        	this.ReadOutput(OutputPath, sql, "Closed RMA");	
        	dgClosedRMA.DataMember = "Closed RMA";
        	dgClosedRMA.DataSource = ds; 
        	
        	sql = "select * from [Activity YTD$]";
        	this.ReadOutput(OutputPath, sql, "Activity YTD");	
        	dgActivityYTD.DataMember = "Activity YTD";
        	dgActivityYTD.DataSource = ds; 
        	
        	sql = "select * from [Activity YTD Based On OTD$]";
        	this.ReadOutput(OutputPath, sql, "Activity YTD Based On OTD");	
        	dgActivityYTDBasedOnOTD.DataMember = "Activity YTD Based On OTD";
        	dgActivityYTDBasedOnOTD.DataSource = ds; 
        	
        	sql = "select * from [WeeklyActivity-AE$]";
        	this.ReadOutput(OutputPath, sql, "WeeklyActivity-AE");	
        	dgWeeklyActivityAE.DataMember = "WeeklyActivity-AE";
        	dgWeeklyActivityAE.DataSource = ds; 
        	
        	sql = "select * from [WeeklyActivity-RFS$]";
        	this.ReadOutput(OutputPath, sql, "WeeklyActivity-RFS");	
        	dgWeeklyActivityRFS.DataMember = "WeeklyActivity-RFS";
        	dgWeeklyActivityRFS.DataSource = ds;         	
        		
        	sql = "select * from [Cockpit AE uploading$]";
        	this.ReadOutput(OutputPath, sql, "Cockpit AE uploading");	
        	dgCockpitAEUploading.DataMember = "Cockpit AE uploading";
        	dgCockpitAEUploading.DataSource = ds; 
        	
        	sql = "select * from [Cockpit RFS uploading$]";
        	this.ReadOutput(OutputPath, sql, "Cockpit RFS uploading");	
        	dgCockpitR4SUploading.DataMember = "Cockpit RFS uploading";
        	dgCockpitR4SUploading.DataSource = ds; 
        	
        	sql = "select * from [Cockpit Not in OTD$]";
        	this.ReadOutput(OutputPath, sql, "Cockpit Not in OTD");	
        	dgCockpitNotInOTD.DataMember = "Cockpit Not in OTD";
        	dgCockpitNotInOTD.DataSource = ds; 

        	sql = "select * from [OTD Not in Cockpit$]";
        	this.ReadOutput(OutputPath, sql, "OTD Not in Cockpit");	
        	dgOTDNotInCockpit.DataMember = "OTD Not in Cockpit";
        	dgOTDNotInCockpit.DataSource = ds; 

        	sql = "select * from [CockpitVsOTDAnalysis$]";
        	this.ReadOutput(OutputPath, sql, "CockpitVsOTDAnalysis");	
        	dgCockpitVsOTDAnalysis.DataMember = "CockpitVsOTDAnalysis";
        	dgCockpitVsOTDAnalysis.DataSource = ds; 
		}		
	}
}
