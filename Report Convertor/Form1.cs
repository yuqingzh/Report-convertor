/*
 * Created by SharpDevelop.
 * User: yuqingzh
 * Date: 2011-6-16
 * Time: 15:26
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
	/// Description of Form1.
	/// </summary>
	public partial class frmInput : Form
	{
		public static DataSet ds = new DataSet();
		
		public static bool cbOTDReportChecked = false;
		public static bool cbNonDespReportChecked = false;
		public static bool cbOpenOrderReportChecked = false;
		public static bool cbWeeklyOTDAnalysisChecked = false;
		public static bool cbMonthlyOTDAnalysisChecked = false;
		public static bool cbClosedRMAReportChecked = false;		
		public static bool cbActivityYTDChecked = false;
		public static bool cbOTDvsCockpitChecked = false;		
		
		public static bool Input1TL9000Exist = false;
		public static bool Input2TWCloseExist = false;
		public static bool Input3TWOpenExist = false;
		public static bool Input4KOREAExist = false;
		public static bool Input5NZExist = false;
		public static bool Input6NonDespExist = false;		
		public static bool Input7OpenOrderOnelogExist = false;
		public static bool Input8OpenOrderNonOnelogExist = false;
		public static bool OpenOrdereSparesNewExist = false;
		public static bool Input9MonthlyOTDExist = false;
		public static bool WeeklyOTDExist = false;
		public static bool ClosedRMAExist = false;		
		public static bool ActivityYTDExist = false;
		public static bool ActivityYTDBasedOnOTDExist = false;
		public static bool WeeklyActivityBasedOnOTDExist = false;
		public static bool ClosedRMAeSparesNewExist = false;
		public static bool OTDvsCockpitExist = false;
		
		public frmInput(Report_Convertor.frmContainer parent)
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
			this.MdiParent = parent;
			
//			ReadAllInput();
		}
		
		void ReadInput(string inputPath, string sql, string tableName)
		{
			string strCon;		
			OleDbConnection conn;
			
			//if ( Office2010Exists() && inputPath.Contains("xlsx") )
			if ( Office2010Exists() && inputPath.Contains("xlsx") )
			{
				strCon = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source=" + 
										inputPath + ";Extended Properties=\"Excel 12.0; HDR=YES; IMEX=1\"";
			}
			else if ( Office2007Exists() && inputPath.Contains("xlsx") )
			{
				strCon = "Provider = Microsoft.ACE.OLEDB.12.0; Data Source=" + 
										inputPath + ";Extended Properties=\"Excel 12.0; HDR=YES; IMEX=1\"";
			}
			else
			{
				strCon = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source =" +
										inputPath + ";Extended Properties=\"Excel 8.0; HDR=Yes; IMEX=1\"";
			}		
			
			
			conn = new OleDbConnection(strCon);
        	try
        	{
        		if (conn.State == ConnectionState.Closed)
        		{
        			conn.Open();
        		}
        		OleDbDataAdapter da = new OleDbDataAdapter(sql, strCon);
        		da.Fill(ds, tableName);    
        	}        	
        	catch//(OleDbException E)
        	{
        		string text = "Error occurred when reading input file. Please check the sheet and attribute name.\n\n" +
        					  "Input file: " + inputPath + "\n\n" +
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
		
		public void ReadAllInput()
		{
			string InputPath;			
        	string sql;
        	
        	ds.Clear();
        	ds.Tables.Clear();
        	
        	ModifyTypeGuessRows( false );
        	
        	// Sales Orders SoldTo - mapping list for CockpitUploading
        	string InputPath_Sales_Orders_SoldTo_mapping_list = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\Common\\Sales Orders SoldTo - mapping list");
        	if ( (cbOTDReportChecked || cbWeeklyOTDAnalysisChecked) && File.Exists(InputPath_Sales_Orders_SoldTo_mapping_list))
            {                      
   		
        		sql = "select [Sales Orders SoldTo - Customer Name],[Sold-to code] from [AE customer only$]";
        		this.ReadInput(InputPath_Sales_Orders_SoldTo_mapping_list, sql, "Sales_Orders_SoldTo_mapping_list");	
            }
        	
        	// Cares Product
        	string InputPath_Cares_Product = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\Common\\Cares Product");
        	if ( (cbOTDReportChecked || cbWeeklyOTDAnalysisChecked) && File.Exists(InputPath_Cares_Product))
            {                      
   		
        		sql = "select * from [Report 1$]";
        		this.ReadInput(InputPath_Cares_Product, sql, "Cares_Product");	
            }
        	
        	// Input1TL9000
        	InputPath = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\OTD\\Input1-TL9000");
        	string InputPath_IB_ALL_INDO_14022013 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\Common\\IB_ALL_INDO_14022013");
        	if (cbOTDReportChecked && File.Exists(InputPath) && File.Exists(InputPath_IB_ALL_INDO_14022013))
            {                      
            	sql = "select * from [Sheet1$]";
        		this.ReadInput(InputPath, sql, "Input1TL9000");			
        		this.dgInput1TL9000.DataMember = "Input1TL9000";
        		this.dgInput1TL9000.DataSource = ds;
        		
        		sql = "select [Material],[Response Code] from [RAW$]";
        		this.ReadInput(InputPath_IB_ALL_INDO_14022013, sql, "IB_ALL_INDO_14022013");	
        		
        		Input1TL9000Exist = true;
            }
        	else
        	{
        		Input1TL9000Exist = false;
        	}
        	
			
			//Input2TWClose
			InputPath = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\OTD\\Input2-TW-Close");
        	if (cbOTDReportChecked && File.Exists(InputPath))
            { 
				sql = "select * from [Taiwan Sample$]";
				this.ReadInput(InputPath, sql, "Input2TWClose");
				this.dgInput2TWClose.DataMember = "Input2TWClose";
        		this.dgInput2TWClose.DataSource = ds;
        		
        		Input2TWCloseExist = true;
        	}
        	else
        	{
        		Input2TWCloseExist = false;
        	}
			
			//Input3TWOpen
			InputPath = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\OTD\\Input3-TW-Open");
			if (cbOTDReportChecked && File.Exists(InputPath))
            { 
        		sql = "select * from [Taiwan Sample$]";
				this.ReadInput(InputPath, sql, "Input3TWOpen");
				this.dgInput3TWOpen.DataMember = "Input3TWOpen";
        		this.dgInput3TWOpen.DataSource = ds;
        		
        		Input3TWOpenExist = true;
			}
			else
			{
				Input3TWOpenExist = false;
			}
			
			//Input4KOREA
			InputPath = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\OTD\\Input4-KOREA");
			if (cbOTDReportChecked && File.Exists(InputPath))
            { 
        		sql = "select * from [Sheet1$]";
        		this.ReadInput(InputPath, sql, "Input4KOREA");
				this.dgInput4KOREA.DataMember = "Input4KOREA";
        		this.dgInput4KOREA.DataSource = ds;
        		
        		Input4KOREAExist = true;
			}
			else
			{
				Input4KOREAExist = false;
			}
			
			//Input5NZ
			InputPath = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\OTD\\Input5-NZ");
			if (cbOTDReportChecked && File.Exists(InputPath))
            { 
           		sql = "select * from [Promised List$]";
        		this.ReadInput(InputPath, sql, "Input5NZ");
        		this.dgInput5NZ.DataMember = "Input5NZ";
        		this.dgInput5NZ.DataSource = ds;
        		
        		Input5NZExist = true;
			}
			else
			{
				Input5NZExist = false;
			}
        	
        	//Input6NonDesp        		
        	InputPath = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\NonDesp\\Input6-NonDesp");
        	if (cbNonDespReportChecked && File.Exists(InputPath))
            { 
        		sql = "select * from [Sheet1$]";
        		this.ReadInput(InputPath, sql, "Input6NonDesp");			
        		this.dgInput6NonDesp.DataMember = "Input6NonDesp";
        		this.dgInput6NonDesp.DataSource = ds;
        		
        		Input6NonDespExist = true;
        	}
        	else
        	{
        		Input6NonDespExist = false;
        	}
        	
        	//Input7OpenOrderOnelog        		
        	InputPath = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\OpenOrder\\Input7-OpenOrderOnelog");
        	if (cbOpenOrderReportChecked && File.Exists(InputPath))
            { 
        		sql = "select * from [Sheet1$]";
        		this.ReadInput(InputPath, sql, "Input7OpenOrderOnelog");			
        		this.dgInput7OpenOrderOnelog.DataMember = "Input7OpenOrderOnelog";
        		this.dgInput7OpenOrderOnelog.DataSource = ds;
        		
        		Input7OpenOrderOnelogExist = true;
        	}
        	else
        	{
        		Input7OpenOrderOnelogExist = false;
        	}
        	
        	//Input8OpenOrderOnelog        		
        	InputPath = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\OpenOrder\\Input8-NZ");
        	if (cbOpenOrderReportChecked && File.Exists(InputPath))
            { 
        		sql = "select * from [Open Order$]";
        		this.ReadInput(InputPath, sql, "Input8OpenOrderNonOnelog");			
        		this.dgInput8OpenOrderNonOnelog.DataMember = "Input8OpenOrderNonOnelog";
        		this.dgInput8OpenOrderNonOnelog.DataSource = ds;
        		
        		Input8OpenOrderNonOnelogExist = true;
        	}
        	else
        	{
        		Input8OpenOrderNonOnelogExist = false;
        	}
        	
        	//OpenOrdereSparesNew       		
        	InputPath = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\OpenOrder\\OpenOrder-eSparesNew");
        	string InputPath19 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\Common\\eSparesNew - Country Code List"); 
			string InputPath20 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\OpenOrder\\OpenOrder-eSparesTW");
				
        	if (cbOpenOrderReportChecked &&
        	    File.Exists(InputPath19) &&
        	    File.Exists(InputPath20) &&
        	    File.Exists(InputPath))
            { 
        		sql = "select [Sales Orders - Type], "
        				+ "[Sales Orders - ShipTo Customer Name], "
        				+ "[Sales Orders - Order  Reason], "
        				+ "[Sales Orders Line - Delivery SLA], "	//New column
        				+ "[Sales Orders - Order  Reason (Medium description)], "
        				+ "[Sales Orders - Sales Order Number], "
        				+ "[Sales Orders Line - LineNumber], "
        				+ "[Sales Orders Line - Material], "
        				+ "[Sales Orders Line - SPT], "
        				+ "[IC Delivery Line -  RcvdSerialNumber], "
        				+ "[Sales Orders - Creation Date], "
        				+ "[IC Shipments - Arrival Date], "
        				+ "[Sales Orders Line - Rejection 12 (Medium description)] "
        				+ "from [Taiwan Sample$]";
        		this.ReadInput(InputPath20, sql, "OpenOrdereSparesTW");
        		
        		sql = "select [Sales Orders - RSCIC/RSLC (Sales Organisation)], "
        				+ "[Customer Status (Closed / Open)], "
        				+ "[Sales Orders - Type], "
        				+ "[Sales Orders SoldTo - Customer Name], "
        				+ "[Sales Orders - Order  Reason], "
        				+ "[Sales Orders - Order Reason Service Description], "
        				+ "[Sales Orders - Order Reason eSpares SLA], "
        				+ "[Sales Orders - Sales Order Number], "
        				+ "[Sales Orders Line - LineNumber], "
        				+ "[Sales Orders Line - Material], "
        				+ "[Sales Orders Line - SPT], "
        				+ "[IC Delivery Line -  RcvdSerialNumber], "
        				+ "[Work Orders  Header - Vendor Name], "
        				+ "[Sales Orders Line - Creation Date], "
        				+ "[IC Shipments - Arrival Date], "
        				+ "[Customer Due Date], "
        				+ "[TAT Status], "
        				+ "[Customer OPEN TAT] "
        				+ "from [RAW Data$]";
        		this.ReadInput(InputPath, sql, "OpenOrdereSparesNew");
        		//	Add a new column of COUNTRY
				DataColumn dc = new DataColumn("COUNTRY");	
				frmInput.ds.Tables["OpenOrdereSparesNew"].Columns.Add(dc);
        		this.dgOpenOrdereSparesNew.DataMember = "OpenOrdereSparesNew";
        		this.dgOpenOrdereSparesNew.DataSource = ds;
        		
        		sql = "select * from [Sheet1$]";
        		this.ReadInput(InputPath19, sql, "OpenOrder - eSparesNew - Country Code List");
        		
        		OpenOrdereSparesNewExist = true;
        	}
        	else
        	{
        		OpenOrdereSparesNewExist = false;
        	}
        	
        	//OpenLogReport_Excluding_List
        	InputPath = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\OpenOrder\\OpenLogReport_Excluding_List");
        	if (cbOpenOrderReportChecked && File.Exists(InputPath) && (Input8OpenOrderNonOnelogExist || Input7OpenOrderOnelogExist) )
            { 
        		sql = "select * from [RMA_Cancelled_List$]";
        		this.ReadInput(InputPath, sql, "RMACancelledList");	
        	}
        	
        	
        	//Input9MonthlyOTD    		
        	InputPath = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\MonthlyOTDAnalysis\\MonthlyOTD");
        	if (cbMonthlyOTDAnalysisChecked && File.Exists(InputPath))
            { 
        		sql = "select * from [OTD$]";
        		this.ReadInput(InputPath, sql, "Input9MonthlyOTD");			
        		this.dgInput9MonthlyOTDOnelog.DataMember = "Input9MonthlyOTD";
        		this.dgInput9MonthlyOTDOnelog.DataSource = ds;
        		
       		
        		Input9MonthlyOTDExist = true;
        	}
        	else
        	{
        		Input9MonthlyOTDExist = false;
        	}
        	
        	//MonthlyOTDKeyCustomerList
        	InputPath = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\MonthlyOTDAnalysis\\Monthly OTD-top 10 customer_ list");
        	if (cbMonthlyOTDAnalysisChecked && File.Exists(InputPath) && Input9MonthlyOTDExist)
            { 
        		sql = "select * from [Sheet1$]";
        		this.ReadInput(InputPath, sql, "MonthlyOTDKeyCustomerList");	
        	}
        	
        	//MonthlyOTDTWCustomerNameTranslation
        	InputPath = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\MonthlyOTDAnalysis\\Taiwan customer name_ translation");
        	if (cbMonthlyOTDAnalysisChecked &&File.Exists(InputPath) && Input9MonthlyOTDExist)
            { 
        		sql = "select * from [Sheet1$]";
        		this.ReadInput(InputPath, sql, "MonthlyOTDTWCustomerNameTranslation");	
        	}
        	
        	//WeeklyOTD    		
        	InputPath = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\WeeklyOTDAnalysis\\WeeklyOTD");
        	if (cbWeeklyOTDAnalysisChecked && File.Exists(InputPath))
            { 
        		sql = "select * from [OTD$]";
        		this.ReadInput(InputPath, sql, "WeeklyOTD");			
        		this.dgWeeklyOTDOnelog.DataMember = "WeeklyOTD";
        		this.dgWeeklyOTDOnelog.DataSource = ds;
        		
        		WeeklyOTDExist = true;
        	}
        	else
        	{
        		WeeklyOTDExist = false;
        	}
        	
        	//WeeklyOTDKeyCustomerList
        	InputPath = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\WeeklyOTDAnalysis\\Weekly OTD-key customer list");
        	if (cbWeeklyOTDAnalysisChecked && File.Exists(InputPath) && WeeklyOTDExist)
            { 
        		sql = "select * from [Sheet1$]";
        		this.ReadInput(InputPath, sql, "WeeklyOTDKeyCustomerList");	
        	}
        	
        	//MonthlyActuallyClosedRMA
        	if (cbClosedRMAReportChecked)
        	{
        		string InputPath1 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\ClosedRMA\\ClosedRMA-ASB");
        		string InputPath2 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\ClosedRMA\\ClosedRMA-Citadel Report");
				string InputPath3 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\ClosedRMA\\ClosedRMA-eSpares TW");
				string InputPath4 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\ClosedRMA\\ClosedRMA-NZ");
				string InputPath5 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\ClosedRMA\\ClosedRMA-Ormes");
				string InputPath6 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\ClosedRMA\\Ref\\VC Catalogue - RESO EMEA");
        		string InputPath7 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\ClosedRMA\\Ref\\VC Catalogue - RESO AMERICAS NAR");
				string InputPath8 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\ClosedRMA\\Ref\\VC Catalogue - RESO APAC CHINA QD");
				string InputPath9 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\ClosedRMA\\Ref\\VC Catalogue - RESO APAC INDIA");
				string InputPath10 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\ClosedRMA\\Ref\\VC Catalogue - RESO China (SHA)");
				string InputPath11 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\ClosedRMA\\Ref\\Excluding List");
				string InputPath12 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\ClosedRMA\\Ref\\Stinger product code mapping list");
				string InputPath13 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\ClosedRMA\\Ref\\APAC Country Code List");
				string InputPath14 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\ClosedRMA\\Ref\\Repairer To SubGroup Mapping List - NZ");
				string InputPath15 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\ClosedRMA\\Ref\\Repairer To SubGroup Mapping List - TW");
				string InputPath16 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\ClosedRMA\\ClosedRMA-eSpares_New"); 
				string InputPath17 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\ClosedRMA\\Ref\\Repairer To SubGroup Mapping List - eSpares New"); 
				string InputPath18 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\Common\\eSparesNew - Country Code List"); 

        		if (File.Exists(InputPath1) &&
				    File.Exists(InputPath2) &&
				    File.Exists(InputPath3) &&
				    File.Exists(InputPath4) &&
				    File.Exists(InputPath5) &&
				    File.Exists(InputPath6) &&
				    File.Exists(InputPath7) &&
				    File.Exists(InputPath8) &&
				    File.Exists(InputPath9) &&
				    File.Exists(InputPath10) &&
				    File.Exists(InputPath11) &&
				    File.Exists(InputPath12) &&
				    File.Exists(InputPath13) &&
				    File.Exists(InputPath14) &&
				    File.Exists(InputPath15) &&
				    File.Exists(InputPath16) &&
				    File.Exists(InputPath17) &&
				    File.Exists(InputPath18))
           		{ 
					DataColumn dc;		
					ClosedRMAExist = true;
					
        			sql = "select [Country],[Leave SHA airport date],"
        				 +"[Received P/N] from [Sheet1$]";
        			this.ReadInput(InputPath1, sql, "ClosedRMA-ASB");	
        			//	Add a new column of repairer sub-group	
					dc = new DataColumn("Repairer sub-group");			
					frmInput.ds.Tables["ClosedRMA-ASB"].Columns.Add(dc);
					//	Add a new column of repair cost
					dc = new DataColumn("Repair cost");	
					frmInput.ds.Tables["ClosedRMA-ASB"].Columns.Add(dc);
        			this.dgClosedRMAASB.DataMember = "ClosedRMA-ASB";
        			this.dgClosedRMAASB.DataSource = ds;
        		
        			sql = "select [Country], [Rcv_Pno], [Rcv_Model], [Repairer],"
        				+"[Repair_Code] from [Sheet1$]";
        			this.ReadInput(InputPath2, sql, "ClosedRMA-Citadel Report");
        			//	Add a new column of repairer sub-group	
					dc = new DataColumn("Repairer sub-group");			
					frmInput.ds.Tables["ClosedRMA-Citadel Report"].Columns.Add(dc);
					//	Add a new column of repair cost
					dc = new DataColumn("Repair cost");	
					frmInput.ds.Tables["ClosedRMA-Citadel Report"].Columns.Add(dc);
        			this.dgClosedRMACitadelReport.DataMember = "ClosedRMA-Citadel Report";
        			this.dgClosedRMACitadelReport.DataSource = ds;
        			
        			sql = "select [IR Delivery Header - Vendor Name],"
        				+"[Sales Orders - ShipTo Customer Name],"
        				+"[IR Delivery Line - Rcvd Material],"
        				+"[IR Delivery Header -  GI Date],"
        				+"[Work Orders  Line - Repair Code] from [Taiwan Sample$]";
        			this.ReadInput(InputPath3, sql, "ClosedRMA-eSpares TW");				
					//	Add a new column of repairer sub-group	
					dc = new DataColumn("Repairer sub-group");			
					frmInput.ds.Tables["ClosedRMA-eSpares TW"].Columns.Add(dc);
					//	Add a new column of repair cost
					dc = new DataColumn("Repair cost");			
					frmInput.ds.Tables["ClosedRMA-eSpares TW"].Columns.Add(dc);			
        			this.dgClosedRMAeSparesTW.DataMember = "ClosedRMA-eSpares TW";
        			this.dgClosedRMAeSparesTW.DataSource = ds;
        			
        			//sql = "select [Manufacturers PartNo], [ExtRep], [RMA1], [Repairer1], [Repairer2] from [Sheet1$]";
        			sql = "select [Manufacturers PartNo], [ExtRep], [Repairer] from [Sheet1$]";
        			this.ReadInput(InputPath4, sql, "ClosedRMA-NZ");
					//	Add a new column of repairer sub-group	
					dc = new DataColumn("Repairer sub-group");			
					frmInput.ds.Tables["ClosedRMA-NZ"].Columns.Add(dc);
					//	Add a new column of repair cost
					dc = new DataColumn("Repair cost");	
					frmInput.ds.Tables["ClosedRMA-NZ"].Columns.Add(dc);
        			this.dgClosedRMANZ.DataMember = "ClosedRMA-NZ";
        			this.dgClosedRMANZ.DataSource = ds;
        			
        			sql = "select [CUST_Country], [OC_GI_Date], [ITEM_FRU] from [RFR APAC$]";
        			this.ReadInput(InputPath5, sql, "ClosedRMA-Ormes");	
        			//	Add a new column of repairer sub-group	
					dc = new DataColumn("Repairer sub-group");			
					frmInput.ds.Tables["ClosedRMA-Ormes"].Columns.Add(dc);
					//	Add a new column of repair cost
					dc = new DataColumn("Repair cost");	
					frmInput.ds.Tables["ClosedRMA-Ormes"].Columns.Add(dc);
        			this.dgClosedRMAOrmes.DataMember = "ClosedRMA-Ormes";
        			this.dgClosedRMAOrmes.DataSource = ds;        	
        			
        			sql = "select [Sales Orders - RSCIC/RSLC (Sales Organisation)], "
        				+ "[Sales Orders SoldTo - Customer Name],"
        				+ "[IR Delivery Header -  GI Date], " 
        				+ "[Work Orders  Header - Vendor Name], "
        				+ "[IR Delivery Line - Rcvd Material] "
        				+ "from [RAW Data$]";
        			this.ReadInput(InputPath16, sql, "ClosedRMA-eSpares New");
        			//	Add a new column of COUNTRY
					dc = new DataColumn("COUNTRY");	
					frmInput.ds.Tables["ClosedRMA-eSpares New"].Columns.Add(dc);
        			//	Add a new column of repairer sub-group	
					dc = new DataColumn("Repairer sub-group");			
					frmInput.ds.Tables["ClosedRMA-eSpares New"].Columns.Add(dc);
					//	Add a new column of repair cost
					dc = new DataColumn("Repair cost");	
					frmInput.ds.Tables["ClosedRMA-eSpares New"].Columns.Add(dc);
        			this.dgeSparesNew.DataMember = "ClosedRMA-eSpares New";
        			this.dgeSparesNew.DataSource = ds;
        			
        		
        			sql = "select [RES_IDENTIFIER],[VC_RFR],[VC_RFR_NFF] from [VC_APAC$]";
        			this.ReadInput(InputPath6, sql, "VC Catalogue - RESO EMEA - VC_APAC"); 
        			
        			sql = "select [RES_IDENTIFIER],[VC_RFR],[VC_RFR_NFF] from [VC_APAC FCA$]";
        			this.ReadInput(InputPath6, sql, "VC Catalogue - RESO EMEA - VC_APAC FCA"); 
        			
        			sql = "select [RES_IDENTIFIER],[VC_RFR_60],[VC_NFF] from [VC_ASIA_PACIFIC$]";
        			this.ReadInput(InputPath7, sql, "VC Catalogue - RESO AMERICAS NAR");        			
        			
        			sql = "select [RES_IDENTIFIER],[VC_RFR],[VC_RFR_NFF] from [ASIA PACIFIC$]";
        			this.ReadInput(InputPath8, sql, "VC Catalogue - RESO APAC CHINA QD");        			
        			
        			sql = "select [RES_ID],[VC_RFR_60],[VC NFF] from [VC-APAC$]";
        			this.ReadInput(InputPath9, sql, "VC Catalogue - RESO APAC INDIA");        			

					sql = "select [RES_IDENTIFIER],[PART_NUMBER],[VC_RFR60],[VC_NFF] from [APAC$]";
        			this.ReadInput(InputPath10, sql, "VC Catalogue - RESO China (SHA)");        			
        			        
        			sql = "select * from [TW for test or B4R$]";
        			this.ReadInput(InputPath11, sql, "Excluding List TW");
        			
        			sql = "select * from [Sheet1$]";
        			this.ReadInput(InputPath12, sql, "Stinger product code mapping list");        			
        			
        			sql = "select [APAC Country ISO CODE] from [Sheet1$]";
        			this.ReadInput(InputPath13, sql, "APAC Country Code List");
        			
        			sql = "select * from [Sheet1$]";
        			this.ReadInput(InputPath14, sql, "Repairer To SubGroup Mapping List - NZ");
        			
        			sql = "select * from [Sheet1$]";
        			this.ReadInput(InputPath15, sql, "Repairer To SubGroup Mapping List - TW");
     			
        			sql = "select * from [Sheet1$]";
        			this.ReadInput(InputPath17, sql, "Repairer To SubGroup Mapping List - eSpares New");
        			
        			sql = "select * from [Sheet1$]";
        			this.ReadInput(InputPath18, sql, "eSparesNew - Country Code List");

        		}
        		else
        		{
        			MessageBox.Show( "Lack of files for ClosedRMA!" );
        			ClosedRMAExist = false;
        		}
        	}
        	else
        	{
        		ClosedRMAExist = false;
        	}
        	

        	//ActivityYTD
        	if (cbActivityYTDChecked)
        	{
        		string InputPath1 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\ActivityYTDReport\\KOREA-YTD-Base on RMA creation date");
        		string InputPath2 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\ActivityYTDReport\\KOREA-YTD-Base on shipping date");
				string InputPath3 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\ActivityYTDReport\\NZ-YTD");
				string InputPath4 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\ActivityYTDReport\\TW-IC-YTD");
				string InputPath5 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\ActivityYTDReport\\OTD-Onelog-TWOnly-YTD");
				string InputPath6 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\ActivityYTDReport\\CitadelShipped-YTD");
				string InputPath7 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\ActivityYTDReport\\CitadelReceived-YTD");
				string InputPath8 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\ActivityYTDReport\\Ticket_Type_Sla ALu_for Citadel shipped");
				string InputPath9 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\ActivityYTDReport\\eSparesNew");
				string InputPath10 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\Common\\eSparesNew - Country Code List"); 


        		if (File.Exists(InputPath1) &&
				    File.Exists(InputPath2) &&
				    File.Exists(InputPath3) &&
				    File.Exists(InputPath4) &&
				    File.Exists(InputPath5) &&
				    File.Exists(InputPath6) &&
				    File.Exists(InputPath7) &&
				    File.Exists(InputPath8) &&
				    File.Exists(InputPath9) &&
				    File.Exists(InputPath10))
           		{ 	
					ActivityYTDExist = true;
					
        			sql = "select [Create_Date],[Customer Name] from [Sheet1$]";
        			this.ReadInput(InputPath1, sql, "KOREA-YTD-Base on RMA creation date");        			
        		
        			sql = "select [Ship_Date_Repl],[Create_Date],[Customer Name],[Delivery Status] from [Sheet1$]";
        			this.ReadInput(InputPath2, sql, "KOREA-YTD-Base on shipping date"); 
        			
        			sql = "select [Customer TAT],[Dispatched Date to Customer],[Target Due Date],[Customer Request Date],[Faulty Board Received Date from Customer],[Country],[Customer Name] from [Monthly Closed Order$]";
        			this.ReadInput(InputPath3, sql, "NZ-YTD"); 
        			
        			sql = "select [Cust Name],[Received] from [Received Raw$]";
        			this.ReadInput(InputPath3, sql, "NZ-Received-YTD"); 
        			
        			sql = "select [IC Shipments - Arrival Date],[Sales Orders - ShipTo Customer Name] from [Taiwan Sample$]";
        			this.ReadInput(InputPath4, sql, "TW-IC-YTD"); 
        			
        			sql = "select [Country],[Create_Date],[Customer Name],[Ship_Date_Repl],[Delivery Status],[Rcv_Date_Def],[SLA] from [Onelog$]";
        			this.ReadInput(InputPath5, sql, "OTD-Onelog-TWOnly-YTD"); 
        			
        			sql = "select [Country],[Customer_Name],[Ship_Date_Repl],[Create_Date],[Rcv_Date_Def],[Ticket_Type],[Service_Type] from [Sheet1$]";
        			this.ReadInput(InputPath6, sql, "CitadelShipped-YTD"); 
        			
        			sql = "select [Country],[Customer_Name],[Rcv_Date_Def] from [Sheet1$]";
        			this.ReadInput(InputPath7, sql, "CitadelReceived-YTD"); 

					sql = "select [Id],[Description],[Response] from [Sheet1$]";
        			this.ReadInput(InputPath8, sql, "Ticket_Type_Sla ALu_for Citadel"); 
        			
        			sql = "select [Sales Orders - RSCIC/RSLC (Sales Organisation)], "
        				+ "[Sales Orders SoldTo - Customer Name],"
        				+ "[Sales Orders - Order  Reason],"
        				+ "[Sales Orders Line - Creation Date],"
        				+ "[IC Shipments - Arrival Date],"
        				+ "[TAT Status],"
        				+ "[OC Delivery Header -  GI_Date],"
        				+ "[IR Delivery Header -  GI Date] " 
        				+ "from [RAW Data$]";
        			this.ReadInput(InputPath9, sql, "ActivityYTD-eSpares New"); 
        			
        			sql = "select * from [Sheet1$]";
        			this.ReadInput(InputPath10, sql, "ActivityYTD-eSparesNew - Country Code List");
        		}
        		else
        		{
        			//MessageBox.Show( "Lack of files for Activity YTD!" );
        			ActivityYTDExist = false;
        		}
        	}
        	else
        	{
        		ActivityYTDExist = false;
        	}     
        	
        	//Activity YTD Based On OTD
        	if (cbActivityYTDChecked)
        	{
        		string InputPath1 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\ActivityYTDReport\\OTD");
								
        		if (File.Exists(InputPath1))
           		{ 	
					ActivityYTDBasedOnOTDExist = true;
					
        			sql = "select [Country],[Customer Name],[Target Compl Date],[Create_Date],[Ship_Date_Repl],[Delivery Status],[Rcv_Date_Def],[SLA] from [Onelog$]";
        			this.ReadInput(InputPath1, sql, "OTD-Onelog");        			
        		
        			sql = "select [Country],[Customer Name],[Target Due Date],[Customer Request Date],[Dispatched Date to Customer],[Delivery Status],[Faulty Board Received Date from Customer],[SLA] from [Non-Onelog$]";
        			this.ReadInput(InputPath1, sql, "OTD-NonOnelog"); 
        			
           		}
        		else
        		{
        			//MessageBox.Show( "Lack of files for Activity YTD Based On OTD!" );
        			ActivityYTDBasedOnOTDExist = false;
        		}
        	}
        	else
        	{
        		ActivityYTDBasedOnOTDExist = false;
        	}   
        	
        	//Weekly Activity YTD Based On OTD
        	if (cbActivityYTDChecked)
        	{
        		string InputPath1 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\ActivityYTDReport\\WeeklyOTD");
								
        		if (File.Exists(InputPath1))
           		{ 	
					WeeklyActivityBasedOnOTDExist = true;
					
        			sql = "select [Service Type],[Sales Orders - ShipTo Country],[Sales Orders SoldTo - Customer Name],[DUE DATE (FINAL)],[SLA START DATE (FINAL)],[DELIVERY DATE (FINAL)],[Net OTD Status (1=OK; 0=NOK)] from [OTD$]";    			
        			this.ReadInput(InputPath1, sql, "WeeklyOTDActivity");
           		}
        		else
        		{
        			//MessageBox.Show( "Lack of files for Weekly Activity Based On OTD!" );
        			WeeklyActivityBasedOnOTDExist = false;
        		}
        	}
        	else
        	{
        		WeeklyActivityBasedOnOTDExist = false;
        	}
        	
        	//Cockpit VS OTD
        	if (cbOTDvsCockpitChecked)
        	{
        		string InputPath1 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\CockpitVsOTD\\YTD OTD Report");
				string InputPath2 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\CockpitVsOTD\\AP AE Cockpit");
				string InputPath3 = SetInputPath(System.Environment.CurrentDirectory + "\\Data\\CockpitVsOTD\\AP RFS Cockpit");
				
        		if (File.Exists(InputPath1) &&
				    File.Exists(InputPath2) &&
				    File.Exists(InputPath3))
           		{ 	
					OTDvsCockpitExist = true;
					
        			//sql = "select [SO_SOLINE],[WK],[Target Compl Date],[Country],[Customer Name],[Comments from In-country ],[L-1],[L-2] from [Onelog$]";
        			sql = "select [Unique Identification],[WK],[Target Compl Date],[Country],[Customer Name],[Comments from In-country ],[L-1],[L-2] from [Onelog$]";
        			this.ReadInput(InputPath1, sql, "OTD-Onelog");        			
        		
        			//sql = "select [SO_SOLINE],[WK] from [Non-Onelog$]";
        			sql = "select [RMA#],[SLA],[WK],[Target Due Date],[Country],[Customer Name],[Comments / Reason],[L-1],[L-2] from [Non-Onelog$]";
        			this.ReadInput(InputPath1, sql, "OTD-NonOnelog"); 
        			
        			sql = "select [SO_SOLINE],[Due Week],[CUSTOMER_DUE_DATE_DAY],[SO_SHIPTO_COUNTRY],[CUSTOMER],[NET_OTD_STATUS],[OTDC_COMMENTS],[OTDC_CORRECTION],[OTDC_FAILURE_DESCRIPTION],[OTDC_FAILURE_REASON],[LAST_FILE],[LAST_UPDATED] from [Volume_data$]";
        			this.ReadInput(InputPath2, sql, "AP-AE-Cockpit"); 
        			
        			sql = "select [SO_SOLINE],[Due Week],[CUSTOMER_DUE_DATE_DAY],[SO_SHIPTO_COUNTRY],[CUSTOMER],[NET_OTD_STATUS],[OTDC_COMMENTS],[OTDC_CORRECTION],[OTDC_FAILURE_DESCRIPTION],[OTDC_FAILURE_REASON],[LAST_FILE],[LAST_UPDATED] from [Volume_data$]";
        			this.ReadInput(InputPath3, sql, "AP-RFS-Cockpit"); 
        			
           		}
        		else
        		{
        			//MessageBox.Show( "Lack of files for Activity YTD Based On OTD!" );
        			OTDvsCockpitExist = false;
        		}
        	}
        	else
        	{
        		OTDvsCockpitExist = false;
        	}
		}
				
		private string SetInputPath(string inputPath)
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
		
		private bool Office2003Exists()
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
		
		private bool Office2007Exists()
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
		
		private bool Office2010Exists()
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
		private bool ModifyTypeGuessRows(bool setDefault)
		{
		    int toSetValue = 0;
			string JetKeyRoot = "SOFTWARE\\Microsoft\\Jet\\4.0\\Engines\\Excel";
			string ACEKeyRoot = "SOFTWARE\\Microsoft\\Office\\12.0\\Access Connectivity Engine\\Engines\\Excel";
			string ACEKeyRoot10 = "SOFTWARE\\Microsoft\\Office\\14.0\\Access Connectivity Engine\\Engines\\Excel";
	  		
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
		    }
		    catch(Exception ex)
		    {
//		        MessageBox.Show("Registry update failed："+ ex.Message);
		        return false;
		    }
		    
		    return true;
		}
		
//		private void SetNullText(DataGrid dg)
//		{
//    		DataGridColumnStyle myGridColumn;
//    		myGridColumn = dg.TableStyles[0].GridColumnStyles[0];
//    		myGridColumn.NullText = "";
//		}
	}
}
