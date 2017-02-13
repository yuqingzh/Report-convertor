/*
 * Created by SharpDevelop.
 * User: yuqingzh
 * Date: 2011-6-17
 * Time: 12:21
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Data;
using System.Data.Sql;
using System.Data.OleDb;

namespace Report_Convertor
{
	
	/// <summary>
	/// Description of Class1.
	/// </summary>
	///
	public enum Input
	{
		INPUT1TL9000,
		INPUT2TWCLOSE,
		INPUT3TWOPEN,
		INPUT4KOREA,
		INPUT5NZ
	};
	
	public class OTD
	{
		
		public OTD()
		{
			
		}
		
		private bool CustomerNameFilter4TW(string customerName)
		{
			if ( customerName == "台灣國際標準電子股份有限公司" ||	
			     customerName == "內政部消防署" ||
			     customerName.ToUpper().Contains("ALCATEL") ||
			     customerName.ToUpper().Contains("LUCENT"))
			{
				return false;
			}
			else
			{
				return true;
			}
		}

		private bool CurrentWeekFilter(DateTime date)
		{
			if ( date >= frmConfig.CurrentWeekStartDate &&
			    date < frmConfig.CurrentWeekEndDate.AddDays(1) )
			{
				return true;
			}
			else
			{
				return false;
			}
		}
		
		private bool PassdueDateFilter(DateTime date)
		{
			
			if ( date >= frmConfig.PassDueStartDate &&
			    date <= frmConfig.PassDueEndDate )
			{
				return true;
			}
			else
			{
				return false;
			}
		}

		private string CalcSLANumberofDays4TW(string strOrderReason, string strDeliverySLA)
		{
			string ret = "";
			
			if (strDeliverySLA != "")
			{
				try
				{
					ret = Convert.ToInt32(strDeliverySLA).ToString(); //blank out when non digital.
				}
				catch{	
				}
			}
			else
			{
				string SLA_2;
			
				if (strOrderReason == "")
				{
					return ret;
				}
						
				try
				{
					if ( strOrderReason == "R12")
					{
						ret = "180";
					}
					else if (strOrderReason == "H08")
					{
						ret = "0.33";
					}
					else if (strOrderReason == "H02")
					{
						ret = "0.08";
					}
					else {
					SLA_2 = strOrderReason.Substring(1, strOrderReason.Length - 1);
							
					switch (SLA_2)
					{
						case "C0":
							ret = "120";
							break;
						case "D1":
							ret = "1"; 
							break;
						case "F0":
							ret = "150"; 
							break;
						case "SU":
							ret = "0"; //Luis - Need to check
							break;
						default:
							try {
								ret = Convert.ToInt32(SLA_2).ToString();
							}
							catch{
							}
							break;
					}//switch
				}
				} //strOrderReason != "R12"
				catch{
				}				
			}
			
			return ret;
		}
		
		private string CalcTATuponActualClose(string SLA, string Ship_Date_Repl, string Rcv_Date_Def, string Create_Date)
		{
			string ret= "";
			try
			{
				DateTime dtShip_Date_Repl, dtRcv_Date_Def, dtCreate_Date;
										
				if (SLA == "R4S")
				{
					if (Ship_Date_Repl == "" || 
							Rcv_Date_Def == "" )
					{
						ret = "N/A";
					}
					else
					{
						dtShip_Date_Repl = Convert.ToDateTime(Ship_Date_Repl).Date;
						dtRcv_Date_Def = Convert.ToDateTime(Rcv_Date_Def).Date;
						ret = (dtShip_Date_Repl.Subtract(dtRcv_Date_Def)).Days.ToString();
					}
				}
				else if (SLA == "D+" || SLA == "H+")
				{
					if (Ship_Date_Repl == "" || 
							Create_Date == "")
					{
						ret = "N/A";
					}
					else
					{
						dtShip_Date_Repl = Convert.ToDateTime(Ship_Date_Repl).Date;
						dtCreate_Date = Convert.ToDateTime(Create_Date).Date;
						ret = (dtShip_Date_Repl.Subtract(dtCreate_Date)).Days.ToString();
					}
				}
			}
			catch
			{
					
			}
			
			return ret;
		}
		
		private string CalcTATuponActualClose4Korea(string SLA, string Ship_Date_Repl, string Rcv_Date_Def, string Create_Date)
		{
			string ret= "";
			try
			{
				DateTime dtShip_Date_Repl, dtRcv_Date_Def, dtCreate_Date;
										
				if (SLA == "R4S")
				{
					if (Ship_Date_Repl == "")
					{
						ret = "N/A";
					}
					else if (Ship_Date_Repl != "" && 
							Rcv_Date_Def != "" )
					{
						dtShip_Date_Repl = Convert.ToDateTime(Ship_Date_Repl).Date;
						dtRcv_Date_Def = Convert.ToDateTime(Rcv_Date_Def).Date;
						ret = (dtShip_Date_Repl.Subtract(dtRcv_Date_Def)).Days.ToString();
					}
					else if (Ship_Date_Repl != "" &&
							Create_Date != "" )					
					{
						dtShip_Date_Repl = Convert.ToDateTime(Ship_Date_Repl).Date;
						dtCreate_Date = Convert.ToDateTime(Create_Date).Date;
						ret = (dtShip_Date_Repl.Subtract(dtCreate_Date)).Days.ToString();
					}
					else if (Rcv_Date_Def == "" && 
							Create_Date == "" )	
					{
						ret = "N/A";
					}
				}
				else if (SLA == "D+" || SLA == "H+")
				{
					if (Ship_Date_Repl == "" || 
							Create_Date == "")
					{
						ret = "N/A";
					}
					else
					{
						dtShip_Date_Repl = Convert.ToDateTime(Ship_Date_Repl).Date;
						dtCreate_Date = Convert.ToDateTime(Create_Date).Date;
						ret = (dtShip_Date_Repl.Subtract(dtCreate_Date)).Days.ToString();
					}
				}
			}
			catch
			{
					
			}
			
			return ret;
		}
				
		private DateTime CalcTargetDueDate(string strOrderReason, string strDeliverySLA, string GIDate, string CreationDate, string uniqueID)
		{		
			DateTime ret = new DateTime(1900, 1, 1, 0, 0, 0);
						
			double days = -10000; //flag
			string SLA_1, SLA_2;
			
			if (strOrderReason == "")
			{
				return ret;
			}
						
			try
			{
				SLA_1 = strOrderReason.Substring(0, 1);

				if (strDeliverySLA != "")
				{
					try {
						days = Convert.ToDouble(strDeliverySLA);
					}
					catch{
					}
				}
			
				if (days == -10000)
				{
					if ( strOrderReason == "R12" )
					{
						days = 180;
					}
					else if ( strOrderReason == "H08" )
					{
						days = 0.33;
					}
					else if (strOrderReason == "H02")
					{
						days = 0.08;
					}
					else {
					SLA_2 = strOrderReason.Substring(1, strOrderReason.Length - 1);
							
					switch (SLA_2)
					{
						case "C0":
							days = 120;
							break;
						case "D1":
							days = 1; 
							break;
						case "F0":
							days = 150; 
							break;
						case "SU":
							days = 0; //Luis - Need to check
							break;
						default:
							try {
								days = Convert.ToDouble(SLA_2);
							}
							catch{
								MessageBox.Show("Cannot convert " + strOrderReason + " to customer due date in eSpares TW.");
							}
							break;
					} //strOrderReason != "R12"
					}//switch
				} //if (days == -10000)
			
				if (SLA_1 == "A" || SLA_1 == "H") //AE
				{
					if ( CreationDate.ToString() == "" )	
					{
						return ret;		//exclude the record
					}
					ret = (Convert.ToDateTime(CreationDate)).AddDays(days);				
				}
				else
				{
					if ( GIDate.ToString() == "" )	
					{
						return ret;		//exclude the record
					}
					ret = (Convert.ToDateTime(GIDate)).AddDays(days);
				}
			} //try
			catch
			{
				string msg = "Onelog: Error Occurred when processing 'Target Due Date' field "
						+ "in the record 'Unique Identification' = " + uniqueID;

					MessageBox.Show(msg);
			}
			return ret;
		}


		private string findSoldToCode(string CustomerName)
		{
			foreach (DataRow dr in frmInput.ds.Tables["Sales_Orders_SoldTo_mapping_list"].Rows)
			{
				if (CustomerName == dr["Sales Orders SoldTo - Customer Name"].ToString()) 
				{
					return dr["Sold-to code"].ToString();										
				}
			}
			
			return "Need Manual Checking";
		}
		
		private void findProductNameAndBizLine(string Material, out string ProductName, out string BizLine)
		{
			foreach (DataRow dr in frmInput.ds.Tables["Cares_Product"].Rows)
			{
				if (Material == dr["CARES Product - Part Code"].ToString()) 
				{
					ProductName = dr["CARES Product - BD Code"].ToString();
					BizLine = dr["CARES Product - Technology Name"].ToString();
					return;
				}
			}
			
			foreach (DataRow dr in frmInput.ds.Tables["Cares_Product"].Rows)
			{
				if (Material == dr["CARES Product - RES Identifier"].ToString()) 
				{
					ProductName = dr["CARES Product - BD Code"].ToString();
					BizLine = dr["CARES Product - Technology Name"].ToString();
					return;										
				}
			}
	
			ProductName = "";
			BizLine = "";
			return;
		}
		
		private string getEnglishMonthName(string month)
		{
			switch(month)
			{
				case "1":
					return "January";
				case "2":
					return "February";
				case "3":
					return "March";	
				case "4":
					return "April";	
				case "5":
					return "May";
				case "6":
					return "June";
				case "7":
					return "July";	
				case "8":
					return "August";	
				case "9":
					return "September";	
				case "10":
					return "October";	
				case "11":
					return "November";	
				case "12":
					return "December";	
			}
			return "";
		}
				
		public void SetValues4Input2TWClose()
		{
			DataSet srcDs, destDs;
			srcDs = frmInput.ds;
			destDs = frmOutput.ds;
		
			foreach (DataRow srcDr in srcDs.Tables["Input2TWClose"].Rows)
			{		
				DateTime TargetComplDate = CalcTargetDueDate(srcDr["Sales Orders - Order  Reason"].ToString(),
				                                             srcDr["Sales Orders Line - Delivery SLA"].ToString(),
				                                             srcDr["IC Shipments - Arrival Date"].ToString(),
				                                             srcDr["Sales Orders - Creation Date"].ToString(),
				                                             srcDr["Sales Orders - Sales Order Number"].ToString() + "\\" 
				                                             + srcDr["Sales Orders Line - LineNumber"].ToString() );
//			
				if ( TargetComplDate.Year == 1900 )
				{
					continue; //"GI DATA" is empty, just exclude
				}
				
				if (this.CustomerNameFilter4TW(srcDr["Sales Orders - ShipTo Customer Name"].ToString()) == true &&
				    this.CurrentWeekFilter(TargetComplDate) == true)
				{				                               
					DataRow dr = destDs.Tables["OTD"].NewRow();
											
					try
					{
						if ( srcDr["Sales Orders - Order  Reason"].ToString().Substring(0, 1) == "A" ) 
						{
							dr["Service Type"] = "D+";
						}
						else if ( srcDr["Sales Orders - Order  Reason"].ToString().Substring(0, 1) == "H" ) 
						{
							dr["Service Type"] = "H+";
						}
						else 
						{
							dr["Service Type"] = "RFS";
						}
					}
					catch
					{
						string msg = "OTD-TWClose: Error Occurred when processing 'SLA' field "
						+ "in the record 'Unique Identification' = "
						+ srcDr["Sales Orders - Sales Order Number"] + "\\" 
						+ srcDr["Sales Orders Line - LineNumber"];

						MessageBox.Show(msg);
					}
					
					dr["Contractual Y/N ?"] = "Y";
					dr["Customer Type"] = "External";
					dr["DUE DATE (FINAL)"] = TargetComplDate;
					dr["Sales Orders - ShipTo Country"] = "TW";
					dr["Sales Orders SoldTo - Customer Name"] = srcDr["Sales Orders - ShipTo Customer Name"];
					
					try
					{
						if (srcDr["OC Delivery Header -  GI_Date"].ToString() == "")
						{
							dr["Net OTD Status (1=OK; 0=NOK)"] = "0";
						}
						else if ( TargetComplDate >= Convert.ToDateTime(srcDr["OC Delivery Header -  GI_Date"].ToString()) )
						{
							dr["Net OTD Status (1=OK; 0=NOK)"] = "1";
						}
						else{
							dr["Net OTD Status (1=OK; 0=NOK)"] = "0";
						}
					}
					catch
					{
						string msg = "OTD-TWClose: Error Occurred when processing 'In TAT' field "
						+ "in the record 'Unique Identification' = "
						+ srcDr["Sales Orders - Sales Order Number"] + "\\" 
						+ srcDr["Sales Orders Line - LineNumber"];

						MessageBox.Show(msg);
					}
				
					dr["Sales Orders - Sales Order Number"] = srcDr["Sales Orders - Sales Order Number"];
					dr["Sales Orders Line - LineNumber"] = srcDr["Sales Orders Line - LineNumber"];
					dr["IC Delivery Line -  RcvdMaterial"] = srcDr["IC Delivery Line -  RcvdMaterial"];
					dr["IC Delivery Line -  RcvdSerialNumber"] = srcDr["IC Delivery Line -  RcvdSerialNumber"];
					
					if (dr["Service Type"].ToString() == "H+" || dr["Service Type"].ToString() == "D+")
					{
						dr["SLA START DATE (FINAL)"] = srcDr["Sales Orders - Creation Date"];
					}
					else if (dr["Service Type"].ToString() == "RFS")
					{
						dr["SLA START DATE (FINAL)"] = srcDr["IC Shipments - Arrival Date"];
					}
				
					
					dr["OC Delivery Line -  SentMaterial"] = srcDr["OC Delivery Line -  SentMaterial"];
					dr["OC Delivery Line -  SentSerialNumber"] = srcDr["OC Delivery Line -  SentSerialNumber"];
					dr["Cares-Customer Requested Date"] = "";
					dr["DELIVERY DATE (FINAL)"] = srcDr["OC Delivery Header -  GI_Date"];
					dr["Sales Orders - RSCIC/RSLC (Sales Organisation)"] = "TW01";
					
					
					
					dr["Sales Orders - Sold To"] = findSoldToCode(srcDr["Sales Orders - ShipTo Customer Name"].ToString());
					dr["Work Orders  Header - Vendor Name"] = "";
					dr["Sales Orders Line - Customer Warranty"] = "";
					dr["Merged SPT"] = srcDr["Sales Orders Line - SPT"];
					dr["Cares-SA"] = "";
					dr["OTDC_Correction"] = "";
					dr["OTDC_Failure Reason"] = "";
					dr["OTDC_Failure Description"] = "";
					dr["Bulk OTD Status (1=OK; 0=NOK)"] = dr["Net OTD Status (1=OK; 0=NOK)"];
					dr["OTDC_Comments"] = "";
					dr["SO CREATION DATE"] = srcDr["Sales Orders - Creation Date"];
					dr["SLA Calculation (FINAL)"] = CalcSLANumberofDays4TW(srcDr["Sales Orders - Order  Reason"].ToString(),
					                       						srcDr["Sales Orders Line - Delivery SLA"].ToString());
					dr["TAT"] = "";
					dr["Over Due (Days)"] = "";
					string productName = "";
					string bizLine = "";
					findProductNameAndBizLine(srcDr["Sales Orders Line - Material"].ToString(), out productName, out bizLine);
					dr["Product Name"] = productName;
					dr["Business Line"] = bizLine;
					dr["Sales Orders - ShipTo Customer Name"] = srcDr["Sales Orders - ShipTo Customer Name"];
					dr["SO-SO LINE"] = dr["Sales Orders - Sales Order Number"].ToString() + "-" + dr["Sales Orders Line - LineNumber"].ToString();
					dr["DUE WEEK"] = frmConfig.CurrentWeek;
					dr["DUE MONTH"] = getEnglishMonthName(TargetComplDate.Month.ToString());
					dr["Sales Orders - RMA"] = srcDr["Sales Orders - RMA"];
					dr["Cares-Customer Ticket"] = "";
					dr["Cares-Customer Reference"] = "";
					dr["Work Orders  Line - Repair Code"] = "";
					dr["REPAIR TAT"] = "";
					dr["POD STATUS"] = "";
					dr["Work Orders  Line - Sequence"] = "";
					dr["COUNTRY"] = "TW";
					dr["Sales Orders Line  - Product hierarchy Description"] = "";
					dr["SLA SUSPENSION DAYS"] = "";
			
					destDs.Tables["OTD"].Rows.Add(dr);
				}
			}
		}
				
		public void SetValues4Input3TWOpen()
		{
			DataSet srcDs, destDs;
			srcDs = frmInput.ds;
			destDs = frmOutput.ds;
			
			foreach (DataRow srcDr in srcDs.Tables["Input3TWOpen"].Rows)
			{
				DateTime TargetComplDate = CalcTargetDueDate(srcDr["Sales Orders - Order  Reason"].ToString(),
				                                             srcDr["Sales Orders Line - Delivery SLA"].ToString(),
				                                             srcDr["IC Shipments - Arrival Date"].ToString(),
				                                             srcDr["Sales Orders - Creation Date"].ToString(),
				                                             srcDr["Sales Orders - Sales Order Number"].ToString() + "\\" 
				                                             + srcDr["Sales Orders Line - LineNumber"].ToString() );
//			
				if ( TargetComplDate.Year == 1900 )
				{
					continue; //"GI DATA" is empty, just exclude
				}
				
				if (this.CustomerNameFilter4TW(srcDr["Sales Orders - ShipTo Customer Name"].ToString()) == true &&
				    this.CurrentWeekFilter(TargetComplDate) == true)
				{				                               
					DataRow dr = destDs.Tables["OTD"].NewRow();
											
					try
					{
						if ( srcDr["Sales Orders - Order  Reason"].ToString().Substring(0, 1) == "A" ) 
						{
							dr["Service Type"] = "D+";
						}
						else if ( srcDr["Sales Orders - Order  Reason"].ToString().Substring(0, 1) == "H" ) 
						{
							dr["Service Type"] = "H+";
						}
						else 
						{
							dr["Service Type"] = "RFS";
						}
					}
					catch
					{
						string msg = "OTD-TWOpen: Error Occurred when processing 'SLA' field "
						+ "in the record 'Unique Identification' = "
						+ srcDr["Sales Orders - Sales Order Number"] + "\\" 
						+ srcDr["Sales Orders Line - LineNumber"];

						MessageBox.Show(msg);
					}
					
					dr["Contractual Y/N ?"] = "Y";
					dr["Customer Type"] = "External";
					dr["DUE DATE (FINAL)"] = TargetComplDate;
					dr["Sales Orders - ShipTo Country"] = "TW";
					dr["Sales Orders SoldTo - Customer Name"] = srcDr["Sales Orders - ShipTo Customer Name"];
					
					try
					{
						if (srcDr["OC Delivery Header -  GI_Date"].ToString() == "")
						{
							dr["Net OTD Status (1=OK; 0=NOK)"] = "0";
						}
						else if ( TargetComplDate >= Convert.ToDateTime(srcDr["OC Delivery Header -  GI_Date"].ToString()) )
						{
							dr["Net OTD Status (1=OK; 0=NOK)"] = "1";
						}
						else{
							dr["Net OTD Status (1=OK; 0=NOK)"] = "0";
						}
					}
					catch
					{
						string msg = "OTD-TWOpen: Error Occurred when processing 'Net OTD Status (1=OK; 0=NOK)' field "
						+ "in the record 'Unique Identification' = "
						+ srcDr["Sales Orders - Sales Order Number"] + "\\" 
						+ srcDr["Sales Orders Line - LineNumber"];

						MessageBox.Show(msg);
					}
				
					dr["Sales Orders - Sales Order Number"] = srcDr["Sales Orders - Sales Order Number"];
					dr["Sales Orders Line - LineNumber"] = srcDr["Sales Orders Line - LineNumber"];
					dr["IC Delivery Line -  RcvdMaterial"] = srcDr["IC Delivery Line -  RcvdMaterial"];
					dr["IC Delivery Line -  RcvdSerialNumber"] = srcDr["IC Delivery Line -  RcvdSerialNumber"];
					
					if (dr["Service Type"].ToString() == "H+" || dr["Service Type"].ToString() == "D+")
					{
						dr["SLA START DATE (FINAL)"] = srcDr["Sales Orders - Creation Date"];
					}
					else if (dr["Service Type"].ToString() == "RFS")
					{
						dr["SLA START DATE (FINAL)"] = srcDr["IC Shipments - Arrival Date"];
					}
				
					
					dr["OC Delivery Line -  SentMaterial"] = srcDr["OC Delivery Line -  SentMaterial"];
					dr["OC Delivery Line -  SentSerialNumber"] = srcDr["OC Delivery Line -  SentSerialNumber"];
					dr["Cares-Customer Requested Date"] = "";
					dr["DELIVERY DATE (FINAL)"] = srcDr["OC Delivery Header -  GI_Date"];
					dr["Sales Orders - RSCIC/RSLC (Sales Organisation)"] = "TW01";
					dr["Sales Orders - Sold To"] = findSoldToCode(srcDr["Sales Orders - ShipTo Customer Name"].ToString());
					dr["Work Orders  Header - Vendor Name"] = "";
					dr["Sales Orders Line - Customer Warranty"] = "";
					dr["Merged SPT"] = srcDr["Sales Orders Line - SPT"];
					dr["Cares-SA"] = "";
					dr["OTDC_Correction"] = "";
					dr["OTDC_Failure Reason"] = "";
					dr["OTDC_Failure Description"] = "";
					dr["Bulk OTD Status (1=OK; 0=NOK)"] = dr["Net OTD Status (1=OK; 0=NOK)"];
					dr["OTDC_Comments"] = "";
					dr["SO CREATION DATE"] = srcDr["Sales Orders - Creation Date"];
					dr["SLA Calculation (FINAL)"] = CalcSLANumberofDays4TW(srcDr["Sales Orders - Order  Reason"].ToString(),
					                       						srcDr["Sales Orders Line - Delivery SLA"].ToString());
					dr["TAT"] = "";
					dr["Over Due (Days)"] = "";
					string productName = "";
					string bizLine = "";
					findProductNameAndBizLine(srcDr["Sales Orders Line - Material"].ToString(), out productName, out bizLine);
					dr["Product Name"] = productName;
					dr["Business Line"] = bizLine;
					dr["Sales Orders - ShipTo Customer Name"] = srcDr["Sales Orders - ShipTo Customer Name"];
					dr["SO-SO LINE"] = dr["Sales Orders - Sales Order Number"].ToString() + "-" + dr["Sales Orders Line - LineNumber"].ToString();
					dr["DUE WEEK"] = frmConfig.CurrentWeek;
					dr["DUE MONTH"] = getEnglishMonthName(TargetComplDate.Month.ToString());
					dr["Sales Orders - RMA"] = srcDr["Sales Orders - RMA"];
					dr["Cares-Customer Ticket"] = "";
					dr["Cares-Customer Reference"] = "";
					dr["Work Orders  Line - Repair Code"] = "";
					dr["REPAIR TAT"] = "";
					dr["POD STATUS"] = "";
					dr["Work Orders  Line - Sequence"] = "";
					dr["COUNTRY"] = "TW";
					dr["Sales Orders Line  - Product hierarchy Description"] = "";
					dr["SLA SUSPENSION DAYS"] = "";
			
					destDs.Tables["OTD"].Rows.Add(dr);
				}
			}
		}
				
		public void SetValues4Input4KOREA()
		{
			DataSet srcDs, destDs;
			srcDs = frmInput.ds;
			destDs = frmOutput.ds;
			
//			destDs.Tables["Onelog"].Clear();
			
			foreach (DataRow srcDr in srcDs.Tables["Input4KOREA"].Rows)
			{ 
				if (this.CustomerNameFilter4TW(srcDr["Sales Orders - ShipTo Customer Name"].ToString()) == true)
				{
					DataRow dr = destDs.Tables["OTD"].NewRow();
											
					dr["Service Type"] = srcDr["Service Type"];
					dr["Contractual Y/N ?"] = srcDr["Contractual Y/N ?"];
					dr["Customer Type"] = srcDr["Customer Type"];
					dr["DUE DATE (FINAL)"] = srcDr["DUE DATE (FINAL)"];
					dr["Sales Orders - ShipTo Country"] = "KR";
					dr["Sales Orders SoldTo - Customer Name"] = srcDr["Sales Orders SoldTo - Customer Name"];
					dr["Net OTD Status (1=OK; 0=NOK)"] = srcDr["Bulk OTD Status (1=OK; 0=NOK)"];
					dr["Sales Orders - Sales Order Number"] = srcDr["Sales Orders - Sales Order Number"];
					dr["Sales Orders Line - LineNumber"] = srcDr["Sales Orders Line - LineNumber"];
					dr["IC Delivery Line -  RcvdMaterial"] = srcDr["IC Delivery Line -  RcvdMaterial"];
					dr["IC Delivery Line -  RcvdSerialNumber"] = srcDr["IC Delivery Line -  RcvdSerialNumber"];
					dr["SLA START DATE (FINAL)"] = srcDr["SLA START DATE (FINAL)"];
					dr["OC Delivery Line -  SentMaterial"] = srcDr["OC Delivery Line -  SentMaterial"];
					dr["OC Delivery Line -  SentSerialNumber"] = srcDr["OC Delivery Line -  SentSerialNumber"];
					dr["Cares-Customer Requested Date"] = srcDr["Cares-Customer Requested Date"];
					dr["DELIVERY DATE (FINAL)"] = srcDr["DELIVERY DATE (FINAL)"];
					dr["Sales Orders - RSCIC/RSLC (Sales Organisation)"] = srcDr["Sales Orders - RSCIC/RSLC (Sales Organisation)"];
					dr["Sales Orders - Sold To"] = findSoldToCode(srcDr["Sales Orders SoldTo - Customer Name"].ToString());
					dr["Work Orders  Header - Vendor Name"] = srcDr["Work Orders  Header - Vendor Name"];
					dr["Sales Orders Line - Customer Warranty"] = srcDr["Sales Orders Line - Customer Warranty"];
					dr["Merged SPT"] = srcDr["Merged SPT"];
					dr["Cares-SA"] = srcDr["Cares-SA"];
					dr["OTDC_Correction"] = srcDr["OTDC_Correction"];
					dr["OTDC_Failure Reason"] = srcDr["OTDC_Failure Reason"];
					dr["OTDC_Failure Description"] = srcDr["OTDC_Failure Reason"];
					dr["Bulk OTD Status (1=OK; 0=NOK)"] = srcDr["Bulk OTD Status (1=OK; 0=NOK)"];
					dr["OTDC_Comments"] = srcDr["OTDC_Comments"];
					dr["SO CREATION DATE"] = srcDr["SO CREATION DATE"];
					dr["SLA Calculation (FINAL)"] = srcDr["SLA Calculation (FINAL)"];
					dr["TAT"] = srcDr["TAT"];
					dr["Over Due (Days)"] = srcDr["Over Due (Days)"];
					string productName = "";
					string bizLine = "";
					findProductNameAndBizLine(srcDr["IC Delivery Line -  RcvdMaterial"].ToString(), out productName, out bizLine);
					dr["Product Name"] = productName;
					dr["Business Line"] = bizLine;
					dr["Sales Orders - ShipTo Customer Name"] = srcDr["Sales Orders - ShipTo Customer Name"];
					dr["SO-SO LINE"] = "";
					
					try
					{
						dr["DUE MONTH"] = getEnglishMonthName(Convert.ToDateTime(srcDr["DUE DATE (FINAL)"].ToString()).Month.ToString());
						//dr["DUE WEEK"] = frmConfig.CurrentWeek;
					}
					catch
					{
					}
					dr["DUE WEEK"] = frmConfig.CurrentWeek;
					//dr["DUE MONTH"] = srcDr["DUE MONTH"];
					dr["Sales Orders - RMA"] = srcDr["Sales Orders - RMA"];
					dr["Cares-Customer Ticket"] = srcDr["Cares-Customer Ticket"];
					dr["Cares-Customer Reference"] = srcDr["Cares-Customer Reference"];
					dr["Work Orders  Line - Repair Code"] = srcDr["Work Orders  Line - Repair Code"];
					dr["REPAIR TAT"] = srcDr["REPAIR TAT"];
					dr["POD STATUS"] = srcDr["POD STATUS"];
					dr["Work Orders  Line - Sequence"] = srcDr["Work Orders  Line - Sequence"];
					dr["COUNTRY"] = "KOREA  REPUBLIC OF"; //srcDr["COUNTRY"];
					dr["Sales Orders Line  - Product hierarchy Description"] = srcDr["Sales Orders Line  - Product hierarchy Description"];
					dr["SLA SUSPENSION DAYS"] = srcDr["SLA SUSPENSION DAYS"];
			
					destDs.Tables["OTD"].Rows.Add(dr);
				}
			}
		}			
			
		public void SetValues4NZ()
		{
			DataSet srcDs, destDs;
			srcDs = frmInput.ds;
			destDs = frmOutput.ds;
			
			foreach (DataRow srcDr in srcDs.Tables["Input5NZ"].Rows)
			{
			  if (this.CustomerNameFilter4TW(srcDr["Sales Orders - ShipTo Customer Name"].ToString()) == true)
			  {
				DataRow dr = destDs.Tables["OTD"].NewRow();
				
				if (srcDr["Customer TAT"].ToString() == "1")
				{
					dr["Service Type"] = "H+";
				}
				else 
				{
					dr["Service Type"] = "RFS";
				}
				
				dr["Contractual Y/N ?"] = "Y";
				dr["Customer Type"] = "External";
				
				try
				{
					if (srcDr["Target Due Date"].ToString() == "")
					{
						DateTime dt = Convert.ToDateTime(srcDr["Customer Request Date"].ToString()).AddDays(1);
						dr["DUE DATE (FINAL)"]  = dt.ToString();
					}
					else
					{
						dr["DUE DATE (FINAL)"]  = srcDr["Target Due Date"].ToString();						
					}
				}
				catch
				{
					string msg = "NonOnelog: Error Occurred when processing 'Target Due Date' field "
						+ "in the record 'Part Number' = " + dr["Part Number"].ToString();

					MessageBox.Show(msg);
				}
				if (srcDr["Country"].ToString() == "PACIFIC ISLANDS")
				{
					dr["Sales Orders - ShipTo Country"] = "PI";
				}
				else if (srcDr["Country"].ToString() == "NEW ZEALAND")
				{
					dr["Sales Orders - ShipTo Country"] = "NZ";
				}
				else
				{
					dr["Sales Orders - ShipTo Country"] = srcDr["Country"];
				}
				dr["Sales Orders SoldTo - Customer Name"] = srcDr["Customer Name"];
					
				try
				{
					if (srcDr["Dispatched Date to Customer"].ToString() == "")
					{
						dr["Net OTD Status (1=OK; 0=NOK)"] = "0";
					}
					else
					{
						DateTime dtTargetDueDate, dtDispatchedDateToCustomer;
						dtTargetDueDate = Convert.ToDateTime(dr["DUE DATE (FINAL)"].ToString()).AddDays(1);
						dtDispatchedDateToCustomer = Convert.ToDateTime(srcDr["Dispatched Date to Customer"].ToString()).AddDays(1);
					
						if (dtDispatchedDateToCustomer > dtTargetDueDate)
						{
							dr["Net OTD Status (1=OK; 0=NOK)"] = "0";
						}
						else
						{
							dr["Net OTD Status (1=OK; 0=NOK)"] = "1";
						}
					}
				}
				catch
				{
					string msg = "NonOnelog: Error Occurred when processing 'Net OTD Status (1=OK; 0=NOK)' field "
						+ "in the record 'Part Number' = " + srcDr["Part Number"].ToString();

					MessageBox.Show(msg);
				}
					
				dr["Sales Orders - Sales Order Number"] = "'" + srcDr["RMA#"].ToString();
				if (dr["Service Type"].ToString() == "RFS")
				{
					dr["Sales Orders Line - LineNumber"] = "2";
					dr["SLA START DATE (FINAL)"] = srcDr["Customer Request Date"];
				}
				else
				{
					dr["Sales Orders Line - LineNumber"] = "1";
					dr["SLA START DATE (FINAL)"] = srcDr["Faulty Board Received Date from Customer"];
				}
				
				dr["IC Delivery Line -  RcvdMaterial"] = "'" + srcDr["Part Number"].ToString().Replace(" ", "");
				dr["IC Delivery Line -  RcvdSerialNumber"] = srcDr["Serial No"];
				dr["OC Delivery Line -  SentMaterial"] = "'" + srcDr["Part Number"].ToString().Replace(" ", "");
				dr["OC Delivery Line -  SentSerialNumber"] = "";
				dr["Cares-Customer Requested Date"] = "";
				dr["DELIVERY DATE (FINAL)"] = srcDr["Dispatched Date to Customer"];
				dr["Sales Orders - RSCIC/RSLC (Sales Organisation)"] = "NZ01";
				dr["Sales Orders - Sold To"] = findSoldToCode(srcDr["Customer Name"].ToString());
				dr["Work Orders  Header - Vendor Name"] = srcDr["RSLC/Repairer "];
				dr["Sales Orders Line - Customer Warranty"] = srcDr["Warranty Status W/OW"].ToString();
				dr["Merged SPT"] = "";
				dr["Cares-SA"] = "";
				dr["OTDC_Correction"] = "";
				dr["OTDC_Failure Reason"] = "";
				dr["OTDC_Failure Description"] = "";
				dr["Bulk OTD Status (1=OK; 0=NOK)"] = dr["Net OTD Status (1=OK; 0=NOK)"];
				dr["OTDC_Comments"] = "";
				dr["SO CREATION DATE"] = srcDr["Customer Request Date"];
				dr["SLA Calculation (FINAL)"] = srcDr["Customer TAT"];
				dr["TAT"] = "";
				dr["Over Due (Days)"] = srcDr["Days Overdue"];
				string partNumber = srcDr["Part Number"].ToString().Replace(" ", "");
				string productName = "";
				string bizLine = "";
				findProductNameAndBizLine(partNumber, out productName, out bizLine);
				dr["Product Name"] = productName;
				dr["Business Line"] = bizLine;
				dr["Sales Orders - ShipTo Customer Name"] = srcDr["Customer Name"];
				dr["SO-SO LINE"] = dr["Sales Orders - Sales Order Number"].ToString() + "-" + dr["Sales Orders Line - LineNumber"].ToString();
				dr["DUE WEEK"] = frmConfig.CurrentWeek;
				if (dr["DUE DATE (FINAL)"].ToString() != "")
				{
					dr["DUE MONTH"] = getEnglishMonthName(Convert.ToDateTime(dr["DUE DATE (FINAL)"].ToString()).Month.ToString());
				}
				dr["Sales Orders - RMA"] = "";
				dr["Cares-Customer Ticket"] = "";
				dr["Cares-Customer Reference"] = "";
				dr["Work Orders  Line - Repair Code"] = "";
				dr["REPAIR TAT"] = srcDr["Repairer Actual TAT"];
				dr["POD STATUS"] = "";
				dr["Work Orders  Line - Sequence"] = "";
				dr["COUNTRY"] = srcDr["Country"];
				dr["Sales Orders Line  - Product hierarchy Description"] = "";
				dr["SLA SUSPENSION DAYS"] = "";
			
				destDs.Tables["OTD"].Rows.Add(dr);			
			  }
			}
		}
	}
}
