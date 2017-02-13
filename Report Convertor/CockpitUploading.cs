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
	
	public class CockpitUploading
	{
		
		public CockpitUploading()
		{
			
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
				
				
		public void SetValues4Input4KOREA()
		{
			DataSet srcDs, destDs;
			srcDs = frmInput.ds;
			destDs = frmOutput.ds;
			
			foreach (DataRow srcDr in srcDs.Tables["Input4KOREA"].Rows)
			{
				DataRow dr = destDs.Tables["Cockpit RFS uploading"].NewRow();
				
				dr["Service Type (RFS)"] = "RFS";
				dr["Contracted Y/N ?"] = "Y";					
				dr["Customer Type (Internal/External) ?"] = "External";	
				dr["Customer Due Date (Day)"] = srcDr["Target Compl Date"];	
				dr["Sales Orders - ShipTo Country"] = "KOREA  REPUBLIC OF";	
				dr["Sales Orders SoldTo - Customer Name"] = srcDr["Customer Name"];	
				dr["Sales Orders - Sales Order Number"] = srcDr["New Serial Number"];	
				dr["Sales Orders Line - LineNumber"] = "2";	
				dr["Merge R4S Material"] = srcDr["Part Number"];	
				dr["IC Delivery Line -  RcvdSerialNumber"] = srcDr["IC Delivery Line-RcvdSerialNumber"];
				dr["IC Shipments - Arrival Date"] = srcDr["IC Shipments - Arrival Date"];
				dr["Merge Sent Material"] = srcDr["Part Number"];	
				dr["OC Delivery Line -  SentSerialNumber"] = srcDr["New Serial Number"];	
				dr["CARES - Customer Requested Date"] = "";	//??
				dr["POD Date (Final Calculation)"] = srcDr["Ship_Date_Repl"];					
				dr["Merge Sales Organization"] = "";//"KOREA  REPUBLIC OF";	
				dr["Sales Order - Sold To"] = "";	//??
				dr["Work Orders  Header - Vendor Name"] = srcDr["Repairer"];	
				dr["Sales Orders Line - Customer Warranty"] = "N";	
				dr["Merge SPT"] = "";	//??
				dr["CARES - Service Agreement"] = "";	//??	
				dr["OTDC_Correction"] = "";	
				dr["OTDC_Failure Reason"] = "";	
				dr["OTDC_Failure Description"] = "";	
				
				try
				{
					DateTime dt1 = Convert.ToDateTime(srcDr["Target Compl Date"].ToString()).Date;
					DateTime dt2 = Convert.ToDateTime(srcDr["Ship_Date_Repl"].ToString()).Date;
					
					if ( dt1 >= dt2 )
					{
						dr["Net OTD Status (1=OK, 0=NOK)"] = "1";
						dr["Bulk OTD Status (1=OK, 0=NOK)"] = "";
					}
					else {
						dr["Net OTD Status (1=OK, 0=NOK)"] = "0";
						dr["Bulk OTD Status (1=OK, 0=NOK)"] ="";
					}
				}
				catch
				{
//					string msg = "Onelog-KOREA: Error Occurred when processing 'Delivery Status' field "
//						+ "in the record 'Unique Identification' = "
//							+ srcDr["New Serial Number"];
//						
//						MessageBox.Show(msg);
				}				
							
				destDs.Tables["Cockpit RFS uploading"].Rows.Add(dr);
			}
		}
		
		
		public void SetValues4Input5NZ()
		{
			DataSet srcDs, destDs;
			srcDs = frmInput.ds;
			destDs = frmOutput.ds;
			int customerTAT = -1;
			
			foreach (DataRow srcDr in srcDs.Tables["Input5NZ"].Rows)
			{
				customerTAT = -1;
				try {
					customerTAT = Convert.ToInt32(srcDr["Customer TAT"].ToString());
				}
				catch{
					continue;
				}
				
				if ( customerTAT == -1 )
				{
					continue;
				}
				else if ( customerTAT != 1 )
				{
					//++++++++++ R4S +++++++++++++++//	
					DataRow dr = destDs.Tables["Cockpit RFS uploading"].NewRow();
				
					dr["Service Type (RFS)"] = "RFS";
					if ( srcDr["Customer Name"].ToString().Contains("Alcatel") ||
					    srcDr["Customer Name"].ToString().Contains("ALU"))
					{
						dr["Contracted Y/N ?"] = "N";	
						dr["Customer Type (Internal/External) ?"] = "Internal";	
					}
					else {
						dr["Contracted Y/N ?"] = "Y";	
						dr["Customer Type (Internal/External) ?"] = "External";	
					}
					
					dr["Customer Due Date (Day)"] = srcDr["Target Due Date"];	
					dr["Sales Orders - ShipTo Country"] = srcDr["Country"];	
					dr["Sales Orders SoldTo - Customer Name"] = srcDr["Customer Name"];	
					dr["Sales Orders - Sales Order Number"] = srcDr["RMA#"];	
					dr["Sales Orders Line - LineNumber"] = "2";	
					dr["Merge R4S Material"] = srcDr["Part Number"];	
					dr["IC Delivery Line -  RcvdSerialNumber"] = srcDr["Serial No"];	
					dr["IC Shipments - Arrival Date"] = srcDr["Faulty Board Received Date from Customer"];
					dr["Merge Sent Material"] = "";	
					dr["OC Delivery Line -  SentSerialNumber"] = "";	
					dr["CARES - Customer Requested Date"] = srcDr["Customer Request Date"];
					dr["POD Date (Final Calculation)"] = srcDr["Dispatched Date to Customer"];					
					dr["Merge Sales Organization"] = "";//srcDr["Country"];	
					dr["Sales Order - Sold To"] = "";	
					dr["Work Orders  Header - Vendor Name"] = srcDr["RSLC/Repairer "];						
					if (srcDr["Warranty StatusW/OW"].ToString() == "OW")
					{
						dr["Sales Orders Line - Customer Warranty"] = "N";
					}
					else
					{
						dr["Sales Orders Line - Customer Warranty"] = srcDr["Warranty StatusW/OW"];	
					}
					dr["Merge SPT"] = "";	
					dr["CARES - Service Agreement"] = "";	//??	
					dr["OTDC_Correction"] = "";	
					dr["OTDC_Failure Reason"] = "";	
					dr["OTDC_Failure Description"] = "";	
				
					try
					{
						if (srcDr["Dispatched Date to Customer"].ToString() == "")
						{
							dr["Net OTD Status (1=OK, 0=NOK)"] = "0";
							dr["Bulk OTD Status (1=OK, 0=NOK)"] ="";
						}
						else
						{
							DateTime dtTargetDueDate, dtDispatchedDateToCustomer;
						
							if (srcDr["Target Due Date"].ToString() == "")
							{
								dtTargetDueDate = Convert.ToDateTime(srcDr["Customer Request Date"].ToString()).AddDays(1);
							}
							else
							{
								dtTargetDueDate = Convert.ToDateTime(srcDr["Target Due Date"].ToString()).AddDays(1);
							}
													
							dtDispatchedDateToCustomer = Convert.ToDateTime(srcDr["Dispatched Date to Customer"].ToString()).AddDays(1);
					
							if (dtDispatchedDateToCustomer > dtTargetDueDate)
							{
								dr["Net OTD Status (1=OK, 0=NOK)"] = "0";
								dr["Bulk OTD Status (1=OK, 0=NOK)"] ="";
							}
							else
							{
								dr["Net OTD Status (1=OK, 0=NOK)"] = "1";
								dr["Bulk OTD Status (1=OK, 0=NOK)"] = "";
							}
						}
					}
					catch
					{
//						string msg = "NonOnelog: Error Occurred when processing 'On Time' field "
//							+ "in the record 'Part Number' = " + dr["Part Number"].ToString();
//
						MessageBox.Show("Net OTD Status (1=OK, 0=NOK is empty.");
					}
							
					destDs.Tables["Cockpit RFS uploading"].Rows.Add(dr);
				}
				else
				{
					//++++++++++ AE +++++++++++++++//				
					DataRow dr = destDs.Tables["Cockpit AE uploading"].NewRow();
				
					dr["Service Type"] = "H+";
					if ( srcDr["Customer Name"].ToString().Contains("Alcatel") ||
					    srcDr["Customer Name"].ToString().Contains("ALU"))
					{
						dr["Contracted Y/N ?"] = "N";	
						dr["Customer Type (Internal/External) ?"] = "Internal";	
					}
					else {
						dr["Contracted Y/N ?"] = "Y";	
						dr["Customer Type (Internal/External) ?"] = "External";	
					}
					
					try
					{
						DateTime dt = Convert.ToDateTime(srcDr["Customer Request Date"].ToString()).AddDays(1);
						dr["Customer Due Date (Day)"]  = dt.ToString();
					}
					catch{	}
				
					dr["Sales Orders - ShipTo Country"] = srcDr["Country"];		
					dr["Sales Orders SoldTo - Customer Name"] = srcDr["Customer Name"];	
					dr["Sales Orders - Sales Order Number"] = srcDr["RMA#"];	
					dr["Sales Orders Line - LineNumber"] = "1";	
					dr["IC Delivery Line -  RcvdMaterial"] = srcDr["Part Number"];	
					dr["IC Delivery Line -  RcvdSerialNumber"] = srcDr["Serial No"];	
//					dr["IC Shipments - Arrival Date"] = srcDr["Faulty Board Received Date from Customer"];
					dr["Merge AE Material"] = dr["IC Delivery Line -  RcvdMaterial"];//"";	
					dr["OC Delivery Line -  SentSerialNumber"] = "";		
					dr["CARES - Customer Requested Date"] = srcDr["Customer Request Date"];
					dr["POD Date (Final Calculation)"] = srcDr["Dispatched Date to Customer"];					
					dr["Merge Sales Organization"] = srcDr["Country"];	
					dr["Sales Order - Sold To"] = this.findSoldToCode(dr["Sales Orders SoldTo - Customer Name"].ToString());//"";
					dr["Work Orders  Header - Vendor Name"] = srcDr["RSLC/Repairer "];	
					if (srcDr["Warranty StatusW/OW"].ToString() == "OW")
					{
						dr["Sales Orders Line - Customer Warranty"] = "N";
					}
					else
					{
						dr["Sales Orders Line - Customer Warranty"] = srcDr["Warranty StatusW/OW"];	
					}
					dr["Merge SPT"] = "";	
					dr["CARES - Service Agreement"] = "";	
					dr["OTDC_Correction"] = "";	
					dr["OTDC_Failure Reason"] = "";	
					dr["OTDC_Failure Description"] = "";	
				
					try
					{
						if (srcDr["Dispatched Date to Customer"].ToString() == "")
						{
							dr["Net OTD Status (1=OK, 0=NOK)"] = "0";
							dr["Bulk OTD Status (1=OK, 0=NOK)"] ="";
						}
						else
						{
							DateTime dtTargetDueDate, dtDispatchedDateToCustomer;
						
							if (srcDr["Target Due Date"].ToString() == "")
							{
								dtTargetDueDate = Convert.ToDateTime(srcDr["Customer Request Date"].ToString()).AddDays(1);
							}
							else
							{
								dtTargetDueDate = Convert.ToDateTime(dr["Target Due Date"].ToString()).AddDays(1);
							}
													
							dtDispatchedDateToCustomer = Convert.ToDateTime(srcDr["Dispatched Date to Customer"].ToString()).AddDays(1);
					
							if (dtDispatchedDateToCustomer > dtTargetDueDate)
							{
								dr["Net OTD Status (1=OK, 0=NOK)"] = "0";
								dr["Bulk OTD Status (1=OK, 0=NOK)"] ="";
							}
							else
							{
								dr["Net OTD Status (1=OK, 0=NOK)"] = "1";
								dr["Bulk OTD Status (1=OK, 0=NOK)"] = "";
							}
						}
					}
					catch
					{
//						string msg = "NonOnelog: Error Occurred when processing 'On Time' field "
//							+ "in the record 'Part Number' = " + dr["Part Number"].ToString();
//
//						MessageBox.Show(msg);
					}
							
					destDs.Tables["Cockpit AE uploading"].Rows.Add(dr);
				
				}
				
			}
		}
		
		//Weekly OTD
		public void SetValues4WeeklyOTD()
		{
			DataSet srcDs, destDs;
			srcDs = frmInput.ds;
			destDs = frmOutput.ds;
			
			foreach (DataRow srcDr in srcDs.Tables["WeeklyOTDOnelog"].Rows)
			{
				if ( srcDr["SLA"].ToString() == "R4S" || srcDr["SLA"].ToString() == "Per Use")
				{
					//++++++++++ R4S or Per Use +++++++++++++++//	
					DataRow dr = destDs.Tables["Cockpit RFS uploading"].NewRow();
				
					dr["Service Type (RFS)"] = "RFS";
					if ( srcDr["SLA"].ToString() == "R4S" )
					{
						dr["Contracted Y/N ?"] = "Y";
					}
					else
					{
						dr["Contracted Y/N ?"] = "N";
					}
					dr["Customer Type (Internal/External) ?"] = "External";	
										
					dr["Customer Due Date (Day)"] = srcDr["Target Compl Date"];	
					dr["Sales Orders - ShipTo Country"] = srcDr["Country"];	
					dr["Sales Orders SoldTo - Customer Name"] = srcDr["Customer Name"];	
					dr["Net OTD Status (1=OK, 0=NOK)"] = srcDr["In TAT"];	
					dr["Sales Orders - Sales Order Number"] = srcDr["Rma#"];	
					dr["Sales Orders Line - LineNumber"] = srcDr["Line#"];	
					dr["Merge R4S Material"] = srcDr["Exp_Part_No"];	
					dr["IC Delivery Line -  RcvdSerialNumber"] = "";	
					dr["IC Shipments - Arrival Date"] = srcDr["Rcv_Date_Def"];
					dr["Merge Sent Material"] = "";	
					dr["OC Delivery Line -  SentSerialNumber"] = "";	
					dr["CARES - Customer Requested Date"] = "";
					dr["POD Date (Final Calculation)"] = srcDr["Ship_Date_Repl"];					
					dr["Merge Sales Organization"] = "";//srcDr["Country"];	
					dr["Sales Order - Sold To"] = "";	
					dr["Work Orders  Header - Vendor Name"] = srcDr["Repairer"];	
					dr["Sales Orders Line - Customer Warranty"] = ""; //srcDr["Warranty_Status"];	
					dr["Merge SPT"] = "";	
					dr["CARES - Service Agreement"] = "";	//??	
					dr["OTDC_Correction"] = "";	
					dr["OTDC_Failure Reason"] = srcDr["L-1"];	
					dr["OTDC_Failure Description"] = srcDr["L-2"];	
					dr["Bulk OTD Status (1=OK, 0=NOK)"] ="";
												
					destDs.Tables["Cockpit RFS uploading"].Rows.Add(dr);
				}
				else if ( srcDr["SLA"].ToString() == "H+" || srcDr["SLA"].ToString() == "D+" ) 
				{
					//++++++++++ AE +++++++++++++++//				
					DataRow dr = destDs.Tables["Cockpit AE uploading"].NewRow();
				
					dr["Service Type"] = srcDr["SLA"];
					dr["Contracted Y/N ?"] = "Y";	
					dr["Customer Type (Internal/External) ?"] = "External";	
										
					dr["Customer Due Date (Day)"] = srcDr["Target Compl Date"];	
					dr["Sales Orders - ShipTo Country"] = srcDr["Country"];	
					dr["Sales Orders SoldTo - Customer Name"] = srcDr["Customer Name"];	
					dr["Net OTD Status (1=OK, 0=NOK)"] = srcDr["In TAT"];	
					dr["Sales Orders - Sales Order Number"] = srcDr["Rma#"];	
					dr["Sales Orders Line - LineNumber"] = srcDr["Line#"];		
					dr["IC Delivery Line -  RcvdMaterial"] = srcDr["Exp_Part_No"];	
					dr["IC Delivery Line -  RcvdSerialNumber"] = "";	
					//dr["IC Shipments - Arrival Date"] = srcDr["Rcv_Date_Def"];
					dr["Merge AE Material"] = dr["IC Delivery Line -  RcvdMaterial"];//"";
					dr["OC Delivery Line -  SentSerialNumber"] = "";	
					dr["CARES - Customer Requested Date"] = "";
					dr["POD Date (Final Calculation)"] = srcDr["Ship_Date_Repl"];					
					dr["Merge Sales Organization"] = "";//srcDr["Country"];	
					dr["Sales Order - Sold To"] = this.findSoldToCode(dr["Sales Orders SoldTo - Customer Name"].ToString());//"";	
					dr["Work Orders  Header - Vendor Name"] = srcDr["Repairer"];	
					dr["Sales Orders Line - Customer Warranty"] = ""; //srcDr["Warranty_Status"];	
					dr["Merge SPT"] = "";	
					dr["CARES - Service Agreement"] = "";	//??	
					dr["OTDC_Correction"] = "";	
					dr["OTDC_Failure Reason"] = srcDr["L-1"];	
					dr["OTDC_Failure Description"] = srcDr["L-2"];	
					dr["Bulk OTD Status (1=OK, 0=NOK)"] ="";
							
					destDs.Tables["Cockpit AE uploading"].Rows.Add(dr);
				
				}
				
			}
			
			
			//+++++++++++++++++++
			foreach (DataRow srcDr in srcDs.Tables["WeeklyOTDNonOnelog"].Rows)
			{
				if ( srcDr["SLA"].ToString() == "R4S" || srcDr["SLA"].ToString() == "Per Use")
				{
					//++++++++++ R4S or Per Use +++++++++++++++//	
					DataRow dr = destDs.Tables["Cockpit RFS uploading"].NewRow();
				
					dr["Service Type (RFS)"] = "RFS";
					if ( srcDr["SLA"].ToString() == "R4S" )
					{
						dr["Contracted Y/N ?"] = "Y";
					}
					else
					{
						dr["Contracted Y/N ?"] = "N";
					}
				
					dr["Customer Type (Internal/External) ?"] = "External";	
										
					dr["Customer Due Date (Day)"] = srcDr["Target Due Date"];	
					dr["Sales Orders - ShipTo Country"] = srcDr["Country"];	
					dr["Sales Orders SoldTo - Customer Name"] = srcDr["Customer Name"];	
					dr["Net OTD Status (1=OK, 0=NOK)"] = srcDr["On Time"];	
					dr["Sales Orders - Sales Order Number"] = srcDr["RMA#"];	
					dr["Sales Orders Line - LineNumber"] = 2;	
					dr["Merge R4S Material"] = srcDr["Part Number"];	
					dr["IC Delivery Line -  RcvdSerialNumber"] = srcDr["Serial No"];	
					dr["IC Shipments - Arrival Date"] = srcDr["Faulty Board Received Date from Customer"];
					dr["Merge Sent Material"] = "";	
					dr["OC Delivery Line -  SentSerialNumber"] = "";	
					dr["CARES - Customer Requested Date"] = srcDr["Customer Request Date"];
					dr["POD Date (Final Calculation)"] = srcDr["Dispatched Date to Customer"];					
					dr["Merge Sales Organization"] = "";//srcDr["Country"];	
					dr["Sales Order - Sold To"] = "";	
					dr["Work Orders  Header - Vendor Name"] = srcDr["RSLC/Repairer "];	
					dr["Sales Orders Line - Customer Warranty"] = ""; //srcDr["Warranty StatusW/OW"];	
					dr["Merge SPT"] = "";	
					dr["CARES - Service Agreement"] = "";	//??	
					dr["OTDC_Correction"] = "";	
					dr["OTDC_Failure Reason"] = srcDr["L-1"];	
					dr["OTDC_Failure Description"] = srcDr["L-2"];	
					dr["Bulk OTD Status (1=OK, 0=NOK)"] ="";
												
					destDs.Tables["Cockpit RFS uploading"].Rows.Add(dr);
				}
				else if ( srcDr["SLA"].ToString() == "H+" || srcDr["SLA"].ToString() == "D+" ) 
				{
					//++++++++++ AE +++++++++++++++//				
					DataRow dr = destDs.Tables["Cockpit AE uploading"].NewRow();
				
					dr["Service Type"] = srcDr["SLA"];
					dr["Contracted Y/N ?"] = "Y";	
					dr["Customer Type (Internal/External) ?"] = "External";	
										
					dr["Customer Due Date (Day)"] = srcDr["Target Due Date"];	
					dr["Sales Orders - ShipTo Country"] = srcDr["Country"];	
					dr["Sales Orders SoldTo - Customer Name"] = srcDr["Customer Name"];	
					dr["Net OTD Status (1=OK, 0=NOK)"] = srcDr["On Time"];	
					dr["Sales Orders - Sales Order Number"] = srcDr["RMA#"];	
					dr["Sales Orders Line - LineNumber"] = 1;	
					dr["IC Delivery Line -  RcvdMaterial"] = srcDr["Part Number"];	
					dr["IC Delivery Line -  RcvdSerialNumber"] = srcDr["Serial No"];	
					//dr["IC Shipments - Arrival Date"] = srcDr["Faulty Board Received Date from Customer"];
					dr["Merge AE Material"] = dr["IC Delivery Line -  RcvdMaterial"];//"";	
					dr["OC Delivery Line -  SentSerialNumber"] = "";	
					dr["CARES - Customer Requested Date"] = srcDr["Customer Request Date"];
					dr["POD Date (Final Calculation)"] = srcDr["Dispatched Date to Customer"];					
					dr["Merge Sales Organization"] = "";//srcDr["Country"];	
					dr["Sales Order - Sold To"] = this.findSoldToCode(dr["Sales Orders SoldTo - Customer Name"].ToString());//"";	
					dr["Work Orders  Header - Vendor Name"] = srcDr["RSLC/Repairer "];	
					dr["Sales Orders Line - Customer Warranty"] = ""; //srcDr["Warranty StatusW/OW"];	
					dr["Merge SPT"] = "";	
					dr["CARES - Service Agreement"] = "";	//??	
					dr["OTDC_Correction"] = "";	
					dr["OTDC_Failure Reason"] = srcDr["L-1"];	
					dr["OTDC_Failure Description"] = srcDr["L-2"];	
					dr["Bulk OTD Status (1=OK, 0=NOK)"] ="";
							
					destDs.Tables["Cockpit AE uploading"].Rows.Add(dr);
				
				}				
			}
		}
	}
}
