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
	
	public class Onelog
	{
		
		public Onelog()
		{
			
		}
		
		private bool CountryFilter4TL9000(string country)
		{
//			if (country == "PHILIPPINES")
//			{
//				return false;
//			}
//			else
			{
				return true;
			}
		}
		
		private bool CustomerNameFilter4TL9000(string customerName)
		{
			if ( customerName.ToUpper().Contains("ALCATEL") ||
			     customerName.ToUpper().Contains("LUCENT") )
			{
				return false;
			}
			else
			{
				return true;
			}
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
		
		private string CalcSLANumberofDays4TL9000(string country, string customerName, string ticketType, string Fmr_BU, string Exp_Part_No, string SLA)
		{
			string ret = "Need Manu Checking";
			
			switch (country.ToUpper())
			{
				case "MALAYSIA":
					if (customerName.Contains("MAXIS"))
					{
						if (Fmr_BU.Contains("ALCATEL"))
						{
							ret = "85";
						}
						else if (Fmr_BU.Contains("LUCENT"))
						{
							ret = "90";
						}
				 	}
					else
					{
						ret = "90";
					}
					break;	
					
				case "INDONESIA":
					string working_code = "";
					bool found = false;
					
					foreach (DataRow dr in frmInput.ds.Tables["IB_ALL_INDO_14022013"].Rows)
					{
						if (Exp_Part_No == dr["Material"].ToString()) 
						{
							working_code = dr["Response Code"].ToString();
							found = true;
							break;						
						}
					}
					
					if (found == false)
					{
						working_code = ticketType;
					}
					
					try
					{
						string temp;
						if (working_code.Substring(working_code.Length - 1, 1) == "H")
						{
							temp = working_code.Substring(0, working_code.Length - 1);
							double d = Math.Round((double)(Convert.ToDouble(temp)/24), 2);							
//							double d = (double)(Convert.ToInt32(temp)/24);
							ret = d.ToString();
						}
						else 
						{
							temp = working_code.Substring(0, working_code.Length - 1);
							ret = Convert.ToInt32(temp).ToString();
						}						
					}
					catch{
					}
					break;
					
				case "AUSTRALIA":
					try
					{
						string temp;
						if (ticketType.Substring(ticketType.Length - 1, 1) == "H")
						{
							ret = "1";
						}
						else 
						{
							temp = ticketType.Substring(0, ticketType.Length - 1);
							ret = Convert.ToInt32(temp).ToString();
						}			
					}
					catch{
					}
					break;
				
				case "BRUNEI DARUSSALAM":
				case "HONG KONG":
					if (SLA == "R4S")
					{
						ret = "90"; //R4R 90 days regardless of products
					}
					else
					{
						ret = "Need Manu Checking";
					}
					break;
					
				default:
					break;
			}
			
			return ret;
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
				
//		private DateTime CalcTargetDueDate(string strShipToCustomerName4, string GIDate, string CreationDate, string uniqueID)
//		{		
//			DateTime ret = new DateTime(1900, 1, 1, 0, 0, 0);
//			
//			if (strShipToCustomerName4 == "")
//			{
////				string msg = "Onelog: 'Sales Orders - ShipTo Customer Name4' is empty "
////						+ "in the record 'Unique Identification' = " + uniqueID
////						+ ", the record is to be excluded.";
////
////				MessageBox.Show(msg);
//				return ret;
//			}
//			
//			try
//			{
//			string SLA_1, SLA_2;
//			SLA_1 = strShipToCustomerName4.Substring(0, 2);
//			
//			if ( strShipToCustomerName4.Substring(strShipToCustomerName4.Length - 1, 1) == "D" )
//			{
//				SLA_2 = strShipToCustomerName4.Substring(2, strShipToCustomerName4.Length - 3);
//			}
//			else
//			{
//				SLA_2 = strShipToCustomerName4.Substring(2, strShipToCustomerName4.Length - 2);
//			}
//			
//			int days = 0;
//				
//			switch (SLA_2)
//			{
//				case "NB":
//					days = 1;
//					break;
//				default:
//			
//					try {
//						days = Convert.ToInt32(SLA_2);
//					}
//					catch
//					{
//						MessageBox.Show("Cannot convert " + SLA_2 + " to Integer when processing 'Sales Orders - ShipTo Customer Name4'.");
//					}
//					
//					break;
//			}
//				
//			if (SLA_1 != "AE") //RR or RE
//			{
//				if ( GIDate.ToString() == "" )	
//				{
////					string msg = "Onelog: 'IC GI Date' is empty "
////						+ "in the record 'Unique Identification' = " + uniqueID
////						+ ", the record is to be excluded.";
////
////					MessageBox.Show(msg);
//					return ret;		//exclude the record
//				}
//				
//				ret = (Convert.ToDateTime(GIDate)).AddDays(days);
//			}
//			else
//			{
//				if ( CreationDate.ToString() == "" )	
//				{
////					string msg = "Onelog: 'Creation Date' is empty "
////						+ "in the record 'Unique Identification' = " + uniqueID
////						+ ", the record is to be excluded.";
////
////					MessageBox.Show(msg);
//					return ret;		//exclude the record
//				}
//				
//				ret = (Convert.ToDateTime(CreationDate)).AddDays(days);
//			}
//			
//			}
//			catch
//			{
//				string msg = "Onelog: Error Occurred when processing 'Target Due Date' field "
//						+ "in the record 'Unique Identification' = " + uniqueID;
//
//					MessageBox.Show(msg);
//			}
//			return ret;
//		}

		private string SetRU(string country)
		{
			string ret;
			
			switch (country.ToUpper())
			{
				case "AUSTRALIA":
				case "NEW ZEALAND":			
				case "PACIFIC ISLANDS":		
				case "PAPUA NEW GUINEA":		
					
					ret = "ANZ";
					break;
					
				case "JAPAN":
				case "KOREA":
				case "KOREA  REPUBLIC OF":
				case "KOREA REPUBLIC OF":
				case "HONG KONG":
				case "TAIWAN":
					
					ret = "NA";
					break;
					
				case "SINGAPORE":
				case "MALAYSIA":
				case "THAILAND":
				case "INDONESIA":
				case "BRUNEI DARUSSALAM":
				case "VIET NAM":
				case "VIETNAM":
					
					ret = "SSEA";
					break;
					
				default:
					ret = "";
					break;
			}
			
			return ret;
		}
		
		string CalcSLA4TL9000(string serviceType, string siteType, string	ticketType, string country,
		               string customerName, string customerType, string	model, string deliveryStatus)
		{
			string ret = "";
			
			if (serviceType == "AE")
			{
				//RULE #1
				if (country == "THAILAND" && customerName.Contains("TRUE MOVE"))
				{
					ret = "R4S";
				}
				
				//RULE #2
				else if (country == "MALAYSIA" && customerName.Contains("CELCOM"))
				{
					ret = "R4S";
				}
				
				//RULE #3
				else if (country == "INDONESIA" && 
				         (siteType.Contains("H") || ticketType.Contains("H")) )
				{
					ret = "H+";
				}
				
				//RULE #4
				else if (country == "KOREA  REPUBLIC OF" )
				{
					if (ticketType == "1D")
					{
						ret = "D+";
					}
					else if ( siteType.Contains("RR") ) //Luis -
					{
						ret = "R4S";
					}
				}
				  
				//RULE #5
				else if ( country == "JAPAN" )
				{
					if (siteType.Contains("H") && ticketType.Contains("H"))	
					{
						ret = "H+";
					}
					else if (customerName.Contains("AT&T SPIDER"))
					{
						ret = "H+";
					}
					else if (customerName.Contains("NTT PC"))
					{
				    		if (model == "11411B" || model == "11411" || 
					         	model == "11594" || model == "11075L" )
							{
								ret = "R4S";
							}
							else
							{
								ret = "D+";
							}
					}
					else if (customerName.Contains("SOFTBANK TELECOM (RS&R)"))
					{
						ret = "R4S";
					}
					else
					{
						ret = "";
					}
				}
				
				//RULE #6
				else if (country == "AUSTRALIA")
				{
					if (ticketType != "1D")
					{
						ret = "";
					}
					else 
					{
//						if (customerName.Contains("OPTUS") || customerName.Contains("TELSTRA"))
//						{
//							ret = "D+";
//						}
//						else 
						{
							ret = "D+";
						}
					}
					    
				}
				//RULE #7
				else if (siteType.Contains("H") && ticketType.Contains("H"))	
				{
					ret = "H+";
				}				
				//RULE #8
				else
				{
					ret = "";
				}
			}
			else // !=AE
			{
				//RULE #1
				if (country == "JAPAN" && customerName.Contains("SOFTBANK TELECOM (RS&R)"))
				{
					if (deliveryStatus == "On Time")
					{
						ret = "R4S";
					}
					else //Past Due
					{
						ret = "Per Use";
					}
				}
				
				//RULE #2
				else if (customerType == "RRN" || customerType == "AEN")
				{
					ret = "Per Use";
				}
				
				//RULE #3
				else if (country == "AUSTRALIA" && 
				         customerName.Contains("OPTUS") &&
				         ticketType == "1D")
				{
					ret = "";
				}
				
				//RULE #4 default to R4S
				else
				{
					ret = "R4S";
				}
			}
			
			return ret;
		}

		public void SetValues4Input1TL9000()
		{
			DataSet srcDs, destDs;
			srcDs = frmInput.ds;
			destDs = frmOutput.ds;
			
//			destDs.Tables["Onelog"].Clear();
			
			foreach (DataRow srcDr in srcDs.Tables["Input1TL9000"].Rows)
			{
				if ( this.CountryFilter4TL9000(srcDr["Country"].ToString()) == true &&
				    this.CustomerNameFilter4TL9000(srcDr["Customer Name"].ToString()) == true)
				{
					DataRow dr = destDs.Tables["Onelog"].NewRow();
//					dr["SLA"] = "aa";				
//					dr["Delivery Status"] = "bb";
					dr["Exp_Part_No"] = "'" + srcDr["Exp_Part_No"].ToString();
					dr["Model"] = srcDr["Model"];	
					
					if ( srcDr["Country"].ToString() == "TAIWAN" )
					{
						dr["Country"] = "TW";
					}
					else
					{
						dr["Country"] = srcDr["Country"]; 
					}
					
					dr["Fmr BU"] = srcDr["Fmr BU"];
					dr["Customer Name"] = srcDr["Customer Name"];	
					dr["Cust_Type"] = srcDr["Cust_Type"];	
					dr["SiteType"] = srcDr["SiteType"];	
					dr["Ticket_Type"] = srcDr["Ticket_Type"];	
					dr["Service_Type"] = srcDr["Service_Type"];	
					dr["Warranty_Status"] = srcDr["Warranty_Status"];	
					dr["Rma#"] = "'" + srcDr["Rma#"];	
					dr["Line#"] = srcDr["Line#"];	
					dr["Create_Date"] = srcDr["Create_Date"];	
					dr["Rcv_Date_Def"] = srcDr["Rcv_Date_Def"];	 
					dr["Ship_Date_Repl"] = srcDr["Ship_Date_Repl"];	
					dr["Target Compl Date"] = srcDr["Targe Compl Date"];	
					dr["TAT"] = srcDr["TAT"];	
					dr["In TAT"] = srcDr["InTat"];	
					dr["Repairer"] = srcDr["Repairer"];	
					dr["Defective TAT"] = srcDr["Defective TAT"].ToString();	
					dr["Repair TAT"] = srcDr["Repair TAT"].ToString();	 
					dr["Order TAT"] = srcDr["Order TAT"];	
					dr["Despatch Store"] = srcDr["Despatch Store"];	
					dr["Repair Comments"] = srcDr["Repair Comments"];	
					dr["Ticket_No"] = "'" + srcDr["Ticket_No"];	
					dr["RU"] = this.SetRU(srcDr["Country"].ToString());
					dr["WK"] = frmConfig.CurrentWeek;
//					dr["Comments from In-country "] = "";	
//					dr["SLA of Indosat"] = "";	
					dr["Clarify_Date"] = srcDr["Clarify_Date"];	
					dr["Clarify_Time"] = srcDr["Clarify_Time"];	
					dr["Signed_By"] = srcDr["Signed_By"];	
					dr["Signed_Date"] = srcDr["Signed_Date"];	
					dr["Unique Identification"] = srcDr["Rma#"] + "\\" + srcDr["Line#"]; 	
//					dr["Connote"] = "";	
//					dr["Status"] = ""; 
					
					if (dr["In TAT"].ToString() == "1")
					{
						dr["Delivery Status"] = "On Time";
					}
					else {
						dr["Delivery Status"] = "Past Due";
					}
	
					//dr["SLA"]
					dr["SLA"] = this.CalcSLA4TL9000( dr["Service_Type"].ToString(),
					             dr["SiteType"].ToString(),
					             dr["Ticket_Type"].ToString(),
					             dr["Country"].ToString(),
					             dr["Customer Name"].ToString(),
					             dr["Cust_Type"].ToString(),
					             dr["Model"].ToString(),
					             dr["Delivery Status"].ToString()
					            );
					
//					dr["SLA number of days"]  = srcDr["Customer TAT"].ToString();				//Newly added
					dr["SLA number of days"]  = CalcSLANumberofDays4TL9000(
					             									dr["Country"].ToString(),
					            	 								dr["Customer Name"].ToString(),
					             									dr["Ticket_Type"].ToString(),
					             									dr["Fmr BU"].ToString(),					             									
					             									srcDr["Exp_Part_No"].ToString(),
					             									dr["SLA"].ToString()
					           										);
		
					dr["TAT upon actual close"] = CalcTATuponActualClose(dr["SLA"].ToString(),
					                                                     dr["Ship_Date_Repl"].ToString(),
					                                                     dr["Rcv_Date_Def"].ToString(),
					                                                     dr["Create_Date"].ToString());
					
					destDs.Tables["Onelog"].Rows.Add(dr);	
				
				}
			}
		}
		
		public void SetValues4Input2TWClose()
		{
			DataSet srcDs, destDs;
			srcDs = frmInput.ds;
			destDs = frmOutput.ds;
			
//			destDs.Tables["Onelog"].Clear();
			
			foreach (DataRow srcDr in srcDs.Tables["Input2TWClose"].Rows)
			{		
				DateTime TargetComplDate = CalcTargetDueDate(srcDr["Sales Orders - Order  Reason"].ToString(),
				                                             srcDr["Sales Orders Line - Delivery SLA"].ToString(),
				                                             srcDr["IC Shipments - Arrival Date"].ToString(),
				                                             srcDr["Sales Orders - Creation Date"].ToString(),
				                                             srcDr["Sales Orders - Sales Order Number"].ToString() + "\\" 
				                                             + srcDr["Sales Orders Line - LineNumber"].ToString() );
//				DateTime TargetComplDate = CalcTargetDueDate(srcDr["Sales Orders - ShipTo Customer Name4"].ToString(),
//				                                             srcDr["IC Shipments - Arrival Date"].ToString(),
//				                                             srcDr["Sales Orders - Creation Date"].ToString(),
//				                                             srcDr["Sales Orders - Sales Order Number"].ToString() + "\\" 
//				                                             + srcDr["Sales Orders Line - LineNumber"].ToString() );
				if ( TargetComplDate.Year == 1900 )
				{
					continue; //"GI DATA" is empty, just exclude
				}
				
				if (this.CustomerNameFilter4TW(srcDr["Sales Orders - ShipTo Customer Name"].ToString()) == true &&
				    this.CurrentWeekFilter(TargetComplDate) == true)
				{				                               
					DataRow dr = destDs.Tables["Onelog"].NewRow();
					//dr["SLA"] = "aa";				
					//dr["Delivery Status"] = "bb";
					dr["Exp_Part_No"] = "'" + srcDr["Sales Orders Line - Material"].ToString();
					dr["Model"] = srcDr["Sales Orders Line - Material Description"];	
					dr["Country"] = "TW";
					dr["Fmr BU"] = "";	
					dr["Customer Name"] = srcDr["Sales Orders - ShipTo Customer Name"];	
					dr["Cust_Type"] = "";	
					dr["SiteType"] = "";	
					dr["Ticket_Type"] = srcDr["Sales Orders Line - Delivery SLA"];	//New column
//					dr["Service_Type"] = srcDr["Sales Orders - ShipTo Customer Name4"];	
					dr["Service_Type"] = srcDr["Sales Orders - Order  Reason"];	
					dr["Warranty_Status"] = "";	
					dr["Rma#"] = "'" + srcDr["Sales Orders Line - SPT"].ToString();
					dr["Line#"] = srcDr["Sales Orders Line - LineNumber"].ToString();	
					dr["Create_Date"] = srcDr["Sales Orders - Creation Date"];	
					dr["Rcv_Date_Def"] = srcDr["IC Shipments - Arrival Date"];	
					dr["Ship_Date_Repl"] = srcDr["OC Delivery Header -  GI_Date"];	
				
					dr["Target Compl Date"] = TargetComplDate;	 // bb column				
					//dr["TAT"] = "";	
				
					try
					{
						if ( TargetComplDate >= Convert.ToDateTime(srcDr["OC Delivery Header -  GI_Date"].ToString()) )
						{
							dr["In TAT"] = 1;	 //bd column
						}
						else{
							dr["In TAT"] = 0;
						}
					}
					catch
					{
						string msg = "Onelog-TWClose: Error Occurred when processing 'In TAT' field "
						+ "in the record 'Unique Identification' = "
						+ srcDr["Sales Orders - Sales Order Number"] + "\\" 
						+ srcDr["Sales Orders Line - LineNumber"];

						MessageBox.Show(msg);
						dr["In TAT"] = 0;
					}
				
//					dr["Repairer"] = "";	
//					dr["Defective TAT"] = "";	
//					dr["Repair TAT"] = "";	
//					dr["Order TAT"] = "";	
//					dr["Despatch Store"] = "";	
//					dr["Repair Comments"] = "";	
//					dr["Ticket_No"] = "";	
					dr["RU"] = "NA";	
					dr["WK"] = frmConfig.CurrentWeek;	
//					dr["Comments from In-country "] = "";	
//					dr["SLA of Indosat"] = "";	
//					dr["Clarify_Date"] = "";	
//					dr["Clarify_Time"] = "";	
//					dr["Signed_By"] = "";	
//					dr["Signed_Date"] = "";	
					dr["Unique Identification"] = srcDr["Sales Orders - Sales Order Number"] + "\\" + srcDr["Sales Orders Line - LineNumber"]; 	
//					dr["Connote"] = "";	
//					dr["Status"] = ""; 
							
					
					
					if (dr["In TAT"].ToString() == "1")
					{
						dr["Delivery Status"] = "On Time";
					}
					else {
						dr["Delivery Status"] = "Past Due";
					}
						
					try
					{
						//dr["SLA"] = 
//						if ( srcDr["Sales Orders - ShipTo Customer Name4"].ToString().Substring(0, 2) == "AE" )
//						if ( dr["Service_Type"].ToString().Substring(0, 1) == "A" ||
//						     dr["Service_Type"].ToString().Substring(0, 1) == "E" ) 
						if ( srcDr["Sales Orders - Order  Reason"].ToString().Substring(0, 1) == "A" ) 
						{
							dr["SLA"] = "D+";
						}
						else if ( srcDr["Sales Orders - Order  Reason"].ToString().Substring(0, 1) == "H" ) 
						{
							dr["SLA"] = "H+";
						}
						else 
						{
							dr["SLA"] = "R4S";
						}
					}
					catch
					{
						string msg = "Onelog-TWClose: Error Occurred when processing 'SLA' field "
						+ "in the record 'Unique Identification' = "
						+ srcDr["Sales Orders - Sales Order Number"] + "\\" 
						+ srcDr["Sales Orders Line - LineNumber"];

						MessageBox.Show(msg);
					}
					

					dr["SLA number of days"] = CalcSLANumberofDays4TW(srcDr["Sales Orders - Order  Reason"].ToString(),
					                       						srcDr["Sales Orders Line - Delivery SLA"].ToString());
												
					dr["TAT upon actual close"] = CalcTATuponActualClose(dr["SLA"].ToString(),
					                                                     dr["Ship_Date_Repl"].ToString(),
					                                                     dr["Rcv_Date_Def"].ToString(),
					                                                     dr["Create_Date"].ToString());
					destDs.Tables["Onelog"].Rows.Add(dr);
				}
			}
		}
				
		public void SetValues4Input3TWOpen()
		{
			DataSet srcDs, destDs;
			srcDs = frmInput.ds;
			destDs = frmOutput.ds;
			
//			destDs.Tables["Onelog"].Clear();
			
			foreach (DataRow srcDr in srcDs.Tables["Input3TWOpen"].Rows)
			{
								
				DateTime TargetComplDate = CalcTargetDueDate(srcDr["Sales Orders - Order  Reason"].ToString(),				                                             
				                                             srcDr["Sales Orders Line - Delivery SLA"].ToString(),
				                                             srcDr["IC Shipments - Arrival Date"].ToString(),
				                                             srcDr["Sales Orders - Creation Date"].ToString(),
				                                             srcDr["Sales Orders - Sales Order Number"].ToString() + "\\" 
				                                             + srcDr["Sales Orders Line - LineNumber"].ToString() );
//				DateTime TargetComplDate = CalcTargetDueDate(srcDr["Sales Orders - ShipTo Customer Name4"].ToString(),
//				                                             srcDr["IC Shipments - Arrival Date"].ToString(),
//				                                             srcDr["Sales Orders - Creation Date"].ToString(),
//				                                             srcDr["Sales Orders - Sales Order Number"].ToString() + "\\" 
//				                                             + srcDr["Sales Orders Line - LineNumber"].ToString() );
				if ( TargetComplDate.Year == 1900 )
				{
					continue; //"GI DATA" is empty, just exclude
				}
				
				if (this.CustomerNameFilter4TW(srcDr["Sales Orders - ShipTo Customer Name"].ToString()) == true &&
				    ( this.CurrentWeekFilter(TargetComplDate) == true || this.PassdueDateFilter(TargetComplDate) == true) )
				{				                               
					DataRow dr = destDs.Tables["Onelog"].NewRow();
					//dr["SLA"] = "aa";				
					//dr["Delivery Status"] = "bb";
					dr["Exp_Part_No"] = "'" + srcDr["Sales Orders Line - Material"].ToString();
					dr["Model"] = srcDr["Sales Orders Line - Material Description"];	
					dr["Country"] = "TW";
					dr["Fmr BU"] = "";	
					dr["Customer Name"] = srcDr["Sales Orders - ShipTo Customer Name"];	
					dr["Cust_Type"] = "";	
					dr["SiteType"] = "";						
					dr["Ticket_Type"] = srcDr["Sales Orders Line - Delivery SLA"];	//New column 
//					dr["Service_Type"] = srcDr["Sales Orders - ShipTo Customer Name4"];
					dr["Service_Type"] = srcDr["Sales Orders - Order  Reason"];	
					dr["Warranty_Status"] = "";	
					dr["Rma#"] = "'" + srcDr["Sales Orders Line - SPT"].ToString();
					dr["Line#"] = srcDr["Sales Orders Line - LineNumber"].ToString();	
					dr["Create_Date"] = srcDr["Sales Orders - Creation Date"];	
					dr["Rcv_Date_Def"] = srcDr["IC Shipments - Arrival Date"];	
					dr["Ship_Date_Repl"] = srcDr["OC Delivery Header -  GI_Date"];	
				
					dr["Target Compl Date"] = TargetComplDate;	 // bb column				
					//dr["TAT"] = "";	
					try
					{						
						if ( TargetComplDate >= Convert.ToDateTime(srcDr["OC Delivery Header -  GI_Date"].ToString()) )
						{
							dr["In TAT"] = 1;	 //bd column
						}
						else{
							dr["In TAT"] = 0;
						}
					}
					catch
					{
//						string msg = "Onelog-TW: Error Occurred when processing 'In TAT' field "
//						+ "in the record 'Unique Identification' = "
//						+ srcDr["Sales Orders - Sales Order Number"] + "\\" 
//						+ srcDr["Sales Orders Line - LineNumber"];
//
//						MessageBox.Show(msg);
						dr["In TAT"] = 0;
					}
				
//					dr["Repairer"] = "";	
//					dr["Defective TAT"] = "";	
//					dr["Repair TAT"] = "";	
//					dr["Order TAT"] = "";	
//					dr["Despatch Store"] = "";	
//					dr["Repair Comments"] = "";	
//					dr["Ticket_No"] = "";	
					dr["RU"] = "NA";	
					dr["WK"] = frmConfig.CurrentWeek;	
					if ( this.CurrentWeekFilter(TargetComplDate) == true ) 
					{
						dr["Comments from In-country "] = frmConfig.CurrentWeek + "-open";	
					}
					else {
						dr["Comments from In-country "] = "Before " + frmConfig.CurrentWeek + "-open";
					}
//					dr["SLA of Indosat"] = "";	
//					dr["Clarify_Date"] = "";	
//					dr["Clarify_Time"] = "";	
//					dr["Signed_By"] = "";	
//					dr["Signed_Date"] = "";	
					dr["Unique Identification"] = srcDr["Sales Orders - Sales Order Number"] + "\\" + srcDr["Sales Orders Line - LineNumber"]; 	
//					dr["Connote"] = "";	
//					dr["Status"] = ""; 
							
//					dr["SLA"] = 
					
					if (dr["In TAT"].ToString() == "1")
					{
						dr["Delivery Status"] = "On Time";
					}
					else {
						dr["Delivery Status"] = "Past Due";
					}
								
					try
					{
						//dr["SLA"] = 
//						if ( srcDr["Sales Orders - ShipTo Customer Name4"].ToString().Substring(0, 2) == "AE" )
//						if ( dr["Service_Type"].ToString().Substring(0, 1) == "A" ||
//						     dr["Service_Type"].ToString().Substring(0, 1) == "E" ) 
						if ( srcDr["Sales Orders - Order  Reason"].ToString().Substring(0, 1) == "A" ) 
						{
							dr["SLA"] = "D+";
						}
						else if ( srcDr["Sales Orders - Order  Reason"].ToString().Substring(0, 1) == "H" ) 
						{
							dr["SLA"] = "H+";
						}
						else 
						{
							dr["SLA"] = "R4S";
						}
					}
					catch
					{
						string msg = "Onelog-TW: Error Occurred when processing 'SLA' field "
						+ "in the record 'Unique Identification' = "
						+ srcDr["Sales Orders - Sales Order Number"] + "\\" 
						+ srcDr["Sales Orders Line - LineNumber"];

						MessageBox.Show(msg);	
					}
					
					dr["SLA number of days"] = CalcSLANumberofDays4TW(srcDr["Sales Orders - Order  Reason"].ToString(),
					                      						srcDr["Sales Orders Line - Delivery SLA"].ToString());
												
					dr["TAT upon actual close"] = CalcTATuponActualClose(dr["SLA"].ToString(),
					                                                     dr["Ship_Date_Repl"].ToString(),
					                                                     dr["Rcv_Date_Def"].ToString(),
					                                                     dr["Create_Date"].ToString());
					
					destDs.Tables["Onelog"].Rows.Add(dr);
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
				DataRow dr = destDs.Tables["Onelog"].NewRow();
//				dr["SLA"] = "aa";				
//				dr["Delivery Status"] = "bb";
				dr["Exp_Part_No"] = "'" + srcDr["Part Number"].ToString();
				dr["Model"] = srcDr["Model"];	
				if ( srcDr["Country"].ToString() == "KOREA" )
				{
					dr["Country"] = "KOREA  REPUBLIC OF";
				}
				else
				{
					dr["Country"] = srcDr["Country"]; 
				}
				dr["Fmr BU"] = srcDr["Fmr BU"];	
				dr["Customer Name"] = srcDr["Customer Name"];	
				dr["Cust_Type"] = srcDr["Cust_Type"];	
				dr["SiteType"] = srcDr["SLA"];	
				dr["Ticket_Type"] = srcDr["Ticket_Type"];	
				dr["Service_Type"] = srcDr["Service_Type"];	
				dr["Warranty_Status"] = srcDr["Warranty_Status"];	
				dr["Rma#"] = "'" + srcDr["Rma#"].ToString();
				dr["Line#"] = "";	
				dr["Create_Date"] = srcDr["Create_Date"];	
				dr["Rcv_Date_Def"] = srcDr["IC Shipments - Arrival Date"];
				dr["Ship_Date_Repl"] = srcDr["Ship_Date_Repl"];	
				dr["Target Compl Date"] = srcDr["Target Compl Date"];	
				dr["TAT"] = srcDr["TAT"];	
				dr["In TAT"] = srcDr["In TAT"];	
				dr["Repairer"] = srcDr["Repairer"];	
				dr["Defective TAT"] = srcDr["Defective TAT"];	
				dr["Repair TAT"] = srcDr["Repair TAT"];	
				dr["Order TAT"] = srcDr["Order TAT"];	
				dr["Despatch Store"] = srcDr["Despatch Store"];	
				dr["Repair Comments"] = srcDr["Repair Comments"];	
				dr["Ticket_No"] = "'" + srcDr["Ticket_No"];	
				dr["RU"] = "NA";	
				dr["WK"] = "wk" + srcDr["WK"].ToString();
//				dr["Comments from In-country "] =  "";	
//				dr["SLA of Indosat"] = "";	
//				dr["Clarify_Date"] = "";	
//				dr["Clarify_Time"] = "";	
//				dr["Signed_By"] = "";	
//				dr["Signed_Date"] = "";	
				dr["Unique Identification"] = srcDr["New Serial Number"]; 	
//				dr["Connote"] = "";	
//				dr["Status"] = ""; 
					
//				if (dr["In TAT"].ToString() == "1")
//				{
//					dr["Delivery Status"] = "On Time";
//				}
//				else {
//					dr["Delivery Status"] = "Past Due";
//				}
				
				try
				{
					DateTime dt1 = Convert.ToDateTime(dr["Target Compl Date"].ToString()).Date;
					DateTime dt2 = Convert.ToDateTime(dr["Ship_Date_Repl"].ToString()).Date;
					
					if ( dt1 >= dt2 )
					{
						dr["Delivery Status"] = "On Time";
						dr["In TAT"] = "1";
					}
					else {
						dr["Delivery Status"] = "Past Due";
						dr["In TAT"] = "0";
					}
				}
				catch
				{
					string msg = "Onelog-KOREA: Error Occurred when processing 'Delivery Status' field "
						+ "in the record 'Unique Identification' = "
							+ srcDr["New Serial Number"];
						
						MessageBox.Show(msg);
				}
			
				
					try
					{
						//dr["SLA"] = 
						if ( dr["SiteType"].ToString().Substring(0, 1) == "R" )	//SiteType ??
						{
							dr["SLA"] = "R4S";
						}
						else 
						{
							dr["SLA"] = "";
						}
					}
					catch
					{
						string msg = "Onelog-KOREA: Error Occurred when processing 'SLA' field "
						+ "in the record 'Unique Identification' = "
							+ srcDr["New Serial Number"];
						
						MessageBox.Show(msg);
					}
					
				try
				{
					string SLA_2 = srcDr["SLA"].ToString().Substring(1, srcDr["SLA"].ToString().Length - 1);
					dr["SLA number of days"] = Convert.ToInt32(SLA_2).ToString();
				}
				catch{
				}
												
				dr["TAT upon actual close"] = CalcTATuponActualClose4Korea(dr["SLA"].ToString(),
					                                                     dr["Ship_Date_Repl"].ToString(),
					                                                     dr["Rcv_Date_Def"].ToString(),
					                                                     dr["Create_Date"].ToString());
					
				destDs.Tables["Onelog"].Rows.Add(dr);
			}
		}			
				
	}
}
