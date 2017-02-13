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
	public class OpenOrderOnelog
	{
		public OpenOrderOnelog()
		{
			
		}
		
		private string SetOverdueGroup(string OverdueDays, string SLADefination)
		{
			string ret = "";
			try
			{
				if (OverdueDays == "")
				{
					ret = "NA";
				}
				else if (Convert.ToInt32(OverdueDays) == 0)
				{
					ret = "NA";
				}				
				else if ( SLADefination == "AE" )
				{
					int i = Convert.ToInt32(OverdueDays);
					
					if ( i > 0 && i <= 15 )
					{
						ret = "'1-15";
					}
					else if ( i > 15 && i <= 30 )
					{
						ret = "'16-30";
					}
					else if ( i > 30 )
					{
						ret = "30+";
					}
				} 
				else // SLADefination == "R4S"
				{
					int i = Convert.ToInt32(OverdueDays);
					
					if ( i > 0 && i <= 30 )
					{
						ret = "'1-30";
					}
					else if ( i > 30 && i <= 60 )
					{
						ret = "'31-60";
					}
					else if ( i > 60 && i <= 90 )
					{
						ret = "'61-90";
					}
					else if ( i > 90 )
					{
						ret = "90+";
					}
				}
			}
			catch
			{
				
			}
			
			return ret;
		}
	
		string CalcSLADefination(string serviceType, string siteType, string ticketType, string country,
		               string customerName, string	model)
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
					ret = "AE";
				}
				
				//RULE #4
				else if (country == "KOREA  REPUBLIC OF" )
				{
					if (ticketType == "1D")
					{
						ret = "AE";
					}
					else if ( siteType.Contains("R4R") ) //Luis -
					{
						ret = "R4S";
					}
				}
				  
				//RULE #5
				else if ( country == "JAPAN" )
				{
					if (siteType.Contains("H") && ticketType.Contains("H"))	
					{
						ret = "AE";
					}
					else if (customerName.Contains("AT&T SPIDER"))
					{
						ret = "AE";
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
								ret = "AE";
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
							ret = "AE";
						}
					}
					    
				}
				//RULE #7
				else if (siteType.Contains("H") && ticketType.Contains("H"))	
				{
					ret = "AE";
				}		
				else
				//RULE #8
				{
					ret = "";
				}
			}
			else // !=AE
			{				
				//RULE #3
				if (country == "AUSTRALIA" && 
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
		
		private string SetRU(string country)
		{
			string ret;
			
			switch (country)
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
				case "HONG KONG":
					
					ret = "NA";
					break;
					
				case "SINGAPORE":
				case "MALAYSIA":
				case "THAILAND":
				case "INDONESIA":
				case "BRUNEI DARUSSALAM":
					
					ret = "SSEA";
					break;
					
				default:
					ret = "";
					break;
			}
			
			return ret;
		}
		
		private bool RMACancelledFilter(string Identification)
		{				
			foreach (DataRow filterDr in frmInput.ds.Tables["RMACancelledList"].Rows)
			{
				if (Identification.ToString() == filterDr["Identification"].ToString() )
				{
					return true;
				}
			}
			
			return false;
		}
		
		private bool CustomerNameFilter(string customerName)
		{
			if ( customerName == "台灣國際標準電子股份有限公司" ||	
			     customerName == "內政部消防署" ||
			     customerName.ToUpper().Contains("ALCATEL") ||
			     customerName.ToUpper().Contains("LUCENT"))
			{
				return true;
			}
			else
			{
				return false;
			}
		}
		
				
		public void SetValues()
		{
			DataSet srcDs, destDs;
			srcDs = frmInput.ds;
			destDs = frmOutput.ds;
			
			foreach (DataRow srcDr in srcDs.Tables["Input7OpenOrderOnelog"].Rows)
			{
				if ( RMACancelledFilter(srcDr["RMA_No"] + "\\" + srcDr["Line_No"]) == true ) //In cancelled list
				{
					continue;
				}
				if ( CustomerNameFilter(srcDr["Customer_Name"].ToString()) == true )
				{
					continue;
				}
				
				DataRow dr = destDs.Tables["OpenOrderOnelog"].NewRow();
				
					
				dr["Overdue"] = srcDr["Overdue"];					
				if (dr["Overdue"].ToString() == "" )
				{
					dr["Delivery Status"] = "Not yet due";
				}
				else if (Convert.ToInt32(dr["Overdue"].ToString()) > 0)
				{
						dr["Delivery Status"] = "Out of TAT";
				}
				else {
						dr["Delivery Status"] = "Not yet due";
				}
				
				dr["Rcvd_Part_No"] = "'" + srcDr["Rcvd_Part_No"];	
				dr["Model_Code"] = srcDr["Model_Code"];	
//				dr["Fmr BU"] = srcDr[""];	
				dr["Country"] = srcDr["Country"];	
				dr["Customer_Name"] = srcDr["Customer_Name"];	
				dr["cust_type"] = srcDr["cust_type"];	
				dr["Site_Type"] = srcDr["Site_Type"];	
				dr["Warranty_Status"] = srcDr["Warranty_Status"];	
				dr["RMA_No"] = "'" + srcDr["RMA_No"];	
				dr["Line_No"] = srcDr["Line_No"];	
				dr["Ticket_Type"] = srcDr["Ticket_Type"];	
				dr["Service_Level"] = srcDr["Service_Level"];	
				dr["Repairer_Name"] = srcDr["Repairer_Name"];	
				dr["Rcvd_Serial_No"] = "'" + srcDr["Rcvd_Serial_No"];	
				dr["Created_On"] = srcDr["Created_On"];	
				dr["Order_Date"] = srcDr["Order_Date"];	
				dr["Target_Compl_Date"] = srcDr["Target_Compl_Date"];	
				dr["Comments_About_Parts_Used"] = srcDr["Comments_About_Parts_Used"];	
				dr["Workshop"] = srcDr["Workshop"];	
				dr["Repair_Code"] = srcDr["Repair_Code"];	
				dr["Current_Repair_Status"] = srcDr["Current_Repair_Status"];	
				dr["Fault_Text"] = srcDr["Fault_Text"];	
				dr["Rcvd_Date_From_Rep"] = srcDr["Rcvd_Date_From_Rep"];	
				dr["Repairer_Ref"] = srcDr["Repairer_Ref"];	
				dr["Current_Status"] = srcDr["Current_Status"];	
				dr["Current_Repair_Status"] = srcDr["Current_Repair_Status"];	
				dr["Repair_Rma_No"] = srcDr["Repair_Rma_No"];	
				dr["Current_Repair_Status_Date"] = srcDr["Current_Repair_Status_Date"];	
				dr["Repair_Comment_Code"] = srcDr["Repair_Comment_Code"];	
				dr["Repair_Comment_Desc"] = srcDr["Repair_Comment_Desc"];	
				dr["Latest_Rst_Date"] = srcDr["Latest_Rst_Date"];	
				dr["Exp_Return_Date"] = srcDr["Exp_Return_Date"];	
//				dr["RU"] = this.SetRU(srcDr["Country"].ToString());		//need to check
//				dr["Contract or not"] = srcDr[""];	
				dr["Identification"] = srcDr["RMA_No"] + "\\" + srcDr["Line_No"]; 	
//				dr["Closed Date/comments"] = srcDr[""];	
//				dr["In Country comments"] = srcDr[""];
				
				dr["SLA"] = this.CalcSLADefination( dr["Service_Level"].ToString(),
					             dr["Site_Type"].ToString(),
					             dr["Ticket_Type"].ToString(),
					             dr["Country"].ToString(),
					             dr["Customer_Name"].ToString(),
					             dr["Model_Code"].ToString()
					            );
				dr["Overdue Group"] = SetOverdueGroup(dr["Overdue"].ToString(), dr["SLA"].ToString());

				
				destDs.Tables["OpenOrderOnelog"].Rows.Add(dr);
			}
		}
	}
}
