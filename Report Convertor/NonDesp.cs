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
	public class NonDesp
	{
		public NonDesp()
		{
			
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
		
		private string SetOverdueGroup(string OverdueDays, string SLADefination)
		{
			string ret = "";
			try
			{
				if (OverdueDays == "TargetDateNotDefined")
				{
					ret = "TargetDateNotDefined";
				}
				else if (OverdueDays == "NA" || OverdueDays == "0")
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
		
		private bool Filter(string customerName, string country, string ticketType)
		{
			
			if (customerName.ToUpper().Contains("ALCATEL") ||
			    customerName.ToUpper().Contains("LUCENT"))
			{
				return false;
			}
			
			if (country == "JAPAN" && ticketType == "L")
			{
				return false;
			}
			
			return true;
		}

		
		public void SetValues()
		{
			DataSet srcDs, destDs;
			srcDs = frmInput.ds;
			destDs = frmOutput.ds;
			
			foreach (DataRow srcDr in srcDs.Tables["Input6NonDesp"].Rows)
			{
				if ( false == this.Filter(srcDr["Customer"].ToString(), 
				                 srcDr["Country"].ToString(),
				                 srcDr["Ticket_Type"].ToString()) )
				{
					continue;
				}
				                 
				DataRow dr = destDs.Tables["NonDesp"].NewRow();

				dr["Country"] = srcDr["Country"];
				dr["Sequence #"] = srcDr["Sequence #"];
				dr["Ref#1"] = srcDr["Ref#1"];
				dr["Ref#2"] = srcDr["Ref#2"];
				dr["Site Type"] = srcDr["Site Type"];
				dr["Customer"] = srcDr["Customer"];
				dr["Site No"] = srcDr["Site No"];
				dr["Fmr Bus Unit"] = srcDr["Fmr Bus Unit"];
				dr["Part Number"] = srcDr["Part Number"];
				dr["Model"] = srcDr["Model"];
				dr["Description"] = srcDr["Description"];
				dr["Product Type"] = srcDr["Product Type"];
				dr["Rma No"] = "'" + srcDr["Rma No"];
				dr["Line No"] = srcDr["Line No"];
				dr["Service Level"] = srcDr["Service Level"];
				dr["Cover Type"] = srcDr["Cover Type"];
				dr["Ticket_Type"] = srcDr["Ticket_Type"];
				dr["Dispatch Store"] = srcDr["Dispatch Store"];
				dr["WC Date"] = srcDr["WC Date"];
				dr["Created On"] = srcDr["Created On"];
				dr["Ord Date"] = srcDr["Ord Date"];
				dr["Release Date"] = srcDr["Release Date"];
				dr["Target Compl Date"] = srcDr["Target Compl Date"];
				dr["Days Open"] = srcDr["Days Open"];
				dr["QoH"] = srcDr["QoH"];
				dr["Order I/p"] = srcDr["Order I/p"];
				dr["Order B/o"] = srcDr["Order B/o"];
				dr["Expected Serial Number"] = srcDr["Expected Serial Number"];
				dr["Received Part Number"] = srcDr["Received Part Number"];
				dr["Received Serial Number"] = srcDr["Received Serial Number"];
				dr["Delivery Name"] = srcDr["Delivery Name"];
				dr["Delivery Address"] = srcDr["Delivery Address"];
				dr["Delivery Town"] = srcDr["Delivery Town"];
				dr["Attention Person"] = srcDr["Attention Person"];
				dr["Current Status"] = srcDr["Current Status"];
				dr["Call Status"] = srcDr["Call Status"];
				dr["Def Rec Date"] = srcDr["Def Rec Date"];
				dr["Workshop"] = srcDr["Workshop"];
				dr["Repairer Name"] = srcDr["Repairer Name"];
				dr["Rep RMA"] = srcDr["Rep RMA"];
				dr["Rep Ref"] = srcDr["Rep Ref"];
				dr["Desp To Rep"] = srcDr["Desp To Rep"];
				dr["Rep Status"] = srcDr["Rep Status"];
				dr["Exp Return"] = srcDr["Exp Return"];
				dr["Desp From Rep"] = srcDr["Desp From Rep"];
				dr["Header_Comments"] = srcDr["Header_Comments"];
				dr["Line_Comments"] = srcDr["Line_Comments"];
				
				try
				{
	//				dr["SLA"]		
					if ( dr["Service Level"].ToString() != "AE" )
					{
						if ( srcDr["Target Compl Date"].ToString() == "" )
						{
							dr["SLA"] = 0;
						}
						else
						{
							DateTime dtTargetComplDate = Convert.ToDateTime(srcDr["Target Compl Date"].ToString());
							DateTime dtDefRecDate = Convert.ToDateTime(srcDr["Def Rec Date"].ToString());
							dr["SLA"] = dtTargetComplDate.Subtract(dtDefRecDate).Duration().Days;
						}
					}
					else
					{
						if ( srcDr["Target Compl Date"].ToString() == "" )
						{
							dr["SLA"] = 0;
						}
						else
						{
							DateTime dtTargetComplDate = Convert.ToDateTime(srcDr["Target Compl Date"].ToString());
							DateTime dtOrdDate = Convert.ToDateTime(srcDr["Ord Date"].ToString());
							dr["SLA"] = dtTargetComplDate.Subtract(dtOrdDate).Duration().Days;
						}
					}
				}
				catch
				{
					string msg = "Non-Desp: Error Occurred when processing 'SLA' field "
						+ "in the record 'Part Number' = " + dr["Part Number"].ToString();

					MessageBox.Show(msg);
				}

				
//				dr["SLA Definition"]
				
				dr["SLA Definition"] = this.CalcSLADefination( dr["Service Level"].ToString(),
					             dr["Site Type"].ToString(),
					             dr["Ticket_Type"].ToString(),
					             dr["Country"].ToString(),
					             dr["Customer"].ToString(),
					             dr["Model"].ToString()
					            );
				
				
//				dr["Overdue Days"];	
				try
				{
					DateTime dtOrdDate, dtCreateDate;
					Int32 daysOpen, SLA;
				
					if (dr["Target Compl Date"].ToString() == "" )
					{
					    dr["Overdue Days"] = "TargetDateNotDefined";
					}
					else
					{		
						if (dr["SLA Definition"].ToString() == "AE" &&
						    srcDr["Ord Date"].ToString() != "" && //OrdDate > CreateDate
						    srcDr["Created On"].ToString() == "")
						{
							TimeSpan ts = frmConfig.CurrentWeekTuesday.Subtract(Convert.ToDateTime(srcDr["Target Compl Date"].ToString()));
							dr["Overdue Days"] = ts.Days;
						}
						else if (srcDr["Ord Date"].ToString() == "" ||
						    	 srcDr["Created On"].ToString() == "")
						{
							daysOpen = Convert.ToInt32(dr["Days Open"].ToString());
							SLA = Convert.ToInt32(dr["SLA"].ToString());
							
							if ( daysOpen <= SLA )
							{
								dr["Overdue Days"] = "NA";
							}
							else
							{
								dr["Overdue Days"] = daysOpen - SLA;
							}
							
						}
						else 
						{
							dtOrdDate = Convert.ToDateTime(srcDr["Ord Date"].ToString());						
							dtCreateDate = Convert.ToDateTime(srcDr["Created On"].ToString());						
						
							if ( dr["SLA Definition"].ToString() == "AE" && dtOrdDate > dtCreateDate)
							{
								TimeSpan ts = frmConfig.CurrentWeekTuesday.Subtract(Convert.ToDateTime(srcDr["Target Compl Date"].ToString()));
								dr["Overdue Days"] = ts.Days;								
							}
							else
							{
								daysOpen = Convert.ToInt32(dr["Days Open"].ToString());
								SLA = Convert.ToInt32(dr["SLA"].ToString());
						
								if ( daysOpen <= SLA )
								{
									dr["Overdue Days"] = "NA";
								}
								else
								{
									dr["Overdue Days"] = daysOpen - SLA;
								}
							}
						}
						
					}
				}
				catch
				{
					string msg = "Non-Desp: Error Occurred when processing 'Overdue Days' field "
						+ "in the record 'Part Number' = " + dr["Part Number"].ToString();

					MessageBox.Show(msg);
				}

				
				//dr["Overdue Group"
				dr["Overdue Group"] = SetOverdueGroup(dr["Overdue Days"].ToString(), dr["SLA Definition"].ToString());

								
				destDs.Tables["NonDesp"].Rows.Add(dr);
			}
		}
	}
}
