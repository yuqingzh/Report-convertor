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
	public class OpenOrdereSpares
	{
		public OpenOrdereSpares()
		{
			
		}
		
		public void SetValues()
		{
			SetValues4eSparesNew();
			SetValues4eSparesTW();
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
				else if (Convert.ToInt32(OverdueDays) <= 0)
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
		
		private string GetCountryName4eSparesNew(string CountryCode)
		{
			string ret = "";
			
			foreach (DataRow dr in frmInput.ds.Tables["OpenOrder - eSparesNew - Country Code List"].Rows)
			{
				if (CountryCode.ToUpper() == dr["COUNTRY CODE"].ToString().ToUpper())
			  	{
			  		ret = dr["COUNTRY NAME"].ToString();
			  		break;
			  	}
			}
			
			if (CountryCode != "" && ret == "")
			{
				ret = "Country name not found!";
			}
			
			return ret;
			
		}
				
		public void SetValues4eSparesNew()
		{
			DataSet srcDs, destDs;
			srcDs = frmInput.ds;
			destDs = frmOutput.ds;
			
			foreach (DataRow srcDr in srcDs.Tables["OpenOrdereSparesNew"].Rows)
			{
				DataRow dr = destDs.Tables["OpenOrdereSparesNew"].NewRow();
												
				//Filter 'Customer Status (Closed / Open)'
				if (srcDr["Customer Status (Closed / Open)"].ToString().ToUpper().Trim() != "OPEN")
				{
					continue;
				}
				
				if ( CustomerNameFilter(srcDr["Sales Orders SoldTo - Customer Name"].ToString()) == true )
				{
					continue;
				}
				 
				dr["Sales Orders - RSCIC/RSLC (Sales Organisation)"] = srcDr["Sales Orders - RSCIC/RSLC (Sales Organisation)"];
				dr["COUNTRY"] = GetCountryName4eSparesNew(dr["Sales Orders - RSCIC/RSLC (Sales Organisation)"].ToString());
				dr["Customer Status (Closed / Open)"] = srcDr["Customer Status (Closed / Open)"];	
				dr["Sales Orders - Type"] = srcDr["Sales Orders - Type"];
				dr["Sales Orders SoldTo - Customer Name"] = srcDr["Sales Orders SoldTo - Customer Name"];
				dr["Sales Orders - Order  Reason"] = srcDr["Sales Orders - Order  Reason"];
				dr["Sales Orders - Order Reason Service Description"] = srcDr["Sales Orders - Order Reason Service Description"];
				dr["Sales Orders - Order Reason eSpares SLA"]  = srcDr["Sales Orders - Order Reason eSpares SLA"];
				dr["Sales Orders - Sales Order Number"]  = srcDr["Sales Orders - Sales Order Number"];
				dr["Sales Orders Line - LineNumber"]  = srcDr["Sales Orders Line - LineNumber"];
				dr["Sales Orders Line - Material"]  = srcDr["Sales Orders Line - Material"];
				dr["Sales Orders Line - SPT"]  = srcDr["Sales Orders Line - SPT"];
				dr["IC Delivery Line -  RcvdSerialNumber"]  = srcDr["IC Delivery Line -  RcvdSerialNumber"];
				dr["Work Orders  Header - Vendor Name"]  = srcDr["Work Orders  Header - Vendor Name"];
				dr["Sales Orders Line - Creation Date"]  = srcDr["Sales Orders Line - Creation Date"];
				dr["IC Shipments - Arrival Date"]  = srcDr["IC Shipments - Arrival Date"];
				dr["Customer Due Date"]  = srcDr["Customer Due Date"];
				dr["Customer OPEN TAT"]  = srcDr["Customer OPEN TAT"];
				dr["TAT Status"]  = srcDr["TAT Status"];

				
				try
				{
//					country = TH01 / THAILAND, customer contains ‘TRUE MOVE’, sales orders – order reason=AE30, the SLA = R4S)
					if ( srcDr["Sales Orders - RSCIC/RSLC (Sales Organisation)"].ToString() == "TH01" &&
						 srcDr["Sales Orders - Order  Reason"].ToString() == "A30")
					{
						dr["SLA"] = "R4S";
					}
					else if ( srcDr["Sales Orders - Order  Reason"].ToString().Substring(0, 1) == "A" ||
					          srcDr["Sales Orders - Order  Reason"].ToString().Substring(0, 1) == "E") //srcDr["Sales Orders - Order  Reason"];
					{
						dr["SLA"] = "AE";
					}
					else 
					{
						dr["SLA"] = "R4S";
					}
				}
				catch
				{
						string msg = "OpenOrdereSparesNew: Error Occurred when processing 'SLA' field "
							+ "Sales Orders - Order  Reason' = "
							+ srcDr["Sales Orders - Order  Reason"].ToString();

						MessageBox.Show(msg);
				}
				
//				if (dr["Customer Due Date"].ToString() == "" ||
//				    dr["Customer Due Date"].ToString() == "Faulty Board not received")
//				{
//					dr["Overdue"] = "Not applicable. Cannot get 'Customer Due Date'.";
//				}
//				else
//				{
//					try
//					{
//						TimeSpan ts = DateTime.Now.Date.Subtract(Convert.ToDateTime(dr["Customer Due Date"].ToString()));
//						dr["Overdue"] = ts.Days.ToString();
//					}
//					catch
//					{
//						dr["Overdue"] = "Not applicable";
//					}
//				}
				
				if (dr["TAT Status"].ToString() == "Out of TAT")
				{
					dr["Overdue"] = srcDr["Customer OPEN TAT"];
				}

				
//				if (dr["Overdue"].ToString().Contains("Not applicable"))
//				{
//					dr["Delivery Status"] = "Not applicable";
//				}
				if ( dr["Overdue"].ToString() == "" )
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
					
				dr["Overdue Group"] = SetOverdueGroup(dr["Overdue"].ToString(), dr["SLA"].ToString());

				
				destDs.Tables["OpenOrdereSparesNew"].Rows.Add(dr);
			}
		}
		
//====================================================================================================================

		private DateTime CalcCustomerDueDate4eSparesTW(string strOrderReason, string strDeliverySLA, string ArrivalDate, string CreationDate)
		{
			DateTime ret = new DateTime(1900, 1, 1, 0, 0, 0);
						
			int days = -10000; //flag
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
						days = Convert.ToInt32(strDeliverySLA);
					}
					catch{
					}
				}
			
				if (days == -10000)
				{
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
								days = Convert.ToInt32(SLA_2);
							}
							catch{
								MessageBox.Show("Cannot convert " + strOrderReason + " to customer due date in eSpares TW.");
							}
							break;
					}//switch
				} //if (days == -10000)
	
				if (SLA_1 == "A" || SLA_1 == "H") //SLA type is AE
				{
					if ( CreationDate.ToString() == "" )	
					{
						return ret;		//exclude the record
					}
				
					ret = (Convert.ToDateTime(CreationDate).Date).AddDays(days);
				}
				else	//SLA type is R4S
				{
					if ( ArrivalDate.ToString() == "" )	
					{
						return ret;		//exclude the record
					}
				
					ret = (Convert.ToDateTime(ArrivalDate).Date).AddDays(days);
				}			
			}
			catch
			{
				string msg = "OpenOrder-eSparesTW: Error Occurred when processing 'Customer Due Date' field "
						+ "'Sales Orders - Order  Reason' = " + strOrderReason + "\n" 
						+ " Creation Date = " + CreationDate + "\n" 
						+ " Arrival Date = " + ArrivalDate;

					MessageBox.Show(msg);
			}
			
			return ret;
		}
		public void SetValues4eSparesTW()
		{
			DataSet srcDs, destDs;
			srcDs = frmInput.ds;
			destDs = frmOutput.ds;
			
			foreach (DataRow srcDr in srcDs.Tables["OpenOrdereSparesTW"].Rows)
			{
				DataRow dr = destDs.Tables["OpenOrdereSparesNew"].NewRow();
												
				//Filter 'Customer Status (Closed / Open)'
				if (srcDr["IC Shipments - Arrival Date"].ToString() == "")
				{
					continue;
				}
				
				if ( CustomerNameFilter(srcDr["Sales Orders - ShipTo Customer Name"].ToString()) == true )
				{
					continue;
				}
				
				if (srcDr["Sales Orders Line - Rejection 12 (Medium description)"].ToString().ToLower().Contains("direct delivery") ||
				    srcDr["Sales Orders Line - Rejection 12 (Medium description)"].ToString().ToLower().Contains("delete line") ||
				    srcDr["Sales Orders Line - Rejection 12 (Medium description)"].ToString().ToLower().Contains("served already") )
				{
					continue;
				}    
				    
				dr["Sales Orders - RSCIC/RSLC (Sales Organisation)"] = "";
				dr["COUNTRY"] = "TW";
				dr["Customer Status (Closed / Open)"] = "Open";	
				dr["Sales Orders - Type"] = srcDr["Sales Orders - Type"];
				dr["Sales Orders SoldTo - Customer Name"] = srcDr["Sales Orders - ShipTo Customer Name"];
				dr["Sales Orders - Order  Reason"] = srcDr["Sales Orders - Order  Reason"];
				dr["Sales Orders - Order Reason Service Description"] = srcDr["Sales Orders - Order  Reason (Medium description)"];
				dr["Sales Orders - Order Reason eSpares SLA"]  = "";
				dr["Sales Orders - Sales Order Number"]  = srcDr["Sales Orders - Sales Order Number"];
				dr["Sales Orders Line - LineNumber"]  = srcDr["Sales Orders Line - LineNumber"];
				dr["Sales Orders Line - Material"]  = srcDr["Sales Orders Line - Material"];
				dr["Sales Orders Line - SPT"]  = srcDr["Sales Orders Line - SPT"];
				dr["IC Delivery Line -  RcvdSerialNumber"]  = srcDr["IC Delivery Line -  RcvdSerialNumber"];
				dr["Work Orders  Header - Vendor Name"]  = "";
				dr["Sales Orders Line - Creation Date"]  = srcDr["Sales Orders - Creation Date"];
				dr["IC Shipments - Arrival Date"]  = srcDr["IC Shipments - Arrival Date"];
				dr["Sales Orders Line - Delivery SLA"] = srcDr["Sales Orders Line - Delivery SLA"];	//New column
				DateTime customerDueDate = CalcCustomerDueDate4eSparesTW(dr["Sales Orders - Order  Reason"].ToString(),
				                                                         dr["Sales Orders Line - Delivery SLA"].ToString(),
				                                                         dr["IC Shipments - Arrival Date"].ToString(),
				                                                         dr["Sales Orders Line - Creation Date"].ToString());
				
				if ( customerDueDate.Year == 1900 || customerDueDate.Year == 2010 || customerDueDate.Year == 2011 )
				{
					continue; //ust exclude
				}
				
				dr["Customer Due Date"]  = customerDueDate.ToString();
								
				try
				{
					if ( srcDr["Sales Orders - Order  Reason"].ToString().Substring(0, 1) == "A" ||
					     srcDr["Sales Orders - Order  Reason"].ToString().Substring(0, 1) == "H") //srcDr["Sales Orders - Order  Reason"];	
					{
						dr["SLA"] = "AE";
					}
					else 
					{
						dr["SLA"] = "R4S";
					}
				}
				catch
				{
						string msg = "OpenOrdereSparesNew: Error Occurred when processing 'SLA' field "
							+ "Sales Orders - Order  Reason' = "
							+ srcDr["Sales Orders - Order  Reason"].ToString();

						MessageBox.Show(msg);
				}
				
				try
				{
					if ( customerDueDate.Date <= DateTime.Now.Date )
					{
						dr["Delivery Status"] = "Out of TAT";
					}
					else
					{
						dr["Delivery Status"] = "Not yet due";
					}
				}
				catch
				{
					
				}
				
								
			
				if (dr["Delivery Status"].ToString() == "Not yet due")
				{
					dr["Overdue"] = "";
				}
				else if (dr["Delivery Status"].ToString() == "Out of TAT")
				{
					if (dr["SLA"].ToString() == "AE")
					{
						try
						{
						
							TimeSpan ts = DateTime.Now.Date.Subtract(Convert.ToDateTime(dr["Sales Orders Line - Creation Date"].ToString()));
							dr["Overdue"] = ts.Days;								
						}
						catch{}
						
					}
					else if (dr["SLA"].ToString() == "R4S")
					{
						try
						{
						
							TimeSpan ts = DateTime.Now.Date.Subtract(Convert.ToDateTime(dr["IC Shipments - Arrival Date"].ToString()));
							dr["Overdue"] = ts.Days;								
						}
						catch{}
					}
				}
				
				dr["Overdue Group"] = SetOverdueGroup(dr["Overdue"].ToString(), dr["SLA"].ToString());
				
				destDs.Tables["OpenOrdereSparesNew"].Rows.Add(dr);
			}
		}
	}
}
