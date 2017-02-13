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
using System.Collections;
using System.Collections.Specialized;
using System.IO;

namespace Report_Convertor
{
	public class WeeklyActivityYTDBasedOnOTD
	{
		public WeeklyActivityYTDBasedOnOTD()
		{
			
		}
			
		public void SetValues()
		{
			string LogFileName = System.Environment.CurrentDirectory + "\\Log-Weekly Activity Based On OTD.txt";
			FileStream fs = new FileStream(LogFileName, FileMode.Create,FileAccess.Write );
			fs.Close(); 
			
			SetOTDData("AE");
			SetOTDData("RFS");
		}
		

//========================================================================================
		private void SetOTDData(string serviceType)
		{
			StringCollection countryList, customerList;
			DataRow destDR;
			
			countryList = GetCountryCollection4OTD();
			foreach (string country in countryList)
			{
				int dueWeeklySubTotal = 0;
		    	int dispatchWeeklySubTotal = 0;
		    	int onTimeWeeklySubTotal = 0; 
	    	
		    	string OTSWeeklySubTotal = "";
		    	
				customerList = GetCustomerCollection4OTD(country);
				foreach (string customer in customerList)
				{		
					int due, dispatch, onTime;
		        	string OTS;		        	
		        	DataTable dt;
		        	
		        	if (serviceType == "RFS")
					{
						dt = frmOutput.ds.Tables["WeeklyActivity-RFS"];
					}
					else
					{
						dt = frmOutput.ds.Tables["WeeklyActivity-AE"];
					}
					destDR = dt.NewRow();
					destDR["COUNTRY"] = country;
					destDR["CUSTOMER"] = customer;
			
					CountOTD(serviceType, country, customer, out due, out dispatch, out onTime, out OTS);
		        	destDR["WEEKLY DUE"] = due;
					destDR["WEEKLY DISPATCHED SUM"] = dispatch.ToString();
		        	destDR["WEEKLY ON-TIME"] = onTime.ToString();
					destDR["WEEKLY OTS% OVERALL"] = OTS;
					
					dueWeeklySubTotal += due;
					dispatchWeeklySubTotal += dispatch;
					onTimeWeeklySubTotal += onTime;
			
					dt.Rows.Add(destDR);	
				}
				
				SetSubTotal4OTD( serviceType,
				        dueWeeklySubTotal,
		    			dispatchWeeklySubTotal,
		    			onTimeWeeklySubTotal,   
		    			OTSWeeklySubTotal);
			}		
		}
		
		private void CountOTD(string serviceType,
		                         			string country,
		                                    string customerName,
			  		                    	out int due,
		                                    out int dispatch, 
		                                    out int onTime,
		                                    out string OTS)
		{
			due = 0;
			dispatch = 0;
			onTime = 0;
		    OTS = "";
		    
			foreach (DataRow dr in frmInput.ds.Tables["WeeklyOTDActivity"].Rows)
			{
			  	if (dr["Sales Orders - ShipTo Country"].ToString() == country &&
				    dr["Sales Orders SoldTo - Customer Name"].ToString() == customerName)
			  	{
					if (serviceType == "AE" && (dr["Service Type"].ToString() == "D+" || dr["Service Type"].ToString() == "H+") ||
					    serviceType == dr["Service Type"].ToString()) //RFS
			  		{
			  			due++;			  		
			  			
			  			if ( dr["Net OTD Status (1=OK; 0=NOK)"].ToString() == "1")
			  			{
			  				onTime++;
			  			}	
			  			
			  			if ( dr["DELIVERY DATE (FINAL)"].ToString() != "" )
			  			{
			  				dispatch++;
			  			} 			
			  		}
				}
			}	
			
			if ( due == 0 )
			{
				OTS = "NA";
			}
			else
			{
				double d = Math.Round((double)onTime/due, 3) * 100;
				OTS = d.ToString() + "%";
			}
		}
		
		private bool CustomerNameFilter(string customerName)
		{
			if ( customerName == "台灣國際標準電子股份有限公司" ||	
			     customerName == "內政部消防署"  ||
			     customerName.ToUpper().Contains("ALCATEL") ||
			     customerName.ToUpper().Contains("LUCENT") )
			{
				return true;
			}
			else
			{
				return false;
			}
		}

		private StringCollection GetCustomerCollection4OTD(string country)
		{
			StringCollection ret = new StringCollection();
			string customer;
			
			foreach (DataRow dr in frmInput.ds.Tables["WeeklyOTDActivity"].Rows)
			{
				customer = dr["Sales Orders SoldTo - Customer Name"].ToString();
				if ( ( dr["Sales Orders - ShipTo Country"].ToString() == country ) &&
				     ( ret.Contains(customer) == false ) &&
				     ( CustomerNameFilter(customer) == false ) )
				{
					ret.Add(customer);
				}
			}
			
			return ret;
		}
			
		private StringCollection GetCountryCollection4OTD()
		{
			StringCollection ret = new StringCollection();
			string country;
			
			foreach (DataRow dr in frmInput.ds.Tables["WeeklyOTDActivity"].Rows)
			{
				country = dr["Sales Orders - ShipTo Country"].ToString();
				if (! ret.Contains(country))
				{
					ret.Add(country);
				}
			}
			
			return ret;
		}
			
		private void SetSubTotal4OTD(string serviceType, 
		                             int dueWeeklySubTotal,
		    						int dispatchWeeklySubTotal,
		    						int onTimeWeeklySubTotal,
		    						string OTSWeeklySubTotal)
		{
			DataRow destDR;	
			
			if ( dueWeeklySubTotal == 0 )
			{
				OTSWeeklySubTotal = "NA";
			}
			else
			{
				double d = Math.Round((double)onTimeWeeklySubTotal/dueWeeklySubTotal, 3) * 100;
				OTSWeeklySubTotal = d.ToString() + "%";
			}
			
			DataTable dt;
			
			if (serviceType == "RFS")
			{
				dt = frmOutput.ds.Tables["WeeklyActivity-RFS"];
			}
			else
			{
				dt = frmOutput.ds.Tables["WeeklyActivity-AE"];
			}
							
			destDR = dt.NewRow();
			destDR["CUSTOMER"] = "Country SubTotal";
					
			destDR["WEEKLY DUE"] = dueWeeklySubTotal;
			destDR["WEEKLY DISPATCHED SUM"] = dispatchWeeklySubTotal.ToString();
		    destDR["WEEKLY ON-TIME"] = onTimeWeeklySubTotal.ToString();
			destDR["WEEKLY OTS% OVERALL"] = OTSWeeklySubTotal;
			dt.Rows.Add(destDR);	
		}
					
		private void WriteLog(string msg)
		{
			string LogFileName = System.Environment.CurrentDirectory + "\\Log-Activity YTD Based On OTD.txt";
			StreamWriter writer = File.AppendText(LogFileName);
			
			try
			{
				writer.WriteLine(msg + "\r\n");
			}
			catch 
			{
			}
			finally
			{
				writer.Close();
			} 
		}
	}
}
	