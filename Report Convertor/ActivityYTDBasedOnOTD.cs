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
	public class ActivityYTDBasedOnOTD
	{
		private DateTime reportOTDMonthlyStartDate, reportOTDMonthlyEndDate, reportOTDYearlyStartDate, reportOTDYearlyEndDate;
		
		public ActivityYTDBasedOnOTD(DateTime startDate, DateTime endDate)
		{
			reportOTDMonthlyStartDate = startDate;
			reportOTDMonthlyEndDate = endDate;
			reportOTDYearlyStartDate = new DateTime(reportOTDMonthlyStartDate.Year, 1, 1, 0, 0, 0);
			reportOTDYearlyEndDate = reportOTDMonthlyEndDate;
		}
		
		public void SetValues()
		{
			string LogFileName = System.Environment.CurrentDirectory + "\\Log-Activity YTD Based On OTD.txt";
			FileStream fs = new FileStream(LogFileName, FileMode.Create,FileAccess.Write );
			fs.Close(); 
			
			SetOTDOnelogData();
			SetOTDNonOnelogData();
		}
		

//========================================================================================
		private void SetOTDOnelogData()
		{
			StringCollection countryList, customerList;
			DataRow destDR;
			
			countryList = GetCountryCollection4OTDOnelog();
			foreach (string country in countryList)
			{
				int dueMonthlySubTotal = 0;
		    	int dispatchMonthlySubTotal = 0;
		    	int onTimeMonthlySubTotal = 0; 
		    	int tatMonthlySubTotal = 0;
		    
		    	int dueYearlySubTotal = 0;
		    	int dispatchYearlySubTotal = 0;
		    	int onTimeYearlySubTotal = 0; 
		    	int tatYearlySubTotal = 0;
		    
		    	string AverageTATOverallMonthlySubTotal = "";
		    	string OTSMonthlySubTotal = "";
		    	string AverageTATOverallYearlySubTotal = "";
		    	string OTSYearlySubTotal = "";
		    	
				customerList = GetCustomerCollection4OTDOnelog(country);
				foreach (string customer in customerList)
				{		
					int due, dispatch, onTime, tat;
		        	string AverageTATOverall, OTS;
		        	
					destDR = frmOutput.ds.Tables["Activity YTD Based On OTD"].NewRow();
					destDR["COUNTRY"] = country;
					destDR["CUSTOMER"] = customer;
					
					CountOTDOnelog(country, customer, reportOTDMonthlyStartDate, reportOTDMonthlyEndDate, out due, out dispatch, out onTime, out tat, out AverageTATOverall, out OTS);
		        	destDR["MONTHLY DUE"] = due;
					destDR["MONTHLY DISPATCHED SUM"] = dispatch.ToString();
		        	destDR["MONTHLY ON-TIME"] = onTime.ToString();
		        	destDR["MONTHLY TAT SUM"] = tat.ToString();
		        	destDR["MONTHLY AVG TAT"] = AverageTATOverall;
					destDR["MONTHLY OTS% OVERALL"] = OTS;
					
					dueMonthlySubTotal += due;
					dispatchMonthlySubTotal += dispatch;
					onTimeMonthlySubTotal += onTime;
					tatMonthlySubTotal += tat;
					
					CountOTDOnelog(country, customer, reportOTDYearlyStartDate, reportOTDYearlyEndDate, out due, out dispatch, out onTime, out tat, out AverageTATOverall, out OTS);
		        	destDR["YEARLY DUE"] = due;
					destDR["YEARLY DISPATCHED SUM"] = dispatch.ToString();
		        	destDR["YEARLY ON-TIME"] = onTime.ToString();
		        	destDR["YEARLY TAT SUM"] = tat.ToString();
		        	destDR["YEARLY AVG TAT"] = AverageTATOverall;
					destDR["YEARLY OTS% OVERALL"] = OTS;
					
					dueYearlySubTotal += due;
					dispatchYearlySubTotal += dispatch;
					onTimeYearlySubTotal += onTime;
					tatYearlySubTotal += tat;
				
					frmOutput.ds.Tables["Activity YTD Based On OTD"].Rows.Add(destDR);	
				}
				
				SetSubTotal4OTD( dueMonthlySubTotal,
		    			dispatchMonthlySubTotal,
		    			onTimeMonthlySubTotal,
		    			tatMonthlySubTotal,		    
						dueYearlySubTotal,
		    			dispatchYearlySubTotal,
		    			onTimeYearlySubTotal,
		    			tatYearlySubTotal,
		    			AverageTATOverallMonthlySubTotal, 
		    			OTSMonthlySubTotal, 
		    			AverageTATOverallYearlySubTotal, 
		    			OTSYearlySubTotal );
			}		
		}
		
		private void CountOTDOnelog(string country,
		                                    string customerName,
		                                    DateTime reportStartDate,
			  		                    	DateTime reportEndDate,
			  		                    	out int due,
		                                    out int dispatch, 
		                                    out int onTime,
		                                    out int tat,
		                                    out string AverageTATOverall,
		                                    out string OTS)
		{
			due = 0;
			dispatch = 0;
			onTime = 0;
		    AverageTATOverall = "";
		    tat = 0;
		    OTS = "";
		    
			foreach (DataRow dr in frmInput.ds.Tables["OTD-Onelog"].Rows)
			{
			  	if (dr["Country"].ToString() == country &&
				    dr["Customer Name"].ToString() == customerName)
			  	{
					int flag = IsDateInDomain( dr["Target Compl Date"].ToString(), 
			  		                    reportStartDate,
			  		                    reportEndDate );
					if ( flag == -1 )
					{
			  			string msg = "OTD-Onelog: Please check if 'Target Compl Date' is valid.\n" +
			  						"Country: " + country + "\n" +
			  						"Customer: " + customerName + "\n" +
			  						"Target Compl Date: " + dr["Target Compl Date"].ToString();
						WriteLog(msg);
						
						continue;
			  		}
			  		else if ( flag == 1 )
			  		{
			  			due++;			  		
			  			
			  			if ( dr["Delivery Status"].ToString() == "On Time")
			  			{
			  				onTime++;
			  			}	
			  			
			  			if ( dr["Ship_Date_Repl"].ToString() != "" )
			  			{
			  				dispatch++;
			  			}
			  			
			  			int tempTat = 0;
			  			if (dr["SLA"].ToString() == "D+" || dr["SLA"].ToString() == "H+")
			  			{
			  				tempTat = CalculateTAT( dr["Ship_Date_Repl"].ToString(), dr["Create_Date"].ToString() );
			  			}
			  			else if (dr["SLA"].ToString() == "R4S")
			  			{
			  				tempTat = CalculateTAT( dr["Ship_Date_Repl"].ToString(), dr["Rcv_Date_Def"].ToString() );
			  			}
			  			
			  			if (tempTat == -10000)
			  			{
			  					string msg = "OTD-Onelog: Error occurred when calculating TAT.\n" +
			  						"Country: " + country + "\n" +
			  						"Customer: " + customerName + "\n" +
			  						"Ship_Date_Repl: " + dr["Ship_Date_Repl"].ToString() + "\n" +			  						
			  						"Create_Date: " + dr["Create_Date"].ToString() + "\n" +
			  						"Rcv_Date_Def: " + dr["Rcv_Date_Def"].ToString();
								WriteLog(msg);
			  			}		
			  			else
			  			{
			  				tat += tempTat;
			  			}
			  		}
			  	}
			}	
			
			if ( dispatch == 0 )
			{
				AverageTATOverall = "NA";				
			}
			else 
			{
				AverageTATOverall = Convert.ToInt32(Math.Round((double)tat/dispatch, 0)).ToString();
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

		private StringCollection GetCustomerCollection4OTDOnelog(string country)
		{
			StringCollection ret = new StringCollection();
			string customer;
			
			foreach (DataRow dr in frmInput.ds.Tables["OTD-Onelog"].Rows)
			{
				customer = dr["Customer Name"].ToString();
				if ( ( dr["Country"].ToString() == country ) &&
				     ( ret.Contains(customer) == false ) &&
				     ( CustomerNameFilter(customer) == false ) )
				{
					ret.Add(customer);
				}
			}
			
			return ret;
		}
			
		private StringCollection GetCountryCollection4OTDOnelog()
		{
			StringCollection ret = new StringCollection();
			string country;
			
			foreach (DataRow dr in frmInput.ds.Tables["OTD-Onelog"].Rows)
			{
				country = dr["Country"].ToString();
				if (! ret.Contains(country))
				{
					ret.Add(country);
				}
			}
			
			return ret;
		}
			
		
//========================================================================================
		private void SetOTDNonOnelogData()
		{
			StringCollection countryList, customerList;
			DataRow destDR;
			
			countryList = GetCountryCollection4OTDNonOnelog();
			foreach (string country in countryList)
			{
				int dueMonthlySubTotal = 0;
		    	int dispatchMonthlySubTotal = 0;
		    	int onTimeMonthlySubTotal = 0; 
		    	int tatMonthlySubTotal = 0;
		    
		    	int dueYearlySubTotal = 0;
		    	int dispatchYearlySubTotal = 0;
		    	int onTimeYearlySubTotal = 0; 
		    	int tatYearlySubTotal = 0;
		    
		    	string AverageTATOverallMonthlySubTotal = "";
		    	string OTSMonthlySubTotal = "";
		    	string AverageTATOverallYearlySubTotal = "";
		    	string OTSYearlySubTotal = "";
		    	
				customerList = GetCustomerCollection4OTDNonOnelog(country);
				foreach (string customer in customerList)
				{		
					int due, dispatch, onTime, tat;
		        	string AverageTATOverall, OTS;
		        	
					destDR = frmOutput.ds.Tables["Activity YTD Based On OTD"].NewRow();
					destDR["COUNTRY"] = country;
					destDR["CUSTOMER"] = customer;
					
					CountOTDNonOnelog(country, customer, reportOTDMonthlyStartDate, reportOTDMonthlyEndDate, out due, out dispatch, out onTime, out tat, out AverageTATOverall, out OTS);
		        	destDR["MONTHLY DUE"] = due;
					destDR["MONTHLY DISPATCHED SUM"] = dispatch.ToString();
		        	destDR["MONTHLY ON-TIME"] = onTime.ToString();
		        	destDR["MONTHLY TAT SUM"] = tat.ToString();
		        	destDR["MONTHLY AVG TAT"] = AverageTATOverall;
					destDR["MONTHLY OTS% OVERALL"] = OTS;
					
					dueMonthlySubTotal += due;
					dispatchMonthlySubTotal += dispatch;
					onTimeMonthlySubTotal += onTime;
					tatMonthlySubTotal += tat;
					
					CountOTDNonOnelog(country, customer, reportOTDYearlyStartDate, reportOTDYearlyEndDate, out due, out dispatch, out onTime, out tat, out AverageTATOverall, out OTS);
		        	destDR["YEARLY DUE"] = due;
					destDR["YEARLY DISPATCHED SUM"] = dispatch.ToString();
		        	destDR["YEARLY ON-TIME"] = onTime.ToString();
		        	destDR["YEARLY TAT SUM"] = tat.ToString();
		        	destDR["YEARLY AVG TAT"] = AverageTATOverall;
					destDR["YEARLY OTS% OVERALL"] = OTS;
					
					dueYearlySubTotal += due;
					dispatchYearlySubTotal += dispatch;
					onTimeYearlySubTotal += onTime;
					tatYearlySubTotal += tat;
				
					frmOutput.ds.Tables["Activity YTD Based On OTD"].Rows.Add(destDR);	
				}
				
				SetSubTotal4OTD( dueMonthlySubTotal,
		    			dispatchMonthlySubTotal,
		    			onTimeMonthlySubTotal,
		    			tatMonthlySubTotal,		    
						dueYearlySubTotal,
		    			dispatchYearlySubTotal,
		    			onTimeYearlySubTotal,
		    			tatYearlySubTotal,
		    			AverageTATOverallMonthlySubTotal, 
		    			OTSMonthlySubTotal, 
		    			AverageTATOverallYearlySubTotal, 
		    			OTSYearlySubTotal );
			}		
		}
		
		private void CountOTDNonOnelog(string country,
		                                    string customerName,
		                                    DateTime reportStartDate,
			  		                    	DateTime reportEndDate,
			  		                    	out int due,
		                                    out int dispatch, 
		                                    out int onTime,
		                                    out int tat,
		                                    out string AverageTATOverall,
		                                    out string OTS)
		{
			due = 0;
			dispatch = 0;
			onTime = 0;
		    AverageTATOverall = "";
		    tat = 0;
		    OTS = "";
		    
			foreach (DataRow dr in frmInput.ds.Tables["OTD-NonOnelog"].Rows)
			{
			  	if (dr["Country"].ToString() == country &&
				    dr["Customer Name"].ToString() == customerName)
			  	{
					int flag = IsDateInDomain( dr["Target Due Date"].ToString(), 
			  		                    reportStartDate,
			  		                    reportEndDate );
					if ( flag == -1 )
					{
			  			string msg = "OTD-Onelog: Please check if 'Target Due Date' is valid.\n" +
			  						"Country:" + country + "\n" +
			  						"Customer: " + customerName + "\n" +
			  						"Target Due Date: " + dr["Target Due Date"].ToString();
						WriteLog(msg);
						
						continue;
			  		}
			  		else if ( flag == 1 )
			  		{
			  			due++;			  		
			  			
			  			if ( dr["Delivery Status"].ToString() == "On Time")
			  			{
			  				onTime++;
			  			}	
			  			
			  			if ( dr["Dispatched Date to Customer"].ToString() != "" )
			  			{
			  				dispatch++;
			  			}
			  			
			  			int tempTat = 0;
			  			
			  			if (dr["SLA"].ToString() == "D+" || dr["SLA"].ToString() == "H+")
			  			{
			  				tempTat = CalculateTAT( dr["Dispatched Date to Customer"].ToString(), dr["Customer Request Date"].ToString() );
			  			}
			  			else if (dr["SLA"].ToString() == "R4S")
			  			{
			  				tempTat = CalculateTAT( dr["Dispatched Date to Customer"].ToString(), dr["Faulty Board Received Date from Customer"].ToString() );
			  			}
			  			
//			  			if (tempTat < 0)
//			  			{
//			  				string temp1 = dr["Dispatched Date to Customer"].ToString();
//			  				string temp2 = dr["Customer Request Date"].ToString();
//			  				string temp3 = dr["Faulty Board Received Date from Customer"].ToString();
//			  				string temp4 = dr["SLA"].ToString();
//			  			}
			  			
			  			if (tempTat == -10000)
			  			{
			  					string msg = "OTD-NonOnelog: Error occurred when calculating TAT.\n" +
			  						"Country: " + country + "\n" +
			  						"Customer: " + customerName + "\n" +
			  						"Dispatched Date to Customer: " + dr["Dispatched Date to Customer"].ToString() + "\n" +			  						
			  						"Customer Request Date: " + dr["Customer Request Date"].ToString() + "\n" +
			  						"Faulty Board Received Date from Customer: " + dr["Faulty Board Received Date from Customer"].ToString();
								WriteLog(msg);
			  			}
			  			else
			  			{
			  				tat += tempTat;
			  			}
			  		}
			  	}
			}	
			
			if ( dispatch == 0 )
			{
				AverageTATOverall = "NA";				
			}
			else 
			{
				AverageTATOverall = Convert.ToInt32(Math.Round((double)tat/dispatch, 0)).ToString();
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

		private StringCollection GetCustomerCollection4OTDNonOnelog(string country)
		{
			StringCollection ret = new StringCollection();
			string customer;
			
			foreach (DataRow dr in frmInput.ds.Tables["OTD-NonOnelog"].Rows)
			{
				customer = dr["Customer Name"].ToString();
				if ( ( dr["Country"].ToString() == country ) &&
				     ( ret.Contains(customer) == false ) && 
				     ( CustomerNameFilter(customer) == false ) )
				{
					ret.Add(customer);
				}
			}
			
			return ret;
		}
			
		private StringCollection GetCountryCollection4OTDNonOnelog()
		{
			StringCollection ret = new StringCollection();
			string country;
			
			foreach (DataRow dr in frmInput.ds.Tables["OTD-NonOnelog"].Rows)
			{
				country = dr["Country"].ToString();
				if ((! ret.Contains(country)) && country != "")
				{
					ret.Add(country);
				}
			}
			
			return ret;
		}
//=====================================================================================
						
		// -10000 --> error
		private int CalculateTAT(string end, string start)
		{
			try
			{
				DateTime startDate = Convert.ToDateTime(start).Date;
				DateTime endDate = Convert.ToDateTime(end).Date;						
						
				TimeSpan ts = endDate.Subtract(startDate);
				return ts.Days;								
			}
			catch
			{
				return -10000;
			}
		}
		
		// 0 --> false
		// 1 --> true
		// -1--> error
		private int IsDateInDomain(string dateString, DateTime startDate, DateTime endDate)
		{
			DateTime dt;
			
			try
			{
				dt = Convert.ToDateTime(dateString).Date;
			}
			catch
			{
				return -1;
			}
			
			if ( dt >= startDate && dt < endDate.AddDays(1) )
			{
				return 1;
			}
			
			return 0;
		}
		
		private void SetSubTotal4OTD( int dueMonthlySubTotal,
		    						int dispatchMonthlySubTotal,
		    						int onTimeMonthlySubTotal,
		    						int tatMonthlySubTotal,		    
									int dueYearlySubTotal,
		    						int dispatchYearlySubTotal,
		    						int onTimeYearlySubTotal,
		    						int tatYearlySubTotal,
		    						string AverageTATOverallMonthlySubTotal, 
		    						string OTSMonthlySubTotal, 
		    						string AverageTATOverallYearlySubTotal, 
		    						string OTSYearlySubTotal )
		{
			DataRow destDR;
			
			if ( dispatchMonthlySubTotal == 0 )
			{
				AverageTATOverallMonthlySubTotal = "NA";
			}
			else 
			{
				AverageTATOverallMonthlySubTotal = Convert.ToInt32(Math.Round((double)tatMonthlySubTotal/dispatchMonthlySubTotal, 0)).ToString();
			}	
			
			if ( dueMonthlySubTotal == 0 )
			{
				OTSMonthlySubTotal = "NA";
			}
			else
			{
				double d = Math.Round((double)onTimeMonthlySubTotal/dueMonthlySubTotal, 3) * 100;
				OTSMonthlySubTotal = d.ToString() + "%";
			}
				
			
			if ( dispatchYearlySubTotal == 0 )
			{
				AverageTATOverallYearlySubTotal = "NA";				
			}
			else 
			{
				AverageTATOverallYearlySubTotal = Convert.ToInt32(Math.Round((double)tatYearlySubTotal/dispatchYearlySubTotal, 0)).ToString();
			}	
			
			if ( dueYearlySubTotal == 0 )
			{
				OTSYearlySubTotal = "NA";
			}
			else
			{
				double d = Math.Round((double)onTimeYearlySubTotal/dueYearlySubTotal, 3) * 100;
				OTSYearlySubTotal = d.ToString() + "%";
			}
				
			
			destDR = frmOutput.ds.Tables["Activity YTD Based On OTD"].NewRow();
			destDR["CUSTOMER"] = "Country SubTotal";
					
			destDR["MONTHLY DUE"] = dueMonthlySubTotal;
			destDR["MONTHLY DISPATCHED SUM"] = dispatchMonthlySubTotal.ToString();
		    destDR["MONTHLY ON-TIME"] = onTimeMonthlySubTotal.ToString();
		    destDR["MONTHLY TAT SUM"] = tatMonthlySubTotal.ToString();
		    destDR["MONTHLY AVG TAT"] = AverageTATOverallMonthlySubTotal;
			destDR["MONTHLY OTS% OVERALL"] = OTSMonthlySubTotal;
			destDR["YEARLY DUE"] = dueYearlySubTotal;
			destDR["YEARLY DISPATCHED SUM"] = dispatchYearlySubTotal.ToString();
		    destDR["YEARLY ON-TIME"] = onTimeYearlySubTotal.ToString();
		    destDR["YEARLY TAT SUM"] = tatYearlySubTotal.ToString();
		    destDR["YEARLY AVG TAT"] = AverageTATOverallYearlySubTotal;
			destDR["YEARLY OTS% OVERALL"] = OTSYearlySubTotal;
			frmOutput.ds.Tables["Activity YTD Based On OTD"].Rows.Add(destDR);	
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
	