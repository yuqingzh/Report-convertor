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
	

	public class ActivityYTD
	{
		private DateTime reportMonthlyStartDate, reportMonthlyEndDate, reportYearlyStartDate, reportYearlyEndDate;
		public ActivityYTD(DateTime startDate, DateTime endDate)
		{
			reportMonthlyStartDate = startDate;
			reportMonthlyEndDate = endDate;
			reportYearlyStartDate = new DateTime(reportMonthlyStartDate.Year, 1, 1, 0, 0, 0);
			reportYearlyEndDate = reportMonthlyEndDate;
		}
		
		public void SetValues()
		{
			string LogFileName = System.Environment.CurrentDirectory + "\\Log-Activity YTD.txt";
			FileStream fs = new FileStream(LogFileName, FileMode.Create,FileAccess.Write );
			fs.Close(); 
			
			SetTWData();
			SetKoreaData();			
			SetNZData();
			SetCitadelData();
			SeteSparesNewData();
		}
		
		private string GetCountryName4eSparesNew(string CountryCode)
		{
			string ret = "";
			
			foreach (DataRow dr in frmInput.ds.Tables["ActivityYTD-eSparesNew - Country Code List"].Rows)
			{
				if (CountryCode.ToUpper() == dr["COUNTRY CODE"].ToString().ToUpper())
			  	{
			  		ret = dr["COUNTRY NAME"].ToString().ToUpper();
			  		break;
			  	}
			}
			
			if (CountryCode != "" && ret == "")
			{
				ret = "Country name not found!";
			}
			
			return ret;
			
		}
				
		private StringCollection GetCountryCollection4eSparesNew()
		{
			StringCollection ret = new StringCollection();
			string country;
			
			foreach (DataRow dr in frmInput.ds.Tables["ActivityYTD-eSpares New"].Rows)
			{
				country = GetCountryName4eSparesNew(dr["Sales Orders - RSCIC/RSLC (Sales Organisation)"].ToString()).ToUpper();
				if ( country.ToUpper().Contains("KOREA"))
				{
					continue;
				}
		
				if ((! ret.Contains(country)) && country != "")
				{
					ret.Add(country);
				}
			}
			
			return ret;
		}
		
		private StringCollection GetCustomerCollection4eSparesNew(string country)
		{
			StringCollection ret = new StringCollection();
			string customer;
			
			foreach (DataRow dr in frmInput.ds.Tables["ActivityYTD-eSpares New"].Rows)
			{
				customer = dr["Sales Orders SoldTo - Customer Name"].ToString();
				if ( ( country.ToUpper() == GetCountryName4eSparesNew(dr["Sales Orders - RSCIC/RSLC (Sales Organisation)"].ToString()).ToUpper() ) &&
				     ( ret.Contains(customer) == false ) && 
				     ( CustomerNameFilter(customer) == false ) )
				{
					ret.Add(customer);
				}
			}		
			
			return ret;
		}
		
		private void SeteSparesNewData()
		{
			StringCollection countryList, customerList;
			DataRow destDR;
			
			countryList = GetCountryCollection4eSparesNew();
			foreach (string country in countryList)
			{
				if (GetCountryCollection4Citadel().Contains(country))
				{
					//Has been processed in SetCitadelData() 
					continue;
				}
				
				int receiptMonthlySubTotal = 0;
		    	int dispatchMonthlySubTotal = 0;
		    	int dispatchWithinSLAMonthlySubTotal = 0; 
		    	int tatMonthlySubTotal = 0;
		    
		    	int receiptYearlySubTotal = 0;
		    	int dispatchYearlySubTotal = 0;
		    	int dispatchWithinSLAYearlySubTotal = 0; 
		    	int tatYearlySubTotal = 0;
		    
		    	string AverageTATOverallMonthlySubTotal = "";
		    	string OTSMonthlySubTotal = "";
		    	string AverageTATOverallYearlySubTotal = "";
		    	string OTSYearlySubTotal = "";
		    	
				customerList = GetCustomerCollection4eSparesNew(country);
				foreach (string customer in customerList)
				{		
					int receipt, dispatch, dispatchWithinSLA, tat;
		        	string AverageTATOverall, OTS;
		        	
					destDR = frmOutput.ds.Tables["Activity YTD"].NewRow();
					destDR["COUNTRY"] = country;
					destDR["CUSTOMER"] = customer;
					
					receipt = CountReceiptVolume4eSparesNew(country, customer, reportMonthlyStartDate, reportMonthlyEndDate);
		        	destDR["MONTHLY RECEIPT"] = receipt;
					CountDispatchVolume4eSparesNew(country, customer, reportMonthlyStartDate, reportMonthlyEndDate, out dispatch, out dispatchWithinSLA, out tat, out AverageTATOverall, out OTS);
		        	destDR["MONTHLY DISPATCH"] = dispatch.ToString();
		        	destDR["MONTHLY DISPATCH WITHIN SLA"] = dispatchWithinSLA.ToString();
		        	destDR["MONTHLY TAT"] = tat.ToString();
		        	destDR["MONTHLY AVG TAT OVERALL"] = AverageTATOverall;
					destDR["MONTHLY OTS%"] = OTS;
					
					receiptMonthlySubTotal += receipt;
					dispatchMonthlySubTotal += dispatch;
					dispatchWithinSLAMonthlySubTotal += dispatchWithinSLA;
					tatMonthlySubTotal += tat;
					
					receipt = CountReceiptVolume4eSparesNew(country, customer, reportYearlyStartDate, reportYearlyEndDate);
		        	destDR["YTD RECEIPT"] = receipt;
					CountDispatchVolume4eSparesNew(country, customer, reportYearlyStartDate, reportYearlyEndDate, out dispatch, out dispatchWithinSLA, out tat, out AverageTATOverall, out OTS);
		        	destDR["YTD DISPATCH"] = dispatch.ToString();
		        	destDR["YTD DISPATCH WITHIN SLA"] = dispatchWithinSLA.ToString();
		        	destDR["YTD TAT"] = tat.ToString();
		        	destDR["YTD AVG TAT OVERALL"] = AverageTATOverall;
					destDR["YTD OTS%"] = OTS;
					
					receiptYearlySubTotal += receipt;
					dispatchYearlySubTotal += dispatch;
					dispatchWithinSLAYearlySubTotal += dispatchWithinSLA;
					tatYearlySubTotal += tat;
				
					frmOutput.ds.Tables["Activity YTD"].Rows.Add(destDR);	
				}
				
				SetSubTotal( receiptMonthlySubTotal,
		    			dispatchMonthlySubTotal,
		    			dispatchWithinSLAMonthlySubTotal,
		    			tatMonthlySubTotal,		    
						receiptYearlySubTotal,
		    			dispatchYearlySubTotal,
		    			dispatchWithinSLAYearlySubTotal,
		    			tatYearlySubTotal,
		    			AverageTATOverallMonthlySubTotal, 
		    			OTSMonthlySubTotal, 
		    			AverageTATOverallYearlySubTotal, 
		    			OTSYearlySubTotal );
			}		
		}
		
		private int CountReceiptVolume4eSparesNew(string country, 
		                                         string customerName, 
		                                         DateTime reportStartDate,
			  		                    		 DateTime reportEndDate)
		{
			int ret = 0;
			foreach (DataRow dr in frmInput.ds.Tables["ActivityYTD-eSpares New"].Rows)
			{
			  	if ( ( country.ToUpper() == GetCountryName4eSparesNew(dr["Sales Orders - RSCIC/RSLC (Sales Organisation)"].ToString()) ) && 
				     dr["Sales Orders SoldTo - Customer Name"].ToString() == customerName)
			  	{
					if ( dr["Sales Orders - Order  Reason"].ToString().Trim() == "")
					{
						continue;
					}
					
					if ( ( dr["Sales Orders - Order  Reason"].ToString().Substring(0, 1) == "A" ||
					       dr["Sales Orders - Order  Reason"].ToString().Substring(0, 1) == "H" ) &&
					          !( dr["Sales Orders - RSCIC/RSLC (Sales Organisation)"].ToString() == "TH01" &&
					      		 dr["Sales Orders - Order  Reason"].ToString() == "A30") ) 
					{
						//dr["SLA"] = "AE";
						int flag = IsDateInDomain( dr["Sales Orders Line - Creation Date"].ToString(), 
			  		                    reportStartDate,
			  		                    reportEndDate );
						if ( flag == -1 )
						{
			  				string msg = "ActivityYTD-eSpares New: Please check if 'Sales Orders Line - Creation Date' is valid.\n" +
			  						"Customer: " + customerName + "\n" +
			  						"Sales Orders Line - Creation Date: " + dr["Sales Orders Line - Creation Date"].ToString();
							WriteLog(msg);
						
							continue;
			  			}
			  			else if ( flag == 1 )
			  			{
			  				ret++;
			  			}
					}
					else 
					{
						//dr["SLA"] = "R4S";
						int flag = IsDateInDomain( dr["IC Shipments - Arrival Date"].ToString(), 
			  		                    reportStartDate,
			  		                    reportEndDate );
						if ( flag == -1 )
						{
			  				string msg = "ActivityYTD-eSpares New: Please check if 'IC Shipments - Arrival Date' is valid.\n" +
			  						"Customer: " + customerName + "\n" +
			  						"IC Shipments - Arrival Date: " + dr["IC Shipments - Arrival Date"].ToString();
							WriteLog(msg);
						
							continue;
			  			}
			  			else if ( flag == 1 )
			  			{
			  				ret++;
			  			}
					}
			  		
			  	}
			}	
			
			return ret;
		}
		
		private void CountDispatchVolume4eSparesNew(string country,
		                                    string customerName,
		                                    DateTime reportStartDate,
			  		                    	DateTime reportEndDate,
		                                    out int dispatch, 
		                                    out int dispatchWithinSLA,
		                                    out int tat,
		                                    out string AverageTATOverall,
		                                    out string OTS)
		{
			dispatch = 0;
			dispatchWithinSLA = 0;
		    AverageTATOverall = "";
		    tat = 0;
		    OTS = "";
		    
			foreach (DataRow dr in frmInput.ds.Tables["ActivityYTD-eSpares New"].Rows)
			{
				
			  	if ( ( country.ToUpper() == GetCountryName4eSparesNew(dr["Sales Orders - RSCIC/RSLC (Sales Organisation)"].ToString()) ) && 
				     dr["Sales Orders SoldTo - Customer Name"].ToString() == customerName)
			  	{
					if ( dr["Sales Orders - Order  Reason"].ToString().Trim() == "")
					{
						continue;
					}
					
			  		int flag = IsDateInDomain( dr["OC Delivery Header -  GI_Date"].ToString(), 
			  		                    reportStartDate,
			  		                    reportEndDate );
					if ( flag == -1 )
					{
			  			string msg = "ActivityYTD-eSpares New: Please check if 'OC Delivery Header -  GI_Date' is valid.\n" +
			  						"Country: " + country + "\n" +
			  						"Customer: " + customerName + "\n" +
			  						"OC Delivery Header -  GI_Date: " + dr["OC Delivery Header -  GI_Date"].ToString();
						WriteLog(msg);
						
						continue;
			  		}
			  		else if ( flag == 1 )
			  		{
			  			dispatch++;
			  						  						
			  			if ( dr["TAT Status"].ToString() == "In TAT")
			  			{
			  				dispatchWithinSLA++;
			  			}	
			  			
			  			int currTAT = 0;
			  			
			  			if ( ( dr["Sales Orders - Order  Reason"].ToString().Substring(0, 1) == "A" ||
					       	   dr["Sales Orders - Order  Reason"].ToString().Substring(0, 1) == "H" ) &&
					          !( dr["Sales Orders - RSCIC/RSLC (Sales Organisation)"].ToString() == "TH01" &&
					             dr["Sales Orders - Order  Reason"].ToString() == "A30") )
						{
							//dr["SLA"] = "AE";
							currTAT = CalculateTAT( dr["OC Delivery Header -  GI_Date"].ToString(), dr["Sales Orders Line - Creation Date"].ToString() );	
			  							  			
			  				if (currTAT == -10000)
			  				{
			  					string msg = "ActivityYTD-eSpares New: Error occurred when calculating TAT.\n" +
			  						"Country: " + country + "\n" + 
			  						"Customer: " + customerName + "\n" +
			  						"OC Delivery Header -  GI_Date: " + dr["OC Delivery Header -  GI_Date"].ToString() + "\n" +	 
			  						"Sales Orders Line - Creation Date: " + dr["Sales Orders Line - Creation Date"].ToString();
								WriteLog(msg);
			  				}		
			  				else
			  				{
			  					tat += currTAT;
			  				}
						}
						else 
						{
							//dr["SLA"] = "R4S";
							currTAT = CalculateTAT( dr["OC Delivery Header -  GI_Date"].ToString(), dr["IC Shipments - Arrival Date"].ToString() );	
			  							  			
			  				if (currTAT == -10000)
			  				{
			  					string msg = "ActivityYTD-eSpares New: Error occurred when calculating TAT.\n" +
			  						"Country: " + country + "\n" + 
			  						"Customer: " + customerName + "\n" +
			  						"OC Delivery Header -  GI_Date: " + dr["OC Delivery Header -  GI_Date"].ToString() + "\n" +	 
			  						"IC Shipments - Arrival Date: " + dr["IC Shipments - Arrival Date"].ToString();
								WriteLog(msg);
			  				}		
			  				else
			  				{
			  					tat += currTAT;
			  				}
						}
			  		
			  		}
			  	}
			}	
			
			if ( dispatch == 0 )
			{
				AverageTATOverall = "NA";
				OTS = "NA";
			}
			else 
			{
				AverageTATOverall = Convert.ToInt32(Math.Round((double)tat/dispatch, 0)).ToString();
				double d = Math.Round((double)dispatchWithinSLA/dispatch, 3) * 100;
				OTS = d.ToString() + "%";
			}			
		}
//=============================================================================================================		
		private void SetNZData()
		{
			StringCollection customerList;
			DataRow destDR;
			
			int receiptMonthlySubTotal = 0;
		    int dispatchMonthlySubTotal = 0;
		    int dispatchWithinSLAMonthlySubTotal = 0; 
		    int tatMonthlySubTotal = 0;
		    
		    int receiptYearlySubTotal = 0;
		    int dispatchYearlySubTotal = 0;
		    int dispatchWithinSLAYearlySubTotal = 0; 
		    int tatYearlySubTotal = 0;
		    
		    string AverageTATOverallMonthlySubTotal = "";
		    string OTSMonthlySubTotal = "";
		    string AverageTATOverallYearlySubTotal = "";
		    string OTSYearlySubTotal = "";
		    
			customerList = GetCustomerCollection4NZ();
			foreach (string customer in customerList)
			{
		        int receipt;
		        int dispatch, dispatchWithinSLA, tat;
		    	string AverageTATOverall, OTS;
		    	
				destDR = frmOutput.ds.Tables["Activity YTD"].NewRow();
				destDR["COUNTRY"] = "NEW ZEALAND";
				destDR["CUSTOMER"] = customer;
				
				receipt = CountReceiptVolume4NZ(customer, reportMonthlyStartDate, reportMonthlyEndDate);		        
				destDR["MONTHLY RECEIPT"] = receipt;
				CountDispatchVolume4NZ(customer, reportMonthlyStartDate, reportMonthlyEndDate, out dispatch, out dispatchWithinSLA, out tat, out AverageTATOverall, out OTS);
		        destDR["MONTHLY DISPATCH"] = dispatch.ToString();
		        destDR["MONTHLY DISPATCH WITHIN SLA"] = dispatchWithinSLA.ToString();
		        destDR["MONTHLY TAT"] = tat.ToString();
		        destDR["MONTHLY AVG TAT OVERALL"] = AverageTATOverall;
				destDR["MONTHLY OTS%"] = OTS;				
				
				receiptMonthlySubTotal += receipt;
				dispatchMonthlySubTotal += dispatch;
				dispatchWithinSLAMonthlySubTotal += dispatchWithinSLA;
				tatMonthlySubTotal += tat;
				
				receipt = CountReceiptVolume4NZ(customer, reportYearlyStartDate, reportYearlyEndDate);
				destDR["YTD RECEIPT"] = receipt;
				CountDispatchVolume4NZ(customer, reportYearlyStartDate, reportYearlyEndDate, out dispatch, out dispatchWithinSLA, out tat, out AverageTATOverall, out OTS);
		        destDR["YTD DISPATCH"] = dispatch.ToString();
		        destDR["YTD DISPATCH WITHIN SLA"] = dispatchWithinSLA.ToString();
		        destDR["YTD TAT"] = tat.ToString();
		        destDR["YTD AVG TAT OVERALL"] = AverageTATOverall;
				destDR["YTD OTS%"] = OTS;
					
				receiptYearlySubTotal += receipt;
				dispatchYearlySubTotal += dispatch;
				dispatchWithinSLAYearlySubTotal += dispatchWithinSLA;
				tatYearlySubTotal += tat;
				
				frmOutput.ds.Tables["Activity YTD"].Rows.Add(destDR);
			}	
			
			SetSubTotal( receiptMonthlySubTotal,
		    			dispatchMonthlySubTotal,
		    			dispatchWithinSLAMonthlySubTotal,
		    			tatMonthlySubTotal,		    
						receiptYearlySubTotal,
		    			dispatchYearlySubTotal,
		    			dispatchWithinSLAYearlySubTotal,
		    			tatYearlySubTotal,
		    			AverageTATOverallMonthlySubTotal, 
		    			OTSMonthlySubTotal, 
		    			AverageTATOverallYearlySubTotal, 
		    			OTSYearlySubTotal );
		}
		
		private int CountReceiptVolume4NZ(string customerName, DateTime reportStartDate, DateTime reportEndDate)
		{
			int ret = 0;
			foreach (DataRow dr in frmInput.ds.Tables["NZ-Received-YTD"].Rows)
			{
			  	if (dr["Cust Name"].ToString() == customerName)
			  	{
			  		int flag = IsDateInDomain( dr["Received"].ToString(), 
			  		                    reportStartDate,
			  		                    reportEndDate );
					if ( flag == -1 )
					{
			  			string msg = "NZ-YTD: Please check if 'Received' is valid.\n" +
			  						"Customer: " + customerName + "\n" +
			  						"Received: " + dr["Received"].ToString();
						WriteLog(msg);
						
						continue;
			  		}
			  		else if ( flag == 1 )
			  		{
			  			ret++;
			  		}
			  	}
			}				
			return ret;
		}
		
		private void CountDispatchVolume4NZ(string customerName, 
		                                    DateTime reportStartDate,
			  		                    	DateTime reportEndDate,
		                                    out int dispatch, 
		                                    out int dispatchWithinSLA,
		                                    out int tat,
		                                    out string AverageTATOverall,
		                                    out string OTS)
		{
			dispatch = 0;
			dispatchWithinSLA = 0;
		    AverageTATOverall = "";
		    OTS = "";
		    tat = 0;
		    
			foreach (DataRow dr in frmInput.ds.Tables["NZ-YTD"].Rows)
			{
			  	if (dr["Country"].ToString() == "NEW ZEALAND" &&
				    dr["Customer Name"].ToString() == customerName)
			  	{
			  		int flag = IsDateInDomain( dr["Dispatched Date to Customer"].ToString(), 
			  		                    reportStartDate,
			  		                    reportEndDate );
					if ( flag == -1 )
					{
			  			string msg = "NZ-YTD: Please check if 'Dispatched Date to Customer' is valid.\n" +
			  						"Country: NEW ZEALAND \n" + 
			  						"Customer: " + customerName + "\n" +
			  						"Dispatched Date to Customer: " + dr["Dispatched Date to Customer"].ToString();
						WriteLog(msg);
						
						continue;
			  		}
			  		else if ( flag == 1 )
			  		{
			  			dispatch++;
			  			
			  			if ( CalculateTAT( dr["Dispatched Date to Customer"].ToString(), dr["Target Due Date"].ToString() ) <= 0 )
			  			{
			  				dispatchWithinSLA++;
			  			}	
			  			
			  			int currTAT = 0;
			  			if (dr["Customer TAT"].ToString() == "1")
			  			{
			  				currTAT = CalculateTAT( dr["Dispatched Date to Customer"].ToString(), dr["Customer Request Date"].ToString() );
			  			}
			  			else
			  			{
			  				currTAT = CalculateTAT( dr["Dispatched Date to Customer"].ToString(), dr["Faulty Board Received Date from Customer"].ToString() );
			  			}
			  			
			  			if (currTAT == -10000)
			  			{
			  					string msg = "NZ-YTD: Error occurred when calculating TAT.\n" +
			  						"Country: NEW ZEALAND \n" + 
			  						"Customer: " + customerName + "\n" +
			  						"Dispatched Date to Customer: " + dr["Dispatched Date to Customer"].ToString() + "\n" +			  						
			  						"Faulty Board Received Date from Customer: " + dr["Faulty Board Received Date from Customer"].ToString() + "\n" +
			  						"Customer Request Date: " + dr["Customer Request Date"].ToString();
								WriteLog(msg);
			  			}	
			  			else
			  			{
			  				tat += currTAT;
			  			}
			  		}
			  	}
			}	
			
			if ( dispatch == 0 )
			{
				AverageTATOverall = "NA";
				OTS = "NA";
			}
			else 
			{
				AverageTATOverall = Convert.ToInt32(Math.Round((double)tat/dispatch, 0)).ToString();
				double d = Math.Round((double)dispatchWithinSLA/dispatch, 3) * 100;
				OTS = d.ToString() + "%";
			}			
		}
		

//================================================================================================
		private void SetCitadelData()
		{
			StringCollection countryList, customerList;
			DataRow destDR;
			
			countryList = GetCountryCollection4Citadel();
			foreach (string country in countryList)
			{
				int receiptMonthlySubTotal = 0;
		    	int dispatchMonthlySubTotal = 0;
		    	int dispatchWithinSLAMonthlySubTotal = 0; 
		    	int tatMonthlySubTotal = 0;
		    
		    	int receiptYearlySubTotal = 0;
		    	int dispatchYearlySubTotal = 0;
		    	int dispatchWithinSLAYearlySubTotal = 0; 
		    	int tatYearlySubTotal = 0;
		    
		    	string AverageTATOverallMonthlySubTotal = "";
		    	string OTSMonthlySubTotal = "";
		    	string AverageTATOverallYearlySubTotal = "";
		    	string OTSYearlySubTotal = "";
		    	
//				customerList = GetCustomerCollection4Citadel(country);
		    	customerList = GetCustomerCollection4CitadelAndeSparesNew(country);
				foreach (string customer in customerList)
				{		
					int receipt, dispatch, dispatchWithinSLA, tat;
		        	string AverageTATOverall, OTS;
		        	
					destDR = frmOutput.ds.Tables["Activity YTD"].NewRow();
					destDR["COUNTRY"] = country;
					destDR["CUSTOMER"] = customer;
					
					receipt = CountReceiptVolume4Citadel(country, customer, reportMonthlyStartDate, reportMonthlyEndDate);
		        	destDR["MONTHLY RECEIPT"] = receipt;
					CountDispatchVolume4Citadel(country, customer, reportMonthlyStartDate, reportMonthlyEndDate, out dispatch, out dispatchWithinSLA, out tat, out AverageTATOverall, out OTS);
		        	destDR["MONTHLY DISPATCH"] = dispatch.ToString();
		        	destDR["MONTHLY DISPATCH WITHIN SLA"] = dispatchWithinSLA.ToString();
		        	destDR["MONTHLY TAT"] = tat.ToString();
		        	destDR["MONTHLY AVG TAT OVERALL"] = AverageTATOverall;
					destDR["MONTHLY OTS%"] = OTS;
					
					receiptMonthlySubTotal += receipt;
					dispatchMonthlySubTotal += dispatch;
					dispatchWithinSLAMonthlySubTotal += dispatchWithinSLA;
					tatMonthlySubTotal += tat;
					
					receipt = CountReceiptVolume4Citadel(country, customer, reportYearlyStartDate, reportYearlyEndDate);
		        	destDR["YTD RECEIPT"] = receipt;
					CountDispatchVolume4Citadel(country, customer, reportYearlyStartDate, reportYearlyEndDate, out dispatch, out dispatchWithinSLA, out tat, out AverageTATOverall, out OTS);
		        	destDR["YTD DISPATCH"] = dispatch.ToString();
		        	destDR["YTD DISPATCH WITHIN SLA"] = dispatchWithinSLA.ToString();
		        	destDR["YTD TAT"] = tat.ToString();
		        	destDR["YTD AVG TAT OVERALL"] = AverageTATOverall;
					destDR["YTD OTS%"] = OTS;
					
					receiptYearlySubTotal += receipt;
					dispatchYearlySubTotal += dispatch;
					dispatchWithinSLAYearlySubTotal += dispatchWithinSLA;
					tatYearlySubTotal += tat;
				
					frmOutput.ds.Tables["Activity YTD"].Rows.Add(destDR);	
				}
				
				SetSubTotal( receiptMonthlySubTotal,
		    			dispatchMonthlySubTotal,
		    			dispatchWithinSLAMonthlySubTotal,
		    			tatMonthlySubTotal,		    
						receiptYearlySubTotal,
		    			dispatchYearlySubTotal,
		    			dispatchWithinSLAYearlySubTotal,
		    			tatYearlySubTotal,
		    			AverageTATOverallMonthlySubTotal, 
		    			OTSMonthlySubTotal, 
		    			AverageTATOverallYearlySubTotal, 
		    			OTSYearlySubTotal );
			}		
		}
		
		private int CountReceiptVolume4Citadel(string country, 
		                                         string customerName, 
		                                         DateTime reportStartDate,
			  		                    		 DateTime reportEndDate)
		{
			int ret = 0;
			foreach (DataRow dr in frmInput.ds.Tables["CitadelReceived-YTD"].Rows)
			{
			  	if ( dr["Country"].ToString().ToUpper() == country.ToUpper() && dr["Customer_Name"].ToString() == customerName)
			  	{
			  		int flag = IsDateInDomain( dr["Rcv_Date_Def"].ToString(), 
			  		                    reportStartDate,
			  		                    reportEndDate );
					if ( flag == -1 )
					{
			  			string msg = "CitadelReceived-YTD: Please check if 'Rcv_Date_Def' is valid.\n" +
			  						"Customer: " + customerName + "\n" +
			  						"Rcv_Date_Def: " + dr["Rcv_Date_Def"].ToString();
						WriteLog(msg);
						
						continue;
			  		}
			  		else if ( flag == 1 )
			  		{
			  			ret++;
			  		}
			  	}
			}	
			
			if (GetCountryCollection4eSparesNew().Contains(country.ToUpper()))
			{
				ret += CountReceiptVolume4eSparesNew(country, customerName, reportStartDate, reportEndDate);
			}
			
			return ret;
		}
		
		private void CountDispatchVolume4Citadel(string country,
		                                    string customerName,
		                                    DateTime reportStartDate,
			  		                    	DateTime reportEndDate,
		                                    out int dispatch, 
		                                    out int dispatchWithinSLA,
		                                    out int tat,
		                                    out string AverageTATOverall,
		                                    out string OTS)
		{
			dispatch = 0;
			dispatchWithinSLA = 0;
		    AverageTATOverall = "";
		    tat = 0;
		    OTS = "";
		    
			foreach (DataRow dr in frmInput.ds.Tables["CitadelShipped-YTD"].Rows)
			{
			  	if (dr["Country"].ToString().ToUpper() == country.ToUpper() &&
				    dr["Customer_Name"].ToString() == customerName)
			  	{
			  		int flag = IsDateInDomain( dr["Ship_Date_Repl"].ToString(), 
			  		                    reportStartDate,
			  		                    reportEndDate );
					if ( flag == -1 )
					{
			  			string msg = "CitadelShipped-YTD: Please check if 'Ship_Date_Repl' is valid.\n" +
			  						"Country: " + country + "\n" +
			  						"Customer: " + customerName + "\n" +
			  						"Ship_Date_Repl: " + dr["Ship_Date_Repl"].ToString();
						WriteLog(msg);
						
						continue;
			  		}
			  		else if ( flag == 1 )
			  		{
			  			dispatch++;
			  			
			  			string Target_Compl_Date = CalculateCitadelTarget_Compl_Date(dr["Create_Date"].ToString(),
			  			                                                             dr["Rcv_Date_Def"].ToString(),
			  			                                                             dr["Service_Type"].ToString(),
			  			                                                             dr["Ticket_Type"].ToString());
			  			
			  			if (Target_Compl_Date == "")
			  			{
			  				string msg = "CitadelShipped-YTD: Target_Compl_Date is set to empty.\n" +
			  						"Country: " + country + "\n" + 
			  						"Customer: " + customerName;
								WriteLog(msg);
			  			}
			  			
			  			if ( CalculateTAT( dr["Ship_Date_Repl"].ToString(), Target_Compl_Date ) <= 0)
			  			{
			  				dispatchWithinSLA++;
			  			}	
			  			
			  			int currTAT = 0;
			  			if (dr["Rcv_Date_Def"].ToString() != "" )
			  			{
			  				currTAT = CalculateTAT( dr["Ship_Date_Repl"].ToString(), dr["Rcv_Date_Def"].ToString() );	
			  			}
			  			else {
			  				currTAT = CalculateTAT( dr["Ship_Date_Repl"].ToString(), dr["Create_Date"].ToString() );	
			  			}
			  			
			  			
			  			if (currTAT == -10000)
			  			{
			  					string msg = "CitadelShipped-YTD: Error occurred when calculating TAT.\n" +
			  						"Country: " + country + "\n" + 
			  						"Customer: " + customerName + "\n" +
			  						"Ship_Date_Repl: " + dr["Ship_Date_Repl"].ToString() + "\n" +	
			  						"Create_Date: " + dr["Create_Date"].ToString() + "\n" + 
			  						"Rcv_Date_Def: " + dr["Rcv_Date_Def"].ToString();
								WriteLog(msg);
			  			}		
			  			else
			  			{
			  				tat += currTAT;
			  			}
			  		}
			  	}
			}	
			
			if (GetCountryCollection4eSparesNew().Contains(country.ToUpper()))
			{
				int dispatcheSparesNew = 0; 
		    	int dispatchWithinSLAeSparesNew = 0;
		   	 	int tateSparesNew = 0;
		    	string AverageTATOveralleSparesNew = "";
		    	string OTSeSparesNew = "";
		                                    	
				CountDispatchVolume4eSparesNew(country, customerName, reportStartDate, reportEndDate, 
		                                out dispatcheSparesNew, out dispatchWithinSLAeSparesNew, out tateSparesNew, 
		                                out AverageTATOveralleSparesNew, out OTSeSparesNew);
		        	
				dispatch += dispatcheSparesNew;
				dispatchWithinSLA += dispatchWithinSLAeSparesNew;
				tat += tateSparesNew;
			}
			
			if ( dispatch == 0 )
			{
				AverageTATOverall = "NA";
				OTS = "NA";
			}
			else 
			{
				AverageTATOverall = Convert.ToInt32(Math.Round((double)tat/dispatch, 0)).ToString();
				double d = Math.Round((double)dispatchWithinSLA/dispatch, 3) * 100;
				OTS = d.ToString() + "%";
			}			
		}
		
		private string CalculateCitadelTarget_Compl_Date(string Create_Date,
			  			                                     string Rcv_Date_Def,
			  			                                     string Service_Type,
			  			                                     string Ticket_Type)
		{
			DateTime dt;
			string ret = "";
			
			if ( Service_Type.Contains("AE") )
			{
				try
				{
					dt = Convert.ToDateTime(Create_Date).Date;
					dt = dt.AddDays(GetDaysPerTicketType(Ticket_Type));
					ret = dt.Date.ToString();
				}
				catch
				{
					string msg = "CitadelShipped-YTD: Error occurred when calculating Target_Compl_Date.\n" +
			  						"Create_Date: " + Create_Date + "\n" + 
			  						"Ticket_Type: " + Ticket_Type + "\n" +
									"Add Days: " + GetDaysPerTicketType(Ticket_Type).ToString();
					
					WriteLog(msg);
				}
			}
			else if ( Service_Type.Contains("R4") )
			{
				try
				{
					if (Rcv_Date_Def != "")
					{
						dt = Convert.ToDateTime(Rcv_Date_Def).Date;
						dt = dt.AddDays(GetDaysPerTicketType(Ticket_Type));
						ret = dt.Date.ToString();
					}
					else{
						dt = Convert.ToDateTime(Create_Date).Date;
						dt = dt.AddDays(GetDaysPerTicketType(Ticket_Type));
						ret = dt.Date.ToString();
					}
					
				}
				catch
				{
					string msg = "CitadelShipped-YTD: Error occurred when calculating Target_Compl_Date.\n" +
			  						"Rcv_Date_Def: " + Rcv_Date_Def + "\n" + 
			  						"Ticket_Type: " + Ticket_Type + "\n" +
								 "Add Days: " + GetDaysPerTicketType(Ticket_Type).ToString();
					
					WriteLog(msg);
				}
			} 
			
			return ret;
		}
		
		private int GetDaysPerTicketType(string Ticket_Type)
		{
			int ret = 0;
			foreach (DataRow dr in frmInput.ds.Tables["Ticket_Type_Sla ALu_for Citadel"].Rows)
			{
				if (Ticket_Type.Contains("(" + dr["Id"].ToString() + ")")  ||
				     Ticket_Type.Contains(dr["Description"].ToString()))
				{
					ret = Convert.ToInt32(dr["Response"].ToString());
					break;
				}
			}
			
			if ( ret == 0 )
			{
				string msg = "CitadelShipped-YTD: Error occurred when calculating TicketType.\n" +
			  						"Ticket_Type: " + Ticket_Type;
					
				WriteLog(msg);
			}
			
			return ret;
		}
		
//================================================================================================

		private void SetKoreaData()
		{
			StringCollection customerList;
			DataRow destDR;
			
			int receiptMonthlySubTotal = 0;
		    int dispatchMonthlySubTotal = 0;
		    int dispatchWithinSLAMonthlySubTotal = 0; 
		    int tatMonthlySubTotal = 0;
		    
		    int receiptYearlySubTotal = 0;
		    int dispatchYearlySubTotal = 0;
		    int dispatchWithinSLAYearlySubTotal = 0; 
		    int tatYearlySubTotal = 0;
		    
		    string AverageTATOverallMonthlySubTotal = "";
		    string OTSMonthlySubTotal = "";
		    string AverageTATOverallYearlySubTotal = "";
		    string OTSYearlySubTotal = "";
		    
			customerList = GetCustomerCollection4Korea();
			foreach (string customer in customerList)
			{	
				int receipt, dispatch, dispatchWithinSLA, tat;
		        string AverageTATOverall, OTS;
		        
				destDR = frmOutput.ds.Tables["Activity YTD"].NewRow();
				destDR["COUNTRY"] = "KOREA  REPUBLIC OF";
				destDR["CUSTOMER"] = customer;
				
				receipt = CountReceiptVolume4Korea(customer, reportMonthlyStartDate, reportMonthlyEndDate);
				destDR["MONTHLY RECEIPT"] = receipt;
				CountDispatchVolume4Korea(customer, reportMonthlyStartDate, reportMonthlyEndDate, out dispatch, out dispatchWithinSLA, out tat, out AverageTATOverall, out OTS);
		        destDR["MONTHLY DISPATCH"] = dispatch.ToString();
		        destDR["MONTHLY DISPATCH WITHIN SLA"] = dispatchWithinSLA.ToString();
		        destDR["MONTHLY TAT"] = tat.ToString();
		        destDR["MONTHLY AVG TAT OVERALL"] = AverageTATOverall;
				destDR["MONTHLY OTS%"] = OTS;
				
				receiptMonthlySubTotal += receipt;
				dispatchMonthlySubTotal += dispatch;
				dispatchWithinSLAMonthlySubTotal += dispatchWithinSLA;
				tatMonthlySubTotal += tat;
					
				receipt = CountReceiptVolume4Korea(customer, reportYearlyStartDate, reportYearlyEndDate);
		        destDR["YTD RECEIPT"] = receipt;
				CountDispatchVolume4Korea(customer, reportYearlyStartDate, reportYearlyEndDate, out dispatch, out dispatchWithinSLA, out tat, out AverageTATOverall, out OTS);
		        destDR["YTD DISPATCH"] = dispatch.ToString();
		        destDR["YTD DISPATCH WITHIN SLA"] = dispatchWithinSLA.ToString();
		        destDR["YTD TAT"] = tat.ToString();
		        destDR["YTD AVG TAT OVERALL"] = AverageTATOverall;
				destDR["YTD OTS%"] = OTS;
					
				receiptYearlySubTotal += receipt;
				dispatchYearlySubTotal += dispatch;
				dispatchWithinSLAYearlySubTotal += dispatchWithinSLA;
				tatYearlySubTotal += tat;
				
				frmOutput.ds.Tables["Activity YTD"].Rows.Add(destDR);	
			}		
			
			SetSubTotal( receiptMonthlySubTotal,
		    			dispatchMonthlySubTotal,
		    			dispatchWithinSLAMonthlySubTotal,
		    			tatMonthlySubTotal,		    
						receiptYearlySubTotal,
		    			dispatchYearlySubTotal,
		    			dispatchWithinSLAYearlySubTotal,
		    			tatYearlySubTotal,
		    			AverageTATOverallMonthlySubTotal, 
		    			OTSMonthlySubTotal, 
		    			AverageTATOverallYearlySubTotal, 
		    			OTSYearlySubTotal );
		}
		
		private int CountReceiptVolume4Korea(string customerName, DateTime reportStartDate, DateTime reportEndDate)
		{
			int ret = 0;
			foreach (DataRow dr in frmInput.ds.Tables["KOREA-YTD-Base on RMA creation date"].Rows)
			{
			  	if (dr["Customer Name"].ToString() == customerName)
			  	{
			  		int flag = IsDateInDomain( dr["Create_Date"].ToString(), 
			  		                    reportStartDate,
			  		                    reportEndDate );
					if ( flag == -1 )
					{
			  			string msg = "KOREA-YTD-Base on RMA creation date: Please check if 'Create_Date' is valid.\n" +
			  						"Customer: " + customerName + "\n" +
			  						"Create_Date: " + dr["Create_Date"].ToString();
						WriteLog(msg);
						
						continue;
			  		}
			  		else if ( flag == 1 )
			  		{
			  			ret++;
			  		}
			  	}
			}	
			
			ret += CountReceiptVolume4Citadel("Korea Republic Of".ToUpper(), customerName, reportStartDate, reportEndDate);
			
			ret += CountReceiptVolume4eSparesNew("KOREA", customerName, reportStartDate, reportEndDate);
			
							
			return ret;
		}
		
		private void CountDispatchVolume4Korea(string customerName, 
		                                       DateTime reportStartDate,
			  		                    	   DateTime reportEndDate,
		                                    out int dispatch, 
		                                    out int dispatchWithinSLA,
		                                    out int tat,
		                                    out string AverageTATOverall,
		                                    out string OTS)
		{
			dispatch = 0;
			dispatchWithinSLA = 0;
		    AverageTATOverall = "";
		    OTS = "";
		    tat = 0;
		    
		    foreach (DataRow dr in frmInput.ds.Tables["KOREA-YTD-Base on shipping date"].Rows)
			{
			  	if ( dr["Customer Name"].ToString() == customerName )
			  	{
			  		int flag = IsDateInDomain( dr["Ship_Date_Repl"].ToString(), 
			  		                    reportStartDate,
			  		                    reportEndDate );
					if ( flag == -1 )
					{
			  			string msg = "KOREA-YTD-Base on shipping date: Please check if 'Ship_Date_Repl' is valid.\n" + 
			  						"Customer: " + customerName + "\n" +
			  						"Ship_Date_Repl: " + dr["Ship_Date_Repl"].ToString();
						WriteLog(msg);
						
						continue;
			  		}
			  		else if ( flag == 1 )
			  		{
			  			dispatch++;
			  			
			  			if ( dr["Delivery Status"].ToString() == "On Time" &&
			  			     dr["Ship_Date_Repl"].ToString() != "")
			  			{
			  				dispatchWithinSLA++;
			  			}	
			  			
			  			int currTAT = 0;
			  			currTAT = CalculateTAT( dr["Ship_Date_Repl"].ToString(), dr["Create_Date"].ToString() );			  			
			  			
			  			if (currTAT == -10000)
			  			{
			  					string msg = "KOREA-YTD-Base on shipping date: Error occurred when calculating TAT.\n" +
			  						"Country: KOREA  REPUBLIC OF \n" + 
			  						"Customer: " + customerName + "\n" +
			  						"Ship_Date_Repl: " + dr["Ship_Date_Repl"].ToString() + "\n" +			  						
			  						"Create_Date: " + dr["Create_Date"].ToString();
								WriteLog(msg);
			  			}	
			  			else
			  			{
			  				tat += currTAT;
			  			}
			  		}
			  	}
			}	
		    
		    int dispatchCitadel = 0; 
		    int dispatchWithinSLACitadel = 0;
		    int tatCitadel = 0;
		    string AverageTATOverallCitadel = "";
		    string OTSCitadel = "";
		                                    	
		    CountDispatchVolume4Citadel("Korea Republic Of".ToUpper(), customerName, reportStartDate, reportEndDate,
		                                out dispatchCitadel, out dispatchWithinSLACitadel, out tatCitadel, 
		                                out AverageTATOverallCitadel, out OTSCitadel);
		        	
			dispatch += dispatchCitadel;
			dispatchWithinSLA += dispatchWithinSLACitadel;
			tat += tatCitadel;
			
			int dispatcheSparesNew = 0; 
		    int dispatchWithinSLAeSparesNew = 0;
		    int tateSparesNew = 0;
		    string AverageTATOveralleSparesNew = "";
		    string OTSeSparesNew = "";
		                                    	
			CountDispatchVolume4eSparesNew("KOREA", customerName, reportStartDate, reportEndDate, 
		                                out dispatcheSparesNew, out dispatchWithinSLAeSparesNew, out tateSparesNew, 
		                                out AverageTATOveralleSparesNew, out OTSeSparesNew);
		        	
			dispatch += dispatcheSparesNew;
			dispatchWithinSLA += dispatchWithinSLAeSparesNew;
			tat += tateSparesNew;
			
			if ( dispatch == 0 )
			{
				AverageTATOverall = "NA";
				OTS = "NA";
			}
			else 
			{
				AverageTATOverall = Convert.ToInt32(Math.Round((double)tat/dispatch, 0)).ToString();
				double d = Math.Round((double)dispatchWithinSLA/dispatch, 3) * 100;
				OTS = d.ToString() + "%";
			}			
		}
		
//====================================================================================================

		private void SetTWData()
		{
			StringCollection customerList;
			DataRow destDR;
			
			int receiptMonthlySubTotal = 0;
		    int dispatchMonthlySubTotal = 0;
		    int dispatchWithinSLAMonthlySubTotal = 0; 
		    int tatMonthlySubTotal = 0;
		    
		    int receiptYearlySubTotal = 0;
		    int dispatchYearlySubTotal = 0;
		    int dispatchWithinSLAYearlySubTotal = 0; 
		    int tatYearlySubTotal = 0;
		    
		    string AverageTATOverallMonthlySubTotal = "";
		    string OTSMonthlySubTotal = "";
		    string AverageTATOverallYearlySubTotal = "";
		    string OTSYearlySubTotal = "";
		    
			customerList = GetCustomerCollection4TW("TW-IC-YTD", 
			                                     "Sales Orders - ShipTo Customer Name",
			                                     "TW",
			                                     "OTD-Onelog-TWOnly-YTD", 
			                                     "Country",
			                                     "Customer Name");
			foreach (string customer in customerList)
			{					
				int receipt, dispatch, dispatchWithinSLA, tat;
		        string AverageTATOverall, OTS;
		        
				destDR = frmOutput.ds.Tables["Activity YTD"].NewRow();
				destDR["COUNTRY"] = "TW";
				destDR["CUSTOMER"] = customer;
				
				receipt = CountReceiptVolume4TW(customer, reportMonthlyStartDate, reportMonthlyEndDate);
		        destDR["MONTHLY RECEIPT"] = receipt;
				CountDispatchVolume4TW(customer, reportMonthlyStartDate, reportMonthlyEndDate, out dispatch, out dispatchWithinSLA, out tat, out AverageTATOverall, out OTS);
		        destDR["MONTHLY DISPATCH"] = dispatch.ToString();
		        destDR["MONTHLY DISPATCH WITHIN SLA"] = dispatchWithinSLA.ToString();
		        destDR["MONTHLY TAT"] = tat.ToString();
		        destDR["MONTHLY AVG TAT OVERALL"] = AverageTATOverall;
				destDR["MONTHLY OTS%"] = OTS;
					
				receiptMonthlySubTotal += receipt;
				dispatchMonthlySubTotal += dispatch;
				dispatchWithinSLAMonthlySubTotal += dispatchWithinSLA;
				tatMonthlySubTotal += tat;
				
				receipt = CountReceiptVolume4TW(customer, reportYearlyStartDate, reportYearlyEndDate);
		        destDR["YTD RECEIPT"] = receipt;
				CountDispatchVolume4TW(customer, reportYearlyStartDate, reportYearlyEndDate, out dispatch, out dispatchWithinSLA, out tat, out AverageTATOverall, out OTS);
		        destDR["YTD DISPATCH"] = dispatch.ToString();
		        destDR["YTD DISPATCH WITHIN SLA"] = dispatchWithinSLA.ToString();
		        destDR["YTD TAT"] = tat.ToString();
		        destDR["YTD AVG TAT OVERALL"] = AverageTATOverall;
				destDR["YTD OTS%"] = OTS;
					
				receiptYearlySubTotal += receipt;
				dispatchYearlySubTotal += dispatch;
				dispatchWithinSLAYearlySubTotal += dispatchWithinSLA;
				tatYearlySubTotal += tat;
				
				frmOutput.ds.Tables["Activity YTD"].Rows.Add(destDR);
			}		
			
			SetSubTotal( receiptMonthlySubTotal,
		    			dispatchMonthlySubTotal,
		    			dispatchWithinSLAMonthlySubTotal,
		    			tatMonthlySubTotal,		    
						receiptYearlySubTotal,
		    			dispatchYearlySubTotal,
		    			dispatchWithinSLAYearlySubTotal,
		    			tatYearlySubTotal,
		    			AverageTATOverallMonthlySubTotal, 
		    			OTSMonthlySubTotal, 
		    			AverageTATOverallYearlySubTotal, 
		    			OTSYearlySubTotal );
		}
		
		private int CountReceiptVolume4TW(string customerName, DateTime reportStartDate, DateTime reportEndDate)
		{
			int ret = 0;
			foreach (DataRow dr in frmInput.ds.Tables["TW-IC-YTD"].Rows)
			{
			  	if (dr["Sales Orders - ShipTo Customer Name"].ToString() == customerName)
			  	{
			  		int flag = IsDateInDomain( dr["IC Shipments - Arrival Date"].ToString(), 
			  		                    reportStartDate,
			  		                    reportEndDate );
					if ( flag == -1 )
					{
			  			string msg = "TW-IC-YTD: Please check if 'IC Shipments - Arrival Date' is valid.\n" +
			  						"Customer: " + customerName + "\n" +
			  						"IC Shipments - Arrival Date: " + dr["IC Shipments - Arrival Date"].ToString();
						WriteLog(msg);
						
						continue;
			  		}
			  		else if ( flag == 1 )
			  		{
			  			ret++;
			  		}
			  	}
			}	
			
			return ret;
		}
		
		private void CountDispatchVolume4TW(string customerName, 
		                                    DateTime reportStartDate,
			  		                    	DateTime reportEndDate,
		                                    out int dispatch, 
		                                    out int dispatchWithinSLA,
		                                    out int tat,
		                                    out string AverageTATOverall,
		                                    out string OTS)
		{
			dispatch = 0;
			dispatchWithinSLA = 0;
		    AverageTATOverall = "";
		    OTS = "";
		    tat = 0;
		    
			foreach (DataRow dr in frmInput.ds.Tables["OTD-Onelog-TWOnly-YTD"].Rows)
			{
			  	if (dr["Country"].ToString().ToUpper() == "TW" &&
				    dr["Customer Name"].ToString() == customerName)
			  	{
			  		int flag = IsDateInDomain( dr["Ship_Date_Repl"].ToString(), 
			  		                    reportStartDate,
			  		                    reportEndDate );
					if ( flag == -1 )
					{
			  			string msg = "OTD-Onelog-TWOnly-YTD: Please check if 'Ship_Date_Repl' is valid.\n" +
			  						"Country: TW \n" + 
			  						"Customer: " + customerName + "\n" +
			  						"Ship_Date_Repl: " + dr["Ship_Date_Repl"].ToString();
						WriteLog(msg);
						
						continue;
			  		}
			  		else if ( flag == 1 )
			  		{
			  			dispatch++;
			  			
			  			if ( dr["Delivery Status"].ToString() == "On Time" &&
			  			     dr["Ship_Date_Repl"].ToString() != "" )
			  			{
			  				dispatchWithinSLA++;
			  			}	
			  			
			  			int currTAT = 0;
			  			if (dr["SLA"].ToString() == "D+" || dr["SLA"].ToString() == "H+")
			  			{
			  				currTAT = CalculateTAT( dr["Ship_Date_Repl"].ToString(), dr["Create_Date"].ToString() );
			  			}
			  			else if (dr["SLA"].ToString() == "R4S")
			  			{
			  				currTAT = CalculateTAT( dr["Ship_Date_Repl"].ToString(), dr["Rcv_Date_Def"].ToString() );
			  			}
			  			
			  			if (currTAT == -10000)
			  			{
			  					string msg = "OTD-Onelog-TWOnly-YTD: Error occurred when calculating TAT.\n" +
			  						"Country: TW \n" + 
			  						"Customer: " + customerName + "\n" +
			  						"Ship_Date_Repl: " + dr["Ship_Date_Repl"].ToString() + "\n" +			  						
			  						"Create_Date: " + dr["Create_Date"].ToString() + "\n" +
			  						"Rcv_Date_Def: " + dr["Rcv_Date_Def"].ToString();
								WriteLog(msg);
			  			}			
			  			else
			  			{
			  				tat += currTAT;
			  			}
			  		}
			  	}
			}	
			
			if ( dispatch == 0 )
			{
				AverageTATOverall = "NA";
				OTS = "NA";
			}
			else 
			{
				AverageTATOverall = Convert.ToInt32(Math.Round((double)tat/dispatch, 0)).ToString();
				double d = Math.Round((double)dispatchWithinSLA/dispatch, 3) * 100;
				OTS = d.ToString() + "%";
			}			
		}
		
//=======================================================================================================
				
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
		
		private StringCollection GetCustomerCollection4NZ()
		{
			StringCollection ret = new StringCollection();
			string customer;
			
			foreach (DataRow dr in frmInput.ds.Tables["NZ-YTD"].Rows)
			{
				customer = dr["Customer Name"].ToString();
				if ( ret.Contains(customer) == false && 
				     ( CustomerNameFilter(customer) == false ) )
				{
					ret.Add(customer);
				}
			}
			
			foreach (DataRow dr in frmInput.ds.Tables["NZ-Received-YTD"].Rows)
			{
				customer = dr["Cust Name"].ToString();
				if ( ret.Contains(customer) == false && 
				     ( CustomerNameFilter(customer) == false ) )
				{
					ret.Add(customer);
				}
			}
			
			return ret;
		}
		
		private StringCollection GetCustomerCollection4CitadelAndeSparesNew(string country)
		{
			StringCollection ret = new StringCollection();
			string customer;
			
			foreach (DataRow dr in frmInput.ds.Tables["CitadelReceived-YTD"].Rows)
			{
				customer = dr["Customer_Name"].ToString();
				if ( ( dr["Country"].ToString().ToUpper() == country.ToUpper() ) &&
				     ( ret.Contains(customer) == false ) && 
				     ( CustomerNameFilter(customer) == false ) )
				{
					ret.Add(customer);
				}
			}
			
			foreach (DataRow dr in frmInput.ds.Tables["CitadelShipped-YTD"].Rows)
			{
				customer = dr["Customer_Name"].ToString();
				if ( ( dr["Country"].ToString().ToUpper() == country.ToUpper() ) &&
				     ( ret.Contains(customer) == false ) && 
				     ( CustomerNameFilter(customer) == false ) )
				{
					ret.Add(customer);
				}
			}
			
			foreach (DataRow dr in frmInput.ds.Tables["ActivityYTD-eSpares New"].Rows)
			{
				customer = dr["Sales Orders SoldTo - Customer Name"].ToString();
				if ( ( country.ToUpper() == GetCountryName4eSparesNew(dr["Sales Orders - RSCIC/RSLC (Sales Organisation)"].ToString()).ToUpper() ) &&
				     ( ret.Contains(customer) == false ) && 
				     ( CustomerNameFilter(customer) == false ) )
				{
					ret.Add(customer);
				}
			}		
			
			return ret;
		}
		
		private StringCollection GetCustomerCollection4TW(string tableName1,		                                               
		                                               string customerColumnName1,
		                                               string country,
		                                               string tableName2,
		                                               string countryColumnName2,
		                                               string customerColumnName2
		                                              )
		{
			StringCollection ret = new StringCollection();
			string customer;
			
			foreach (DataRow dr in frmInput.ds.Tables[tableName1].Rows)
			{
				customer = dr[customerColumnName1].ToString();
				if ( ( ret.Contains(customer) == false ) && 
				     ( CustomerNameFilter(customer) == false ) )
				{
					ret.Add(customer);
				}
			}
			
			if (tableName2 != "" && customerColumnName2 != "")
			{
				foreach (DataRow dr in frmInput.ds.Tables[tableName2].Rows)
				{
					customer = dr[customerColumnName2].ToString();
					if (( dr[countryColumnName2].ToString().ToUpper() == country.ToUpper() ) &&
					    ( ret.Contains(customer) == false ) &&
				     	( CustomerNameFilter(customer) == false ))
					{
						ret.Add(customer);
					}
				}
			}
			
			return ret;
		}
		
		private StringCollection GetCustomerCollection4Korea()
		{
			StringCollection ret = new StringCollection();
			string customer;
			
			foreach (DataRow dr in frmInput.ds.Tables["KOREA-YTD-Base on RMA creation date"].Rows)
			{
				customer = dr["Customer Name"].ToString();
				if ( ret.Contains(customer) == false && 
				     ( CustomerNameFilter(customer) == false ) )
				{
					ret.Add(customer);
				}
			}
			
			foreach (DataRow dr in frmInput.ds.Tables["KOREA-YTD-Base on shipping date"].Rows)
			{
				customer = dr["Customer Name"].ToString();
				if ( ret.Contains(customer) == false && 
				     ( CustomerNameFilter(customer) == false ) )
				{
					ret.Add(customer);
				}
			}
			
			foreach (DataRow dr in frmInput.ds.Tables["CitadelReceived-YTD"].Rows)
			{
				customer = dr["Customer_Name"].ToString();
				if ( ( dr["Country"].ToString().ToUpper() == "Korea Republic Of".ToUpper() ) &&
				     ( ret.Contains(customer) == false ) && 
				     ( CustomerNameFilter(customer) == false ) )
				{
					ret.Add(customer);
				}
			}
			
			foreach (DataRow dr in frmInput.ds.Tables["CitadelShipped-YTD"].Rows)
			{
				customer = dr["Customer_Name"].ToString();
				if ( ( dr["Country"].ToString().ToUpper() == "Korea Republic Of".ToUpper() ) &&
				     ( ret.Contains(customer) == false ) && 
				     ( CustomerNameFilter(customer) == false ) )
				{
					ret.Add(customer);
				}
			}
			
			foreach (DataRow dr in frmInput.ds.Tables["ActivityYTD-eSpares New"].Rows)
			{
				customer = dr["Sales Orders SoldTo - Customer Name"].ToString();
				if ( ( dr["Sales Orders - RSCIC/RSLC (Sales Organisation)"].ToString().Contains("KR01") ) && //KOREA
				     ( ret.Contains(customer) == false ) && 
				     ( CustomerNameFilter(customer) == false ) )
				{
					ret.Add(customer);
				}
			}
			
			return ret;
		}
		
		private StringCollection GetCountryCollection4Citadel()
		{
			StringCollection ret = new StringCollection();
			string country;
			
			foreach (DataRow dr in frmInput.ds.Tables["CitadelReceived-YTD"].Rows)
			{
				country = dr["Country"].ToString().ToUpper();
				if ( country == "Korea Republic Of".ToUpper() || country == "TW" || 
				     country == "Taiwan".ToUpper() ||country == "NEW ZEALAND")
				{
					continue;
				}
				
				if ((! ret.Contains(country)) && country != "")
				{
					ret.Add(country);
				}
			}
			
			foreach (DataRow dr in frmInput.ds.Tables["CitadelShipped-YTD"].Rows)
			{
				country = dr["Country"].ToString().ToUpper();
				if ( country == "Korea Republic Of".ToUpper() || country == "TW" || 
				     country == "Taiwan".ToUpper() ||country == "NEW ZEALAND")
				{
					continue;
				}
				
				if (! ret.Contains(country))
				{
					ret.Add(country);
				}
			}
			
			return ret;
		}
		
		private void SetSubTotal( int receiptMonthlySubTotal,
		    						int dispatchMonthlySubTotal,
		    						int dispatchWithinSLAMonthlySubTotal,
		    						int tatMonthlySubTotal,		    
									int receiptYearlySubTotal,
		    						int dispatchYearlySubTotal,
		    						int dispatchWithinSLAYearlySubTotal,
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
				OTSMonthlySubTotal = "NA";
			}
			else 
			{
				AverageTATOverallMonthlySubTotal = Convert.ToInt32(Math.Round((double)tatMonthlySubTotal/dispatchMonthlySubTotal, 0)).ToString();
				double d = Math.Round((double)dispatchWithinSLAMonthlySubTotal/dispatchMonthlySubTotal, 3) * 100;
				OTSMonthlySubTotal = d.ToString() + "%";
			}	
			
			if ( dispatchYearlySubTotal == 0 )
			{
				AverageTATOverallYearlySubTotal = "NA";
				OTSYearlySubTotal = "NA";
			}
			else 
			{
				AverageTATOverallYearlySubTotal = Convert.ToInt32(Math.Round((double)tatYearlySubTotal/dispatchYearlySubTotal, 0)).ToString();
				double d = Math.Round((double)dispatchWithinSLAYearlySubTotal/dispatchYearlySubTotal, 3) * 100;
				OTSYearlySubTotal = d.ToString() + "%";
			}	
			
			destDR = frmOutput.ds.Tables["Activity YTD"].NewRow();
			destDR["CUSTOMER"] = "Country SubTotal";
			destDR["MONTHLY RECEIPT"] = receiptMonthlySubTotal;
			destDR["MONTHLY DISPATCH"] = dispatchMonthlySubTotal.ToString();
		    destDR["MONTHLY DISPATCH WITHIN SLA"] = dispatchWithinSLAMonthlySubTotal.ToString();
		    destDR["MONTHLY TAT"] = tatMonthlySubTotal.ToString();
		    destDR["MONTHLY AVG TAT OVERALL"] = AverageTATOverallMonthlySubTotal;
			destDR["MONTHLY OTS%"] = OTSMonthlySubTotal;	
			destDR["YTD RECEIPT"] = receiptYearlySubTotal;
			destDR["YTD DISPATCH"] = dispatchYearlySubTotal.ToString();
		    destDR["YTD DISPATCH WITHIN SLA"] = dispatchWithinSLAYearlySubTotal.ToString();
		    destDR["YTD TAT"] = tatYearlySubTotal.ToString();
		    destDR["YTD AVG TAT OVERALL"] = AverageTATOverallYearlySubTotal;
			destDR["YTD OTS%"] = OTSYearlySubTotal;
			
			frmOutput.ds.Tables["Activity YTD"].Rows.Add(destDR);
		}
		private void WriteLog(string msg)
		{
			string LogFileName = System.Environment.CurrentDirectory + "\\Log-Activity YTD.txt";
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
