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
	public class MonthlyOTD
	{
		public MonthlyOTD()
		{
			
		}
				
//Korea Telecom Codes -- FULL
//===================================
//(HMT001) HOLIM TECHNOLOGY
//(KTF001) KTF-KTTECH
//(ORI001) ORIENT TELECOM CO  LTD
//KT
//(SUN001) SUNGHWA TELECOMMUNICATIONS
//(001117) KT CORPORATION
//(000627) KOREA TELECOM
//	
//SK Group Codes -- FULL
//==================================
//(000712) HANARO TELECOM
//(001137) SAMHWA TELECOM INDUSTRY CO LTD
//(SMT001) SAMHWA TELECOM
//(001260) VINE TELECOM
//(SAN001) SANDUEL NETWORKS
//	
//Korea Telecom Codes -- Half
//LG Nortel
//(001256) LGNORTEL ??LG-ERICSSON??
//	
//SK Group Codes -- Half
//LG Nortel
//(001256) LGNORTEL ?? LG-ERICSSON??
//==================================

private bool IsKT(string customerName, int LGNortelOnTimeRecordNumber, int LGEricssonOnTimeRecordNumber,
                 int LGNortelPastDueRecordNumber, int LGEricssonPastDueRecordNumber, bool OnTime)
{
	if ( customerName.ToUpper().Contains("HOLIM TECHNOLOGY")  ||
		 customerName.ToUpper().Contains("KTF-KTTECH")  ||
		 customerName.ToUpper().Contains("ORIENT TELECOM")  ||		 
		 customerName.ToUpper().Contains("SUNGHWA TELECOMMUNICATIONS")  ||
		 customerName.ToUpper().Contains("KOREA TELECOM")  ||
		 ( customerName.ToUpper().Contains("KT") && ( customerName != "SKT") && (! customerName.ToUpper().Contains("BFKT")) )  ||
		 ( customerName.ToUpper().Contains("LG") && customerName.ToUpper().Contains("NORTEL") && (OnTime == true && LGNortelOnTimeRecordNumber % 2 == 1) ) ||
		 ( customerName.ToUpper().Contains("LG") && customerName.ToUpper().Contains("NORTEL") && (OnTime == false && LGNortelPastDueRecordNumber % 2 == 1) ) ||
		 ( customerName.ToUpper().Contains("LG") && ( customerName.ToUpper().Contains("ERICSSON") || customerName.ToUpper().Contains("ERRICSON") ) && (OnTime == true && LGEricssonOnTimeRecordNumber % 2 == 1) ) ||
		 ( customerName.ToUpper().Contains("LG") && ( customerName.ToUpper().Contains("ERICSSON") || customerName.ToUpper().Contains("ERRICSON") ) && (OnTime == false && LGEricssonPastDueRecordNumber % 2 == 1) ))
	{
		return true;
	}
	else
	{
		return false;
	}
}

private bool IsSK(string customerName, int LGNortelOnTimeRecordNumber, int LGEricssonOnTimeRecordNumber,
                 int LGNortelPastDueRecordNumber, int LGEricssonPastDueRecordNumber, bool OnTime)
{
	if ( ( customerName.ToUpper().Contains("HANARO TELECOM") ) ||
		 ( customerName.ToUpper().Contains("SAMHWA TELECOM INDUSTRY") ) ||
		 ( customerName.ToUpper().Contains("(SMT001) SAMHWA TELECOM") ) ||
		 ( customerName.ToUpper().Contains("VINE TELECOM") ) ||
		 ( customerName.ToUpper().Contains("SANDUEL NETWORKS") )||
		 ( customerName.ToUpper().Contains("SANDEUL NETWORKS") )||
		 ( customerName.ToUpper().Contains("SK BROADBAND") )||
		 ( customerName == "SKT" )  ||
		 ( customerName.ToUpper().Contains("LG") && customerName.ToUpper().Contains("NORTEL") && (OnTime == true && LGNortelOnTimeRecordNumber % 2 == 0) ) ||
		 ( customerName.ToUpper().Contains("LG") && customerName.ToUpper().Contains("NORTEL") && (OnTime == false && LGNortelPastDueRecordNumber % 2 == 0) ) ||
		 ( customerName.ToUpper().Contains("LG") && ( customerName.ToUpper().Contains("ERICSSON") || customerName.ToUpper().Contains("ERRICSON") ) && (OnTime == true && LGEricssonOnTimeRecordNumber % 2 == 0) ) ||
		 ( customerName.ToUpper().Contains("LG") && ( customerName.ToUpper().Contains("ERICSSON") || customerName.ToUpper().Contains("ERRICSON") ) && (OnTime == false && LGEricssonPastDueRecordNumber % 2 == 0) ))
	{
		return true;
	}
	else
	{
		return false;
	}
}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="country"></param>
		/// <param name="customerName"></param>
		/// <param name="due"></param>
		/// <param name="onTime"></param>
		/// <param name="OTS"></param>
		private void CalculateAE(string customerName, string country, 
		                         out string due, out string onTime, out string OTS)
		{
			DataSet srcDs = frmInput.ds;
			
			due = "Not Applicable";
			onTime = "Not Applicable";
			OTS = "Not Applicable";
			
			int iDue = 0, iOnTime = 0;
			
		    int LGNortelOnTimeRecordNumber = 0; // Odd assigned to Korea Telecom
										  // Even assigned to SK
			int LGEricssonOnTimeRecordNumber = 0; // Odd assigned to Korea Telecom
										  // Even assigned to SK
	        int LGNortelPastDueRecordNumber = 0; // Odd assigned to Korea Telecom
										  // Even assigned to SK
			int LGEricssonPastDueRecordNumber = 0; // Odd assigned to Korea Telecom
										  // Even assigned to SK
										  
	        bool bOnTime = false;

			{
			  	foreach (DataRow dr in srcDs.Tables["Input9MonthlyOTD"].Rows)
			  	{
			  		if ( dr["Net OTD Status (1=OK; 0=NOK)"].ToString() == "1" )
					{
						bOnTime = true;
					}	
			  		else
			  		{
			  			bOnTime = false;
			  		}
			  		
	    			if (( dr["Service Type"].ToString() == "D+" || dr["Service Type"].ToString() == "H+" ) &&
			  		    ( dr["Sales Orders SoldTo - Customer Name"].ToString().ToUpper().Contains("LG") && 
			  		     dr["Sales Orders SoldTo - Customer Name"].ToString().ToUpper().Contains("NORTEL") ))
		    		{
			  			if (bOnTime == true)
			  			{
		    				LGNortelOnTimeRecordNumber++;
			  			}
			  			else
			  			{
			  				LGNortelPastDueRecordNumber++;
			  			}
		    		}
		    		
		    		if (( dr["Service Type"].ToString() == "D+" || dr["Service Type"].ToString() == "H+" ) &&
		    		    ( (dr["Sales Orders SoldTo - Customer Name"].ToString().ToUpper().Contains("ERICSSON") || dr["Sales Orders SoldTo - Customer Name"].ToString().ToUpper().Contains("ERRICSON") ) &&
			  		       dr["Sales Orders SoldTo - Customer Name"].ToString().ToUpper().Contains("LG" ) ))
					{
			  			if (bOnTime == true)
			  			{
		    				LGEricssonOnTimeRecordNumber++;
			  			}
			  			else
			  			{
			  				LGEricssonPastDueRecordNumber++;
			  			}		    			
		    		}
			  		
			  		
					if (( dr["Service Type"].ToString() == "D+" || dr["Service Type"].ToString() == "H+" ) &&
		    		    ( ( dr["Sales Orders - ShipTo Country"].ToString() == country && dr["Sales Orders SoldTo - Customer Name"].ToString().Contains(customerName) ) ||
		    		      ( customerName == "CHT" && dr["Sales Orders SoldTo - Customer Name"].ToString().Contains("中華電信") ) ||
		    		      ( country == "THAILAND" && customerName.Contains("TRUE MOVE") && dr["Sales Orders SoldTo - Customer Name"].ToString().ToUpper().Contains("BFKT") ) ||
		 				  ( country.Contains("KOREA") && customerName == "KOREA TELECOM" && IsKT(dr["Sales Orders SoldTo - Customer Name"].ToString(), LGNortelOnTimeRecordNumber, LGEricssonOnTimeRecordNumber, LGNortelPastDueRecordNumber, LGEricssonPastDueRecordNumber, bOnTime) ) ||
		    		      ( country.Contains("KOREA") && customerName == "SK" && IsSK(dr["Sales Orders SoldTo - Customer Name"].ToString(), LGNortelOnTimeRecordNumber, LGEricssonOnTimeRecordNumber, LGNortelPastDueRecordNumber, LGEricssonPastDueRecordNumber, bOnTime) ) 
				        ))
					{		  			
						iDue++;
					
						if ( dr["Net OTD Status (1=OK; 0=NOK)"].ToString() == "1" )
						{
							iOnTime++;
						}			
					}	
			  	}
			}
			
			if (iDue != 0)
			{
				double d = Math.Round((double)iOnTime/iDue, 3) * 100;
				OTS = d.ToString() + "%";
			}
			
			due = iDue.ToString();
			onTime = iOnTime.ToString();
		}
		
		/// <summary>
		/// //
		/// </summary>
		/// <param name="country"></param>
		/// <param name="customerName"></param>
		/// <param name="due"></param>
		/// <param name="onTime"></param>
		/// <param name="OTS"></param>
		
		private void CalculateR4S(string customerName, string country,
		                         out string due, out string onTime, out string OTS)
		{
			DataSet srcDs = frmInput.ds;
			
			due = "Not Applicable";
			onTime = "Not Applicable";
			OTS = "Not Applicable";
			
			int iDue = 0, iOnTime = 0;
			
		    int LGNortelOnTimeRecordNumber = 0; // Odd assigned to Korea Telecom
										  // Even assigned to SK
			int LGEricssonOnTimeRecordNumber = 0; // Odd assigned to Korea Telecom
										  // Even assigned to SK
		    int LGNortelPastDueRecordNumber = 0; // Odd assigned to Korea Telecom
										  // Even assigned to SK
			int LGEricssonPastDueRecordNumber = 0; // Odd assigned to Korea Telecom
										  // Even assigned to SK
										  
	        bool bOnTime = false;
			
			{
			  	foreach (DataRow dr in srcDs.Tables["Input9MonthlyOTD"].Rows)
			  	{
			  		if ( dr["Net OTD Status (1=OK; 0=NOK)"].ToString() == "1" )
					{
						bOnTime = true;
					}	
			  		else
			  		{
			  			bOnTime = false;
			  		}
			  		
	    		    if (( dr["Service Type"].ToString() == "RFS" ) &&
		    		    ( dr["Sales Orders SoldTo - Customer Name"].ToString().ToUpper().Contains("LG") && 
			  		      dr["Sales Orders SoldTo - Customer Name"].ToString().ToUpper().Contains("NORTEL") ))
		    		{
			  			if (bOnTime == true)
			  			{
		    				LGNortelOnTimeRecordNumber++;
			  			}
			  			else
			  			{
			  				LGNortelPastDueRecordNumber++;
			  			}
		    		}
		    		
		    		if (( dr["Service Type"].ToString() == "RFS" ) &&
		    		    ( (dr["Sales Orders SoldTo - Customer Name"].ToString().ToUpper().Contains("ERICSSON") || dr["Sales Orders SoldTo - Customer Name"].ToString().ToUpper().Contains("ERRICSON") ) &&
			  		       dr["Sales Orders SoldTo - Customer Name"].ToString().ToUpper().Contains("LG" ) ))
		    		{
			  			if (bOnTime == true)
			  			{
		    				LGEricssonOnTimeRecordNumber++;
			  			}
			  			else
			  			{
			  				LGEricssonPastDueRecordNumber++;
			  			}	
			  		}
		    		
					if (( dr["Service Type"].ToString() == "RFS" ) &&
		    		    ( ( dr["Sales Orders - ShipTo Country"].ToString() == country && dr["Sales Orders SoldTo - Customer Name"].ToString().Contains(customerName) ) ||
		    		      ( customerName == "CHT" && dr["Sales Orders SoldTo - Customer Name"].ToString().Contains("中華電信") ) ||
		    		      ( country == "THAILAND" && customerName.Contains("TRUE MOVE") && dr["Sales Orders SoldTo - Customer Name"].ToString().ToUpper().Contains("BFKT") ) ||
		     		      ( country.Contains("KOREA") && customerName == "KOREA TELECOM" && IsKT(dr["Sales Orders SoldTo - Customer Name"].ToString(), LGNortelOnTimeRecordNumber, LGEricssonOnTimeRecordNumber, LGNortelPastDueRecordNumber, LGEricssonPastDueRecordNumber, bOnTime) ) ||
		    		      ( country.Contains("KOREA") && customerName == "SK" && IsSK(dr["Sales Orders SoldTo - Customer Name"].ToString(), LGNortelOnTimeRecordNumber, LGEricssonOnTimeRecordNumber, LGNortelPastDueRecordNumber, LGEricssonPastDueRecordNumber, bOnTime) ) 
				        ))
					{
						iDue++;
					
						if ( dr["Net OTD Status (1=OK; 0=NOK)"].ToString() == "1" )
						{
							iOnTime++;
						}				
						
					}				
			  	}
			}
			
			if (iOnTime != 0)
			{
				double d = Math.Round((double)iOnTime/iDue, 3) * 100;
				OTS = d.ToString() + "%";
			}
			
			due = iDue.ToString();
			onTime = iOnTime.ToString();
		}
		
		private void Merge(DataTable tempDT, DataTable destDT)
		{
						
			//Merge the temp data table - tempDT
			for (int i = 0; i < tempDT.Rows.Count; i++)
			{
				DataRow tempDR = tempDT.Rows[i];
				//检查此记录是否已处理过	
				bool processed = false;
				foreach (DataRow destDR in destDT.Rows)
				{
					if (tempDR["Country"].ToString() == destDR["Country"].ToString() &&
					    tempDR["Customer Name"].ToString() == destDR["Customer Name"].ToString())
					{
						processed = true;
					}
				}
				
				if (processed == true)
				{
					continue;
				}
				
				DataRow dr = destDT.NewRow();
				dr["Country"] = tempDR["Country"];
				dr["Customer Name"] = tempDR["Customer Name"];
				dr["Due"] = tempDR["Due"];
				dr["On Time"] = tempDR["On Time"];
				dr["OTS %"] = tempDR["OTS %"];
//				dr["Remarks"] = tempDR["Remarks"];
				
				for (int j = i+1; j < tempDT.Rows.Count; j++)
				{
					if (tempDT.Rows[j]["Country"].ToString() == dr["Country"].ToString() &&
					    tempDT.Rows[j]["Customer Name"].ToString() == dr["Customer Name"].ToString())
						{	
					    	int totVolume, onTimeVolume;
							totVolume =  Convert.ToInt32(dr["Due"].ToString()) +
								Convert.ToInt32(tempDT.Rows[j]["Due"].ToString());
							dr["Due"] = totVolume;
							
							onTimeVolume =  Convert.ToInt32(dr["On Time"].ToString()) +
								Convert.ToInt32(tempDT.Rows[j]["On Time"].ToString());
							
							dr["On Time"] = onTimeVolume;
							
							if (totVolume != 0)
							{
								double d = Math.Round((double)onTimeVolume/totVolume, 3) * 100;
								dr["OTS %"] = d.ToString() + "%";
							}
							else 
							{
								dr["OTS %"] = "Not Applicable";
							}
							
//							if (dr["Remarks"].ToString() == "")
//							{
//								dr["Remarks"] = tempDT.Rows[j]["Remarks"].ToString();
//						
//							}
//							else 
//							{
//								dr["Remarks"] = dr["Remarks"].ToString()  + " \\" + tempDT.Rows[j]["Remarks"].ToString() + "\n";
//							}
						}
					
				}
				
				destDT.Rows.Add(dr);
			}
		}
		
		public void SetValues()
		{
			DataSet srcDs, destDs;
			DataTable tempDT;
			srcDs = frmInput.ds;
			destDs = frmOutput.ds;
			
			tempDT = destDs.Tables["MonthlyOTDAEMetrics"].Clone();
			
			foreach (DataRow filterDr in srcDs.Tables["MonthlyOTDKeyCustomerList"].Rows)
			{
				string country, customerName, countryDisplayName, customerDisplayName, due, onTime, OTS;
				country = filterDr["Country"].ToString();
				customerName = filterDr["Customer Name"].ToString();
				countryDisplayName = filterDr["Country Display Name"].ToString();
			    customerDisplayName = filterDr["Customer Display Name"].ToString();		

				CalculateAE(customerName,
			                country,
			            	out due,
			            	out onTime,
			            	out OTS
			           		);
//				DataRow dr = destDs.Tables["MonthlyOTDAEMetrics"].NewRow();
			    DataRow dr = tempDT.NewRow();
				dr["Country"] = countryDisplayName;
				dr["Customer Name"] = customerDisplayName;
				dr["Due"] = due;
				dr["On Time"] = onTime;
				dr["OTS %"] = OTS;
//				destDs.Tables["MonthlyOTDAEMetrics"].Rows.Add(dr);
				tempDT.Rows.Add(dr);
			}
			
			Merge( tempDT, destDs.Tables["MonthlyOTDAEMetrics"]);
			
			tempDT.Clear();
			tempDT = destDs.Tables["MonthlyOTDR4SMetrics"].Clone();
			foreach (DataRow filterDr in srcDs.Tables["MonthlyOTDKeyCustomerList"].Rows)
			{
				string country, customerName, countryDisplayName, customerDisplayName, due, onTime, OTS;
				country = filterDr["Country"].ToString();
				customerName = filterDr["Customer Name"].ToString();
				countryDisplayName = filterDr["Country Display Name"].ToString();
			    customerDisplayName = filterDr["Customer Display Name"].ToString();		

				CalculateR4S(customerName,
			                country,
			            	out due,
			            	out onTime,
			            	out OTS
			           		);
//				DataRow dr = destDs.Tables["MonthlyOTDR4SMetrics"].NewRow();
			    DataRow dr = tempDT.NewRow();
				dr["Country"] = countryDisplayName;
				dr["Customer Name"] = customerDisplayName;
				dr["Due"] = due;
				dr["On Time"] = onTime;
				dr["OTS %"] = OTS;
//				destDs.Tables["MonthlyOTDR4SMetrics"].Rows.Add(dr);
				tempDT.Rows.Add(dr);
			}
			
			Merge( tempDT, destDs.Tables["MonthlyOTDR4SMetrics"]);
			
		}
	

		
		private string TWTranslate(string customerName)
		{
			foreach (DataRow dr in frmInput.ds.Tables["MonthlyOTDTWCustomerNameTranslation"].Rows)
			{
				if (customerName.Contains(dr["Customer Name"].ToString()))
				{
					return dr["Translated Customer Name"].ToString();
				}
			}
			
			return customerName;
		}

//Louis:TBD		
		public void TWTranslateAll()
		{
			DataSet srcDs, destDs;
			srcDs = frmInput.ds;
			destDs = frmOutput.ds;
			
			foreach (DataRow srcDr in srcDs.Tables["Input9MonthlyOTD"].Rows)
			{
				DataRow dr = destDs.Tables["MonthlyOTDTWTranslation"].NewRow();
				if (srcDr["Sales Orders - ShipTo Country"].ToString() == "TW")
				{
					dr["Sales Orders SoldTo - Customer Name"] = TWTranslate(srcDr["Sales Orders SoldTo - Customer Name"].ToString());
				}
				else{
					dr["Sales Orders SoldTo - Customer Name"] = srcDr["Sales Orders SoldTo - Customer Name"];	
				}
															
					dr["Service Type"] = srcDr["Service Type"];
					dr["Contractual Y/N ?"] = srcDr["Contractual Y/N ?"];
					dr["Customer Type"] = srcDr["Customer Type"];
					dr["DUE DATE (FINAL)"] = srcDr["DUE DATE (FINAL)"];
					dr["Sales Orders - ShipTo Country"] = srcDr["Sales Orders - ShipTo Country"];
					//dr["Sales Orders SoldTo - Customer Name"] = srcDr["Sales Orders SoldTo - Customer Name"];
					dr["Net OTD Status (1=OK; 0=NOK)"] = srcDr["Net OTD Status (1=OK; 0=NOK)"];
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
					dr["Sales Orders - Sold To"] = srcDr["Sales Orders - Sold To"];
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
					dr["Product Name"] = srcDr["Product Name"];
					dr["Business Line"] = srcDr["Business Line"];
					dr["Sales Orders - ShipTo Customer Name"] = srcDr["Sales Orders - ShipTo Customer Name"];
					dr["SO-SO LINE"] = srcDr["SO-SO LINE"];
					dr["DUE WEEK"] = srcDr["DUE WEEK"] ;
					dr["DUE MONTH"] = srcDr["DUE MONTH"];
					dr["Sales Orders - RMA"] = srcDr["Sales Orders - RMA"];
					dr["Cares-Customer Ticket"] = srcDr["Cares-Customer Ticket"];
					dr["Cares-Customer Reference"] = srcDr["Cares-Customer Reference"];
					dr["Work Orders  Line - Repair Code"] = srcDr["Work Orders  Line - Repair Code"];
					dr["REPAIR TAT"] = srcDr["REPAIR TAT"];
					dr["POD STATUS"] = srcDr["POD STATUS"];
					dr["Work Orders  Line - Sequence"] = srcDr["Work Orders  Line - Sequence"];
					dr["COUNTRY"] = srcDr["COUNTRY"];
					dr["Sales Orders Line  - Product hierarchy Description"] = srcDr["Sales Orders Line  - Product hierarchy Description"];
					dr["SLA SUSPENSION DAYS"] = srcDr["SLA SUSPENSION DAYS"];
					dr["Remarks"] = srcDr["Remarks"];
//				dr["Clarify_Date"] = srcDr["Clarify_Date"];	
//				try
//				{
//					if (srcDr["Clarify_Time"].ToString() != "")
//					{
//						dr["Clarify_Time"] = (Convert.ToDateTime(srcDr["Clarify_Time"].ToString())).ToLongTimeString().ToString();
//					}
//					else
//					{
//						dr["Clarify_Time"] = srcDr["Clarify_Time"];
//					}
//				}
//				catch
//				{
//					string msg = "Error Occurred when processing 'Clarify_Time' field - " + srcDr["Clarify_Time"].ToString() + ".";
//					MessageBox.Show(msg);
//				}
//			
				
				destDs.Tables["MonthlyOTDTWTranslation"].Rows.Add(dr);
			}
			
		}
	
	}
	
}
