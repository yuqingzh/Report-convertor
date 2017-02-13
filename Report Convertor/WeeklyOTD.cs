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

namespace Report_Convertor
{
	
	/// <summary>
	/// Description of Class1.
	/// </summary>
	///
	public class CommentsInConuntry
	{
		public string Comments;
		public int Counter;
		
		public CommentsInConuntry()
		{
			Comments = "";
			Counter = 0;
		}
		
		public CommentsInConuntry(string comments, int counter)
		{
			Comments = comments;
			Counter = counter;
		}
	}
	
	public class WeeklyOTD
	{
		public WeeklyOTD()
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


		private void CalculateAE(string customerName, 
		                         string country,
		                              out string SLA,
		                              out string due, 
		                         	  out string onTime, 
		                         	  out string OTS,
		                         	  out string Remark)
		{
		    SLA = "";		    
		    OTS = "NA";
		    Remark = "";
		    List<CommentsInConuntry> commentsList = new List<CommentsInConuntry>();
		    bool commentsFound = false;
		    
		    int idue = 0, ionTime = 0;
			
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
		    	foreach (DataRow dr in frmInput.ds.Tables["WeeklyOTD"].Rows)
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
						idue++;
					
						if ( dr["Net OTD Status (1=OK; 0=NOK)"].ToString() == "1" )
						{
							ionTime++;
						}
						
						SLA = dr["Service Type"].ToString();
						
						commentsFound = false;
						if (dr["OTDC_Comments"].ToString() != "")
						{
						  	foreach ( CommentsInConuntry comments in commentsList )
						  	{
								if (comments.Comments == dr["OTDC_Comments"].ToString() ) //Luis
								{
									comments.Counter++;
									commentsFound = true;
								}
						  	}
						  	if (commentsFound == false)
							{
								CommentsInConuntry comments = new CommentsInConuntry(dr["OTDC_Comments"].ToString(), 1);
								commentsList.Add(comments);
							}
						}
					}
				}
		    }
			
			if (idue != 0)
			{
				double d = Math.Round((double)ionTime/idue, 3) * 100;
				OTS = d.ToString() + "%";
			}
			
			due = idue.ToString();
			onTime = ionTime.ToString();
			
			foreach ( CommentsInConuntry comments in commentsList )
			{
				Remark += comments.Counter.ToString() + " \\" + comments.Comments + "\n";
			}
		}                     	  
		    
		
		private void CalculateR4S(string customerName,
		                          string country,
		                              out string SLA,
		                              out string due, 
		                         	  out string onTime, 
		                         	  out string OTS,
		                         	  out string Remark)
		{
		    SLA = "";		    
		    OTS = "NA";
		    Remark = "";
		    List<CommentsInConuntry> commentsList = new List<CommentsInConuntry>();
		    bool commentsFound = false;
		    
			int idue = 0, ionTime = 0;
			
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
		    	foreach (DataRow dr in frmInput.ds.Tables["WeeklyOTD"].Rows)
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
		    		    ( dr["Sales Orders SoldTo - Customer Name"].ToString().ToUpper().Contains("LG") && dr["Sales Orders SoldTo - Customer Name"].ToString().ToUpper().Contains("NORTEL") ))
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
						idue++;
					
						if ( dr["Net OTD Status (1=OK; 0=NOK)"].ToString() == "1" )
						{
							ionTime++;
						}
						
						SLA = dr["Service Type"].ToString();					
						
						commentsFound = false;
						if (dr["OTDC_Comments"].ToString() != "")
						{
						  	foreach ( CommentsInConuntry comments in commentsList )
						  	{
								if (comments.Comments == dr["OTDC_Comments"].ToString() ) //Luis
								{
									comments.Counter++;
									commentsFound = true;
								}
						  	}
						  	if (commentsFound == false)
							{
								CommentsInConuntry comments = new CommentsInConuntry(dr["OTDC_Comments"].ToString(), 1);
								commentsList.Add(comments);
							}
						}
					}
				}
		    }
			
			if (idue != 0)
			{
				double d = Math.Round((double)ionTime/idue, 3) * 100;
				OTS = d.ToString() + "%";
			}
			
			due = idue.ToString();
			onTime = ionTime.ToString();
			
			foreach ( CommentsInConuntry comments in commentsList )
			{
				Remark += comments.Counter.ToString() + " \\" + comments.Comments + "\n";
			}
		}   
		
		
		public void SetValues()
		{
			DataSet srcDs, destDs;
			DataTable tempDT;
			srcDs = frmInput.ds;
			destDs = frmOutput.ds;
			
			tempDT = destDs.Tables["WeeklyOTDAnalysis"].Clone();
			
			foreach (DataRow filterDr in srcDs.Tables["WeeklyOTDKeyCustomerList"].Rows)
			{
				bool recordFound = false;
				string country, customerName, countryDisplayName, customerDisplayName, SLA, due, onTime, OTS, Remark;
				country = filterDr["Country"].ToString();
				customerName = filterDr["Customer Name"].ToString();
				countryDisplayName = filterDr["Country Display Name"].ToString();
			    customerDisplayName = filterDr["Customer Display Name"].ToString();			    
				CalculateAE(customerName,
			                country,
			            	out SLA,
			            	out due,
			            	out onTime,
			            	out OTS,
			            	out Remark
			           		);
				if (due != "0")
				{
					recordFound = true;
//					DataRow dr = destDs.Tables["WeeklyOTDAnalysis"].NewRow();	
					DataRow dr = tempDT.NewRow();
					dr["SLA"] = SLA;
					dr["Country"] = countryDisplayName;
					dr["Customer Name"] = customerDisplayName;
					dr["Total RMA Volume"] = due;
					dr["On Time RMA Volume"] = onTime;
					dr["OTD %"] = OTS;
					dr["Remarks"] = Remark;

//					destDs.Tables["WeeklyOTDAnalysis"].Rows.Add(dr);
					tempDT.Rows.Add(dr);
				}
				
				CalculateR4S(customerName,
			                country,
			            	out SLA,
			            	out due,
			            	out onTime,
			            	out OTS,
			            	out Remark
			           		);
				if (due != "0" || recordFound == false)
				{
//					DataRow dr = destDs.Tables["WeeklyOTDAnalysis"].NewRow();
					DataRow dr = tempDT.NewRow();
					dr["SLA"] = SLA;
					dr["Country"] = countryDisplayName;
					dr["Customer Name"] = customerDisplayName;
					dr["Total RMA Volume"] = due;
					dr["On Time RMA Volume"] = onTime;
					dr["OTD %"] = OTS;
					dr["Remarks"] = Remark;

//					destDs.Tables["WeeklyOTDAnalysis"].Rows.Add(dr);
					tempDT.Rows.Add(dr);
				}				
	
			}	
			
			//Merge the temp data table - tempDT
			for (int i = 0; i < tempDT.Rows.Count; i++)
			{
				DataRow tempDR = tempDT.Rows[i];
				//检查此记录是否已处理过	
				bool processed = false;

				foreach (DataRow destDR in destDs.Tables["WeeklyOTDAnalysis"].Rows)
				{
					if (tempDR["SLA"].ToString() == destDR["SLA"].ToString() &&
					    tempDR["Country"].ToString() == destDR["Country"].ToString() &&
					    tempDR["Customer Name"].ToString() == destDR["Customer Name"].ToString())
					{
						processed = true;
					}
				}
				
				if (processed == true)
				{
					continue;
				}
				
				DataRow dr = destDs.Tables["WeeklyOTDAnalysis"].NewRow();
				dr["SLA"] = tempDR["SLA"];
				dr["Country"] = tempDR["Country"];
				dr["Customer Name"] = tempDR["Customer Name"];
				dr["Total RMA Volume"] = tempDR["Total RMA Volume"];
				dr["On Time RMA Volume"] = tempDR["On Time RMA Volume"];
				dr["OTD %"] = tempDR["OTD %"];
				dr["Remarks"] = tempDR["Remarks"];
				
				for (int j = i+1; j < tempDT.Rows.Count; j++)
				{
					if ((tempDT.Rows[j]["SLA"].ToString() == dr["SLA"].ToString()) && //|| dr["SLA"].ToString() == "" || tempDT.Rows[j]["SLA"].ToString() == "" ) &&//需要合并
					    tempDT.Rows[j]["Country"].ToString() == dr["Country"].ToString() &&
					    tempDT.Rows[j]["Customer Name"].ToString() == dr["Customer Name"].ToString())
						{			
//							if (tempDT.Rows[j]["SLA"].ToString() != "")
//							{
//								dr["SLA"] = tempDT.Rows[j]["SLA"];
//							}

							int totVolume, onTimeVolume;
							totVolume =  Convert.ToInt32(dr["Total RMA Volume"].ToString()) +
								Convert.ToInt32(tempDT.Rows[j]["Total RMA Volume"].ToString());
							dr["Total RMA Volume"] = totVolume;
							
							onTimeVolume =  Convert.ToInt32(dr["On Time RMA Volume"].ToString()) +
								Convert.ToInt32(tempDT.Rows[j]["On Time RMA Volume"].ToString());
							
							dr["On Time RMA Volume"] = onTimeVolume;
														
							if (totVolume != 0)
							{
								double d = Math.Round((double)onTimeVolume/totVolume, 3) * 100;
								dr["OTD %"] = d.ToString() + "%";
							}
							else 
							{
								dr["OTD %"] = "NA";
							}
							
							if (dr["Remarks"].ToString() == "")
							{
								dr["Remarks"] = tempDT.Rows[j]["Remarks"].ToString();
						
							}
							else 
							{
								dr["Remarks"] = dr["Remarks"].ToString()  + " \\" + tempDT.Rows[j]["Remarks"].ToString() + "\n";
							}
						}
					
				}
				
				destDs.Tables["WeeklyOTDAnalysis"].Rows.Add(dr);
				
			}
			
		}
	
	}
	
}
