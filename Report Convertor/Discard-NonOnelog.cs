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
	public class NonOnelog
	{
		public NonOnelog()
		{
			
		}
		
		
		public void SetValues()
		{
			DataSet srcDs, destDs;
			srcDs = frmInput.ds;
			destDs = frmOutput.ds;
			
			foreach (DataRow srcDr in srcDs.Tables["Input5NZ"].Rows)
			{
				DataRow dr = destDs.Tables["NonOneLog"].NewRow();
//				dr["SLA "] = "aa";				
//				dr["Delivery Status"] = "bb";
				dr["Country"] = srcDr["Country"].ToString();
				dr["Product"] = srcDr["Product"].ToString();
				dr["Part Number"] = "'" + srcDr["Part Number"].ToString();
				dr["Description"] = srcDr["Description"].ToString();
				dr["Customer Name"] = srcDr["Customer Name"].ToString();	
				dr["Customer Reference #"] = srcDr["Customer Reference #"].ToString();
				dr["Customer TAT"] = srcDr["Customer TAT"].ToString();
				dr["RMA#"] = "'" + srcDr["RMA#"].ToString();
				dr["Warranty Status_W/OW"] = srcDr["Warranty Status_W/OW"].ToString();
				dr["SPT"]  = srcDr["SPT"].ToString();
				dr["Qty"]  = srcDr["Qty"].ToString();
				dr["Serial No"]  = "'" + srcDr["Serial No"].ToString();
				dr["RSLC/Repairer "]  = srcDr["RSLC/Repairer "].ToString();
				dr["Customer Request Date"]  = srcDr["Customer Request Date"].ToString();
				dr["RMA Request Date*"]  = srcDr["RMA Request Date*"].ToString();
				dr["Faulty Board Received Date from Customer"]  = srcDr["Faulty Board Received Date from Customer"].ToString();
				dr["Ship to RSLC/Repairer Date"]  = srcDr["Ship to RSLC/Repairer Date"].ToString();
				dr["Received Date from RSLC/Repairer"]  = srcDr["Received Date from RSLC/Repairer"].ToString();
				dr["Dispatched Date to Customer"]  = srcDr["Dispatched Date to Customer"].ToString();
				dr["Repairer_RSLC SLA"]  = srcDr["Repairer_RSLC SLA"].ToString();
				dr["Comment from RSLC/Repairer"]  = srcDr["Comment from RSLC/Repairer"].ToString();
				
				try
				{
					if (srcDr["Target Due Date"].ToString() == "")
					{
						DateTime dt = Convert.ToDateTime(srcDr["Customer Request Date"].ToString()).AddDays(1);
						dr["Target Due Date"]  = dt.ToString();
					}
					else
					{
						dr["Target Due Date"]  = srcDr["Target Due Date"].ToString();
					}
				}
				catch
				{
					string msg = "NonOnelog: Error Occurred when processing 'Target Due Date' field "
						+ "in the record 'Part Number' = " + dr["Part Number"].ToString();

					MessageBox.Show(msg);
				}
				dr["Repair Days Overdue"]  = "";
				string test = srcDr["Days Overdue"].ToString();		
				dr["Repair Status"]  = srcDr["Repair Status"].ToString();	
				dr["Customer Status"]  = srcDr["Customer Status"].ToString();
				dr["Repairer Actual TAT"]  = srcDr["Repairer Actual TAT"].ToString();
				dr["Customer Actual TAT"]  = srcDr["Customer Actual TAT"].ToString();
				
				try
				{
					if (srcDr["Dispatched Date to Customer"].ToString() == "")
					{
						dr["On Time"] = 0;
					}
					else
					{
						DateTime dtTargetDueDate, dtDispatchedDateToCustomer;
						dtTargetDueDate = Convert.ToDateTime(dr["Target Due Date"].ToString()).AddDays(1);
						dtDispatchedDateToCustomer = Convert.ToDateTime(srcDr["Dispatched Date to Customer"].ToString()).AddDays(1);
					
						if (dtDispatchedDateToCustomer > dtTargetDueDate)
						{
							dr["On Time"] = 0;
						}
						else
						{
							dr["On Time"] = 1;
						}
					}
				}
				catch
				{
					string msg = "NonOnelog: Error Occurred when processing 'On Time' field "
						+ "in the record 'Part Number' = " + dr["Part Number"].ToString();

					MessageBox.Show(msg);
				}
				dr["RSCIC TAT"]	  = srcDr["RSCIC TAT"].ToString();
				dr["RMA Handling TAT"]	  = srcDr["RMA Handling TAT"].ToString();
				dr["WK"] = frmConfig.CurrentWeek;
				dr["Comments / Reason"]  = srcDr["Comments"].ToString();
				
				if ( dr["Customer TAT"].ToString() == "1")
				{
					dr["SLA"] = "H+";
				}
				else{
					dr["SLA"] = "R4S";
				}
				
				if (dr["On Time"].ToString() == "1")
				{
					dr["Delivery Status"] = "On Time";
				}
				else {
					dr["Delivery Status"] = "Past Due";
				}
								
				
				
				dr["SLA number of days"]  = srcDr["Customer TAT"].ToString();				//Newly added
		
				try
				{
					DateTime dtDispatchedDateToCustomer, dtFaultyBoardReceivedDatefromCustomer, dtCustomerRequestDate;
					if (dr["SLA"].ToString() == "R4S")
					{
						if (srcDr["Dispatched Date to Customer"].ToString() == "" || 
					    	srcDr["Faulty Board Received Date from Customer"].ToString() == "")
						{
							dr["TAT upon actual close"] = "N/A";
						}
						else
						{
							dtDispatchedDateToCustomer = Convert.ToDateTime(srcDr["Dispatched Date to Customer"].ToString()).Date;
							dtFaultyBoardReceivedDatefromCustomer = Convert.ToDateTime(dr["Faulty Board Received Date from Customer"].ToString()).Date;
			
							dr["TAT upon actual close"] = (dtDispatchedDateToCustomer.Subtract(dtFaultyBoardReceivedDatefromCustomer)).Days.ToString();
						}
					}
					else
					{
						if (srcDr["Dispatched Date to Customer"].ToString() == "" || 
					    	srcDr["Customer Request Date"].ToString() == "")
						{
							dr["TAT upon actual close"] = "N/A";
						}
						else
						{
							dtDispatchedDateToCustomer = Convert.ToDateTime(srcDr["Dispatched Date to Customer"].ToString()).Date;
							dtCustomerRequestDate = Convert.ToDateTime(srcDr["Customer Request Date"].ToString()).Date;
					
							dr["TAT upon actual close"] = (dtDispatchedDateToCustomer.Subtract(dtCustomerRequestDate)).Days.ToString();
						}
					}
				}
				catch
				{
					
				}

				destDs.Tables["NonOneLog"].Rows.Add(dr);
			}
		}
	}
}
