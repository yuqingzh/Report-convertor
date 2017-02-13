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
	public class OpenOrderNonOnelog
	{
		public OpenOrderNonOnelog()
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
		
		public void SetValues()
		{
			DataSet srcDs, destDs;
			srcDs = frmInput.ds;
			destDs = frmOutput.ds;
			
			foreach (DataRow srcDr in srcDs.Tables["Input8OpenOrderNonOnelog"].Rows)
			{
				if ( CustomerNameFilter(srcDr["Customer Name"].ToString()) == true )
				{
					continue;
				}
				
				DataRow dr = destDs.Tables["OpenOrderNonOnelog"].NewRow();
								
				dr["Product"] = "'" + srcDr["Product"];
				dr["Part Number"] = "'" + srcDr["Part Number"];
				dr["Description"] = srcDr["Description"];
				dr["Country"] = srcDr["Country"];
				dr["Customer Name"] = srcDr["Customer Name"];	
				dr["Customer Reference #"] = srcDr["Customer Reference #"];
				dr["Customer TAT"] = srcDr["Customer TAT"];
				dr["RMA#"] = "'" + srcDr["RMA#"];
				dr["Warranty Status_W/OW"] = srcDr["Warranty Status_W/OW"];
				dr["SPT"]  = srcDr["SPT"];
				dr["Qty"]  = srcDr["Qty"];
				dr["Serial No"]  = "'" + srcDr["Serial No"];
				dr["RSLC/Repairer "]  = srcDr["RSLC/Repairer "];
				dr["Customer Request Date"]  = srcDr["Customer Request Date"];
				dr["RMA Request Date*"]  = srcDr["RMA Request Date*"];
				dr["Faulty Board Received Date from Customer"]  = srcDr["Faulty Board Received Date from Customer"];
				dr["Ship to RSLC/Repairer Date"]  = srcDr["Ship to RSLC/Repairer Date"];
				dr["Received Date from RSLC/Repairer"]  = srcDr["Received Date from RSLC/Repairer"];
				dr["Dispatched Date to Customer"]  = srcDr["Dispatched Date to Customer"];
				dr["Repairer_RSLC SLA"]  = srcDr["Repairer_RSLC SLA"];
				dr["Comment from RSLC/Repairer"]  = srcDr["Comment from RSLC/Repairer"];

				try
				{
					if (srcDr["Target Due Date"].ToString() == "")
					{
						DateTime dt = Convert.ToDateTime(srcDr["Faulty Board Received Date from Customer"].ToString());
						dt = dt.AddDays(Convert.ToInt32(srcDr["Customer TAT"].ToString()));
						dr["Target Due Date"]  = dt.ToString();
					}
					else
					{
						dr["Target Due Date"]  = srcDr["Target Due Date"];
					}
				}
				catch
				{
					string msg = "OpenOrderNonOnelog: Error Occurred when processing 'Target Due Date' field "
						+ "in the record 'Part Number' = " + dr["Part Number"].ToString();

					MessageBox.Show(msg);
				}
				
				dr["Repair Days Overdue"]  = "";
				dr["Repair Status"]  = srcDr["Repair Status"];	
				dr["Customer Status"]  = srcDr["Customer Status"];
				dr["Repairer Actual TAT"]  = srcDr["Repairer Actual TAT"];
				dr["Customer Actual TAT"]  = srcDr["Customer Actual TAT"];
				
				dr["RSCIC TAT"]  = srcDr["RSCIC TAT"];
				dr["RMA Handling TAT"]  = srcDr["RMA Handling TAT"];
				
				if ( dr["Customer TAT"].ToString() == "1")
				{
					dr["SLA"] = "AE";
				}
				else{
					dr["SLA"] = "R4S";
				}
				
				
				TimeSpan ts = DateTime.Now.Date.Subtract(Convert.ToDateTime(dr["Target Due Date"].ToString()));
				dr["Overdue"] = ts.Days.ToString();
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
					
				dr["Overdue Group"] = SetOverdueGroup(dr["Overdue"].ToString(), dr["SLA"].ToString());

				
				destDs.Tables["OpenOrderNonOnelog"].Rows.Add(dr);
			}
		}
	}
}
