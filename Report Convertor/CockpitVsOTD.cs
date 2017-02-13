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

	
	public class CockpitVsOTD
	{
		public CockpitVsOTD()
		{
			
		}
		
		int OTDNumber = 0;
		int CockpitNumber = 0;
		int CockpitNotInOTDNumber = 0;
		int OTDNotInCockpitNUmber =0;
		
		private void ProcessOTD()
		{
			DataColumn dc1 = new DataColumn("SO_SOLINE");	
			frmInput.ds.Tables["OTD-Onelog"].Columns.Add(dc1);
			DataColumn dc2 = new DataColumn("SO_SOLINE");	
			frmInput.ds.Tables["OTD-NonOnelog"].Columns.Add(dc2);
			
			foreach (DataRow dr1 in frmInput.ds.Tables["OTD-Onelog"].Rows)
			{
				string temp = dr1["Unique Identification"].ToString();
				temp = temp.ToString().Replace("/", "-");
				temp = temp.ToString().Replace("\\", "-");
				dr1["SO_SOLINE"] = temp;
				OTDNumber++;                                                                              
			}
			foreach (DataRow dr2 in frmInput.ds.Tables["OTD-NonOnelog"].Rows)
			{
				string temp = dr2["RMA#"].ToString();
				if (dr2["SLA"].ToString() == "R4S")
				{
					temp = temp + "-2";
				}
				else
				{
					temp = temp + "-1";
				}

				dr2["SO_SOLINE"] = temp;
				OTDNumber++;                                                                              
			}
		}
		
		private void ProcessCockpit()
		{
			foreach (DataRow dr1 in frmInput.ds.Tables["AP-AE-Cockpit"].Rows)
			{
				CockpitNumber++;                                                                              
			}
			foreach (DataRow dr2 in frmInput.ds.Tables["AP-RFS-Cockpit"].Rows)
			{
				CockpitNumber++;                                                                              
			}
		}

		private void CalculateOTDNotInCockpit()
		{
	  
	        bool found = false;
			
			//OTD-Onelog not in Cockpit	        
			foreach (DataRow dr1 in frmInput.ds.Tables["OTD-Onelog"].Rows)
			{
				found = false;
				
				foreach (DataRow dr2 in frmInput.ds.Tables["AP-AE-Cockpit"].Rows)
				{
					if ( dr1["SO_SOLINE"].ToString() == dr2["SO_SOLINE"].ToString())
					{
						found = true;
						break;
					}
				}
				
				if (found == false)
				{
					foreach (DataRow dr3 in frmInput.ds.Tables["AP-RFS-Cockpit"].Rows)
					{
						if ( dr1["SO_SOLINE"].ToString() == dr3["SO_SOLINE"].ToString())
						{
							found = true;
							break;
						}
					}
				}
	    
				if (found == false)
				{
					OTDNotInCockpitNUmber++;
					DataRow dr = frmOutput.ds.Tables["OTD Not in Cockpit"].NewRow();
				
					dr["SO_SOLINE"] = dr1["SO_SOLINE"].ToString();
					dr["WK"] = dr1["WK"].ToString();
					dr["Unique Identification"] = dr1["Unique Identification"].ToString();
					dr["Target Compl Date"] = dr1["Target Compl Date"].ToString();
					dr["Country"] = dr1["Country"].ToString();
					dr["Customer Name"] = dr1["Customer Name"].ToString();
					dr["Comments from In-country "] = dr1["Comments from In-country "].ToString();
					dr["L-1"] = dr1["L-1"].ToString();
					dr["L-2"] = dr1["L-2"].ToString();					
					frmOutput.ds.Tables["OTD Not in Cockpit"].Rows.Add(dr);
				}
			}
			
			//OTD-NonOnelog not in Cockpit
			foreach (DataRow dr1 in frmInput.ds.Tables["OTD-NonOnelog"].Rows)
			{
				found = false;
				
				foreach (DataRow dr2 in frmInput.ds.Tables["AP-AE-Cockpit"].Rows)
				{
					if ( dr1["SO_SOLINE"].ToString() == dr2["SO_SOLINE"].ToString())
					{
						found = true;
						break;
					}
				}
				
				if (found == false)
				{
					foreach (DataRow dr3 in frmInput.ds.Tables["AP-RFS-Cockpit"].Rows)
					{
						if ( dr1["SO_SOLINE"].ToString() == dr3["SO_SOLINE"].ToString())
						{
							found = true;
							break;
						}
					}
				}
	    
				if (found == false)
				{
					OTDNotInCockpitNUmber++;
					DataRow dr = frmOutput.ds.Tables["OTD Not in Cockpit"].NewRow();
				
					dr["SO_SOLINE"] = dr1["SO_SOLINE"].ToString();
					dr["WK"] = dr1["WK"].ToString();
					dr["RMA#"] = dr1["RMA#"].ToString();
					dr["SLA"] = dr1["SLA"].ToString();
					dr["Target Due Date"] = dr1["Target Due Date"].ToString();
					dr["Country"] = dr1["Country"].ToString();
					dr["Customer Name"] = dr1["Customer Name"].ToString();
					dr["Comments / Reason"] = dr1["Comments / Reason"].ToString();
					dr["L-1"] = dr1["L-1"].ToString();
					dr["L-2"] = dr1["L-2"].ToString();
					frmOutput.ds.Tables["OTD Not in Cockpit"].Rows.Add(dr);
				}
			}
			
		}                     	  
		    
 
		private void CalculateCockpitNotInOTD()
		{
	  
	        bool found = false;
			
			//AP-AE-Cockpit not in OTD	        
			foreach (DataRow dr1 in frmInput.ds.Tables["AP-AE-Cockpit"].Rows)
			{
				found = false;
				
				foreach (DataRow dr2 in frmInput.ds.Tables["OTD-Onelog"].Rows)
				{
					if ( dr1["SO_SOLINE"].ToString() == dr2["SO_SOLINE"].ToString())
					{
						found = true;
						break;
					}
				}
				
				if (found == false)
				{
					foreach (DataRow dr3 in frmInput.ds.Tables["OTD-NonOnelog"].Rows)
					{
						if ( dr1["SO_SOLINE"].ToString() == dr3["SO_SOLINE"].ToString())
						{
							found = true;
							break;
						}
					}
				}
	    
				if (found == false)
				{
					CockpitNotInOTDNumber++;
					DataRow dr = frmOutput.ds.Tables["Cockpit Not in OTD"].NewRow();
				
					dr["SO_SOLINE"] = dr1["SO_SOLINE"].ToString();
					dr["Due Week"] = dr1["Due Week"].ToString();
					dr["CUSTOMER_DUE_DATE_DAY"] = dr1["CUSTOMER_DUE_DATE_DAY"].ToString();
					dr["SO_SHIPTO_COUNTRY"] = dr1["SO_SHIPTO_COUNTRY"].ToString();
					dr["CUSTOMER"] = dr1["CUSTOMER"].ToString();
					dr["NET_OTD_STATUS"] = dr1["NET_OTD_STATUS"].ToString();
					dr["OTDC_COMMENTS"] = dr1["OTDC_COMMENTS"].ToString();
					dr["OTDC_CORRECTION"] = dr1["OTDC_CORRECTION"].ToString();
					dr["OTDC_FAILURE_DESCRIPTION"] = dr1["OTDC_FAILURE_DESCRIPTION"].ToString();
					dr["OTDC_FAILURE_REASON"] = dr1["OTDC_FAILURE_REASON"].ToString();
					dr["LAST_FILE"] = dr1["LAST_FILE"].ToString();
					dr["LAST_UPDATED"] = dr1["LAST_UPDATED"].ToString();
					frmOutput.ds.Tables["Cockpit Not in OTD"].Rows.Add(dr);
				}
			}
			
			//Cockpit not in OTD-NonOnelog 
			foreach (DataRow dr1 in frmInput.ds.Tables["AP-RFS-Cockpit"].Rows)
			{
				found = false;
				
				foreach (DataRow dr2 in frmInput.ds.Tables["OTD-Onelog"].Rows)
				{
					if ( dr1["SO_SOLINE"].ToString() == dr2["SO_SOLINE"].ToString())
					{
						found = true;
						break;
					}
				}
				
				if (found == false)
				{
					foreach (DataRow dr3 in frmInput.ds.Tables["OTD-NonOnelog"].Rows)
					{
						if ( dr1["SO_SOLINE"].ToString() == dr3["SO_SOLINE"].ToString())
						{
							found = true;
							break;
						}
					}
				}
	    
				if (found == false)
				{
					CockpitNotInOTDNumber++;
					DataRow dr = frmOutput.ds.Tables["Cockpit Not in OTD"].NewRow();
				
					dr["SO_SOLINE"] = dr1["SO_SOLINE"].ToString();
					dr["Due Week"] = dr1["Due Week"].ToString();
					dr["CUSTOMER_DUE_DATE_DAY"] = dr1["CUSTOMER_DUE_DATE_DAY"].ToString();
					dr["SO_SHIPTO_COUNTRY"] = dr1["SO_SHIPTO_COUNTRY"].ToString();
					dr["CUSTOMER"] = dr1["CUSTOMER"].ToString();
					dr["NET_OTD_STATUS"] = dr1["NET_OTD_STATUS"].ToString();
					dr["OTDC_COMMENTS"] = dr1["OTDC_COMMENTS"].ToString();
					dr["OTDC_CORRECTION"] = dr1["OTDC_CORRECTION"].ToString();
					dr["OTDC_FAILURE_DESCRIPTION"] = dr1["OTDC_FAILURE_DESCRIPTION"].ToString();
					dr["OTDC_FAILURE_REASON"] = dr1["OTDC_FAILURE_REASON"].ToString();
					dr["LAST_FILE"] = dr1["LAST_FILE"].ToString();
					dr["LAST_UPDATED"] = dr1["LAST_UPDATED"].ToString();					
					frmOutput.ds.Tables["Cockpit Not in OTD"].Rows.Add(dr);
				}
			}
			
		}                     	  
		    		
		
		public void SetValues()
		{
			ProcessOTD();
			ProcessCockpit();
			CalculateOTDNotInCockpit();
			CalculateCockpitNotInOTD();
			
			DataRow dr = frmOutput.ds.Tables["CockpitVsOTDAnalysis"].NewRow();
				
			dr["Cockpit Line Number"] = CockpitNumber.ToString();
			dr["OTD Line Number"] = OTDNumber.ToString();
			dr["Cockpit Not in OTD"] = CockpitNotInOTDNumber.ToString();
			dr["OTD Not in Cockpit"] = OTDNotInCockpitNUmber.ToString();

			frmOutput.ds.Tables["CockpitVsOTDAnalysis"].Rows.Add(dr);
		}
			
	}
	
}
