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
	public class ClosedRMA
	{
		public ClosedRMA()
		{
			
		}

		private void LookupVC(string id, string repairer, string country, 
		                        out string  VC_NFF, out string  VC_RFR60, out string remark)
		{			
			VC_NFF = "";
			VC_RFR60 = "";
			remark = "";
			
			string VCID = "";
			
			if (repairer == "RLC-Ormes")
			{
				if (country.ToUpper() == "THAILAND" ||
				    country.ToUpper() == "INDONESIA")
				{
					foreach (DataRow dr in frmInput.ds.Tables["VC Catalogue - RESO EMEA - VC_APAC FCA"].Rows)
					{
						VCID = RemoveBlank(dr["RES_IDENTIFIER"].ToString()); 
						if ( id == VCID )
						{
							VC_NFF = dr["VC_RFR_NFF"].ToString();
							VC_RFR60 = dr["VC_RFR"].ToString();
							return;
						}
						if ( "CISCO" + id == VCID )
						{
							VC_NFF = dr["VC_RFR_NFF"].ToString();
							VC_RFR60 = dr["VC_RFR"].ToString();
							remark = "Looked up by linking 'CISCO' and the original ID. New ID: CISCO" + id;
							return;
						}
					}
					
					foreach (DataRow dr in frmInput.ds.Tables["VC Catalogue - RESO EMEA - VC_APAC FCA"].Rows)
					{
						VCID = RemoveBlank(dr["RES_IDENTIFIER"].ToString());				
						if (CompareString(id, VCID))
						{
							VC_NFF = dr["VC_RFR_NFF"].ToString();
							VC_RFR60 = dr["VC_RFR"].ToString();
							if (id.Length < VCID.Length)
							{
								remark = "Looked up by cutting the length of the 'RES_IDENTIFIER' (" + VCID + ") to the length of the original ID (" + id + ").";
							}
							else if (id.Length > VCID.Length)
							{
								remark = "Looked up by cutting the length of the original ID (" + id + ") to the length of the 'RES_IDENTIFIER' (" + VCID + ").";
							}
							return;
						}			
					}	
					
					//模糊查找
					foreach (DataRow dr in frmInput.ds.Tables["VC Catalogue - RESO EMEA - VC_APAC FCA"].Rows)
					{
						VCID = RemoveBlank(dr["RES_IDENTIFIER"].ToString()); 
						if (VCID.Contains(id))
						{
							VC_NFF = dr["VC_RFR_NFF"].ToString();
							VC_RFR60 = dr["VC_RFR"].ToString();
							remark = "Looked up this record as the 'RES_IDENTIFIER' (" + dr["RES_IDENTIFIER"].ToString()
										+ ") contains the original ID (" + id + ").";
							return;
						}
					}
				}
				else
				{
					foreach (DataRow dr in frmInput.ds.Tables["VC Catalogue - RESO EMEA - VC_APAC"].Rows)
					{
						VCID = RemoveBlank(dr["RES_IDENTIFIER"].ToString()); 
						if (id == VCID)
						{
							VC_NFF = dr["VC_RFR_NFF"].ToString();
							VC_RFR60 = dr["VC_RFR"].ToString();
							return;
						}
						if ( "CISCO" + id == VCID )
						{
							VC_NFF = dr["VC_RFR_NFF"].ToString();
							VC_RFR60 = dr["VC_RFR"].ToString();
							remark = "Looked up by linking 'CISCO' and the original ID. New ID: CISCO" + id;
							return;
						}
					}
					
					foreach (DataRow dr in frmInput.ds.Tables["VC Catalogue - RESO EMEA - VC_APAC"].Rows)
					{
						VCID = RemoveBlank(dr["RES_IDENTIFIER"].ToString());				
						if (CompareString(id, VCID))
						{
							VC_NFF = dr["VC_RFR_NFF"].ToString();
							VC_RFR60 = dr["VC_RFR"].ToString();
							if (id.Length < VCID.Length)
							{
								remark = "Looked up by cutting the length of the 'RES_IDENTIFIER' (" + VCID + ") to the length of the original ID (" + id + ").";
							}
							else if (id.Length > VCID.Length)
							{
								remark = "Looked up by cutting the length of the original ID (" + id + ") to the length of the 'RES_IDENTIFIER' (" + VCID + ").";
							}
							return;
						}			
					}	
					
					//模糊查找
					foreach (DataRow dr in frmInput.ds.Tables["VC Catalogue - RESO EMEA - VC_APAC"].Rows)
					{
						VCID = RemoveBlank(dr["RES_IDENTIFIER"].ToString()); 
						if (VCID.Contains(id))
						{
							VC_NFF = dr["VC_RFR_NFF"].ToString();
							VC_RFR60 = dr["VC_RFR"].ToString();
							remark = "Looked up this record as the 'RES_IDENTIFIER' (" + 
										dr["RES_IDENTIFIER"].ToString() + ") contains the original ID (" + id + ").";
							return;
						}
					}
				}
				
				VC_RFR60 = "309";
				VC_NFF = "309";
				remark = "Manually put 309 for the repair cost";
				return; 
			}
			else if (repairer == "RLC-India")
			{
				foreach (DataRow dr in frmInput.ds.Tables["VC Catalogue - RESO APAC INDIA"].Rows)
				{
					VCID = RemoveBlank(dr["RES_ID"].ToString()); 
					if (id == VCID)
					{
						VC_NFF = dr["VC NFF"].ToString();
						VC_RFR60 = dr["VC_RFR_60"].ToString();
						return;
					}
					if ( "CISCO" + id == VCID )
					{
						VC_NFF = dr["VC NFF"].ToString();
						VC_RFR60 = dr["VC_RFR_60"].ToString();
						remark = "Looked up by linking 'CISCO' and the original ID. New ID: CISCO" + id;
						return;
					}
				}
				
				foreach (DataRow dr in frmInput.ds.Tables["VC Catalogue - RESO APAC INDIA"].Rows)
				{
					VCID = RemoveBlank(dr["RES_ID"].ToString());				
					if (CompareString(id, VCID))
					{
						VC_NFF = dr["VC NFF"].ToString();
						VC_RFR60 = dr["VC_RFR_60"].ToString();
						if (id.Length < VCID.Length)
						{
							remark = "Looked up by cutting the length of the 'RES_ID' (" + VCID + ") to the length of the original ID (" + id + ").";
						}
						else if (id.Length > VCID.Length)
						{
							remark = "Looked up by cutting the length of the original ID (" + id + ") to the length of the 'RES_ID' (" + VCID + ").";
						}
						return;
					}			
				}
				
				//模糊查找
				foreach (DataRow dr in frmInput.ds.Tables["VC Catalogue - RESO APAC INDIA"].Rows)
				{
					VCID = RemoveBlank(dr["RES_ID"].ToString()); 
					if (VCID.Contains(id))
					{
						VC_NFF = dr["VC NFF"].ToString();
						VC_RFR60 = dr["VC_RFR_60"].ToString();
						remark = "Looked up this record as the 'RES_IDENTIFIER' (" + dr["RES_ID"].ToString() + ") contains the original ID (" + id + ").";
						return;
					}
				}
				
				VC_RFR60 = "309";
				VC_NFF = "309";
				remark = "Manually put 309 for the repair cost";
				return; 
			}
			else if (repairer == "RLC-LVL")
			{
				bool found = false;
				foreach (DataRow dr in frmInput.ds.Tables["VC Catalogue - RESO AMERICAS NAR"].Rows)
				{
					VCID = RemoveBlank(dr["RES_IDENTIFIER"].ToString()); 
					if (id == VCID)
					{
						found = true;
						VC_NFF = dr["VC_NFF"].ToString();
						VC_RFR60 = dr["VC_RFR_60"].ToString();
						return;
					}		
					if ( "CISCO" + id == VCID )
					{
						found = true;
						VC_NFF = dr["VC_NFF"].ToString();
						VC_RFR60 = dr["VC_RFR_60"].ToString();
						remark = "Looked up by linking 'CISCO' and the original ID. New ID: CISCO" + id;
						return;
					}
				}
				
				//re-search
				string id1 = "";
				string id2 = "";	
				string id3 = "";
				if (found == false)
				{
					foreach (DataRow dr in frmInput.ds.Tables["Stinger product code mapping list"].Rows)
					{
						if (id == RemoveBlank(dr["Rcv_Pno"].ToString()) ||
						    id == RemoveBlank(dr["Rcv_Model"].ToString()) ||
						    id == RemoveBlank(dr["trim pn"].ToString()))
						{
							id1 = RemoveBlank(dr["Rcv_Pno"].ToString());
							id2 = RemoveBlank(dr["Rcv_Model"].ToString());
							id3 = RemoveBlank(dr["trim pn"].ToString());
							break;								
						}
					}
					if (id1 != "")
					{
						foreach (DataRow dr in frmInput.ds.Tables["VC Catalogue - RESO AMERICAS NAR"].Rows)
						{
							VCID = RemoveBlank(dr["RES_IDENTIFIER"].ToString()); 
							if (id1 == VCID)
							{
								VC_NFF = dr["VC_NFF"].ToString();
								VC_RFR60 = dr["VC_RFR_60"].ToString();
								remark = "Looked up by 'Rcv_Pno' in 'Stinger product code mapping list'";
								return;
							}	
							if ( "CISCO" + id1 == VCID )
							{
								found = true;
								VC_NFF = dr["VC_NFF"].ToString();
								VC_RFR60 = dr["VC_RFR_60"].ToString();
								remark = "Looked up by linking 'CISCO' and 'Rcv_Pno' in 'Stinger product code mapping list'. New ID: CISCO" + id1;
								return;
							}
						}
					}
					if (id2 != "")
					{
						foreach (DataRow dr in frmInput.ds.Tables["VC Catalogue - RESO AMERICAS NAR"].Rows)
						{
							VCID = RemoveBlank(dr["RES_IDENTIFIER"].ToString()); 
							if (id2 == VCID)
							{
								VC_NFF = dr["VC_NFF"].ToString();
								VC_RFR60 = dr["VC_RFR_60"].ToString();
								remark = "Looked up by 'Rcv_Model' in 'Stinger product code mapping list'";
								return;
							}	
							if ( "CISCO" + id2 == VCID )
							{
								found = true;
								VC_NFF = dr["VC_NFF"].ToString();
								VC_RFR60 = dr["VC_RFR_60"].ToString();
								remark = "Looked up by linking 'CISCO' and 'Rcv_Model' in 'Stinger product code mapping list'. New ID: CISCO" + id2;
								return;
							}
						}
					}	
					if (id3 != "")
					{
						foreach (DataRow dr in frmInput.ds.Tables["VC Catalogue - RESO AMERICAS NAR"].Rows)
						{
							VCID = RemoveBlank(dr["RES_IDENTIFIER"].ToString()); 
							if (id3 == VCID)
							{
								VC_NFF = dr["VC_NFF"].ToString();
								VC_RFR60 = dr["VC_RFR_60"].ToString();
								remark = "Looked up by 'trim pn' in 'Stinger product code mapping list'";								
								return;
							}	
							if ( "CISCO" + id3 == VCID )
							{
								found = true;
								VC_NFF = dr["VC_NFF"].ToString();
								VC_RFR60 = dr["VC_RFR_60"].ToString();
								remark = "Looked up by linking 'CISCO' and 'trim pn' in 'Stinger product code mapping list'. New ID: CISCO" + id3;
								return;
							}
						}
					}
				}	
				
				foreach (DataRow dr in frmInput.ds.Tables["VC Catalogue - RESO AMERICAS NAR"].Rows)
				{
					VCID = RemoveBlank(dr["RES_IDENTIFIER"].ToString());				
					if (CompareString(id, VCID))
					{
						VC_NFF = dr["VC_NFF"].ToString();
						VC_RFR60 = dr["VC_RFR_60"].ToString();
						if (id.Length < VCID.Length)
						{
							remark = "Looked up by cutting the length of the 'RES_IDENTIFIER' (" + VCID + ") to the length of the original ID (" + id + ").";
						}
						else if (id.Length > VCID.Length)
						{
							remark = "Looked up by cutting the length of the original ID (" + id + ") to the length of the 'RES_IDENTIFIER' (" + VCID + ").";
						}
						return;
					}		
				}
				
				//模糊查找
				foreach (DataRow dr in frmInput.ds.Tables["VC Catalogue - RESO AMERICAS NAR"].Rows)
				{
					VCID = RemoveBlank(dr["RES_IDENTIFIER"].ToString());
					if (VCID.Contains(id))
					{
						VC_NFF = dr["VC_NFF"].ToString();
						VC_RFR60 = dr["VC_RFR_60"].ToString();
						remark = "Looked up this record as the 'RES_IDENTIFIER' (" + dr["RES_IDENTIFIER"].ToString() + ") contains the original ID (" + id + ").";
						return;
					}
				}
				
				VC_RFR60 = "309";
				VC_NFF = "309";
				remark = "Manually put 309 for the repair cost";
				return; 
			}
			else if (repairer == "RLC-QD")
			{
				foreach (DataRow dr in frmInput.ds.Tables["VC Catalogue - RESO APAC CHINA QD"].Rows)
				{
					VCID = RemoveBlank(dr["RES_IDENTIFIER"].ToString()); 
					if (id == VCID)
					{
						VC_NFF = dr["VC_RFR_NFF"].ToString();
						VC_RFR60 = dr["VC_RFR"].ToString();
						return;
					}
					if ( "CISCO" + id == VCID )
					{
						VC_NFF = dr["VC_RFR_NFF"].ToString();
						VC_RFR60 = dr["VC_RFR"].ToString();
						remark = "Looked up by linking 'CISCO' and the original ID. New ID: CISCO" + id;
						return;
					}
				}
				
				foreach (DataRow dr in frmInput.ds.Tables["VC Catalogue - RESO APAC CHINA QD"].Rows)
				{
					VCID = RemoveBlank(dr["RES_IDENTIFIER"].ToString());				
					if (CompareString(id, VCID))
					{
						VC_NFF = dr["VC_RFR_NFF"].ToString();
						VC_RFR60 = dr["VC_RFR"].ToString();
						if (id.Length < VCID.Length)
						{
							remark = "Looked up by cutting the length of the 'RES_IDENTIFIER' (" + VCID + ") to the length of the original ID (" + id + ").";
						}
						else if (id.Length > VCID.Length)
						{
							remark = "Looked up by cutting the length of the original ID (" + id + ") to the length of the 'RES_IDENTIFIER' (" + VCID + ").";
						}
						return;
					}			
				}
				
				//模糊查找
				foreach (DataRow dr in frmInput.ds.Tables["VC Catalogue - RESO APAC CHINA QD"].Rows)
				{
					VCID = RemoveBlank(dr["RES_IDENTIFIER"].ToString()); 
					if (VCID.Contains(id))
					{
						VC_NFF = dr["VC_RFR_NFF"].ToString();
						VC_RFR60 = dr["VC_RFR"].ToString();
						remark = "Looked up this record as the 'RES_IDENTIFIER' (" + dr["RES_IDENTIFIER"].ToString() + ") contains the original ID (" + id + ").";
						return;
					}
				}
				
				VC_RFR60 = "309";
				VC_NFF = "309";
				remark = "Manually put 309 for the repair cost";
				return; 
			}
			else if (repairer == "RLC-ASB")
			{
				foreach (DataRow dr in frmInput.ds.Tables["VC Catalogue - RESO China (SHA)"].Rows)
				{
					VCID = RemoveBlank(dr["RES_IDENTIFIER"].ToString()); 
					if (id == VCID)
					{
						VC_NFF = dr["VC_NFF"].ToString();
						VC_RFR60 = dr["VC_RFR60"].ToString();
						return;
					}
					if ( "CISCO" + id == VCID )
					{
						VC_NFF = dr["VC_NFF"].ToString();
						VC_RFR60 = dr["VC_RFR60"].ToString();
						remark = "Looked up by linking 'CISCO' and the original ID. New ID: CISCO" + id;
						return;
					}
				}
				
				//re-search
				foreach (DataRow dr in frmInput.ds.Tables["VC Catalogue - RESO China (SHA)"].Rows)
				{
					string partNo = RemoveBlank(dr["PART_NUMBER"].ToString());
					if (id.Length <= partNo.Length)
					{
						if (id == partNo.Substring(0,id.Length))
						{
							VC_NFF = dr["VC_NFF"].ToString();
							VC_RFR60 = dr["VC_RFR60"].ToString();
							remark = "Looked up by 'PART_NUMBER' rather than 'RES_IDENTIFIER'";
							return;
						}
					}
				}
				
				foreach (DataRow dr in frmInput.ds.Tables["VC Catalogue - RESO China (SHA)"].Rows)
				{
					VCID = RemoveBlank(dr["RES_IDENTIFIER"].ToString());				
					if (CompareString(id, VCID))
					{
						VC_NFF = dr["VC_NFF"].ToString();
						VC_RFR60 = dr["VC_RFR60"].ToString();
						if (id.Length < VCID.Length)
						{
							remark = "Looked up by cutting the length of the 'RES_IDENTIFIER' (" + VCID + ") to the length of the original ID (" + id + ").";
						}
						else if (id.Length > VCID.Length)
						{
							remark = "Looked up by cutting the length of the original ID (" + id + ") to the length of the 'RES_IDENTIFIER' (" + VCID + ").";
						}
						return;
					}			
				}
				
				foreach (DataRow dr in frmInput.ds.Tables["VC Catalogue - RESO China (SHA)"].Rows)
				{
					VCID = RemoveBlank(dr["PART_NUMBER"].ToString());				
					if (CompareString(id, VCID))
					{
						VC_NFF = dr["VC_NFF"].ToString();
						VC_RFR60 = dr["VC_RFR60"].ToString();
						if (id.Length < VCID.Length)
						{
							remark = "Looked up by cutting the length of the 'PART_NUMBER' (" + VCID + ") to the length of the original ID (" + id + ").";
						}
						else if (id.Length > VCID.Length)
						{
							remark = "Looked up by cutting the length of the original ID (" + id + ") to the length of the 'PART_NUMBER' (" + VCID + ").";
						}
						return;
					}			
				}
				
				//Re-search in QD VC
				foreach (DataRow dr in frmInput.ds.Tables["VC Catalogue - RESO APAC CHINA QD"].Rows)
				{
					VCID = RemoveBlank(dr["RES_IDENTIFIER"].ToString()); 
					if (id == VCID)
					{
						VC_NFF = dr["VC_RFR_NFF"].ToString();
						VC_RFR60 = dr["VC_RFR"].ToString();
						return;
					}
					if ( "CISCO" + id == VCID )
					{
						VC_NFF = dr["VC_RFR_NFF"].ToString();
						VC_RFR60 = dr["VC_RFR"].ToString();
						remark = "Looked up by linking 'CISCO' and the original ID. New ID: CISCO" + id;
						return;
					}
				}
				
				foreach (DataRow dr in frmInput.ds.Tables["VC Catalogue - RESO APAC CHINA QD"].Rows)
				{
					VCID = RemoveBlank(dr["RES_IDENTIFIER"].ToString());				
					if (CompareString(id, VCID))
					{
						VC_NFF = dr["VC_RFR_NFF"].ToString();
						VC_RFR60 = dr["VC_RFR"].ToString();
						if (id.Length < VCID.Length)
						{
							remark = "Looked up by cutting the length of the 'RES_IDENTIFIER' (" + VCID + ") to the length of the original ID (" + id + ").";
						}
						else if (id.Length > VCID.Length)
						{
							remark = "Looked up by cutting the length of the original ID (" + id + ") to the length of the 'RES_IDENTIFIER' (" + VCID + ").";
						}
						return;
					}			
				}
				
				//模糊查找
				foreach (DataRow dr in frmInput.ds.Tables["VC Catalogue - RESO China (SHA)"].Rows)
				{
					if (RemoveBlank(dr["RES_IDENTIFIER"].ToString()).Contains(id))
					{
						VC_NFF = dr["VC_NFF"].ToString();
						VC_RFR60 = dr["VC_RFR60"].ToString();
						remark = "Looked up this record as the 'RES_IDENTIFIER' (" + dr["RES_IDENTIFIER"].ToString() + ") contains the original ID (" + id + ").";
						return;
					}
				}
				//模糊查找
				foreach (DataRow dr in frmInput.ds.Tables["VC Catalogue - RESO APAC CHINA QD"].Rows)
				{
					if (RemoveBlank(dr["RES_IDENTIFIER"].ToString()).Contains(id))
					{
						VC_NFF = dr["VC_RFR_NFF"].ToString();
						VC_RFR60 = dr["VC_RFR"].ToString();
						remark = "Looked up this record as the 'RES_IDENTIFIER' (" + dr["RES_IDENTIFIER"].ToString() + ") contains the original ID (" + id + ").";
						return;
					}
				}
				
				VC_RFR60 = "309";
				VC_NFF = "309";
				remark = "Manually put 309 for the repair cost";
				return; 
			}
		}
		
		private bool CompareString(string str1, string str2)
		{
			if (str1.Length == 0 || str2.Length == 0)
			{
				return false;
			}
			
			if (str1 == str2)
			{
				return true;
			}
			else if (str1.Length < str2.Length)
			{
				if (str1 == str2.Substring(0, str1.Length))
				{
					return true;
				}
			}
			else if (str1.Length > str2.Length)
			{
				if (str2 == str1.Substring(0, str2.Length))
				{
					return true;
				}
			}
			
			return false;
		}
		
		private string GetRepairerSubGroup4Citadel(string repairer)
		{
			string ret = "";
			string tempStr = "";
			
			if (repairer == "")
			{
				return "N/A";
			}
			
			if (repairer.ToLower().Contains("alcatel cit"))
			{
				ret = "RLC-Ormes";
			}
			else
			{
			  	tempStr = repairer.Substring(1, 6);
			  	switch (tempStr)
			  	{
					case "INGURR":
						ret = "RLC-India";
						break;
					case "USLOUR":
						ret = "RLC-LVL";
						break;
					case "CHLU2R":
						ret = "RLC-QD";
						break;
					case "CHSH1R":
					case "CHSOLR":		
						ret = "RLC-ASB";
						break;
					default:
						break;				
			  	}
			}
			
			if (ret == "")
			{		
				bool foundAPAC = false;
			  	tempStr = repairer.Substring(1, 2);

			  	foreach (DataRow dr in frmInput.ds.Tables["APAC Country Code List"].Rows)
				{
			  		if (dr["APAC Country ISO CODE"].ToString() != "" &&
			  		    tempStr == dr["APAC Country ISO CODE"].ToString())
			  		{
			  			ret = "3rd party within AP";
			  			foundAPAC =true;
			  			break;
			  		}
			  	}
			  	
			  	if (foundAPAC == false)
			  	{
			  		ret = "3rd party outside of AP";
			  	}
			}
			
			if (ret == "")
			{
				ret = "N/A";
			}
			
			return ret;
		}
		
		private string GetRepairerSubGroup4NZ(string repairer)
		{
			string ret = "";
			string workingStr = repairer;
			
			if (workingStr == "")
			{
				ret = "3rd party within AP";
			}
			else if (workingStr.Contains("ALCATEL-LUCENT LOUISVILLE"))
			{
				ret = "RLC-LVL";
			}
			else if (workingStr.Contains("ALCATEL-LUCENT FRANCE"))
			{
				ret = "RLC-Ormes";
			}
			else if (workingStr.Contains("ALCATEL-LUCENT NZ"))
			{
				ret = "3rd party within AP";
			}
			else if (workingStr.Contains("ALCATEL-LUCENT SHANGHAI BELL"))
			{
				ret = "RLC-ASB";
			}
			else if (workingStr.Contains("ALCATEL-LUCENT QINGDAO"))
			{
				ret = "RLC-QD";
			}
			else
			{
			  	foreach (DataRow dr in frmInput.ds.Tables["Repairer To SubGroup Mapping List - NZ"].Rows)
				{
			  		if (workingStr.ToLower().Contains(dr["Country"].ToString().ToLower()))
			  		{
			  			ret = dr["SubGroup"].ToString();
			  			break;
			  		}
			  	}
			}
			
			if (ret == "")
			{
				ret = "N/A";
			}
			
			return ret;
		}
		
		private string GetRepairerSubGroup4TW(string repairer)
		{
			string ret = "";
			
			foreach (DataRow dr in frmInput.ds.Tables["Repairer To SubGroup Mapping List - TW"].Rows)
			{
				if (repairer.ToUpper().Contains(dr["Repairer"].ToString().ToUpper()))
			  	{
			  		ret = dr["SubGroup"].ToString();
			  		break;
			  	}
			}
			
			if (ret == "")
			{
				ret = "N/A";
			}
			
			return ret;
		}
		
		private string GetRepairerSubGroup4eSparesNew(string repairer)
		{
			string ret = "";
			
			foreach (DataRow dr in frmInput.ds.Tables["Repairer To SubGroup Mapping List - eSpares New"].Rows)
			{
				if (repairer.ToUpper().Contains(dr["Repairer"].ToString().ToUpper()))
			  	{
			  		ret = dr["SubGroup"].ToString();
			  		break;
			  	}
			}
			
			if (ret == "")
			{
				ret = "N/A";
			}
			
			return ret;
		}
	
		private void DataEditing4CitadelReport()
		{
			foreach (DataRow dr in frmInput.ds.Tables["ClosedRMA-Citadel Report"].Rows)
			{
				//Filter on column AF (repairer)
				if (dr["Country"].ToString().ToUpper() == "TAIWAN" ||
					dr["Country"].ToString().ToUpper() == "TAÏWAN" ||
					dr["Country"].ToString().ToUpper() == "NEW ZEALAND")
//					dr["Country"].ToString().ToUpper() == "PHILIPPINES")
				{
					dr["Repair cost"] = "Ignore";
					continue;
				}
				//Look up per column to get the repair price, then times column   repair cost
				else
				{					
					dr["repairer sub-group"] = GetRepairerSubGroup4Citadel(dr["Repairer"].ToString());
					string VC_NFF, VC_RFR60, remark;
					string partNo = RemoveBlank(dr["Rcv_Pno"].ToString());
					if (partNo == "")
					{
						continue;
					}
					
					LookupVC(partNo,
					         dr["repairer sub-group"].ToString(), 
					         dr["Country"].ToString(),
					         out VC_NFF, 
					         out VC_RFR60,
					         out remark);
					
					if (remark == "Manually put 309 for the repair cost")
					{
						string rcvModel = RemoveBlank(dr["Rcv_Model"].ToString());
						if (rcvModel != "")
					 	{
										
							LookupVC(rcvModel,
					         	dr["repairer sub-group"].ToString(), 
					         	dr["Country"].ToString(),
					         	out VC_NFF, 
					        	out VC_RFR60,
					         	out remark);
						}
					}
					
					if (dr["Repair_Code"].ToString().Substring(0, 5) == "(NTF)" ||
					    dr["Repair_Code"].ToString().Substring(0, 5) == "(TOK)")
					{
						dr["Repair cost"] = VC_NFF;
					}
					else
					{
						dr["Repair cost"] = VC_RFR60;
					}
					
					DataRow destDR = frmOutput.ds.Tables["Closed RMA-Citadel"].NewRow();
					destDR["Country"] = dr["Country"];
					destDR["Rcv_Pno"] = dr["Rcv_Pno"];
					destDR["Rcv_Model"] = dr["Rcv_Model"];
					destDR["Repair_Code"] = dr["Repair_Code"];
					destDR["Repairer"] = dr["Repairer"];
					destDR["repairer sub-group"] = dr["repairer sub-group"];
					destDR["Repair cost"] = dr["Repair cost"];
					destDR["Remark"] = remark;				
					if (dr["repairer sub-group"].ToString() == "N/A")
					{
						destDR["Remark"] = "Warning: Cannot get repairer sub-group,pls double check!";
					}
					frmOutput.ds.Tables["Closed RMA-Citadel"].Rows.Add(destDR);
				}
			}
		}
		
		private bool IsNum(string str)
		{
		    for(int i = 0; i < str.Length; i++)
		    {
				//if(!Char.IsNumber(str, i))
				if(str[i]<'0' || str[i]>'9')
				{
					return false;
				}
		    }
		    return true;
		}
		
		private bool CustomerNameFilter4TW(string customerName)
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
		
		private void DataEditing4eSpareTWReport ()
		{

			//Filter on column AR (IR delivery header – GI date)
			foreach (DataRow dr in frmInput.ds.Tables["ClosedRMA-eSpares TW"].Rows)
			{
				bool filtered = false;
				
				if (dr["IR Delivery Header - Vendor Name"].ToString().ToUpper().Contains("TAISEL TAIPEI VERIFY CENTER") &&
				    dr["Work Orders  Line - Repair Code"].ToString() == "TOK")
				{
					//Keep this line.
				}
				else
				{
					foreach (DataRow tempDr in frmInput.ds.Tables["Excluding List TW"].Rows)
					{
						if ( dr["IR Delivery Header - Vendor Name"].ToString().ToLower().Contains(
							tempDr["IR Delivery Header - Vendor Name"].ToString().ToLower()) )
						{
							//frmInput.ds.Tables["ClosedRMA-eSpares TW"].Rows.Remove(dr);
							dr["Repair cost"] = "Ignore";
							filtered = true;
							continue;
						}						
					}
				}	
				
				if (dr["IR Delivery Header - Vendor Name"].ToString().Contains("堡華工業 BOA HWA"))
				{
					string str = dr["IR Delivery Line - Rcvd Material"].ToString();
					if (str.Length == 9 && IsNum(str))
					{
						filtered = true;
					}
				}

				if (this.CustomerNameFilter4TW(dr["Sales Orders - ShipTo Customer Name"].ToString()) == true)
				{
					filtered = true;
				}
				
				if (filtered)
				{
					continue;
				}
				
				DateTime dt = new DateTime();
				try
				{									
					if (dr["IR Delivery Header -  GI Date"].ToString() == "" )
					{
						//frmInput.ds.Tables["ClosedRMA-eSpares TW"].Rows.Remove(dr);
						dr["Repair cost"] = "Ignore";
						continue;
//						string msg = "Error in ClosedRMA-eSpares TW: 'IR Delivery Header -  GI Date' field is empty. Click OK to continue.";
//						MessageBox.Show(msg);
					}
					else
					{
					    dt = Convert.ToDateTime(dr["IR Delivery Header -  GI Date"].ToString()).Date;
					    if ( dt < frmConfig.ClosedRMAStartDateTW || dt > frmConfig.ClosedRMAEndDateTW )
					    {
					    	//frmInput.ds.Tables["ClosedRMA-eSpares TW"].Rows.Remove(dr);
					    	dr["Repair cost"] = "Ignore";
							continue;
					    }
					    //Look up per column to get the repair price, then times column   repair cost
						else
						{
							dr["repairer sub-group"] = GetRepairerSubGroup4TW(dr["IR Delivery Header - Vendor Name"].ToString());
							string VC_NFF, VC_RFR60, remark;
							string rcvdMaterial = RemoveBlank(dr["IR Delivery Line - Rcvd Material"].ToString());
							if (rcvdMaterial == "")
							{
								continue;
							}
							LookupVC(rcvdMaterial,
					       			dr["repairer sub-group"].ToString(), 
					         		"TW",
					         		out VC_NFF, 
					         		out VC_RFR60,
					         		out remark);
							if (dr["Work Orders  Line - Repair Code"].ToString().Contains("NTF") ||
						    	dr["Work Orders  Line - Repair Code"].ToString().Contains("TOK"))
							{
								dr["Repair cost"] = VC_NFF;
							}
							else
							{
								dr["Repair cost"] = VC_RFR60;
							}
							
							DataRow destDR = frmOutput.ds.Tables["Closed RMA-eSpares TW"].NewRow();
							destDR["Country"] = "Taiwan";
							destDR["IR Delivery Header - Vendor Name"] = dr["IR Delivery Header - Vendor Name"];
							destDR["IR Delivery Line - Rcvd Material"] = dr["IR Delivery Line - Rcvd Material"];
							destDR["IR Delivery Header -  GI Date"] = dr["IR Delivery Header -  GI Date"];
							destDR["Work Orders  Line - Repair Code"] = dr["Work Orders  Line - Repair Code"];
							destDR["repairer sub-group"] = dr["repairer sub-group"];
							destDR["Repair cost"] = dr["Repair cost"];
							destDR["Remark"] = remark;				
							if (dr["repairer sub-group"].ToString() == "N/A")
							{
								destDR["Remark"] = "Warning: Cannot get repairer sub-group,pls double check!";
							}
							frmOutput.ds.Tables["Closed RMA-eSpares TW"].Rows.Add(destDR);
						}
					}
				}
				catch
				{
					string msg = "ClosedRMA-eSpares TW: Error Occurred.";
					MessageBox.Show(msg);
				}
			}
		}
		
		private void DataEditing4ASB ()
		{		
			//Filter on column W (leave SHA airport date), to get the reporting month’s data
			foreach (DataRow dr in frmInput.ds.Tables["ClosedRMA-ASB"].Rows)
			{
				DateTime dt = new DateTime();
				try
				{									
					if (dr["Leave SHA airport date"].ToString() == "" ||
					   dr["Country"].ToString().ToUpper() == "TAIWAN" ||
					   dr["Country"].ToString().ToUpper() == "TAÏWAN" ||
					   dr["Country"].ToString().ToUpper() == "NEW ZEALAND" )
//					dr["Country"].ToString().ToUpper() == "PHILIPPINES")
					{
						
						//frmInput.ds.Tables["ClosedRMA-ASB"].Rows.Remove(dr);
						dr["Repair cost"] = "Ignore";
						continue;
					}
					else					
					{
					    dt = Convert.ToDateTime(dr["Leave SHA airport date"].ToString()).Date;
					    
					    if ( dt < frmConfig.ClosedRMAStartDateASB || dt > frmConfig.ClosedRMAEndDateASB )
					    {
					    	//frmInput.ds.Tables["ClosedRMA-ASB"].Rows.Remove(dr);
					    	dr["Repair cost"] = "Ignore";
							continue;
					    }
					    //Look up per column to get the repair price, then times column   repair cost
						else
						{
							dr["repairer sub-group"] = "RLC-ASB";
							string VC_NFF, VC_RFR60, remark;
							string partNo = RemoveBlank(dr["Received P/N"].ToString());
							if (partNo == "")
							{
								continue;
							}
							LookupVC(partNo,
					         		dr["repairer sub-group"].ToString(), 
					         		dr["Country"].ToString(),
					         		out VC_NFF, 
					         		out VC_RFR60,
					         	 	out remark);
							dr["Repair cost"] = VC_RFR60;	
							
							DataRow destDR = frmOutput.ds.Tables["Closed RMA-ASB"].NewRow();
							destDR["Country"] = dr["Country"];
							destDR["Leave SHA airport date"] = dr["Leave SHA airport date"];
							destDR["Received P/N"] = dr["Received P/N"];
							destDR["repairer sub-group"] = dr["repairer sub-group"];
							destDR["Repair cost"] = dr["Repair cost"];
							destDR["Remark"] = remark;
							if (dr["repairer sub-group"].ToString() == "N/A")
							{
								destDR["Remark"] = "Warning: Cannot get repairer sub-group,pls double check!";
							}
							frmOutput.ds.Tables["Closed RMA-ASB"].Rows.Add(destDR);					
						}
					}
				}
				catch
				{
					string msg = "ClosedRMA-ASB: Error Occurred.";
					MessageBox.Show(msg);
				}
			}
		}
		
		private void DataEditing4Ormes ()
		{
			//Filter on column AR (IR delivery header – GI date)
			foreach (DataRow dr in frmInput.ds.Tables["ClosedRMA-Ormes"].Rows)
			{
				DateTime dt = new DateTime();
				try
				{									
					if (dr["OC_GI_Date"].ToString() == "" ||
					    dr["CUST_Country"].ToString().ToUpper() == "TAIWAN" ||
					    dr["CUST_Country"].ToString().ToUpper() == "TAÏWAN" ||
					    dr["CUST_Country"].ToString().ToUpper() == "NEW ZEALAND")
//					dr["Country"].ToString().ToUpper() == "PHILIPPINES")
					{
						//frmInput.ds.Tables["ClosedRMA-Ormes"].Rows.Remove(dr);
						dr["Repair cost"] = "Ignore";
						continue;
					}
					else
					{
					    dt = Convert.ToDateTime(dr["OC_GI_Date"].ToString()).Date;
					    if ( dt < frmConfig.ClosedRMAStartDateOrmes || dt > frmConfig.ClosedRMAEndDateOrmes )
					    {
					    	//frmInput.ds.Tables["ClosedRMA-Ormes"].Rows.Remove(dr);
					    	dr["Repair cost"] = "Ignore";
							continue;
					    }
					    
					    //Look up per column to get the repair price, then times column   repair cost
						else
						{
							dr["repairer sub-group"] = "RLC-Ormes";
							string VC_NFF, VC_RFR60, remark;
							string item_fru = RemoveBlank(dr["ITEM_FRU"].ToString());
							if (item_fru == "")
							{
								continue;
							}
							LookupVC(item_fru,
					         	dr["repairer sub-group"].ToString(), 
					         	dr["CUST_Country"].ToString(),
					         	out VC_NFF, 
					         	out VC_RFR60,
					         	out remark);
							dr["Repair cost"] = VC_RFR60;	
							
							DataRow destDR = frmOutput.ds.Tables["Closed RMA-Ormes"].NewRow();
							destDR["CUST_Country"] = dr["CUST_Country"];
							destDR["OC_GI_Date"] = dr["OC_GI_Date"];
							destDR["ITEM_FRU"] = dr["ITEM_FRU"];
							destDR["repairer sub-group"] = dr["repairer sub-group"];
							destDR["Repair cost"] = dr["Repair cost"];
							destDR["Remark"] = remark;
							if (dr["repairer sub-group"].ToString() == "N/A")
							{
								destDR["Remark"] = "Warning: Cannot get repairer sub-group,pls double check!";
							}
							frmOutput.ds.Tables["Closed RMA-Ormes"].Rows.Add(destDR);			
						}		    
					}
				}
				catch
				{
					string msg = "ClosedRMA-Ormes: Error Occurred.";
					MessageBox.Show(msg);
				}
			}
			
			//Look up per column to get the repair price, then times column   repair cost
		}
		
		private void DataEditing4NZ ()
		{
			//Look up per column to get the repair price, then times column   repair cost
			foreach (DataRow dr in frmInput.ds.Tables["ClosedRMA-NZ"].Rows)
			{
				dr["repairer sub-group"] = GetRepairerSubGroup4NZ(dr["Repairer"].ToString());
				string VC_NFF, VC_RFR60, remark;				
				//Remove blank in the string
				string partNo = RemoveBlank(dr["Manufacturers PartNo"].ToString());	
				if (partNo == "")
				{
					continue;
				}
				LookupVC(partNo, 
					         dr["repairer sub-group"].ToString(), 
					         "NZ",		
					         out VC_NFF, 
					         out VC_RFR60,
					         out remark);
				dr["Repair cost"] = VC_RFR60;	
				
				DataRow destDR = frmOutput.ds.Tables["Closed RMA-NZ"].NewRow();
				destDR["Country"] = "New Zealand";
				destDR["Manufacturers PartNo"] = dr["Manufacturers PartNo"];
				destDR["ExtRep"] = dr["ExtRep"];
				//destDR["RMA1"] = dr["RMA1"];
				destDR["Repairer"] = dr["Repairer"];
				//destDR["Repairer2"] = dr["Repairer2"];
				destDR["repairer sub-group"] = dr["repairer sub-group"];
				destDR["Repair cost"] = dr["Repair cost"];
				destDR["Remark"] = remark;
				if (dr["repairer sub-group"].ToString() == "N/A")
				{
					destDR["Remark"] = "Warning: Cannot get repairer sub-group,pls double check!";
				}
				frmOutput.ds.Tables["Closed RMA-NZ"].Rows.Add(destDR);				
			}	
		}
		
		private string GetCountryName4eSparesNew(string CountryCode)
		{
			string ret = "";
			
			foreach (DataRow dr in frmInput.ds.Tables["eSparesNew - Country Code List"].Rows)
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
		private void DataEditing4eSparesNew()
		{
			foreach (DataRow dr in frmInput.ds.Tables["ClosedRMA-eSpares New"].Rows)
			{
				try
				{
					DateTime dt = Convert.ToDateTime(dr["IR Delivery Header -  GI Date"].ToString()).Date;
					if ( dt < frmConfig.ClosedRMAStartDateeSparesNew || dt > frmConfig.ClosedRMAEndDateeSparesNew )
					{
						dr["Repair cost"] = "Ignore";
						continue;
					}
				}
				catch
				{
					dr["Repair cost"] = "Ignore";
					continue;
				}
				
				//Filter on column 'Work Orders  Header - Vendor Name' (repairer)
//				if (dr["Work Orders  Header - Vendor Name"].ToString().ToUpper().Contains("ALCATEL LUCENT FRANCE") ||
//				    dr["Work Orders  Header - Vendor Name"].ToString().ToUpper().Contains("CIT") ||
//				    dr["Work Orders  Header - Vendor Name"].ToString().ToUpper().Contains("ORMES") ||
//				    dr["Work Orders  Header - Vendor Name"].ToString().ToUpper().Contains("SHANGHAI BELL") ||
//				    dr["Work Orders  Header - Vendor Name"].ToString().ToUpper().Contains("ASB") )
//				{
//					dr["Repair cost"] = "Ignore";
//					continue;
//				}
//				//Look up per column to get the repair price, then times column  repair cost
//				else
				{					
					dr["repairer sub-group"] = GetRepairerSubGroup4eSparesNew(dr["Work Orders  Header - Vendor Name"].ToString());
					string VC_NFF, VC_RFR60, remark;
					string RcvdMaterial = RemoveBlank(dr["IR Delivery Line - Rcvd Material"].ToString());
					if (RcvdMaterial == "")
					{
						continue;
					}
					
					dr["COUNTRY"] = GetCountryName4eSparesNew(dr["Sales Orders - RSCIC/RSLC (Sales Organisation)"].ToString());
					LookupVC(RcvdMaterial,
					         dr["repairer sub-group"].ToString(), 
					         dr["COUNTRY"].ToString(),
					         out VC_NFF, 
					         out VC_RFR60,
					         out remark);
					
					
					dr["Repair cost"] = VC_RFR60;
										
					DataRow destDR = frmOutput.ds.Tables["Closed RMA-eSpares New"].NewRow();
					destDR["Country"] = dr["COUNTRY"];
					destDR["Customer Name"] = dr["Sales Orders SoldTo - Customer Name"];
					destDR["IR Delivery Line - Rcvd Material"] = dr["IR Delivery Line - Rcvd Material"];
					destDR["IR Delivery Header -  GI Date"] = dr["IR Delivery Header -  GI Date"].ToString();
					destDR["repairer sub-group"] = dr["repairer sub-group"];
					destDR["Repair cost"] = dr["Repair cost"];
					destDR["Remark"] = remark;				
					if (dr["repairer sub-group"].ToString() == "N/A")
					{
						destDR["Remark"] = "Warning: Cannot get repairer sub-group,pls double check!";
					}
					frmOutput.ds.Tables["Closed RMA-eSpares New"].Rows.Add(destDR);
				}
			}
		}
	
		private string RemoveBlank(string str)
		{
			//Remove blank in the string
			string ret = "";
			for (int i=0; i < str.Length; i++)
			{
				if (str.Substring(i, 1) != " ")
				{
					ret += str.Substring(i, 1);
				}
			}
			return ret;
		}
		
		private StringCollection GetCountryCollection()
		{
			StringCollection ret = new StringCollection();
			string country;
			
			ret.Add("NEW ZEALAND");
			ret.Add("TAIWAN");
			
			foreach (DataRow dr in frmOutput.ds.Tables["Closed RMA-ASB"].Rows)
			{
				country = dr["Country"].ToString().ToUpper();
				if ( country == "KOREA, REP. OF" || country == "KOREA REPUBLIC OF" || country == "KOREA  REPUBLIC OF")
				{
					country = "KOREA";
				}
				if ( country == "TAÏWAN")
				{
					country = "TAIWAN";
				}
				if ( country == "VIET NAM")
				{
					country = "VIETNAM";
				}
				if ( country == "HONGKONG")
				{
					country = "HONG KONG";
				}
				
				if (! ret.Contains(country))
				{
					ret.Add(country);
				}
			}
			
			foreach (DataRow dr in frmOutput.ds.Tables["Closed RMA-Ormes"].Rows)
			{
				country = dr["CUST_Country"].ToString().ToUpper();
				if ( country == "KOREA, REP. OF" || country == "KOREA REPUBLIC OF" || country == "KOREA  REPUBLIC OF")
				{
					country = "KOREA";
				}
				if ( country == "TAÏWAN")
				{
					country = "TAIWAN";
				}
				if ( country == "VIET NAM")
				{
					country = "VIETNAM";
				}
				if ( country == "HONGKONG")
				{
					country = "HONG KONG";
				}
				
				if (! ret.Contains(country))
				{
					ret.Add(country);
				}
			}
			
			foreach (DataRow dr in frmOutput.ds.Tables["Closed RMA-Citadel"].Rows)
			{
				country = dr["Country"].ToString().ToUpper();
				if ( country == "KOREA, REP. OF" || country == "KOREA REPUBLIC OF" || country == "KOREA  REPUBLIC OF")
				{
					country = "KOREA";
				}
				if ( country == "TAÏWAN")
				{
					country = "TAIWAN";
				}
				if ( country == "VIET NAM")
				{
					country = "VIETNAM";
				}
				if ( country == "HONGKONG")
				{
					country = "HONG KONG";
				}
				
				if (! ret.Contains(country))
				{
					ret.Add(country);
				}
			}
			
			foreach (DataRow dr in frmOutput.ds.Tables["Closed RMA-eSpares New"].Rows)
			{
				country = dr["Country"].ToString().ToUpper();
				if ( country == "KOREA, REP. OF" || country == "KOREA REPUBLIC OF" || country == "KOREA  REPUBLIC OF")
				{
					country = "KOREA";
				}
				if ( country == "TAÏWAN")
				{
					country = "TAIWAN";
				}
				if ( country == "VIET NAM")
				{
					country = "VIETNAM";
				}
				if ( country == "HONGKONG")
				{
					country = "HONG KONG";
				}
				
				if (! ret.Contains(country))
				{
					ret.Add(country);
				}
			}
			
			return ret;
		}
		
		private void GetQtyAndCost(string country, string repairerSubGroup, out int Qty, out double Cost)
		{
			Qty = 0;
			Cost = 0;
			
			foreach (DataRow dr in frmOutput.ds.Tables["Closed RMA-eSpares New"].Rows)
			{
				if ( ( dr["Country"].ToString().ToUpper() == country.ToUpper() ||
				       (country == "VIETNAM" && dr["Country"].ToString().ToUpper() == "VIET NAM") ||
				       (country == "HONG KONG" && dr["Country"].ToString().ToUpper() == "HONGKONG") ||
				       (country == "KOREA" && dr["Country"].ToString().ToUpper() == "KOREA, REP. OF") ||
				       (country == "KOREA" && dr["Country"].ToString().ToUpper() == "KOREA REPUBLIC OF") ||
				       (country == "KOREA" && dr["Country"].ToString().ToUpper() == "KOREA  REPUBLIC OF") )&&
				     dr["Repairer sub-group"].ToString() == repairerSubGroup)
				{
					Qty++;
					if (dr["Repair cost"].ToString() != "")
					{
						Cost += Convert.ToDouble(dr["Repair cost"].ToString());
						Cost = Math.Round(Cost, 2);
					}
				}
			}
			
			foreach (DataRow dr in frmOutput.ds.Tables["Closed RMA-ASB"].Rows)
			{
				if ( ( dr["Country"].ToString().ToUpper() == country.ToUpper() ||
				       (country == "VIETNAM" && dr["Country"].ToString().ToUpper() == "VIET NAM") ||
				       (country == "TAIWAN" && dr["Country"].ToString().ToUpper() == "TAÏWAN") ||
				       (country == "HONG KONG" && dr["Country"].ToString().ToUpper() == "HONGKONG") ||
				       (country == "KOREA" && dr["Country"].ToString().ToUpper() == "KOREA, REP. OF") ||
				       (country == "KOREA" && dr["Country"].ToString().ToUpper() == "KOREA REPUBLIC OF") ||
				       (country == "KOREA" && dr["Country"].ToString().ToUpper() == "KOREA  REPUBLIC OF") )&&
				     dr["Repairer sub-group"].ToString() == repairerSubGroup)
				{
					Qty++;
					if (dr["Repair cost"].ToString() != "")
					{
						Cost += Convert.ToDouble(dr["Repair cost"].ToString());
						Cost = Math.Round(Cost, 2);
					}
				}
			}
			
			foreach (DataRow dr in frmOutput.ds.Tables["Closed RMA-Ormes"].Rows)
			{
				if ( ( dr["CUST_Country"].ToString().ToUpper() == country.ToUpper() ||
				       (country == "VIETNAM" && dr["CUST_Country"].ToString().ToUpper() == "VIET NAM") ||
				       (country == "TAIWAN" && dr["CUST_Country"].ToString().ToUpper() == "TAÏWAN") ||
				       (country == "HONG KONG" && dr["CUST_Country"].ToString().ToUpper() == "HONGKONG") ||
				       (country == "KOREA" && dr["CUST_Country"].ToString().ToUpper() == "KOREA, REP. OF") ||
				       (country == "KOREA" && dr["CUST_Country"].ToString().ToUpper() == "KOREA REPUBLIC OF") ||
				       (country == "KOREA" && dr["CUST_Country"].ToString().ToUpper() == "KOREA  REPUBLIC OF") )&&
				     dr["Repairer sub-group"].ToString() == repairerSubGroup)
				{
					Qty++;
					if (dr["Repair cost"].ToString() != "")
					{
						Cost += Convert.ToDouble(dr["Repair cost"].ToString());
						Cost = Math.Round(Cost, 2);
					}
				}
			}
			
			foreach (DataRow dr in frmOutput.ds.Tables["Closed RMA-Citadel"].Rows)
			{
				if ( ( dr["Country"].ToString().ToUpper() == country.ToUpper() ||
				       (country == "VIETNAM" && dr["Country"].ToString().ToUpper() == "VIET NAM") ||
				       (country == "TAIWAN" && dr["Country"].ToString().ToUpper() == "TAÏWAN") ||
				       (country == "HONG KONG" && dr["Country"].ToString().ToUpper() == "HONGKONG") ||
				       (country == "KOREA" && dr["Country"].ToString().ToUpper() == "KOREA, REP. OF") ||
				       (country == "KOREA" && dr["Country"].ToString().ToUpper() == "KOREA REPUBLIC OF") ||
				       (country == "KOREA" && dr["Country"].ToString().ToUpper() == "KOREA  REPUBLIC OF") )&&
				     dr["Repairer sub-group"].ToString() == repairerSubGroup)
				{
					Qty++;
					if (dr["Repair cost"].ToString() != "")
					{
						Cost += Convert.ToDouble(dr["Repair cost"].ToString());
						Cost = Math.Round(Cost, 2);
					}
				}
			}
			
			if (country.ToUpper() == "NEW ZEALAND")
			{
				foreach (DataRow dr in frmOutput.ds.Tables["Closed RMA-NZ"].Rows)
				{
					if ( dr["Repairer sub-group"].ToString() == repairerSubGroup )
					{
						Qty++;
						if (dr["Repair cost"].ToString() != "")
						{
							Cost += Convert.ToDouble(dr["Repair cost"].ToString());
							Cost = Math.Round(Cost, 2);
						}
					}
				}
			}
			
			if (country.ToUpper() == "TAIWAN" || country.ToUpper() == "TAÏWAN")
			{
				foreach (DataRow dr in frmOutput.ds.Tables["Closed RMA-eSpares TW"].Rows)
				{
					if ( dr["Repairer sub-group"].ToString() == repairerSubGroup )
					{
						Qty++;
						if (dr["Repair cost"].ToString() != "")
						{
							Cost += Convert.ToDouble(dr["Repair cost"].ToString());
							Cost = Math.Round(Cost, 2);
						}
					}
				}
			}
		}
		private void ClosedRMASummary()
		{
			StringCollection countryList;
			DataRow destDR;
			int Qty, subTotalQty = 0, groundTotalQty = 0;
			double Cost, subTotalCost = 0, groundTotalCost = 0;			
			
			countryList = GetCountryCollection();
			foreach (string country in countryList)
			{	
				subTotalQty = 0;
				subTotalCost = 0;
				
				GetQtyAndCost(country, "RLC-ASB", out Qty, out Cost);				
				subTotalQty += Qty;
				subTotalCost += Cost;	
				destDR = frmOutput.ds.Tables["Closed RMA"].NewRow();
				destDR["Country"] = country;
				destDR["Total AP Repair Volume"] = "Total Volume Repaired in ASB (China)";
				destDR["Qty"] = Qty.ToString();
				destDR["RLC Cost (ITP cost only)"] = Math.Round(Cost, 2).ToString();				
				destDR["Remark"] = "";
				frmOutput.ds.Tables["Closed RMA"].Rows.Add(destDR);					
				
				GetQtyAndCost(country, "RLC-QD", out Qty, out Cost);
				subTotalQty += Qty;
				subTotalCost += Cost;		
				destDR = frmOutput.ds.Tables["Closed RMA"].NewRow();
				destDR["Total AP Repair Volume"] = "Total Volume Repaired in QD (China)";
				destDR["Qty"] = Qty.ToString();
				destDR["RLC Cost (ITP cost only)"] = Math.Round(Cost, 2).ToString();				
				destDR["Remark"] = "";
				frmOutput.ds.Tables["Closed RMA"].Rows.Add(destDR);	
				
				GetQtyAndCost(country, "RLC-India", out Qty, out Cost);
				subTotalQty += Qty;
				subTotalCost += Cost;	
				destDR = frmOutput.ds.Tables["Closed RMA"].NewRow();
				destDR["Total AP Repair Volume"] = "Total Volume Repaired in India";
				destDR["Qty"] = Qty.ToString();
				destDR["RLC Cost (ITP cost only)"] = Math.Round(Cost, 2).ToString();				
				destDR["Remark"] = "";
				frmOutput.ds.Tables["Closed RMA"].Rows.Add(destDR);	
												
				GetQtyAndCost(country, "RLC-Ormes", out Qty, out Cost);
				subTotalQty += Qty;
				subTotalCost += Cost;		
				destDR = frmOutput.ds.Tables["Closed RMA"].NewRow();
				destDR["Total AP Repair Volume"] = "Total Volume Repaired by Ormes";
				destDR["Qty"] = Qty.ToString();
				destDR["RLC Cost (ITP cost only)"] = Math.Round(Cost, 2).ToString();				
				destDR["Remark"] = "";
				frmOutput.ds.Tables["Closed RMA"].Rows.Add(destDR);	
				
				GetQtyAndCost(country, "RLC-LVL", out Qty, out Cost);
				subTotalQty += Qty;
				subTotalCost += Cost;		
				destDR = frmOutput.ds.Tables["Closed RMA"].NewRow();
				destDR["Total AP Repair Volume"] = "Total Volume Repaired by  NAR";
				destDR["Qty"] = Qty.ToString();
				destDR["RLC Cost (ITP cost only)"] = Math.Round(Cost, 2).ToString();				
				destDR["Remark"] = "";
				frmOutput.ds.Tables["Closed RMA"].Rows.Add(destDR);	
				
				GetQtyAndCost(country, "3rd party within AP", out Qty, out Cost);
				subTotalQty += Qty;		
				destDR = frmOutput.ds.Tables["Closed RMA"].NewRow();
				destDR["Total AP Repair Volume"] = "Total Volume Repaired by 3rd in APAC";
				destDR["Qty"] = Qty.ToString();				
				destDR["Remark"] = "";
				frmOutput.ds.Tables["Closed RMA"].Rows.Add(destDR);	
				
				GetQtyAndCost(country, "3rd party outside of AP", out Qty, out Cost);
				subTotalQty += Qty;	
				destDR = frmOutput.ds.Tables["Closed RMA"].NewRow();
				destDR["Total AP Repair Volume"] = "Total Volume Repaired by 3rd Party Repair Vendors outside APAC";
				destDR["Qty"] = Qty.ToString();				
				destDR["Remark"] = "";
				frmOutput.ds.Tables["Closed RMA"].Rows.Add(destDR);	
				
				destDR = frmOutput.ds.Tables["Closed RMA"].NewRow();
				destDR["Total AP Repair Volume"] = "Sub-Total";
				destDR["Qty"] = subTotalQty.ToString();
				destDR["RLC Cost (ITP cost only)"] = Math.Round(subTotalCost, 2).ToString();				
				destDR["Remark"] = "";
				frmOutput.ds.Tables["Closed RMA"].Rows.Add(destDR);	
				
				groundTotalQty += subTotalQty;
				groundTotalCost += Math.Round(subTotalCost, 2);
			}	
			
			destDR = frmOutput.ds.Tables["Closed RMA"].NewRow();
			destDR["Total AP Repair Volume"] = "Ground Total";
			destDR["Qty"] = groundTotalQty.ToString();
			destDR["RLC Cost (ITP cost only)"] = Math.Round(groundTotalCost, 2).ToString();				
			destDR["Remark"] = "";
			frmOutput.ds.Tables["Closed RMA"].Rows.Add(destDR);	
		}
		
		public void SetValues()
		{
			DataEditing4ASB();
			DataEditing4eSpareTWReport();
			DataEditing4CitadelReport();
			DataEditing4Ormes();
			DataEditing4NZ();
			DataEditing4eSparesNew();
			ClosedRMASummary();
		}
	}
}
