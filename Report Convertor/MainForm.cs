/*
 * Created by SharpDevelop.
 * User: yuqingzh
 * Date: 2011-6-16
 * Time: 15:09
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Data;
using System.Data.Sql;
using System.Data.OleDb;

namespace Report_Convertor
{
	/// <summary>
	/// Description of MainForm.
	/// </summary>
	/// 
	public partial class frmContainer : Form
	{
		
		public static Report_Convertor.frmInput Input;
		public static Report_Convertor.frmOutput Output;		
		public static Report_Convertor.frmConfig Config;
			
		public frmContainer()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//		
		}
		
		void FrmContainerLoad(object sender, EventArgs e)
		{
			Input = new frmInput(this);
			Output = new frmOutput(this);		
			
			Input.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			Input.ControlBox = false;
			
			Output.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			Output.ControlBox = false;
			
			Config = new frmConfig();
			Config.Owner = this;
			Config.ControlBox = false;
			Config.ShowDialog();

			if (frmInput.cbClosedRMAReportChecked ||
				frmInput.cbMonthlyOTDAnalysisChecked ||
				frmInput.cbNonDespReportChecked ||
				frmInput.cbOpenOrderReportChecked ||
				frmInput.cbWeeklyOTDAnalysisChecked ||
				frmInput.cbMonthlyOTDAnalysisChecked ||
				frmInput.cbOTDReportChecked ||
				frmInput.cbActivityYTDChecked ||
				frmInput.cbOTDvsCockpitChecked )
			{
				Input.Show();
				Output.Show();				
			}
			
			this.UpdateMenu();
			this.UpdateTab();
		}
		
		public void UpdateMenu()
		{
			if (frmInput.Input1TL9000Exist == false)
			{
				convertInput1TL9000ToolStripMenuItem.Enabled = false;
			}
			else
			{
				convertInput1TL9000ToolStripMenuItem.Enabled = true;
			}
			
			if (frmInput.Input2TWCloseExist == false)
			{
				convertInput2TWCloseToolStripMenuItem.Enabled = false;
			}
			else
			{
				convertInput2TWCloseToolStripMenuItem.Enabled = true;

			}
			
			if (frmInput.Input3TWOpenExist == false)
			{
				convertInput3TWOpenToolStripMenuItem.Enabled = false;
			}
			else
			{
				convertInput3TWOpenToolStripMenuItem.Enabled = true;
			}
			
			if (frmInput.Input4KOREAExist == false)
			{
				convertInput4KOREAToolStripMenuItem.Enabled = false;
			}
			else
			{
				convertInput4KOREAToolStripMenuItem.Enabled = true;
			}
			
			if (frmInput.Input5NZExist == false)
			{
				convertInput5NAToolStripMenuItem.Enabled = false;
			}
			else
			{
				convertInput5NAToolStripMenuItem.Enabled = true;
			}
			
			if (frmInput.Input6NonDespExist == false)
			{
				toolStripMenuItem2.Enabled = false;
			}
			else
			{
				toolStripMenuItem2.Enabled = true;
			}
			
			if ( frmInput.Input1TL9000Exist == false &&
			     frmInput.Input2TWCloseExist == false &&
			     frmInput.Input3TWOpenExist == false &&
			     frmInput.Input5NZExist == false )
			{			     
				toolStripMenuItem5.Enabled = false;
			}
			else{
				toolStripMenuItem5.Enabled = true;
			}
			
			if ( frmInput.Input4KOREAExist == false &&
			     frmInput.Input5NZExist == false &&
			     frmInput.WeeklyOTDExist == false)
			{			     
				cockpitUploading.Enabled = false;
			}
			else{
				cockpitUploading.Enabled = true;
			}
			
			if ( frmInput.Input1TL9000Exist == false &&
			     frmInput.Input2TWCloseExist == false &&
			     frmInput.Input3TWOpenExist == false &&
			     frmInput.Input4KOREAExist == false &&
			     frmInput.Input5NZExist == false )
			{			     
				convetOTDToolStripMenuItem.Enabled = false;
			}
			else{
				convetOTDToolStripMenuItem.Enabled = true;
			}
			
			if ( frmInput.Input7OpenOrderOnelogExist == false )
			{
				convertOpenOrderOneLogToolStripMenuItem.Enabled = false;
			}
			else
			{
				convertOpenOrderOneLogToolStripMenuItem.Enabled = true;
			}
				
			if ( frmInput.Input8OpenOrderNonOnelogExist == false )
			{
				convertOpenOrderNonOneLogToolStripMenuItem.Enabled = false;
			}
			else
			{
				convertOpenOrderNonOneLogToolStripMenuItem.Enabled = true;
			}
			
			if ( frmInput.OpenOrdereSparesNewExist == false )
			{
				convertOpenOrdereSparesNew.Enabled = false;
			}
			else
			{
				convertOpenOrdereSparesNew.Enabled = true;
			}
			
			if ( frmInput.Input7OpenOrderOnelogExist == false &&
			     frmInput.Input8OpenOrderNonOnelogExist == false &&
			     frmInput.OpenOrdereSparesNewExist == false )
			{
				convertOpenOrderToolStripMenuItem.Enabled = false;
			}
			else
			{
				convertOpenOrderToolStripMenuItem.Enabled = true;
			}	
			
			if (frmInput.Input9MonthlyOTDExist == false )
			{
				monthlyOTDAnalysisToolStripMenuItem.Enabled = false;
			}
			else
			{
				monthlyOTDAnalysisToolStripMenuItem.Enabled = true;
			}
			
			if (frmInput.WeeklyOTDExist == false )
			{
				weeklyOTDAnalysisToolStripMenuItem.Enabled = false;
			}
			else
			{
				weeklyOTDAnalysisToolStripMenuItem.Enabled = true;
			}
			
			if (frmInput.ClosedRMAExist == false)
			{
				monthlyActualClosedRMAReportToolStripMenuItem.Enabled = false;
			}
			else
			{
				monthlyActualClosedRMAReportToolStripMenuItem.Enabled = true;
			}
			
			if (frmInput.ActivityYTDExist == false)
			{
				activityYTDReportToolStripMenuItem.Enabled = false;
			}
			else
			{
				activityYTDReportToolStripMenuItem.Enabled = true;
			}
			
			if (frmInput.ActivityYTDBasedOnOTDExist == false)
			{
				this.activityYTDBasedOnOTDToolStripMenuItem.Enabled = false;
			}
			else
			{
				this.activityYTDBasedOnOTDToolStripMenuItem.Enabled = true;
			}
			
			if (frmInput.ActivityYTDBasedOnOTDExist == false || 
			    frmInput.ActivityYTDExist == false)
			{
				this.activityYTDAllToolStripMenuItem.Enabled = false;
			}
			else
			{
				this.activityYTDAllToolStripMenuItem.Enabled = true;
			}

			if (frmInput.WeeklyActivityBasedOnOTDExist == false)
			{
				this.WeeklyActivityReportToolStripMenuItem.Enabled = false;
			}
			else
			{
				this.WeeklyActivityReportToolStripMenuItem.Enabled = true;
			}			
			
			
			if (frmInput.OTDvsCockpitExist == false)
			{
				this.CockpitVsOTDMenuItem.Enabled = false;
			}
			else
			{
				this.CockpitVsOTDMenuItem.Enabled = true;
			}
		}
		
		public void UpdateTab()
		{
			Input.tcInput.TabPages.Clear();
			Output.tcOutput.TabPages.Clear();
			
			if ( frmInput.cbOTDReportChecked )
			{
//				Input.tcInput.TabPages.Add(Input.tpInput1TL9000);
				Input.tcInput.TabPages.Add(Input.tpInput2TWClose);
				Input.tcInput.TabPages.Add(Input.tpInput3TWOpen);
//				Input.tcInput.TabPages.Add(Input.tpInput4KOREA);
				Input.tcInput.TabPages.Add(Input.tpInput5NZ);
				
//				Output.tcOutput.TabPages.Add(Output.tpOpenLog);
//				Output.tcOutput.TabPages.Add(Output.tpNonOnelog);
//				Output.tcOutput.TabPages.Add(Output.tpCockpitAEUploading);
//				Output.tcOutput.TabPages.Add(Output.tpCockpitR4SUploading);
				Output.tcOutput.TabPages.Add(Output.tpOTD);
			}
			
			if ( frmInput.cbNonDespReportChecked )
			{
				Input.tcInput.TabPages.Add(Input.tpInput6NonDesp);
				Output.tcOutput.TabPages.Add(Output.tpNonDesp);
			}
			
			if ( frmInput.cbOpenOrderReportChecked )
			{
				Input.tcInput.TabPages.Add(Input.tp7OpenOrderOnelog);
				Input.tcInput.TabPages.Add(Input.tp8OpenOrderNonOnelog);
				Input.tcInput.TabPages.Add(Input.tpOpenOrdereSparesNew);
				
				
				Output.tcOutput.TabPages.Add(Output.tpOpenOrderOnelog);
				Output.tcOutput.TabPages.Add(Output.tpOpenOrderNonOnelog);	
				Output.tcOutput.TabPages.Add(Output.tpOpenOrdereSparesNew);
			}
			
			if ( frmInput.cbWeeklyOTDAnalysisChecked )
			{
				Input.tcInput.TabPages.Add(Input.tpWeeklyOTDOnelog);
				Input.tcInput.TabPages.Add(Input.tpWeeklyOTDNonOnelog);				
				
				Output.tcOutput.TabPages.Add(Output.tpWeeklyOTDAnalysis);
				Output.tcOutput.TabPages.Add(Output.tpCockpitAEUploading);
				Output.tcOutput.TabPages.Add(Output.tpCockpitR4SUploading);
			}	
			
			if ( frmInput.cbMonthlyOTDAnalysisChecked )
			{
				Input.tcInput.TabPages.Add(Input.tpMonthlyOTDOnelog);
				Input.tcInput.TabPages.Add(Input.tpMonthlyOTDNonOnelog);				
				
				Output.tcOutput.TabPages.Add(Output.tpMonthlyOTDAEMetrics);
				Output.tcOutput.TabPages.Add(Output.tpMonthlyOTDR4SMetrics);
				Output.tcOutput.TabPages.Add(Output.tpMonthlyOTDTWTranslation);
			}
			
			if ( frmInput.cbClosedRMAReportChecked )
			{				
				Input.tcInput.TabPages.Add(Input.tpClosedRMACitadelReport);
				Input.tcInput.TabPages.Add(Input.tpClosedRMAOrmes);
				Input.tcInput.TabPages.Add(Input.tpClosedRMANZ);
				Input.tcInput.TabPages.Add(Input.tpClosedRMAASB);				
				Input.tcInput.TabPages.Add(Input.tpClosedRMAeSparesTW);
				Input.tcInput.TabPages.Add(Input.tpeSparesNew);
				
				Output.tcOutput.TabPages.Add(Output.TpClosedRMASummary);				
				Output.tcOutput.TabPages.Add(Output.tpClosedRMACitadel);
				Output.tcOutput.TabPages.Add(Output.tpClosedRMAOrmes);	
				Output.tcOutput.TabPages.Add(Output.tpClosedRMANZ);				
				Output.tcOutput.TabPages.Add(Output.tpClosedRMAASB);
				Output.tcOutput.TabPages.Add(Output.tpClosedRMATW);
				Output.tcOutput.TabPages.Add(Output.tpeSparesNew);
			}			
		
			if ( frmInput.cbActivityYTDChecked )
			{		
				Output.tcOutput.TabPages.Add(Output.tpWeeklyActivityAE);
				Output.tcOutput.TabPages.Add(Output.tpWeeklyActivityRFS);				
				Output.tcOutput.TabPages.Add(Output.tpActivityYTD);
				Output.tcOutput.TabPages.Add(Output.tpActivityYTDBasedOnOTD);
				
			}
			
			if ( frmInput.cbOTDvsCockpitChecked )
			{				
				Output.tcOutput.TabPages.Add(Output.tpCockpitVsOTDAnalysis);
				Output.tcOutput.TabPages.Add(Output.tpCockpitNotInOTD);
				Output.tcOutput.TabPages.Add(Output.tpOTDNotInCockpit);
			}
		}
		
		void InputToolStripMenuItemClick(object sender, EventArgs e)
		{
			foreach (Form childForm in this.MdiChildren )
			{
				if (childForm is Report_Convertor.frmInput)
				{
					childForm.Focus();
					childForm.WindowState = System.Windows.Forms.FormWindowState.Maximized;
					return;
				}
			}
			
			Input = new frmInput(this);	
			
			Input.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			Input.ControlBox = false;
			Input.Show();		
		}
		
		void OutputToolStripMenuItemClick(object sender, EventArgs e)
		{
			foreach (Form childForm in this.MdiChildren )
			{
				if (childForm is Report_Convertor.frmOutput)
				{
					childForm.Focus();
					childForm.WindowState = System.Windows.Forms.FormWindowState.Maximized;
					return;
				}
			}
			
			frmOutput Output = new frmOutput(this);			
			Output.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			Output.ControlBox = false;
			Output.Show();
		}
	
		void ToolStripMenuItem1Click(object sender, EventArgs e)
		{
			try
			{
				bool ret;
				string outputFilePath = System.Environment.CurrentDirectory + "\\Data\\Output.xls";
				
				if (File.Exists(outputFilePath))
                {                      
                    File.Delete(outputFilePath);
                } 

				DataSet2WorkBook dataSet2WorkBook = new DataSet2WorkBook(ref frmOutput.ds, outputFilePath);
				this.Cursor = Cursors.WaitCursor;
				ret = dataSet2WorkBook.ConvertAll();
				this.Cursor = Cursors.Default;
			
				if (ret == true)
				{
					MessageBox.Show("Completed successfully!");
				}
				else{
					MessageBox.Show("Error occurred! Contact your husband ;-)");
				}
			
			}
			catch
			{
				MessageBox.Show("Error occurred! Restart the application and try again!");
			}
		}

		void ExitToolStripMenuItemClick(object sender, EventArgs e)
		{
			Application.Exit();
		}
		
		void ConvertInput1TL9000ToolStripMenuItemClick(object sender, EventArgs e)
		{

		}
		
		void ConvertInput2TWCloseToolStripMenuItemClick(object sender, EventArgs e)
		{
			frmOutput.ds.Clear();
			
			Onelog onelog = new Onelog();
			onelog.SetValues4Input2TWClose();
			Output.Focus();
			Output.tcOutput.SelectTab(Output.tpOpenLog);
		}
		
		void ConvertInput3TWOpenToolStripMenuItemClick(object sender, EventArgs e)
		{
			frmOutput.ds.Clear();
			
			Onelog onelog = new Onelog();
			onelog.SetValues4Input3TWOpen();
			Output.Focus();
			Output.tcOutput.SelectTab(Output.tpOpenLog);
		}
		
		void ConvertInput4KOREAToolStripMenuItemClick(object sender, EventArgs e)
		{

		}
		
		void ConvertInput5NAToolStripMenuItemClick(object sender, EventArgs e)
		{
			frmOutput.ds.Clear();
			
			NonOnelog nonOnelog = new NonOnelog();
        	nonOnelog.SetValues();
        	Output.Focus();
			Output.tcOutput.SelectTab(Output.tpNonOnelog);
		}
		
		void ConvertOTDToolStripMenuItemClick(object sender, EventArgs e)
		{
			frmOutput.ds.Clear();
			
			this.Cursor = Cursors.WaitCursor;
			OTD otd = new OTD();
			if (frmInput.Input5NZExist)
			{				
				otd.SetValues4NZ();
				Output.Focus();
				Output.tcOutput.SelectTab(Output.tpOTD);
			}     
			
			if (frmInput.Input2TWCloseExist)
			{
				otd.SetValues4Input2TWClose();
				Output.Focus();
				Output.tcOutput.SelectTab(Output.tpOTD);
			}
			
			if (frmInput.Input3TWOpenExist)
			{
				otd.SetValues4Input3TWOpen();
				Output.Focus();
				Output.tcOutput.SelectTab(Output.tpOTD);
			}
        	
			this.Cursor = Cursors.Default;
		}
		
		void ToolStripMenuItem2Click(object sender, EventArgs e)
		{
			frmOutput.ds.Clear();
			this.Cursor = Cursors.WaitCursor;
			NonDesp nonDesp = new NonDesp();
			nonDesp.SetValues();
			
			Output.Focus();
			Output.tcOutput.SelectTab(Output.tpNonDesp);
			
			this.Cursor = Cursors.Default;
		}
		
		void ToolStripMenuItem4Click(object sender, EventArgs e)
		{
			Config.ShowDialog();
			this.UpdateMenu();
			this.UpdateTab();
		}
		
		void ToolStripMenuItem5Click(object sender, EventArgs e)
		{
			frmOutput.ds.Clear();			
			
			if (frmInput.Input5NZExist)
			{
				NonOnelog nonOnelog = new NonOnelog();
				nonOnelog.SetValues();
				
				Output.Focus();
				Output.tcOutput.SelectTab(Output.tpNonOnelog);
			}   
			
			Onelog onelog = new Onelog();
			
        	if (frmInput.Input1TL9000Exist)
			{
        		onelog.SetValues4Input1TL9000();
        		Output.Focus();
				Output.tcOutput.SelectTab(Output.tpOpenLog);
			}
			
			if (frmInput.Input2TWCloseExist)
			{
        		onelog.SetValues4Input2TWClose();
        		Output.Focus();
				Output.tcOutput.SelectTab(Output.tpOpenLog);
			}
			
			if (frmInput.Input3TWOpenExist)
			{
				onelog.SetValues4Input3TWOpen();
				Output.Focus();
				Output.tcOutput.SelectTab(Output.tpOpenLog);
			}
        	    	     	
		}
		
		void ToolStripMenuItem6Click(object sender, EventArgs e)
		{
			string strCon = " Provider = Microsoft.Jet.OLEDB.4.0 ; Data Source =" +
										frmOutput.OutputPath + ";Extended Properties=\"Excel 8.0; HDR=Yes; IMEX=2\"";
			frmOutput.conn = new OleDbConnection(strCon);
        	try
        	{
        		if (frmOutput.conn.State == ConnectionState.Closed)
        		{
        			frmOutput.conn.Open();
        		}
        		frmOutput.da.Update(frmOutput.ds, "NonOneLog");
        	}        	
        	catch(OleDbException E)
        	{
        		throw new Exception(E.Message);
        	}
			finally
			{
				if (frmOutput.conn.State == ConnectionState.Open)
				{
 					frmOutput.conn.Close();
				}
			} 
		}
		
		void ConvertOpenOrderOneLogToolStripMenuItemClick(object sender, EventArgs e)
		{
			frmOutput.ds.Clear();
			this.Cursor = Cursors.WaitCursor;
			OpenOrderOnelog openOrderOnelog = new OpenOrderOnelog();
			openOrderOnelog.SetValues();
			
			Output.Focus();
			Output.tcOutput.SelectTab(Output.tpOpenOrderOnelog);
			this.Cursor = Cursors.Default;
		}
		
		void ConvertOpenOrderNonOneLogToolStripMenuItemClick(object sender, EventArgs e)
		{
			frmOutput.ds.Clear();
			this.Cursor = Cursors.WaitCursor;
			OpenOrderNonOnelog openOrderNonOnelog = new OpenOrderNonOnelog();
			openOrderNonOnelog.SetValues();
			
			Output.Focus();
			Output.tcOutput.SelectTab(Output.tpOpenOrderNonOnelog);
			this.Cursor = Cursors.Default;
		}
		
		void ConvertOpenOrdereSparesNewClick(object sender, EventArgs e)
		{
			frmOutput.ds.Clear();
			this.Cursor = Cursors.WaitCursor;
			OpenOrdereSpares openOrdereSparesNew = new OpenOrdereSpares();
			openOrdereSparesNew.SetValues();
			
			Output.Focus();
			Output.tcOutput.SelectTab(Output.tpOpenOrdereSparesNew);
			this.Cursor = Cursors.Default;
		}
		
		void ConvertOpenOrderToolStripMenuItemClick(object sender, EventArgs e)
		{
			frmOutput.ds.Clear();			
			    
			this.Cursor = Cursors.WaitCursor;
			
			if (frmInput.Input8OpenOrderNonOnelogExist)
			{
				OpenOrderNonOnelog openOrderNonOnelog = new OpenOrderNonOnelog();
				openOrderNonOnelog.SetValues();
				
				Output.Focus();
				Output.tcOutput.SelectTab(Output.tpOpenOrderNonOnelog);
			}
			
			if (frmInput.Input7OpenOrderOnelogExist)
			{
				OpenOrderOnelog openOrderOnelog = new OpenOrderOnelog();
				openOrderOnelog.SetValues();
				Output.Focus();
				Output.tcOutput.SelectTab(Output.tpOpenOrderOnelog);
			}		
			
			if (frmInput.OpenOrdereSparesNewExist)
			{
				OpenOrdereSpares openOrdereSparesNew = new OpenOrdereSpares();
				openOrdereSparesNew.SetValues();
				Output.Focus();
				Output.tcOutput.SelectTab(Output.tpOpenOrdereSparesNew);
			}		
			
			this.Cursor = Cursors.Default;
		}
		
		void MonthlyOTDAnalysisToolStripMenuItemClick(object sender, EventArgs e)
		{
			frmOutput.ds.Clear();
			this.Cursor   =   Cursors.WaitCursor;		
			MonthlyOTD monthlyOTD = new MonthlyOTD();
			monthlyOTD.SetValues();
			monthlyOTD.TWTranslateAll();
			
			Output.Focus();
			Output.tcOutput.SelectTab(Output.tpMonthlyOTDAEMetrics);
			this.Cursor   =   Cursors.Default;		
		}
		
		void WeeklyOTDAnalysisToolStripMenuItemClick(object sender, EventArgs e)
		{
			frmOutput.ds.Clear();
			this.Cursor   =   Cursors.WaitCursor;	
			WeeklyOTD weeklyOTD = new WeeklyOTD();
			weeklyOTD.SetValues();
			
			Output.Focus();
			Output.tcOutput.SelectTab(Output.tpWeeklyOTDAnalysis);
			this.Cursor   =   Cursors.Default;	
		}
		
		void MonthlyActualClosedRMAReportToolStripMenuItemClick(object sender, EventArgs e)
		{
			frmOutput.ds.Clear();
			this.Cursor   =   Cursors.WaitCursor;				
			ClosedRMA closeRMA = new ClosedRMA();
			closeRMA.SetValues();
			Output.Focus();
			Output.tcOutput.SelectTab(Output.TpClosedRMASummary);
			this.Cursor   =   Cursors.Default;			
		}
				
		void HelpToolStripMenuItemClick(object sender, EventArgs e)
		{
			frmHelp help = new frmHelp();
			help.Owner = this;
			help.ShowDialog();
		}
		
		void AboutToolStripMenuItem1Click(object sender, EventArgs e)
		{
			string text = "                    Report Convertor v1.0  Copyright (C) 2011              \n\n" +
						  "                           Author:                                         \n" +
						  "                                       Luis Zhang                          \n" +
						  "                                       Becky Liang                         \n\n" + 
				          "                           Helpdesk:                                       \n" +
						  "                                       Luis.Zhang@Alcatel-Lucent.com       \n" +
						  "                                       Becky.Liang@Alcatel-Lucent.com      \n";
			MessageBox.Show(text);
		}

		
		void ActivityYTDReportToolStripMenuItemClick(object sender, EventArgs e)
		{
			frmOutput.ds.Clear();
			this.Cursor   =   Cursors.WaitCursor;				
			ActivityYTD activityYTD = new ActivityYTD(frmConfig.ActivityYTDStartDate, frmConfig.ActivityYTDEndDate);
			activityYTD.SetValues();
			Output.Focus();
			Output.tcOutput.SelectTab(Output.tpActivityYTD);
			this.Cursor   =   Cursors.Default;			
		}
		
		void ActivityYTDBasedOnOTDToolStripMenuItemClick(object sender, EventArgs e)
		{
			frmOutput.ds.Clear();
			this.Cursor   =   Cursors.WaitCursor;				
			ActivityYTDBasedOnOTD activityYTDBasedOnOTD = new ActivityYTDBasedOnOTD(frmConfig.ActivityYTDBasedOnOTDStartDate, frmConfig.ActivityYTDBasedOnOTDEndDate);
			activityYTDBasedOnOTD.SetValues();
			Output.Focus();
			Output.tcOutput.SelectTab(Output.tpActivityYTD);
			this.Cursor   =   Cursors.Default;	
		}
		
		void ActivityYTDAllToolStripMenuItemClick(object sender, EventArgs e)
		{
//			frmOutput.ds.Clear();
//			this.Cursor   =   Cursors.WaitCursor;				
//			ActivityYTD activityYTD = new ActivityYTD(frmConfig.ActivityYTDStartDate, frmConfig.ActivityYTDEndDate);
//			activityYTD.SetValues();
//			ActivityYTDBasedOnOTD activityYTDBasedOnOTD = new ActivityYTDBasedOnOTD(frmConfig.ActivityYTDBasedOnOTDStartDate, frmConfig.ActivityYTDBasedOnOTDEndDate);
//			activityYTDBasedOnOTD.SetValues();
//			Output.Focus();
//			Output.tcOutput.SelectTab(Output.tpActivityYTD);
//			this.Cursor   =   Cursors.Default;	
		}
		
		
		
		void CockpitUploadingClick(object sender, EventArgs e)
		{
			frmOutput.ds.Clear();
			this.Cursor   =   Cursors.WaitCursor;				
			CockpitUploading cockpitUploading = new CockpitUploading();
			if ( frmInput.Input4KOREAExist == true )
			{			     
				cockpitUploading.SetValues4Input4KOREA();
			}
			if ( frmInput.Input5NZExist == true )
			{			     
				cockpitUploading.SetValues4Input5NZ();
			}	
			if ( frmInput.WeeklyOTDExist == true )
			{
				cockpitUploading.SetValues4WeeklyOTD();
			}
			Output.Focus();
			Output.tcOutput.SelectTab(Output.tpCockpitAEUploading);
			this.Cursor   =   Cursors.Default;	
		}
		
		void CockpitVsOTDMenuItemClick(object sender, System.EventArgs e)
		{
			frmOutput.ds.Clear();
			this.Cursor   =   Cursors.WaitCursor;				
			CockpitVsOTD cockpitVsOTD = new CockpitVsOTD();
			cockpitVsOTD.SetValues();
			Output.Focus();
			Output.tcOutput.SelectTab(Output.tpCockpitVsOTDAnalysis);
			this.Cursor   =   Cursors.Default;	
		}
				
		
		void WeeklyActivityReportToolStripMenuItemClick(object sender, EventArgs e)
		{
			frmOutput.ds.Clear();
			this.Cursor   =   Cursors.WaitCursor;				
			WeeklyActivityYTDBasedOnOTD weeklyActivity = new WeeklyActivityYTDBasedOnOTD();
			weeklyActivity.SetValues();
			Output.Focus();
			Output.tcOutput.SelectTab(Output.tpWeeklyActivityAE);
			this.Cursor   =   Cursors.Default;	
			
		}
		
		
		void BPOPortionToolStripMenuItemClick(object sender, EventArgs e)
		{
			frmBPO bpo = new frmBPO(this);
			//bpo.Owner = this;
			bpo.Show();
			
		}
		
		void InvoicePortionToolStripMenuItemClick(object sender, EventArgs e)
		{
			frmInvoice invoice = new frmInvoice(this);
			//invoice.Owner = this;
			invoice.Show();
			
		}
		
		void SummaryToolStripMenuItemClick(object sender, EventArgs e)
		{
			frmSummary summary = new frmSummary(this);
			//summary.Owner = this;
			summary.Show();
			
		}
		
		void ExportToolStripMenuItemClick(object sender, EventArgs e)
		{
			try
			{
				bool ret;
				string outputFilePath = System.Environment.CurrentDirectory + "\\Data\\BPO\\Summary-" + 
											DateTime.Now.ToShortDateString() + ".xls";
				
				if (File.Exists(outputFilePath))
                {                      
                    File.Delete(outputFilePath);
                } 

				DataSet2WorkBook dataSet2WorkBook = new DataSet2WorkBook(ref frmSummary.ds, outputFilePath);
				this.Cursor = Cursors.WaitCursor;
				ret = dataSet2WorkBook.ConvertAll();
				this.Cursor = Cursors.Default;
			
				if (ret == true)
				{
					MessageBox.Show("Completed successfully!");
				}
				else{
					MessageBox.Show("Error occurred! Contact your husband ;-)");
				}
			
			}
			catch
			{
				MessageBox.Show("Error occurred! Restart the application and try again!");
			}
		}
	}
}
