/*
 * Created by SharpDevelop.
 * User: yuqingzh
 * Date: 2011-6-16
 * Time: 15:09
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
namespace Report_Convertor
{
	partial class frmContainer
	{
		/// <summary>
		/// Designer variable used to keep track of non-visual components.
		/// </summary>
		private System.ComponentModel.IContainer components = null;
		
		/// <summary>
		/// Disposes resources used by the form.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing) {
				if (components != null) {
					components.Dispose();
				}
			}
			base.Dispose(disposing);
		}
		
		/// <summary>
		/// This method is required for Windows Forms designer support.
		/// Do not change the method contents inside the source code editor. The Forms designer might
		/// not be able to load this method if it was changed manually.
		/// </summary>
		private void InitializeComponent()
		{
			this.menuStrip1 = new System.Windows.Forms.MenuStrip();
			this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.input1ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.Output2ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripMenuItem4 = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
			this.toolStripMenuItem6 = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
			this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.operationToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.convertInput1TL9000ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.convertInput2TWCloseToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.convertInput3TWOpenToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.convertInput4KOREAToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.convertInput5NAToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripMenuItem5 = new System.Windows.Forms.ToolStripMenuItem();
			this.convetOTDToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.cockpitUploading = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripSeparator4 = new System.Windows.Forms.ToolStripSeparator();
			this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripSeparator5 = new System.Windows.Forms.ToolStripSeparator();
			this.convertOpenOrderOneLogToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.convertOpenOrderNonOneLogToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.convertOpenOrdereSparesNew = new System.Windows.Forms.ToolStripMenuItem();
			this.convertOpenOrderToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
			this.weeklyOTDAnalysisToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripSeparator6 = new System.Windows.Forms.ToolStripSeparator();
			this.monthlyOTDAnalysisToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripSeparator7 = new System.Windows.Forms.ToolStripSeparator();
			this.monthlyActualClosedRMAReportToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
			this.activityYTDReportToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.activityYTDBasedOnOTDToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.activityYTDAllToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.WeeklyActivityReportToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.toolStripSeparator8 = new System.Windows.Forms.ToolStripSeparator();
			this.CockpitVsOTDMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
			this.aboutToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
			this.menuStrip1.SuspendLayout();
			this.SuspendLayout();
			// 
			// menuStrip1
			// 
			this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
									this.fileToolStripMenuItem,
									this.operationToolStripMenuItem,
									this.aboutToolStripMenuItem});
			this.menuStrip1.Location = new System.Drawing.Point(0, 0);
			this.menuStrip1.Name = "menuStrip1";
			this.menuStrip1.Size = new System.Drawing.Size(954, 25);
			this.menuStrip1.TabIndex = 0;
			this.menuStrip1.Text = "menuStrip1";
			// 
			// fileToolStripMenuItem
			// 
			this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
									this.input1ToolStripMenuItem,
									this.Output2ToolStripMenuItem,
									this.toolStripMenuItem4,
									this.toolStripSeparator3,
									this.toolStripMenuItem6,
									this.toolStripMenuItem1,
									this.exitToolStripMenuItem});
			this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
			this.fileToolStripMenuItem.Size = new System.Drawing.Size(39, 21);
			this.fileToolStripMenuItem.Text = "File";
			// 
			// input1ToolStripMenuItem
			// 
			this.input1ToolStripMenuItem.Name = "input1ToolStripMenuItem";
			this.input1ToolStripMenuItem.Size = new System.Drawing.Size(116, 22);
			this.input1ToolStripMenuItem.Text = "Input";
			this.input1ToolStripMenuItem.Click += new System.EventHandler(this.InputToolStripMenuItemClick);
			// 
			// Output2ToolStripMenuItem
			// 
			this.Output2ToolStripMenuItem.Name = "Output2ToolStripMenuItem";
			this.Output2ToolStripMenuItem.Size = new System.Drawing.Size(116, 22);
			this.Output2ToolStripMenuItem.Text = "Output";
			this.Output2ToolStripMenuItem.Click += new System.EventHandler(this.OutputToolStripMenuItemClick);
			// 
			// toolStripMenuItem4
			// 
			this.toolStripMenuItem4.Name = "toolStripMenuItem4";
			this.toolStripMenuItem4.Size = new System.Drawing.Size(116, 22);
			this.toolStripMenuItem4.Text = "Config";
			this.toolStripMenuItem4.Click += new System.EventHandler(this.ToolStripMenuItem4Click);
			// 
			// toolStripSeparator3
			// 
			this.toolStripSeparator3.Name = "toolStripSeparator3";
			this.toolStripSeparator3.Size = new System.Drawing.Size(113, 6);
			// 
			// toolStripMenuItem6
			// 
			this.toolStripMenuItem6.Name = "toolStripMenuItem6";
			this.toolStripMenuItem6.Size = new System.Drawing.Size(116, 22);
			this.toolStripMenuItem6.Text = "Save";
			this.toolStripMenuItem6.Visible = false;
			this.toolStripMenuItem6.Click += new System.EventHandler(this.ToolStripMenuItem6Click);
			// 
			// toolStripMenuItem1
			// 
			this.toolStripMenuItem1.Name = "toolStripMenuItem1";
			this.toolStripMenuItem1.Size = new System.Drawing.Size(116, 22);
			this.toolStripMenuItem1.Text = "Export";
			this.toolStripMenuItem1.Click += new System.EventHandler(this.ToolStripMenuItem1Click);
			// 
			// exitToolStripMenuItem
			// 
			this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
			this.exitToolStripMenuItem.Size = new System.Drawing.Size(116, 22);
			this.exitToolStripMenuItem.Text = "Exit";
			this.exitToolStripMenuItem.Click += new System.EventHandler(this.ExitToolStripMenuItemClick);
			// 
			// operationToolStripMenuItem
			// 
			this.operationToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
									this.convertInput1TL9000ToolStripMenuItem,
									this.convertInput2TWCloseToolStripMenuItem,
									this.convertInput3TWOpenToolStripMenuItem,
									this.convertInput4KOREAToolStripMenuItem,
									this.convertInput5NAToolStripMenuItem,
									this.toolStripMenuItem5,
									this.convetOTDToolStripMenuItem,
									this.cockpitUploading,
									this.toolStripSeparator4,
									this.toolStripMenuItem2,
									this.toolStripSeparator5,
									this.convertOpenOrderOneLogToolStripMenuItem,
									this.convertOpenOrderNonOneLogToolStripMenuItem,
									this.convertOpenOrdereSparesNew,
									this.convertOpenOrderToolStripMenuItem,
									this.toolStripSeparator1,
									this.weeklyOTDAnalysisToolStripMenuItem,
									this.toolStripSeparator6,
									this.monthlyOTDAnalysisToolStripMenuItem,
									this.toolStripSeparator7,
									this.monthlyActualClosedRMAReportToolStripMenuItem,
									this.toolStripSeparator2,
									this.activityYTDReportToolStripMenuItem,
									this.activityYTDBasedOnOTDToolStripMenuItem,
									this.activityYTDAllToolStripMenuItem,
									this.WeeklyActivityReportToolStripMenuItem,
									this.toolStripSeparator8,
									this.CockpitVsOTDMenuItem});
			this.operationToolStripMenuItem.Name = "operationToolStripMenuItem";
			this.operationToolStripMenuItem.Size = new System.Drawing.Size(79, 21);
			this.operationToolStripMenuItem.Text = "Operation";
			// 
			// convertInput1TL9000ToolStripMenuItem
			// 
			this.convertInput1TL9000ToolStripMenuItem.Name = "convertInput1TL9000ToolStripMenuItem";
			this.convertInput1TL9000ToolStripMenuItem.Size = new System.Drawing.Size(297, 22);
			this.convertInput1TL9000ToolStripMenuItem.Text = "Convert Input1TL9000";
			this.convertInput1TL9000ToolStripMenuItem.Visible = false;
			this.convertInput1TL9000ToolStripMenuItem.Click += new System.EventHandler(this.ConvertInput1TL9000ToolStripMenuItemClick);
			// 
			// convertInput2TWCloseToolStripMenuItem
			// 
			this.convertInput2TWCloseToolStripMenuItem.Name = "convertInput2TWCloseToolStripMenuItem";
			this.convertInput2TWCloseToolStripMenuItem.Size = new System.Drawing.Size(297, 22);
			this.convertInput2TWCloseToolStripMenuItem.Text = "Convert Input2TWClose";
			this.convertInput2TWCloseToolStripMenuItem.Click += new System.EventHandler(this.ConvertInput2TWCloseToolStripMenuItemClick);
			// 
			// convertInput3TWOpenToolStripMenuItem
			// 
			this.convertInput3TWOpenToolStripMenuItem.Name = "convertInput3TWOpenToolStripMenuItem";
			this.convertInput3TWOpenToolStripMenuItem.Size = new System.Drawing.Size(297, 22);
			this.convertInput3TWOpenToolStripMenuItem.Text = "Convert Input3TWOpen";
			this.convertInput3TWOpenToolStripMenuItem.Click += new System.EventHandler(this.ConvertInput3TWOpenToolStripMenuItemClick);
			// 
			// convertInput4KOREAToolStripMenuItem
			// 
			this.convertInput4KOREAToolStripMenuItem.Name = "convertInput4KOREAToolStripMenuItem";
			this.convertInput4KOREAToolStripMenuItem.Size = new System.Drawing.Size(297, 22);
			this.convertInput4KOREAToolStripMenuItem.Text = "Convert Input4KOREA";
			this.convertInput4KOREAToolStripMenuItem.Visible = false;
			this.convertInput4KOREAToolStripMenuItem.Click += new System.EventHandler(this.ConvertInput4KOREAToolStripMenuItemClick);
			// 
			// convertInput5NAToolStripMenuItem
			// 
			this.convertInput5NAToolStripMenuItem.Name = "convertInput5NAToolStripMenuItem";
			this.convertInput5NAToolStripMenuItem.Size = new System.Drawing.Size(297, 22);
			this.convertInput5NAToolStripMenuItem.Text = "Convert Input5NZ";
			this.convertInput5NAToolStripMenuItem.Click += new System.EventHandler(this.ConvertInput5NAToolStripMenuItemClick);
			// 
			// toolStripMenuItem5
			// 
			this.toolStripMenuItem5.Name = "toolStripMenuItem5";
			this.toolStripMenuItem5.Size = new System.Drawing.Size(297, 22);
			this.toolStripMenuItem5.Text = "Convert OTD w/o KOREA";
			this.toolStripMenuItem5.Click += new System.EventHandler(this.ToolStripMenuItem5Click);
			// 
			// convetOTDToolStripMenuItem
			// 
			this.convetOTDToolStripMenuItem.Name = "convetOTDToolStripMenuItem";
			this.convetOTDToolStripMenuItem.Size = new System.Drawing.Size(297, 22);
			this.convetOTDToolStripMenuItem.Text = "Convert OTD";
			this.convetOTDToolStripMenuItem.Click += new System.EventHandler(this.ConvertOTDToolStripMenuItemClick);
			// 
			// cockpitUploading
			// 
			this.cockpitUploading.Name = "cockpitUploading";
			this.cockpitUploading.Size = new System.Drawing.Size(297, 22);
			this.cockpitUploading.Text = "Cockpit Uploading";
			this.cockpitUploading.Click += new System.EventHandler(this.CockpitUploadingClick);
			// 
			// toolStripSeparator4
			// 
			this.toolStripSeparator4.Name = "toolStripSeparator4";
			this.toolStripSeparator4.Size = new System.Drawing.Size(294, 6);
			// 
			// toolStripMenuItem2
			// 
			this.toolStripMenuItem2.Name = "toolStripMenuItem2";
			this.toolStripMenuItem2.Size = new System.Drawing.Size(297, 22);
			this.toolStripMenuItem2.Text = "Convert NonDesp";
			this.toolStripMenuItem2.Click += new System.EventHandler(this.ToolStripMenuItem2Click);
			// 
			// toolStripSeparator5
			// 
			this.toolStripSeparator5.Name = "toolStripSeparator5";
			this.toolStripSeparator5.Size = new System.Drawing.Size(294, 6);
			// 
			// convertOpenOrderOneLogToolStripMenuItem
			// 
			this.convertOpenOrderOneLogToolStripMenuItem.Name = "convertOpenOrderOneLogToolStripMenuItem";
			this.convertOpenOrderOneLogToolStripMenuItem.Size = new System.Drawing.Size(297, 22);
			this.convertOpenOrderOneLogToolStripMenuItem.Text = "Convert Input7OpenOrderOnelog";
			this.convertOpenOrderOneLogToolStripMenuItem.Click += new System.EventHandler(this.ConvertOpenOrderOneLogToolStripMenuItemClick);
			// 
			// convertOpenOrderNonOneLogToolStripMenuItem
			// 
			this.convertOpenOrderNonOneLogToolStripMenuItem.Name = "convertOpenOrderNonOneLogToolStripMenuItem";
			this.convertOpenOrderNonOneLogToolStripMenuItem.Size = new System.Drawing.Size(297, 22);
			this.convertOpenOrderNonOneLogToolStripMenuItem.Text = "Convert Input8OpenOrderNonOnelog";
			this.convertOpenOrderNonOneLogToolStripMenuItem.Click += new System.EventHandler(this.ConvertOpenOrderNonOneLogToolStripMenuItemClick);
			// 
			// convertOpenOrdereSparesNew
			// 
			this.convertOpenOrdereSparesNew.Name = "convertOpenOrdereSparesNew";
			this.convertOpenOrdereSparesNew.Size = new System.Drawing.Size(297, 22);
			this.convertOpenOrdereSparesNew.Text = "ConvertOpenOrdereSparesNew";
			this.convertOpenOrdereSparesNew.Click += new System.EventHandler(this.ConvertOpenOrdereSparesNewClick);
			// 
			// convertOpenOrderToolStripMenuItem
			// 
			this.convertOpenOrderToolStripMenuItem.Name = "convertOpenOrderToolStripMenuItem";
			this.convertOpenOrderToolStripMenuItem.Size = new System.Drawing.Size(297, 22);
			this.convertOpenOrderToolStripMenuItem.Text = "Convert OpenOrder";
			this.convertOpenOrderToolStripMenuItem.Click += new System.EventHandler(this.ConvertOpenOrderToolStripMenuItemClick);
			// 
			// toolStripSeparator1
			// 
			this.toolStripSeparator1.Name = "toolStripSeparator1";
			this.toolStripSeparator1.Size = new System.Drawing.Size(294, 6);
			// 
			// weeklyOTDAnalysisToolStripMenuItem
			// 
			this.weeklyOTDAnalysisToolStripMenuItem.Name = "weeklyOTDAnalysisToolStripMenuItem";
			this.weeklyOTDAnalysisToolStripMenuItem.Size = new System.Drawing.Size(297, 22);
			this.weeklyOTDAnalysisToolStripMenuItem.Text = "Weekly OTD Analysis";
			this.weeklyOTDAnalysisToolStripMenuItem.Click += new System.EventHandler(this.WeeklyOTDAnalysisToolStripMenuItemClick);
			// 
			// toolStripSeparator6
			// 
			this.toolStripSeparator6.Name = "toolStripSeparator6";
			this.toolStripSeparator6.Size = new System.Drawing.Size(294, 6);
			// 
			// monthlyOTDAnalysisToolStripMenuItem
			// 
			this.monthlyOTDAnalysisToolStripMenuItem.Name = "monthlyOTDAnalysisToolStripMenuItem";
			this.monthlyOTDAnalysisToolStripMenuItem.Size = new System.Drawing.Size(297, 22);
			this.monthlyOTDAnalysisToolStripMenuItem.Text = "Monthly OTD Analysis";
			this.monthlyOTDAnalysisToolStripMenuItem.Click += new System.EventHandler(this.MonthlyOTDAnalysisToolStripMenuItemClick);
			// 
			// toolStripSeparator7
			// 
			this.toolStripSeparator7.Name = "toolStripSeparator7";
			this.toolStripSeparator7.Size = new System.Drawing.Size(294, 6);
			// 
			// monthlyActualClosedRMAReportToolStripMenuItem
			// 
			this.monthlyActualClosedRMAReportToolStripMenuItem.Name = "monthlyActualClosedRMAReportToolStripMenuItem";
			this.monthlyActualClosedRMAReportToolStripMenuItem.Size = new System.Drawing.Size(297, 22);
			this.monthlyActualClosedRMAReportToolStripMenuItem.Text = "Monthly Actual Closed RMA Report";
			this.monthlyActualClosedRMAReportToolStripMenuItem.Click += new System.EventHandler(this.MonthlyActualClosedRMAReportToolStripMenuItemClick);
			// 
			// toolStripSeparator2
			// 
			this.toolStripSeparator2.Name = "toolStripSeparator2";
			this.toolStripSeparator2.Size = new System.Drawing.Size(294, 6);
			// 
			// activityYTDReportToolStripMenuItem
			// 
			this.activityYTDReportToolStripMenuItem.Name = "activityYTDReportToolStripMenuItem";
			this.activityYTDReportToolStripMenuItem.Size = new System.Drawing.Size(297, 22);
			this.activityYTDReportToolStripMenuItem.Text = "Activity YTD Report";
			this.activityYTDReportToolStripMenuItem.Click += new System.EventHandler(this.ActivityYTDReportToolStripMenuItemClick);
			// 
			// activityYTDBasedOnOTDToolStripMenuItem
			// 
			this.activityYTDBasedOnOTDToolStripMenuItem.Name = "activityYTDBasedOnOTDToolStripMenuItem";
			this.activityYTDBasedOnOTDToolStripMenuItem.Size = new System.Drawing.Size(297, 22);
			this.activityYTDBasedOnOTDToolStripMenuItem.Text = "Activity YTD Based On OTD";
			this.activityYTDBasedOnOTDToolStripMenuItem.Click += new System.EventHandler(this.ActivityYTDBasedOnOTDToolStripMenuItemClick);
			// 
			// activityYTDAllToolStripMenuItem
			// 
			this.activityYTDAllToolStripMenuItem.Name = "activityYTDAllToolStripMenuItem";
			this.activityYTDAllToolStripMenuItem.Size = new System.Drawing.Size(297, 22);
			this.activityYTDAllToolStripMenuItem.Text = "Activity YTD All";
			this.activityYTDAllToolStripMenuItem.Visible = false;
			this.activityYTDAllToolStripMenuItem.Click += new System.EventHandler(this.ActivityYTDAllToolStripMenuItemClick);
			// 
			// WeeklyActivityReportToolStripMenuItem
			// 
			this.WeeklyActivityReportToolStripMenuItem.Name = "WeeklyActivityReportToolStripMenuItem";
			this.WeeklyActivityReportToolStripMenuItem.Size = new System.Drawing.Size(297, 22);
			this.WeeklyActivityReportToolStripMenuItem.Text = "Weekly Activity Report";
			this.WeeklyActivityReportToolStripMenuItem.Click += new System.EventHandler(this.WeeklyActivityReportToolStripMenuItemClick);
			// 
			// toolStripSeparator8
			// 
			this.toolStripSeparator8.Name = "toolStripSeparator8";
			this.toolStripSeparator8.Size = new System.Drawing.Size(294, 6);
			// 
			// CockpitVsOTDMenuItem
			// 
			this.CockpitVsOTDMenuItem.Name = "CockpitVsOTDMenuItem";
			this.CockpitVsOTDMenuItem.Size = new System.Drawing.Size(297, 22);
			this.CockpitVsOTDMenuItem.Text = "CockpitVsOTD";
			this.CockpitVsOTDMenuItem.Click += new System.EventHandler(this.CockpitVsOTDMenuItemClick);
			// 
			// aboutToolStripMenuItem
			// 
			this.aboutToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
									this.helpToolStripMenuItem,
									this.aboutToolStripMenuItem1});
			this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
			this.aboutToolStripMenuItem.Size = new System.Drawing.Size(47, 21);
			this.aboutToolStripMenuItem.Text = "Help";
			// 
			// helpToolStripMenuItem
			// 
			this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
			this.helpToolStripMenuItem.Size = new System.Drawing.Size(111, 22);
			this.helpToolStripMenuItem.Text = "Help";
			this.helpToolStripMenuItem.Click += new System.EventHandler(this.HelpToolStripMenuItemClick);
			// 
			// aboutToolStripMenuItem1
			// 
			this.aboutToolStripMenuItem1.Name = "aboutToolStripMenuItem1";
			this.aboutToolStripMenuItem1.Size = new System.Drawing.Size(111, 22);
			this.aboutToolStripMenuItem1.Text = "About";
			this.aboutToolStripMenuItem1.Click += new System.EventHandler(this.AboutToolStripMenuItem1Click);
			// 
			// frmContainer
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(954, 471);
			this.Controls.Add(this.menuStrip1);
			this.IsMdiContainer = true;
			this.MainMenuStrip = this.menuStrip1;
			this.Name = "frmContainer";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Report Convertor";
			this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
			this.Load += new System.EventHandler(this.FrmContainerLoad);
			this.menuStrip1.ResumeLayout(false);
			this.menuStrip1.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();
		}
		private System.Windows.Forms.ToolStripMenuItem WeeklyActivityReportToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem CockpitVsOTDMenuItem;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator8;
		private System.Windows.Forms.ToolStripMenuItem cockpitUploading;
		private System.Windows.Forms.ToolStripMenuItem convertOpenOrdereSparesNew;
		private System.Windows.Forms.ToolStripMenuItem activityYTDAllToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem activityYTDBasedOnOTDToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem activityYTDReportToolStripMenuItem;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
		private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem1;
		private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem monthlyActualClosedRMAReportToolStripMenuItem;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator7;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator6;
		private System.Windows.Forms.ToolStripMenuItem weeklyOTDAnalysisToolStripMenuItem;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator5;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator4;
		private System.Windows.Forms.ToolStripMenuItem monthlyOTDAnalysisToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem convertOpenOrderNonOneLogToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem convertOpenOrderOneLogToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem convertOpenOrderToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem6;
		private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem5;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
		private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem4;
		private System.Windows.Forms.ToolStripMenuItem convetOTDToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem2;
		private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
		private System.Windows.Forms.ToolStripMenuItem convertInput5NAToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem convertInput4KOREAToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem convertInput3TWOpenToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem convertInput2TWCloseToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem convertInput1TL9000ToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
		private System.Windows.Forms.ToolStripMenuItem Output2ToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem operationToolStripMenuItem;
		
		private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem input1ToolStripMenuItem;
		private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
		private System.Windows.Forms.MenuStrip menuStrip1;			
		

	}
}
