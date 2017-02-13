/*
 * Created by SharpDevelop.
 * User: yuqingzh
 * Date: 2011-6-20
 * Time: 11:24
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
namespace Report_Convertor
{
	partial class frmConfig
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
			this.button1 = new System.Windows.Forms.Button();
			this.groupBox3 = new System.Windows.Forms.GroupBox();
			this.cbOTDvsCockpit = new System.Windows.Forms.CheckBox();
			this.cbActivityYTD = new System.Windows.Forms.CheckBox();
			this.cbClosedRMAReport = new System.Windows.Forms.CheckBox();
			this.cbMonthlyOTDAnalysis = new System.Windows.Forms.CheckBox();
			this.cbWeeklyOTDAnalysis = new System.Windows.Forms.CheckBox();
			this.cbOpenOrderReport = new System.Windows.Forms.CheckBox();
			this.cbNonDespReport = new System.Windows.Forms.CheckBox();
			this.cbOTDReport = new System.Windows.Forms.CheckBox();
			this.gbOTD = new System.Windows.Forms.GroupBox();
			this.label7 = new System.Windows.Forms.Label();
			this.tbCurrentWeek = new System.Windows.Forms.TextBox();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label3 = new System.Windows.Forms.Label();
			this.dtpPassDueDateTo = new System.Windows.Forms.DateTimePicker();
			this.label4 = new System.Windows.Forms.Label();
			this.dtpPassDueDateFrom = new System.Windows.Forms.DateTimePicker();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.label2 = new System.Windows.Forms.Label();
			this.dtpCurrentWeekTo = new System.Windows.Forms.DateTimePicker();
			this.label1 = new System.Windows.Forms.Label();
			this.dtpCurrentWeekFrom = new System.Windows.Forms.DateTimePicker();
			this.gbClosedRMA = new System.Windows.Forms.GroupBox();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.label16 = new System.Windows.Forms.Label();
			this.dtpeSparesNewEnd = new System.Windows.Forms.DateTimePicker();
			this.label17 = new System.Windows.Forms.Label();
			this.dtpeSparesNewStart = new System.Windows.Forms.DateTimePicker();
			this.groupBox7 = new System.Windows.Forms.GroupBox();
			this.label9 = new System.Windows.Forms.Label();
			this.dtpOrmesEnd = new System.Windows.Forms.DateTimePicker();
			this.label10 = new System.Windows.Forms.Label();
			this.dtpOrmesStart = new System.Windows.Forms.DateTimePicker();
			this.groupBox6 = new System.Windows.Forms.GroupBox();
			this.label6 = new System.Windows.Forms.Label();
			this.dtpTWEnd = new System.Windows.Forms.DateTimePicker();
			this.label8 = new System.Windows.Forms.Label();
			this.dtpTWStart = new System.Windows.Forms.DateTimePicker();
			this.groupBox8 = new System.Windows.Forms.GroupBox();
			this.label5 = new System.Windows.Forms.Label();
			this.dtpASBEnd = new System.Windows.Forms.DateTimePicker();
			this.label11 = new System.Windows.Forms.Label();
			this.dtpASBStart = new System.Windows.Forms.DateTimePicker();
			this.gbActivityYTD = new System.Windows.Forms.GroupBox();
			this.label12 = new System.Windows.Forms.Label();
			this.dtpActivityYTDEnd = new System.Windows.Forms.DateTimePicker();
			this.label13 = new System.Windows.Forms.Label();
			this.dtpActivityYTDStart = new System.Windows.Forms.DateTimePicker();
			this.gbActivityYTDBasedOnOTD = new System.Windows.Forms.GroupBox();
			this.label14 = new System.Windows.Forms.Label();
			this.dtpActivityYTDBasedOnOTDEnd = new System.Windows.Forms.DateTimePicker();
			this.label15 = new System.Windows.Forms.Label();
			this.dtpActivityYTDBasedOnOTDStart = new System.Windows.Forms.DateTimePicker();
			this.groupBox3.SuspendLayout();
			this.gbOTD.SuspendLayout();
			this.groupBox2.SuspendLayout();
			this.groupBox1.SuspendLayout();
			this.gbClosedRMA.SuspendLayout();
			this.groupBox4.SuspendLayout();
			this.groupBox7.SuspendLayout();
			this.groupBox6.SuspendLayout();
			this.groupBox8.SuspendLayout();
			this.gbActivityYTD.SuspendLayout();
			this.gbActivityYTDBasedOnOTD.SuspendLayout();
			this.SuspendLayout();
			// 
			// button1
			// 
			this.button1.Location = new System.Drawing.Point(261, 412);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(75, 21);
			this.button1.TabIndex = 0;
			this.button1.Text = "OK";
			this.button1.UseVisualStyleBackColor = true;
			this.button1.Click += new System.EventHandler(this.Button1Click);
			// 
			// groupBox3
			// 
			this.groupBox3.Controls.Add(this.cbOTDvsCockpit);
			this.groupBox3.Controls.Add(this.cbActivityYTD);
			this.groupBox3.Controls.Add(this.cbClosedRMAReport);
			this.groupBox3.Controls.Add(this.cbMonthlyOTDAnalysis);
			this.groupBox3.Controls.Add(this.cbWeeklyOTDAnalysis);
			this.groupBox3.Controls.Add(this.cbOpenOrderReport);
			this.groupBox3.Controls.Add(this.cbNonDespReport);
			this.groupBox3.Controls.Add(this.cbOTDReport);
			this.groupBox3.Location = new System.Drawing.Point(41, 330);
			this.groupBox3.Name = "groupBox3";
			this.groupBox3.Size = new System.Drawing.Size(554, 68);
			this.groupBox3.TabIndex = 4;
			this.groupBox3.TabStop = false;
			this.groupBox3.Text = "Input Selection";
			// 
			// cbOTDvsCockpit
			// 
			this.cbOTDvsCockpit.Location = new System.Drawing.Point(425, 40);
			this.cbOTDvsCockpit.Name = "cbOTDvsCockpit";
			this.cbOTDvsCockpit.Size = new System.Drawing.Size(149, 22);
			this.cbOTDvsCockpit.TabIndex = 7;
			this.cbOTDvsCockpit.Text = "OTD vs Cockpit";
			this.cbOTDvsCockpit.UseVisualStyleBackColor = true;
			// 
			// cbActivityYTD
			// 
			this.cbActivityYTD.Location = new System.Drawing.Point(425, 18);
			this.cbActivityYTD.Name = "cbActivityYTD";
			this.cbActivityYTD.Size = new System.Drawing.Size(111, 22);
			this.cbActivityYTD.TabIndex = 6;
			this.cbActivityYTD.Text = "Activity YTD";
			this.cbActivityYTD.UseVisualStyleBackColor = true;
			this.cbActivityYTD.CheckedChanged += new System.EventHandler(this.CbActivityYTDCheckedChanged);
			// 
			// cbClosedRMAReport
			// 
			this.cbClosedRMAReport.Location = new System.Drawing.Point(268, 42);
			this.cbClosedRMAReport.Name = "cbClosedRMAReport";
			this.cbClosedRMAReport.Size = new System.Drawing.Size(149, 22);
			this.cbClosedRMAReport.TabIndex = 5;
			this.cbClosedRMAReport.Text = "Closed RMA Report";
			this.cbClosedRMAReport.UseVisualStyleBackColor = true;
			this.cbClosedRMAReport.CheckedChanged += new System.EventHandler(this.CbClosedRMAReportCheckedChanged);
			// 
			// cbMonthlyOTDAnalysis
			// 
			this.cbMonthlyOTDAnalysis.Location = new System.Drawing.Point(268, 18);
			this.cbMonthlyOTDAnalysis.Name = "cbMonthlyOTDAnalysis";
			this.cbMonthlyOTDAnalysis.Size = new System.Drawing.Size(149, 22);
			this.cbMonthlyOTDAnalysis.TabIndex = 4;
			this.cbMonthlyOTDAnalysis.Text = "Monthly OTD Analysis";
			this.cbMonthlyOTDAnalysis.UseVisualStyleBackColor = true;
			// 
			// cbWeeklyOTDAnalysis
			// 
			this.cbWeeklyOTDAnalysis.Location = new System.Drawing.Point(135, 42);
			this.cbWeeklyOTDAnalysis.Name = "cbWeeklyOTDAnalysis";
			this.cbWeeklyOTDAnalysis.Size = new System.Drawing.Size(149, 22);
			this.cbWeeklyOTDAnalysis.TabIndex = 3;
			this.cbWeeklyOTDAnalysis.Text = "Weekly OTD Analysis";
			this.cbWeeklyOTDAnalysis.UseVisualStyleBackColor = true;
			// 
			// cbOpenOrderReport
			// 
			this.cbOpenOrderReport.Location = new System.Drawing.Point(135, 18);
			this.cbOpenOrderReport.Name = "cbOpenOrderReport";
			this.cbOpenOrderReport.Size = new System.Drawing.Size(144, 22);
			this.cbOpenOrderReport.TabIndex = 2;
			this.cbOpenOrderReport.Text = "Open Order Report";
			this.cbOpenOrderReport.UseVisualStyleBackColor = true;
			// 
			// cbNonDespReport
			// 
			this.cbNonDespReport.Location = new System.Drawing.Point(31, 42);
			this.cbNonDespReport.Name = "cbNonDespReport";
			this.cbNonDespReport.Size = new System.Drawing.Size(144, 22);
			this.cbNonDespReport.TabIndex = 1;
			this.cbNonDespReport.Text = "NonDesp Report";
			this.cbNonDespReport.UseVisualStyleBackColor = true;
			// 
			// cbOTDReport
			// 
			this.cbOTDReport.Location = new System.Drawing.Point(31, 18);
			this.cbOTDReport.Name = "cbOTDReport";
			this.cbOTDReport.Size = new System.Drawing.Size(144, 22);
			this.cbOTDReport.TabIndex = 0;
			this.cbOTDReport.Text = "OTD Report";
			this.cbOTDReport.UseVisualStyleBackColor = true;
			this.cbOTDReport.CheckedChanged += new System.EventHandler(this.CbOTDReportCheckedChanged);
			// 
			// gbOTD
			// 
			this.gbOTD.Controls.Add(this.label7);
			this.gbOTD.Controls.Add(this.tbCurrentWeek);
			this.gbOTD.Controls.Add(this.groupBox2);
			this.gbOTD.Controls.Add(this.groupBox1);
			this.gbOTD.Location = new System.Drawing.Point(41, 12);
			this.gbOTD.Name = "gbOTD";
			this.gbOTD.Size = new System.Drawing.Size(255, 176);
			this.gbOTD.TabIndex = 15;
			this.gbOTD.TabStop = false;
			this.gbOTD.Text = "OTD Report";
			// 
			// label7
			// 
			this.label7.Location = new System.Drawing.Point(39, 16);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(74, 21);
			this.label7.TabIndex = 21;
			this.label7.Text = "Report Week";
			// 
			// tbCurrentWeek
			// 
			this.tbCurrentWeek.Location = new System.Drawing.Point(127, 14);
			this.tbCurrentWeek.Name = "tbCurrentWeek";
			this.tbCurrentWeek.Size = new System.Drawing.Size(94, 21);
			this.tbCurrentWeek.TabIndex = 20;
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.label3);
			this.groupBox2.Controls.Add(this.dtpPassDueDateTo);
			this.groupBox2.Controls.Add(this.label4);
			this.groupBox2.Controls.Add(this.dtpPassDueDateFrom);
			this.groupBox2.Location = new System.Drawing.Point(39, 100);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(194, 66);
			this.groupBox2.TabIndex = 17;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Previous Weeks";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(22, 42);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(36, 21);
			this.label3.TabIndex = 4;
			this.label3.Text = "To";
			// 
			// dtpPassDueDateTo
			// 
			this.dtpPassDueDateTo.Enabled = false;
			this.dtpPassDueDateTo.Location = new System.Drawing.Point(61, 42);
			this.dtpPassDueDateTo.Name = "dtpPassDueDateTo";
			this.dtpPassDueDateTo.Size = new System.Drawing.Size(126, 21);
			this.dtpPassDueDateTo.TabIndex = 5;
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(22, 23);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(36, 21);
			this.label4.TabIndex = 2;
			this.label4.Text = "From";
			// 
			// dtpPassDueDateFrom
			// 
			this.dtpPassDueDateFrom.Location = new System.Drawing.Point(61, 20);
			this.dtpPassDueDateFrom.Name = "dtpPassDueDateFrom";
			this.dtpPassDueDateFrom.Size = new System.Drawing.Size(126, 21);
			this.dtpPassDueDateFrom.TabIndex = 4;
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.label2);
			this.groupBox1.Controls.Add(this.dtpCurrentWeekTo);
			this.groupBox1.Controls.Add(this.label1);
			this.groupBox1.Controls.Add(this.dtpCurrentWeekFrom);
			this.groupBox1.Location = new System.Drawing.Point(39, 35);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(194, 66);
			this.groupBox1.TabIndex = 3;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Current Week";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(23, 43);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(36, 21);
			this.label2.TabIndex = 4;
			this.label2.Text = "To";
			// 
			// dtpCurrentWeekTo
			// 
			this.dtpCurrentWeekTo.Location = new System.Drawing.Point(59, 42);
			this.dtpCurrentWeekTo.Name = "dtpCurrentWeekTo";
			this.dtpCurrentWeekTo.Size = new System.Drawing.Size(126, 21);
			this.dtpCurrentWeekTo.TabIndex = 2;
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(23, 20);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(36, 21);
			this.label1.TabIndex = 2;
			this.label1.Text = "From";
			// 
			// dtpCurrentWeekFrom
			// 
			this.dtpCurrentWeekFrom.Location = new System.Drawing.Point(59, 21);
			this.dtpCurrentWeekFrom.Name = "dtpCurrentWeekFrom";
			this.dtpCurrentWeekFrom.Size = new System.Drawing.Size(126, 21);
			this.dtpCurrentWeekFrom.TabIndex = 1;
			// 
			// gbClosedRMA
			// 
			this.gbClosedRMA.Controls.Add(this.groupBox4);
			this.gbClosedRMA.Controls.Add(this.groupBox7);
			this.gbClosedRMA.Controls.Add(this.groupBox6);
			this.gbClosedRMA.Controls.Add(this.groupBox8);
			this.gbClosedRMA.Location = new System.Drawing.Point(339, 12);
			this.gbClosedRMA.Name = "gbClosedRMA";
			this.gbClosedRMA.Size = new System.Drawing.Size(256, 312);
			this.gbClosedRMA.TabIndex = 16;
			this.gbClosedRMA.TabStop = false;
			this.gbClosedRMA.Text = "Closed RMA";
			// 
			// groupBox4
			// 
			this.groupBox4.Controls.Add(this.label16);
			this.groupBox4.Controls.Add(this.dtpeSparesNewEnd);
			this.groupBox4.Controls.Add(this.label17);
			this.groupBox4.Controls.Add(this.dtpeSparesNewStart);
			this.groupBox4.Location = new System.Drawing.Point(28, 246);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(197, 62);
			this.groupBox4.TabIndex = 21;
			this.groupBox4.TabStop = false;
			this.groupBox4.Text = "eSpares New";
			// 
			// label16
			// 
			this.label16.Location = new System.Drawing.Point(23, 40);
			this.label16.Name = "label16";
			this.label16.Size = new System.Drawing.Size(36, 21);
			this.label16.TabIndex = 4;
			this.label16.Text = "To";
			// 
			// dtpeSparesNewEnd
			// 
			this.dtpeSparesNewEnd.Location = new System.Drawing.Point(59, 39);
			this.dtpeSparesNewEnd.Name = "dtpeSparesNewEnd";
			this.dtpeSparesNewEnd.Size = new System.Drawing.Size(126, 21);
			this.dtpeSparesNewEnd.TabIndex = 2;
			// 
			// label17
			// 
			this.label17.Location = new System.Drawing.Point(23, 20);
			this.label17.Name = "label17";
			this.label17.Size = new System.Drawing.Size(36, 21);
			this.label17.TabIndex = 2;
			this.label17.Text = "From";
			// 
			// dtpeSparesNewStart
			// 
			this.dtpeSparesNewStart.Location = new System.Drawing.Point(59, 18);
			this.dtpeSparesNewStart.Name = "dtpeSparesNewStart";
			this.dtpeSparesNewStart.Size = new System.Drawing.Size(126, 21);
			this.dtpeSparesNewStart.TabIndex = 1;
			// 
			// groupBox7
			// 
			this.groupBox7.Controls.Add(this.label9);
			this.groupBox7.Controls.Add(this.dtpOrmesEnd);
			this.groupBox7.Controls.Add(this.label10);
			this.groupBox7.Controls.Add(this.dtpOrmesStart);
			this.groupBox7.Location = new System.Drawing.Point(30, 174);
			this.groupBox7.Name = "groupBox7";
			this.groupBox7.Size = new System.Drawing.Size(197, 62);
			this.groupBox7.TabIndex = 20;
			this.groupBox7.TabStop = false;
			this.groupBox7.Text = "Ormes";
			// 
			// label9
			// 
			this.label9.Location = new System.Drawing.Point(23, 40);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(36, 21);
			this.label9.TabIndex = 4;
			this.label9.Text = "To";
			// 
			// dtpOrmesEnd
			// 
			this.dtpOrmesEnd.Location = new System.Drawing.Point(59, 39);
			this.dtpOrmesEnd.Name = "dtpOrmesEnd";
			this.dtpOrmesEnd.Size = new System.Drawing.Size(126, 21);
			this.dtpOrmesEnd.TabIndex = 2;
			// 
			// label10
			// 
			this.label10.Location = new System.Drawing.Point(23, 20);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(36, 21);
			this.label10.TabIndex = 2;
			this.label10.Text = "From";
			// 
			// dtpOrmesStart
			// 
			this.dtpOrmesStart.Location = new System.Drawing.Point(59, 18);
			this.dtpOrmesStart.Name = "dtpOrmesStart";
			this.dtpOrmesStart.Size = new System.Drawing.Size(126, 21);
			this.dtpOrmesStart.TabIndex = 1;
			// 
			// groupBox6
			// 
			this.groupBox6.Controls.Add(this.label6);
			this.groupBox6.Controls.Add(this.dtpTWEnd);
			this.groupBox6.Controls.Add(this.label8);
			this.groupBox6.Controls.Add(this.dtpTWStart);
			this.groupBox6.Location = new System.Drawing.Point(30, 101);
			this.groupBox6.Name = "groupBox6";
			this.groupBox6.Size = new System.Drawing.Size(197, 62);
			this.groupBox6.TabIndex = 19;
			this.groupBox6.TabStop = false;
			this.groupBox6.Text = "TW";
			// 
			// label6
			// 
			this.label6.Location = new System.Drawing.Point(23, 40);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(36, 21);
			this.label6.TabIndex = 4;
			this.label6.Text = "To";
			// 
			// dtpTWEnd
			// 
			this.dtpTWEnd.Location = new System.Drawing.Point(59, 39);
			this.dtpTWEnd.Name = "dtpTWEnd";
			this.dtpTWEnd.Size = new System.Drawing.Size(126, 21);
			this.dtpTWEnd.TabIndex = 2;
			// 
			// label8
			// 
			this.label8.Location = new System.Drawing.Point(23, 20);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(36, 21);
			this.label8.TabIndex = 2;
			this.label8.Text = "From";
			// 
			// dtpTWStart
			// 
			this.dtpTWStart.Location = new System.Drawing.Point(59, 18);
			this.dtpTWStart.Name = "dtpTWStart";
			this.dtpTWStart.Size = new System.Drawing.Size(126, 21);
			this.dtpTWStart.TabIndex = 1;
			// 
			// groupBox8
			// 
			this.groupBox8.Controls.Add(this.label5);
			this.groupBox8.Controls.Add(this.dtpASBEnd);
			this.groupBox8.Controls.Add(this.label11);
			this.groupBox8.Controls.Add(this.dtpASBStart);
			this.groupBox8.Location = new System.Drawing.Point(30, 26);
			this.groupBox8.Name = "groupBox8";
			this.groupBox8.Size = new System.Drawing.Size(197, 62);
			this.groupBox8.TabIndex = 18;
			this.groupBox8.TabStop = false;
			this.groupBox8.Text = "ASB";
			// 
			// label5
			// 
			this.label5.Location = new System.Drawing.Point(23, 40);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(36, 21);
			this.label5.TabIndex = 4;
			this.label5.Text = "To";
			// 
			// dtpASBEnd
			// 
			this.dtpASBEnd.Location = new System.Drawing.Point(59, 39);
			this.dtpASBEnd.Name = "dtpASBEnd";
			this.dtpASBEnd.Size = new System.Drawing.Size(126, 21);
			this.dtpASBEnd.TabIndex = 2;
			// 
			// label11
			// 
			this.label11.Location = new System.Drawing.Point(23, 20);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(36, 21);
			this.label11.TabIndex = 2;
			this.label11.Text = "From";
			// 
			// dtpASBStart
			// 
			this.dtpASBStart.Location = new System.Drawing.Point(59, 18);
			this.dtpASBStart.Name = "dtpASBStart";
			this.dtpASBStart.Size = new System.Drawing.Size(126, 21);
			this.dtpASBStart.TabIndex = 1;
			// 
			// gbActivityYTD
			// 
			this.gbActivityYTD.Controls.Add(this.label12);
			this.gbActivityYTD.Controls.Add(this.dtpActivityYTDEnd);
			this.gbActivityYTD.Controls.Add(this.label13);
			this.gbActivityYTD.Controls.Add(this.dtpActivityYTDStart);
			this.gbActivityYTD.Location = new System.Drawing.Point(41, 186);
			this.gbActivityYTD.Name = "gbActivityYTD";
			this.gbActivityYTD.Size = new System.Drawing.Size(255, 68);
			this.gbActivityYTD.TabIndex = 20;
			this.gbActivityYTD.TabStop = false;
			this.gbActivityYTD.Text = "Activity YTD Report";
			// 
			// label12
			// 
			this.label12.Location = new System.Drawing.Point(48, 42);
			this.label12.Name = "label12";
			this.label12.Size = new System.Drawing.Size(30, 21);
			this.label12.TabIndex = 4;
			this.label12.Text = "To";
			// 
			// dtpActivityYTDEnd
			// 
			this.dtpActivityYTDEnd.Location = new System.Drawing.Point(84, 41);
			this.dtpActivityYTDEnd.Name = "dtpActivityYTDEnd";
			this.dtpActivityYTDEnd.Size = new System.Drawing.Size(149, 21);
			this.dtpActivityYTDEnd.TabIndex = 2;
			// 
			// label13
			// 
			this.label13.Location = new System.Drawing.Point(47, 20);
			this.label13.Name = "label13";
			this.label13.Size = new System.Drawing.Size(36, 21);
			this.label13.TabIndex = 2;
			this.label13.Text = "From";
			// 
			// dtpActivityYTDStart
			// 
			this.dtpActivityYTDStart.Location = new System.Drawing.Point(83, 18);
			this.dtpActivityYTDStart.Name = "dtpActivityYTDStart";
			this.dtpActivityYTDStart.Size = new System.Drawing.Size(150, 21);
			this.dtpActivityYTDStart.TabIndex = 1;
			// 
			// gbActivityYTDBasedOnOTD
			// 
			this.gbActivityYTDBasedOnOTD.Controls.Add(this.label14);
			this.gbActivityYTDBasedOnOTD.Controls.Add(this.dtpActivityYTDBasedOnOTDEnd);
			this.gbActivityYTDBasedOnOTD.Controls.Add(this.label15);
			this.gbActivityYTDBasedOnOTD.Controls.Add(this.dtpActivityYTDBasedOnOTDStart);
			this.gbActivityYTDBasedOnOTD.Location = new System.Drawing.Point(40, 256);
			this.gbActivityYTDBasedOnOTD.Name = "gbActivityYTDBasedOnOTD";
			this.gbActivityYTDBasedOnOTD.Size = new System.Drawing.Size(256, 68);
			this.gbActivityYTDBasedOnOTD.TabIndex = 21;
			this.gbActivityYTDBasedOnOTD.TabStop = false;
			this.gbActivityYTDBasedOnOTD.Text = "Activity YTD Based On OTD Report";
			// 
			// label14
			// 
			this.label14.Location = new System.Drawing.Point(48, 42);
			this.label14.Name = "label14";
			this.label14.Size = new System.Drawing.Size(30, 21);
			this.label14.TabIndex = 4;
			this.label14.Text = "To";
			// 
			// dtpActivityYTDBasedOnOTDEnd
			// 
			this.dtpActivityYTDBasedOnOTDEnd.Location = new System.Drawing.Point(84, 41);
			this.dtpActivityYTDBasedOnOTDEnd.Name = "dtpActivityYTDBasedOnOTDEnd";
			this.dtpActivityYTDBasedOnOTDEnd.Size = new System.Drawing.Size(149, 21);
			this.dtpActivityYTDBasedOnOTDEnd.TabIndex = 2;
			// 
			// label15
			// 
			this.label15.Location = new System.Drawing.Point(47, 20);
			this.label15.Name = "label15";
			this.label15.Size = new System.Drawing.Size(36, 21);
			this.label15.TabIndex = 2;
			this.label15.Text = "From";
			// 
			// dtpActivityYTDBasedOnOTDStart
			// 
			this.dtpActivityYTDBasedOnOTDStart.Location = new System.Drawing.Point(83, 18);
			this.dtpActivityYTDBasedOnOTDStart.Name = "dtpActivityYTDBasedOnOTDStart";
			this.dtpActivityYTDBasedOnOTDStart.Size = new System.Drawing.Size(150, 21);
			this.dtpActivityYTDBasedOnOTDStart.TabIndex = 1;
			// 
			// frmConfig
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(630, 441);
			this.Controls.Add(this.gbActivityYTDBasedOnOTD);
			this.Controls.Add(this.gbActivityYTD);
			this.Controls.Add(this.gbClosedRMA);
			this.Controls.Add(this.gbOTD);
			this.Controls.Add(this.groupBox3);
			this.Controls.Add(this.button1);
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "frmConfig";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Configuration";
			this.Load += new System.EventHandler(this.FrmConfigLoad);
			this.groupBox3.ResumeLayout(false);
			this.gbOTD.ResumeLayout(false);
			this.gbOTD.PerformLayout();
			this.groupBox2.ResumeLayout(false);
			this.groupBox1.ResumeLayout(false);
			this.gbClosedRMA.ResumeLayout(false);
			this.groupBox4.ResumeLayout(false);
			this.groupBox7.ResumeLayout(false);
			this.groupBox6.ResumeLayout(false);
			this.groupBox8.ResumeLayout(false);
			this.gbActivityYTD.ResumeLayout(false);
			this.gbActivityYTDBasedOnOTD.ResumeLayout(false);
			this.ResumeLayout(false);
		}
		private System.Windows.Forms.CheckBox cbOTDvsCockpit;
		private System.Windows.Forms.DateTimePicker dtpeSparesNewEnd;
		private System.Windows.Forms.DateTimePicker dtpeSparesNewStart;
		private System.Windows.Forms.Label label17;
		private System.Windows.Forms.Label label16;
		private System.Windows.Forms.GroupBox groupBox4;
		private System.Windows.Forms.DateTimePicker dtpActivityYTDBasedOnOTDStart;
		private System.Windows.Forms.Label label15;
		private System.Windows.Forms.DateTimePicker dtpActivityYTDBasedOnOTDEnd;
		private System.Windows.Forms.Label label14;
		private System.Windows.Forms.GroupBox gbActivityYTDBasedOnOTD;
		private System.Windows.Forms.DateTimePicker dtpActivityYTDEnd;
		private System.Windows.Forms.DateTimePicker dtpActivityYTDStart;
		private System.Windows.Forms.Label label13;
		private System.Windows.Forms.Label label12;
		private System.Windows.Forms.GroupBox gbActivityYTD;
		private System.Windows.Forms.CheckBox cbActivityYTD;
		private System.Windows.Forms.DateTimePicker dtpOrmesEnd;
		private System.Windows.Forms.DateTimePicker dtpOrmesStart;
		private System.Windows.Forms.DateTimePicker dtpTWEnd;
		private System.Windows.Forms.DateTimePicker dtpTWStart;
		private System.Windows.Forms.DateTimePicker dtpASBEnd;
		private System.Windows.Forms.DateTimePicker dtpASBStart;
		private System.Windows.Forms.GroupBox gbClosedRMA;
		private System.Windows.Forms.GroupBox gbOTD;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.GroupBox groupBox8;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.GroupBox groupBox6;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.GroupBox groupBox7;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.CheckBox cbWeeklyOTDAnalysis;
		private System.Windows.Forms.CheckBox cbMonthlyOTDAnalysis;
		private System.Windows.Forms.CheckBox cbClosedRMAReport;
		private System.Windows.Forms.CheckBox cbOTDReport;
		private System.Windows.Forms.CheckBox cbNonDespReport;
		private System.Windows.Forms.CheckBox cbOpenOrderReport;
		private System.Windows.Forms.GroupBox groupBox3;
		private System.Windows.Forms.TextBox tbCurrentWeek;
		private System.Windows.Forms.DateTimePicker dtpPassDueDateFrom;
		private System.Windows.Forms.DateTimePicker dtpPassDueDateTo;
		private System.Windows.Forms.DateTimePicker dtpCurrentWeekTo;
		private System.Windows.Forms.DateTimePicker dtpCurrentWeekFrom;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.GroupBox groupBox1;
				
	}
}
