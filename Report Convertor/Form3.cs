/*
 * Created by SharpDevelop.
 * User: yuqingzh
 * Date: 2011-6-20
 * Time: 11:24
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Drawing;
using System.Windows.Forms;

namespace Report_Convertor
{
	/// <summary>
	/// Description of Form3.
	/// </summary>
	public partial class frmConfig : Form
	{				
		public static DateTime ReportDate = new DateTime();
		public static DateTime CurrentWeekStartDate = new DateTime();
		public static DateTime CurrentWeekEndDate = new DateTime();
		public static DateTime PassDueStartDate = new DateTime();
		public static DateTime PassDueEndDate = new DateTime();
		public static string CurrentWeek;
		public static DateTime CurrentWeekTuesday = new DateTime();	//The name is a bit confusing, 
																	//it is the next Tuesday after current week
																	//Usually, it is the report date as the Non-desp 
																	//report should be done on Tuesday.		
																	
		public static DateTime ClosedRMAStartDateASB = new DateTime();
		public static DateTime ClosedRMAEndDateASB = new DateTime();
		public static DateTime ClosedRMAStartDateTW = new DateTime();
		public static DateTime ClosedRMAEndDateTW = new DateTime();
		public static DateTime ClosedRMAStartDateOrmes = new DateTime();
		public static DateTime ClosedRMAEndDateOrmes = new DateTime();
		public static DateTime ClosedRMAStartDateeSparesNew = new DateTime();
		public static DateTime ClosedRMAEndDateeSparesNew = new DateTime();
		
		
		public static DateTime ActivityYTDStartDate = new DateTime();
		public static DateTime ActivityYTDEndDate = new DateTime();
		
		public static DateTime ActivityYTDBasedOnOTDStartDate = new DateTime();
		public static DateTime ActivityYTDBasedOnOTDEndDate = new DateTime();
		
		
		public frmConfig()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
			
			ReportDate = DateTime.Now.Date;
			for ( DateTime dt = ReportDate; dt >= ReportDate.AddDays(-6); dt = dt.AddDays(-1) )
			{
				if (dt.DayOfWeek == DayOfWeek.Sunday)
				{
					CurrentWeekEndDate = dt;
					break;
				}
			}
			CurrentWeekStartDate = CurrentWeekEndDate.AddDays(-6);
			PassDueEndDate = CurrentWeekStartDate.AddDays(-1);
			PassDueStartDate = PassDueEndDate.AddMonths(-1);
			CurrentWeek = WeekOfYear2(CurrentWeekStartDate).ToString();
//			this.cbOTDReport.Checked = true;
//			this.cbActivityYTD.Checked = true;
			
			int startMonth = 0;
			if ( (ReportDate.Month - 1) == 0 )
			{
				startMonth = 12;
			}
			else
			{
				startMonth = ReportDate.Month - 1;
			}
			
			int startYear = 0;
			if ( (ReportDate.Month - 1) == 0 )
			{
				startYear = ReportDate.Year - 1;
			}
			else 
			{
				startYear = ReportDate.Year;
			}
			
			ClosedRMAStartDateASB = new DateTime(startYear, startMonth, 1);
			ClosedRMAEndDateASB = new DateTime(ReportDate.Year, ReportDate.Month, 1).AddDays(-1);
			ClosedRMAStartDateTW = new DateTime(startYear, startMonth, 1);
			ClosedRMAEndDateTW = new DateTime(ReportDate.Year, ReportDate.Month, 1).AddDays(-1);
			ClosedRMAStartDateOrmes = new DateTime(startYear, startMonth, 1);
			ClosedRMAEndDateOrmes = new DateTime(ReportDate.Year, ReportDate.Month, 1).AddDays(-1);
			ClosedRMAStartDateeSparesNew = new DateTime(startYear, startMonth, 1);
			ClosedRMAEndDateeSparesNew = new DateTime(ReportDate.Year, ReportDate.Month, 1).AddDays(-1);

			
			ActivityYTDStartDate = new DateTime(startYear, startMonth, 1);
			ActivityYTDEndDate = new DateTime(ReportDate.Year, ReportDate.Month, 1).AddDays(-1);
			
			ActivityYTDBasedOnOTDStartDate = new DateTime(startYear, startMonth, 1);
			ActivityYTDBasedOnOTDEndDate = new DateTime(ReportDate.Year, ReportDate.Month, 1).AddDays(-1);
		}

		/// <summary>
        /// 根据日期计算日期周数（以周一为一周的第一天）
        /// </summary>
        /// <param name="date">日期</param>
        /// <returns>日期周数</returns>
   
	
		private int WeekOfYear2(DateTime date)
        {

            DayOfWeek dw = (Convert.ToDateTime(string.Format("{0}-1-1 0:0:0", date.Year.ToString()))).DayOfWeek;
            int day = 0;
            switch (dw)
            {
               
                case DayOfWeek.Monday:
                    {
                        day = -1;
                        break;
                    }
                case DayOfWeek.Tuesday:
                    {
                        day = 0;
                        break;
                    }
                case DayOfWeek.Wednesday:
                    {
                        day = 1;
                        break;
                    }
                case DayOfWeek.Thursday:
                    {
                        day = 2;
                        break;
                    }
                case DayOfWeek.Friday:
                    {
                        day = 3;
                        break;
                    }
                case DayOfWeek.Saturday:
                    {
                        day = 4;
                        break;
                    }
                case DayOfWeek.Sunday:
                    {
                        day = 5;
                        break;
                    }
            }
            int week = (date.DayOfYear + day) / 7 + 1;
			//int week = (date.DayOfYear + day) / 7;
				
            return week;
        } 
		
		void FrmConfigLoad(object sender, EventArgs e)
		{
			this.dtpCurrentWeekFrom.Value = CurrentWeekStartDate;
			this.dtpCurrentWeekTo.Value = CurrentWeekEndDate;
			this.dtpPassDueDateFrom.Value = PassDueStartDate;
			this.dtpPassDueDateTo.Value = PassDueEndDate;	
			this.tbCurrentWeek.Text = (WeekOfYear2(CurrentWeekStartDate)).ToString();
			
			this.dtpASBStart.Value = ClosedRMAStartDateASB;
			this.dtpASBEnd.Value = ClosedRMAEndDateASB;
			this.dtpTWStart.Value = ClosedRMAStartDateTW;
			this.dtpTWEnd.Value = ClosedRMAEndDateTW;
			this.dtpOrmesStart.Value = ClosedRMAStartDateOrmes;
			this.dtpOrmesEnd.Value = ClosedRMAEndDateOrmes;
			this.dtpeSparesNewStart.Value = ClosedRMAStartDateeSparesNew;
			this.dtpeSparesNewEnd.Value = ClosedRMAEndDateeSparesNew;
			
			this.dtpActivityYTDStart.Value = ActivityYTDStartDate;
			this.dtpActivityYTDEnd.Value = ActivityYTDEndDate;
			
			this.dtpActivityYTDBasedOnOTDStart.Value = ActivityYTDBasedOnOTDStartDate;
			this.dtpActivityYTDBasedOnOTDEnd.Value = ActivityYTDBasedOnOTDEndDate;
			
			if (cbOTDReport.Checked)
			{
				this.dtpPassDueDateTo.Enabled = true;
				this.dtpPassDueDateFrom.Enabled = true;
				this.dtpCurrentWeekTo.Enabled = true;
				this.dtpCurrentWeekFrom.Enabled = true;
				this.tbCurrentWeek.Enabled = true;				
			}
			else{
				this.dtpPassDueDateTo.Enabled = false;
				this.dtpPassDueDateFrom.Enabled = false;
				this.dtpCurrentWeekTo.Enabled = false;
				this.dtpCurrentWeekFrom.Enabled = false;
				this.tbCurrentWeek.Enabled = false;				
			}
			
			if (cbClosedRMAReport.Checked)
			{
				this.gbClosedRMA.Enabled = true;
			}
			else
			{
				this.gbClosedRMA.Enabled = false;
			}
			
			if (cbActivityYTD.Checked)
			{
				this.gbActivityYTD.Enabled = true;
				this.gbActivityYTDBasedOnOTD.Enabled = true;
			}
			else
			{
				this.gbActivityYTD.Enabled = false;
				this.gbActivityYTDBasedOnOTD.Enabled = false;
			}
		}
		
		void Button1Click(object sender, EventArgs e)
		{			
			CurrentWeekStartDate = this.dtpCurrentWeekFrom.Value.Date;
			CurrentWeekEndDate = this.dtpCurrentWeekTo.Value.Date;
			PassDueStartDate = this.dtpPassDueDateFrom.Value.Date;
			PassDueEndDate = this.dtpPassDueDateTo.Value.Date;		
			
			for ( DateTime dt = CurrentWeekEndDate; dt <= CurrentWeekEndDate.AddDays(6); dt = dt.AddDays(1) )
			{
				if (dt.DayOfWeek == DayOfWeek.Tuesday)
				{
					CurrentWeekTuesday = dt;
					break;
				}
			}
			
			if (cbOTDReport.Checked == true && this.tbCurrentWeek.Text == "")
			{
				MessageBox.Show("Report week number cannot be blank for OTD report.");
				return;
			}
			else
			{
				CurrentWeek = this.tbCurrentWeek.Text;
			}
			
			ClosedRMAStartDateASB = this.dtpASBStart.Value.Date;
			ClosedRMAEndDateASB = this.dtpASBEnd.Value.Date;
			ClosedRMAStartDateTW = this.dtpTWStart.Value.Date;
			ClosedRMAEndDateTW = this.dtpTWEnd.Value.Date;;
			ClosedRMAStartDateOrmes = this.dtpOrmesStart.Value.Date;
			ClosedRMAEndDateOrmes = this.dtpOrmesEnd.Value.Date;
			ClosedRMAStartDateeSparesNew = this.dtpeSparesNewStart.Value.Date;
			ClosedRMAEndDateeSparesNew = this.dtpeSparesNewEnd.Value.Date;
			
			ActivityYTDStartDate = this.dtpActivityYTDStart.Value.Date;
			ActivityYTDEndDate = this.dtpActivityYTDEnd.Value.Date;
			ActivityYTDBasedOnOTDStartDate = this.dtpActivityYTDBasedOnOTDStart.Value.Date;
			ActivityYTDBasedOnOTDEndDate = this.dtpActivityYTDBasedOnOTDEnd.Value.Date;
			
			frmInput.cbClosedRMAReportChecked = cbClosedRMAReport.Checked;
			frmInput.cbMonthlyOTDAnalysisChecked = cbMonthlyOTDAnalysis.Checked;
			frmInput.cbNonDespReportChecked = cbNonDespReport.Checked;
			frmInput.cbOpenOrderReportChecked = cbOpenOrderReport.Checked;
			frmInput.cbWeeklyOTDAnalysisChecked = cbWeeklyOTDAnalysis.Checked;
			frmInput.cbMonthlyOTDAnalysisChecked = cbMonthlyOTDAnalysis.Checked;
			frmInput.cbOTDReportChecked = cbOTDReport.Checked;
			frmInput.cbActivityYTDChecked = cbActivityYTD.Checked;
			frmInput.cbOTDvsCockpitChecked = cbOTDvsCockpit.Checked;
			
			this.Cursor   =   Cursors.WaitCursor;			
			frmContainer.Input.ReadAllInput();
			this.Cursor   =   Cursors.Default;
			
			this.Hide();			
		}
		
		void DtpCurrentWeekFromValueChanged(object sender, System.EventArgs e)
		{
			this.dtpPassDueDateTo.Value = this.dtpCurrentWeekFrom.Value.AddDays(-1);
			this.dtpPassDueDateFrom.Value = this.dtpPassDueDateTo.Value.AddMonths(-1);
			this.dtpCurrentWeekTo.Value = this.dtpCurrentWeekFrom.Value.AddDays(6);
			this.tbCurrentWeek.Text = (WeekOfYear2(this.dtpCurrentWeekFrom.Value.Date)).ToString(); 
		}
		
		void CbOTDReportCheckedChanged(object sender, EventArgs e)
		{
			if (cbOTDReport.Checked)
			{
				this.dtpPassDueDateTo.Enabled = true;
				this.dtpPassDueDateFrom.Enabled = true;
				this.dtpCurrentWeekTo.Enabled = true;
				this.dtpCurrentWeekFrom.Enabled = true;
				this.tbCurrentWeek.Enabled = true;				
			}
			else{
				this.dtpPassDueDateTo.Enabled = false;
				this.dtpPassDueDateFrom.Enabled = false;
				this.dtpCurrentWeekTo.Enabled = false;
				this.dtpCurrentWeekFrom.Enabled = false;
				this.tbCurrentWeek.Enabled = false;				
			}
		}
		
		void CbClosedRMAReportCheckedChanged(object sender, EventArgs e)
		{
			if (cbClosedRMAReport.Checked)
			{
				gbClosedRMA.Enabled = true;
			}
			else
			{
				gbClosedRMA.Enabled = false;
			}
		}
		
		void CbActivityYTDCheckedChanged(object sender, EventArgs e)
		{
			if (cbActivityYTD.Checked)
			{
				this.gbActivityYTD.Enabled = true;
				this.gbActivityYTDBasedOnOTD.Enabled = true;
			}
			else
			{
				this.gbActivityYTD.Enabled = false;
				this.gbActivityYTDBasedOnOTD.Enabled = false;
			}
		}
		
	}
}
