/*
 * Created by SharpDevelop.
 * User: zhangyuqing3
 * Date: 2017-2-22
 * Time: 18:21
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
namespace Report_Convertor
{
	partial class frmSummary
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
			this.dgSummary = new System.Windows.Forms.DataGrid();
			((System.ComponentModel.ISupportInitialize)(this.dgSummary)).BeginInit();
			this.SuspendLayout();
			// 
			// dgSummary
			// 
			this.dgSummary.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
									| System.Windows.Forms.AnchorStyles.Left) 
									| System.Windows.Forms.AnchorStyles.Right)));
			this.dgSummary.DataMember = "";
			this.dgSummary.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.dgSummary.Location = new System.Drawing.Point(12, 12);
			this.dgSummary.Name = "dgSummary";
			this.dgSummary.Size = new System.Drawing.Size(842, 513);
			this.dgSummary.TabIndex = 0;
			// 
			// frmSummary
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(866, 463);
			this.Controls.Add(this.dgSummary);
			this.Name = "frmSummary";
			this.Text = "Summary";
			((System.ComponentModel.ISupportInitialize)(this.dgSummary)).EndInit();
			this.ResumeLayout(false);
		}
		private System.Windows.Forms.DataGrid dgSummary;
	}
}
