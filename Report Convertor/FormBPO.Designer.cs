/*
 * Created by SharpDevelop.
 * User: zhangyuqing3
 * Date: 2017-2-22
 * Time: 15:44
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
namespace Report_Convertor
{
	partial class frmBPO
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
			this.dgBPO = new System.Windows.Forms.DataGrid();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.btUpdate = new System.Windows.Forms.Button();
			this.btInsert = new System.Windows.Forms.Button();
			this.label7 = new System.Windows.Forms.Label();
			this.dtBpoPeriodEnd = new System.Windows.Forms.DateTimePicker();
			this.label6 = new System.Windows.Forms.Label();
			this.dtBpoPeriodStart = new System.Windows.Forms.DateTimePicker();
			this.label5 = new System.Windows.Forms.Label();
			this.dtBpoDate = new System.Windows.Forms.DateTimePicker();
			this.tbAmount = new System.Windows.Forms.TextBox();
			this.label4 = new System.Windows.Forms.Label();
			this.tbBpoNumber = new System.Windows.Forms.TextBox();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.cbBpoCategory = new System.Windows.Forms.ComboBox();
			this.label1 = new System.Windows.Forms.Label();
			this.cbSupplier = new System.Windows.Forms.ComboBox();
			((System.ComponentModel.ISupportInitialize)(this.dgBPO)).BeginInit();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// dgBPO
			// 
			this.dgBPO.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
									| System.Windows.Forms.AnchorStyles.Left) 
									| System.Windows.Forms.AnchorStyles.Right)));
			this.dgBPO.DataMember = "";
			this.dgBPO.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.dgBPO.Location = new System.Drawing.Point(12, 162);
			this.dgBPO.Name = "dgBPO";
			this.dgBPO.Size = new System.Drawing.Size(760, 276);
			this.dgBPO.TabIndex = 8;
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.btUpdate);
			this.groupBox1.Controls.Add(this.btInsert);
			this.groupBox1.Controls.Add(this.label7);
			this.groupBox1.Controls.Add(this.dtBpoPeriodEnd);
			this.groupBox1.Controls.Add(this.label6);
			this.groupBox1.Controls.Add(this.dtBpoPeriodStart);
			this.groupBox1.Controls.Add(this.label5);
			this.groupBox1.Controls.Add(this.dtBpoDate);
			this.groupBox1.Controls.Add(this.tbAmount);
			this.groupBox1.Controls.Add(this.label4);
			this.groupBox1.Controls.Add(this.tbBpoNumber);
			this.groupBox1.Controls.Add(this.label3);
			this.groupBox1.Controls.Add(this.label2);
			this.groupBox1.Controls.Add(this.cbBpoCategory);
			this.groupBox1.Controls.Add(this.label1);
			this.groupBox1.Controls.Add(this.cbSupplier);
			this.groupBox1.Location = new System.Drawing.Point(12, 12);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(760, 144);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			// 
			// btUpdate
			// 
			this.btUpdate.Location = new System.Drawing.Point(662, 114);
			this.btUpdate.Name = "btUpdate";
			this.btUpdate.Size = new System.Drawing.Size(75, 23);
			this.btUpdate.TabIndex = 30;
			this.btUpdate.Text = "Update";
			this.btUpdate.UseVisualStyleBackColor = true;
			this.btUpdate.Click += new System.EventHandler(this.BtUpdateClick);
			// 
			// btInsert
			// 
			this.btInsert.Location = new System.Drawing.Point(561, 114);
			this.btInsert.Name = "btInsert";
			this.btInsert.Size = new System.Drawing.Size(75, 23);
			this.btInsert.TabIndex = 29;
			this.btInsert.Text = "Insert";
			this.btInsert.UseVisualStyleBackColor = true;
			this.btInsert.Click += new System.EventHandler(this.BtInsertClick);
			// 
			// label7
			// 
			this.label7.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
			this.label7.Location = new System.Drawing.Point(535, 88);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(30, 23);
			this.label7.TabIndex = 27;
			this.label7.Text = "--";
			// 
			// dtBpoPeriodEnd
			// 
			this.dtBpoPeriodEnd.Location = new System.Drawing.Point(580, 83);
			this.dtBpoPeriodEnd.Name = "dtBpoPeriodEnd";
			this.dtBpoPeriodEnd.Size = new System.Drawing.Size(157, 21);
			this.dtBpoPeriodEnd.TabIndex = 7;
			// 
			// label6
			// 
			this.label6.Location = new System.Drawing.Point(273, 89);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(75, 23);
			this.label6.TabIndex = 25;
			this.label6.Text = "BPO period";
			// 
			// dtBpoPeriodStart
			// 
			this.dtBpoPeriodStart.Location = new System.Drawing.Point(356, 83);
			this.dtBpoPeriodStart.Name = "dtBpoPeriodStart";
			this.dtBpoPeriodStart.Size = new System.Drawing.Size(141, 21);
			this.dtBpoPeriodStart.TabIndex = 6;
			// 
			// label5
			// 
			this.label5.Location = new System.Drawing.Point(23, 89);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(55, 23);
			this.label5.TabIndex = 23;
			this.label5.Text = "BPO date";
			// 
			// dtBpoDate
			// 
			this.dtBpoDate.Location = new System.Drawing.Point(104, 83);
			this.dtBpoDate.Name = "dtBpoDate";
			this.dtBpoDate.Size = new System.Drawing.Size(150, 21);
			this.dtBpoDate.TabIndex = 5;
			// 
			// tbAmount
			// 
			this.tbAmount.Location = new System.Drawing.Point(625, 48);
			this.tbAmount.Name = "tbAmount";
			this.tbAmount.Size = new System.Drawing.Size(112, 21);
			this.tbAmount.TabIndex = 4;
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(517, 52);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(92, 23);
			this.label4.TabIndex = 20;
			this.label4.Text = "Amount in NZD";
			// 
			// tbBpoNumber
			// 
			this.tbBpoNumber.Location = new System.Drawing.Point(106, 14);
			this.tbBpoNumber.Name = "tbBpoNumber";
			this.tbBpoNumber.Size = new System.Drawing.Size(148, 21);
			this.tbBpoNumber.TabIndex = 1;
			this.tbBpoNumber.Leave += new System.EventHandler(this.TbBpoNumberLeave);
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(23, 17);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(75, 23);
			this.label3.TabIndex = 18;
			this.label3.Text = "BPO Number";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(23, 52);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(83, 23);
			this.label2.TabIndex = 17;
			this.label2.Text = "BPO Category";
			// 
			// cbBpoCategory
			// 
			this.cbBpoCategory.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbBpoCategory.FormattingEnabled = true;
			this.cbBpoCategory.Items.AddRange(new object[] {
									"RMA admin",
									"W/H handling",
									"Repair",
									"Local trans",
									"Int\'l trans"});
			this.cbBpoCategory.Location = new System.Drawing.Point(106, 49);
			this.cbBpoCategory.Name = "cbBpoCategory";
			this.cbBpoCategory.Size = new System.Drawing.Size(150, 20);
			this.cbBpoCategory.TabIndex = 2;
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(273, 52);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(64, 23);
			this.label1.TabIndex = 15;
			this.label1.Text = "Supplier";
			// 
			// cbSupplier
			// 
			this.cbSupplier.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbSupplier.FormattingEnabled = true;
			this.cbSupplier.Items.AddRange(new object[] {
									"Loop Tech",
									"Panalpina",
									"DHL",
									"K+N",
									"Bollore"});
			this.cbSupplier.Location = new System.Drawing.Point(356, 49);
			this.cbSupplier.Name = "cbSupplier";
			this.cbSupplier.Size = new System.Drawing.Size(141, 20);
			this.cbSupplier.TabIndex = 3;
			// 
			// frmBPO
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(784, 450);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.dgBPO);
			this.Name = "frmBPO";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "BPO";
			((System.ComponentModel.ISupportInitialize)(this.dgBPO)).EndInit();
			this.groupBox1.ResumeLayout(false);
			this.groupBox1.PerformLayout();
			this.ResumeLayout(false);
		}
		private System.Windows.Forms.Button btInsert;
		private System.Windows.Forms.Button btUpdate;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.DataGrid dgBPO;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.DateTimePicker dtBpoPeriodEnd;
		private System.Windows.Forms.DateTimePicker dtBpoPeriodStart;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.DateTimePicker dtBpoDate;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.TextBox tbAmount;
		private System.Windows.Forms.TextBox tbBpoNumber;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.ComboBox cbBpoCategory;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.ComboBox cbSupplier;
		

	}
}
