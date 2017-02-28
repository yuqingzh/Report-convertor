/*
 * Created by SharpDevelop.
 * User: zhangyuqing3
 * Date: 2017-2-22
 * Time: 17:58
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
namespace Report_Convertor
{
	partial class frmInvoice
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
			this.tbInvoiceAmount = new System.Windows.Forms.TextBox();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.dgInvoice = new System.Windows.Forms.DataGrid();
			this.tbInvoiceNumber = new System.Windows.Forms.TextBox();
			this.label2 = new System.Windows.Forms.Label();
			this.cbCurrency = new System.Windows.Forms.ComboBox();
			this.label1 = new System.Windows.Forms.Label();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.cbSubmitMonth = new System.Windows.Forms.ComboBox();
			this.btUpdate = new System.Windows.Forms.Button();
			this.btInsert = new System.Windows.Forms.Button();
			this.tbBpoCategoryDetails = new System.Windows.Forms.TextBox();
			this.label9 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.cbBpoCategory = new System.Windows.Forms.ComboBox();
			this.label7 = new System.Windows.Forms.Label();
			this.cbSupplier = new System.Windows.Forms.ComboBox();
			this.label8 = new System.Windows.Forms.Label();
			this.dtInvoiceDate = new System.Windows.Forms.DateTimePicker();
			this.cbBpoNumber = new System.Windows.Forms.ComboBox();
			((System.ComponentModel.ISupportInitialize)(this.dgInvoice)).BeginInit();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// tbInvoiceAmount
			// 
			this.tbInvoiceAmount.Location = new System.Drawing.Point(375, 45);
			this.tbInvoiceAmount.Name = "tbInvoiceAmount";
			this.tbInvoiceAmount.Size = new System.Drawing.Size(130, 21);
			this.tbInvoiceAmount.TabIndex = 4;
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(277, 49);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(92, 23);
			this.label4.TabIndex = 20;
			this.label4.Text = "Invoice $";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(21, 49);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(80, 23);
			this.label3.TabIndex = 18;
			this.label3.Text = "Invoice No.";
			// 
			// dgInvoice
			// 
			this.dgInvoice.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
									| System.Windows.Forms.AnchorStyles.Left) 
									| System.Windows.Forms.AnchorStyles.Right)));
			this.dgInvoice.DataMember = "";
			this.dgInvoice.HeaderForeColor = System.Drawing.SystemColors.ControlText;
			this.dgInvoice.Location = new System.Drawing.Point(12, 166);
			this.dgInvoice.Name = "dgInvoice";
			this.dgInvoice.Size = new System.Drawing.Size(760, 272);
			this.dgInvoice.TabIndex = 12;
			// 
			// tbInvoiceNumber
			// 
			this.tbInvoiceNumber.Location = new System.Drawing.Point(113, 46);
			this.tbInvoiceNumber.Name = "tbInvoiceNumber";
			this.tbInvoiceNumber.Size = new System.Drawing.Size(144, 21);
			this.tbInvoiceNumber.TabIndex = 3;
			this.tbInvoiceNumber.Leave += new System.EventHandler(this.TbInvoiceNumberLeave);
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(532, 48);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(64, 23);
			this.label2.TabIndex = 17;
			this.label2.Text = "Currency";
			// 
			// cbCurrency
			// 
			this.cbCurrency.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbCurrency.FormattingEnabled = true;
			this.cbCurrency.Items.AddRange(new object[] {
									"NZD",
									"USD",
									"EUR"});
			this.cbCurrency.Location = new System.Drawing.Point(616, 45);
			this.cbCurrency.Name = "cbCurrency";
			this.cbCurrency.Size = new System.Drawing.Size(119, 20);
			this.cbCurrency.TabIndex = 5;
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(23, 18);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(84, 23);
			this.label1.TabIndex = 15;
			this.label1.Text = "Submit month";
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.cbBpoNumber);
			this.groupBox1.Controls.Add(this.cbSubmitMonth);
			this.groupBox1.Controls.Add(this.btUpdate);
			this.groupBox1.Controls.Add(this.btInsert);
			this.groupBox1.Controls.Add(this.tbBpoCategoryDetails);
			this.groupBox1.Controls.Add(this.label9);
			this.groupBox1.Controls.Add(this.label5);
			this.groupBox1.Controls.Add(this.label6);
			this.groupBox1.Controls.Add(this.cbBpoCategory);
			this.groupBox1.Controls.Add(this.label7);
			this.groupBox1.Controls.Add(this.cbSupplier);
			this.groupBox1.Controls.Add(this.label8);
			this.groupBox1.Controls.Add(this.dtInvoiceDate);
			this.groupBox1.Controls.Add(this.tbInvoiceAmount);
			this.groupBox1.Controls.Add(this.label4);
			this.groupBox1.Controls.Add(this.tbInvoiceNumber);
			this.groupBox1.Controls.Add(this.label3);
			this.groupBox1.Controls.Add(this.label2);
			this.groupBox1.Controls.Add(this.cbCurrency);
			this.groupBox1.Controls.Add(this.label1);
			this.groupBox1.Location = new System.Drawing.Point(12, 12);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(760, 148);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			// 
			// cbSubmitMonth
			// 
			this.cbSubmitMonth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbSubmitMonth.FormattingEnabled = true;
			this.cbSubmitMonth.Items.AddRange(new object[] {
									"1",
									"2",
									"3",
									"4",
									"5",
									"6",
									"7",
									"8",
									"9",
									"10",
									"11",
									"12"});
			this.cbSubmitMonth.Location = new System.Drawing.Point(113, 14);
			this.cbSubmitMonth.Name = "cbSubmitMonth";
			this.cbSubmitMonth.Size = new System.Drawing.Size(144, 20);
			this.cbSubmitMonth.TabIndex = 1;
			// 
			// btUpdate
			// 
			this.btUpdate.Location = new System.Drawing.Point(660, 109);
			this.btUpdate.Name = "btUpdate";
			this.btUpdate.Size = new System.Drawing.Size(75, 23);
			this.btUpdate.TabIndex = 11;
			this.btUpdate.Text = "Update";
			this.btUpdate.UseVisualStyleBackColor = true;
			this.btUpdate.Click += new System.EventHandler(this.BtUpdateClick);
			// 
			// btInsert
			// 
			this.btInsert.Location = new System.Drawing.Point(559, 109);
			this.btInsert.Name = "btInsert";
			this.btInsert.Size = new System.Drawing.Size(75, 23);
			this.btInsert.TabIndex = 10;
			this.btInsert.Text = "Insert";
			this.btInsert.UseVisualStyleBackColor = true;
			this.btInsert.Click += new System.EventHandler(this.BtInsertClick);
			// 
			// tbBpoCategoryDetails
			// 
			this.tbBpoCategoryDetails.Location = new System.Drawing.Point(170, 109);
			this.tbBpoCategoryDetails.Name = "tbBpoCategoryDetails";
			this.tbBpoCategoryDetails.Size = new System.Drawing.Size(368, 21);
			this.tbBpoCategoryDetails.TabIndex = 9;
			// 
			// label9
			// 
			this.label9.Location = new System.Drawing.Point(21, 112);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(143, 23);
			this.label9.TabIndex = 38;
			this.label9.Text = "BPO Category Details";
			// 
			// label5
			// 
			this.label5.Location = new System.Drawing.Point(21, 79);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(80, 23);
			this.label5.TabIndex = 35;
			this.label5.Text = "BPO Number";
			// 
			// label6
			// 
			this.label6.Location = new System.Drawing.Point(532, 79);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(78, 23);
			this.label6.TabIndex = 34;
			this.label6.Text = "BPO Category";
			// 
			// cbBpoCategory
			// 
			this.cbBpoCategory.Enabled = false;
			this.cbBpoCategory.FormattingEnabled = true;
			this.cbBpoCategory.Items.AddRange(new object[] {
									"RMA admin",
									"W/H handling",
									"Repair",
									"Local trans",
									"Int\'l trans"});
			this.cbBpoCategory.Location = new System.Drawing.Point(616, 77);
			this.cbBpoCategory.Name = "cbBpoCategory";
			this.cbBpoCategory.Size = new System.Drawing.Size(118, 20);
			this.cbBpoCategory.TabIndex = 7;
			// 
			// label7
			// 
			this.label7.Location = new System.Drawing.Point(277, 79);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(64, 23);
			this.label7.TabIndex = 32;
			this.label7.Text = "Supplier";
			// 
			// cbSupplier
			// 
			this.cbSupplier.Enabled = false;
			this.cbSupplier.FormattingEnabled = true;
			this.cbSupplier.Items.AddRange(new object[] {
									"Loop Tech",
									"Panalpina",
									"DHL",
									"K+N",
									"Bollore"});
			this.cbSupplier.Location = new System.Drawing.Point(375, 76);
			this.cbSupplier.Name = "cbSupplier";
			this.cbSupplier.Size = new System.Drawing.Size(129, 20);
			this.cbSupplier.TabIndex = 6;
			// 
			// label8
			// 
			this.label8.Location = new System.Drawing.Point(277, 17);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(82, 23);
			this.label8.TabIndex = 30;
			this.label8.Text = "Invoice date";
			// 
			// dtInvoiceDate
			// 
			this.dtInvoiceDate.Location = new System.Drawing.Point(375, 15);
			this.dtInvoiceDate.Name = "dtInvoiceDate";
			this.dtInvoiceDate.Size = new System.Drawing.Size(150, 21);
			this.dtInvoiceDate.TabIndex = 2;
			// 
			// cbBpoNumber
			// 
			this.cbBpoNumber.Enabled = false;
			this.cbBpoNumber.FormattingEnabled = true;
			this.cbBpoNumber.Location = new System.Drawing.Point(113, 76);
			this.cbBpoNumber.Name = "cbBpoNumber";
			this.cbBpoNumber.Size = new System.Drawing.Size(144, 20);
			this.cbBpoNumber.TabIndex = 39;
			// 
			// frmInvoice
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(784, 450);
			this.Controls.Add(this.dgInvoice);
			this.Controls.Add(this.groupBox1);
			this.Name = "frmInvoice";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Invoice";
			((System.ComponentModel.ISupportInitialize)(this.dgInvoice)).EndInit();
			this.groupBox1.ResumeLayout(false);
			this.groupBox1.PerformLayout();
			this.ResumeLayout(false);
		}
		private System.Windows.Forms.ComboBox cbBpoNumber;
		private System.Windows.Forms.ComboBox cbSubmitMonth;
		private System.Windows.Forms.Button btInsert;
		private System.Windows.Forms.Button btUpdate;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.TextBox tbBpoCategoryDetails;
		private System.Windows.Forms.ComboBox cbSupplier;
		private System.Windows.Forms.ComboBox cbBpoCategory;
		private System.Windows.Forms.DateTimePicker dtInvoiceDate;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.ComboBox cbCurrency;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TextBox tbInvoiceNumber;
		private System.Windows.Forms.DataGrid dgInvoice;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.TextBox tbInvoiceAmount;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
	}
}
