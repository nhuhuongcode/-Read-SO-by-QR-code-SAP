namespace WindowsFormsApp1
{
	partial class Form1
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.gv_rdr1 = new System.Windows.Forms.DataGridView();
			this.lb_docnum = new System.Windows.Forms.Label();
			this.tb_docnum = new System.Windows.Forms.TextBox();
			this.lb_docdate = new System.Windows.Forms.Label();
			this.lb_docduedate = new System.Windows.Forms.Label();
			this.tb_cardname = new System.Windows.Forms.TextBox();
			this.lb_cardname = new System.Windows.Forms.Label();
			this.tb_cardcode = new System.Windows.Forms.TextBox();
			this.lb_cardcode = new System.Windows.Forms.Label();
			this.dt_docdate = new System.Windows.Forms.DateTimePicker();
			this.dt_docduedate = new System.Windows.Forms.DateTimePicker();
			this.bt_Add = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.gv_rdr1)).BeginInit();
			this.SuspendLayout();
			// 
			// gv_rdr1
			// 
			this.gv_rdr1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.gv_rdr1.Location = new System.Drawing.Point(46, 141);
			this.gv_rdr1.Name = "gv_rdr1";
			this.gv_rdr1.RowHeadersWidth = 51;
			this.gv_rdr1.RowTemplate.Height = 24;
			this.gv_rdr1.Size = new System.Drawing.Size(727, 261);
			this.gv_rdr1.TabIndex = 0;
			// 
			// lb_docnum
			// 
			this.lb_docnum.AutoSize = true;
			this.lb_docnum.Location = new System.Drawing.Point(504, 24);
			this.lb_docnum.Name = "lb_docnum";
			this.lb_docnum.Size = new System.Drawing.Size(76, 16);
			this.lb_docnum.TabIndex = 1;
			this.lb_docnum.Text = "Số chứng từ";
			// 
			// tb_docnum
			// 
			this.tb_docnum.Location = new System.Drawing.Point(613, 18);
			this.tb_docnum.Name = "tb_docnum";
			this.tb_docnum.Size = new System.Drawing.Size(103, 22);
			this.tb_docnum.TabIndex = 2;
			// 
			// lb_docdate
			// 
			this.lb_docdate.AutoSize = true;
			this.lb_docdate.Location = new System.Drawing.Point(504, 67);
			this.lb_docdate.Name = "lb_docdate";
			this.lb_docdate.Size = new System.Drawing.Size(84, 16);
			this.lb_docdate.TabIndex = 3;
			this.lb_docdate.Text = "Posting Date";
			// 
			// lb_docduedate
			// 
			this.lb_docduedate.AutoSize = true;
			this.lb_docduedate.Location = new System.Drawing.Point(503, 105);
			this.lb_docduedate.Name = "lb_docduedate";
			this.lb_docduedate.Size = new System.Drawing.Size(89, 16);
			this.lb_docduedate.TabIndex = 6;
			this.lb_docduedate.Text = "Delivery Date";
			// 
			// tb_cardname
			// 
			this.tb_cardname.Location = new System.Drawing.Point(152, 99);
			this.tb_cardname.Name = "tb_cardname";
			this.tb_cardname.Size = new System.Drawing.Size(210, 22);
			this.tb_cardname.TabIndex = 11;
			// 
			// lb_cardname
			// 
			this.lb_cardname.AutoSize = true;
			this.lb_cardname.Location = new System.Drawing.Point(42, 105);
			this.lb_cardname.Name = "lb_cardname";
			this.lb_cardname.Size = new System.Drawing.Size(104, 16);
			this.lb_cardname.TabIndex = 10;
			this.lb_cardname.Text = "Customer Name";
			// 
			// tb_cardcode
			// 
			this.tb_cardcode.Location = new System.Drawing.Point(152, 61);
			this.tb_cardcode.Name = "tb_cardcode";
			this.tb_cardcode.Size = new System.Drawing.Size(100, 22);
			this.tb_cardcode.TabIndex = 9;
			// 
			// lb_cardcode
			// 
			this.lb_cardcode.AutoSize = true;
			this.lb_cardcode.Location = new System.Drawing.Point(43, 67);
			this.lb_cardcode.Name = "lb_cardcode";
			this.lb_cardcode.Size = new System.Drawing.Size(100, 16);
			this.lb_cardcode.TabIndex = 8;
			this.lb_cardcode.Text = "Customer Code";
			// 
			// dt_docdate
			// 
			this.dt_docdate.Location = new System.Drawing.Point(612, 59);
			this.dt_docdate.Name = "dt_docdate";
			this.dt_docdate.Size = new System.Drawing.Size(143, 22);
			this.dt_docdate.TabIndex = 12;
			// 
			// dt_docduedate
			// 
			this.dt_docduedate.CustomFormat = "\"dd.MM.yy\"";
			this.dt_docduedate.Location = new System.Drawing.Point(613, 97);
			this.dt_docduedate.Name = "dt_docduedate";
			this.dt_docduedate.Size = new System.Drawing.Size(142, 22);
			this.dt_docduedate.TabIndex = 13;
			// 
			// bt_Add
			// 
			this.bt_Add.Location = new System.Drawing.Point(669, 424);
			this.bt_Add.Name = "bt_Add";
			this.bt_Add.Size = new System.Drawing.Size(104, 42);
			this.bt_Add.TabIndex = 14;
			this.bt_Add.Text = "Add";
			this.bt_Add.UseVisualStyleBackColor = true;
			this.bt_Add.Click += new System.EventHandler(this.bt_Add_Click);
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(824, 504);
			this.Controls.Add(this.bt_Add);
			this.Controls.Add(this.dt_docduedate);
			this.Controls.Add(this.dt_docdate);
			this.Controls.Add(this.tb_cardname);
			this.Controls.Add(this.lb_cardname);
			this.Controls.Add(this.tb_cardcode);
			this.Controls.Add(this.lb_cardcode);
			this.Controls.Add(this.lb_docduedate);
			this.Controls.Add(this.lb_docdate);
			this.Controls.Add(this.tb_docnum);
			this.Controls.Add(this.lb_docnum);
			this.Controls.Add(this.gv_rdr1);
			this.Name = "Form1";
			this.Text = "Form1";
			this.Load += new System.EventHandler(this.Form1_Load);
			((System.ComponentModel.ISupportInitialize)(this.gv_rdr1)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.DataGridView gv_rdr1;
		private System.Windows.Forms.Label lb_docnum;
		private System.Windows.Forms.TextBox tb_docnum;
		private System.Windows.Forms.Label lb_docdate;
		private System.Windows.Forms.Label lb_docduedate;
		private System.Windows.Forms.TextBox tb_cardname;
		private System.Windows.Forms.Label lb_cardname;
		private System.Windows.Forms.TextBox tb_cardcode;
		private System.Windows.Forms.Label lb_cardcode;
		private System.Windows.Forms.DateTimePicker dt_docdate;
		private System.Windows.Forms.DateTimePicker dt_docduedate;
		private System.Windows.Forms.Button bt_Add;
	}
}

