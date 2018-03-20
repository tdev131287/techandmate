namespace TSCPPT_Addin
{
    partial class frmChartcalc
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
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btn_Cancel = new System.Windows.Forms.Button();
            this.btn_Caption = new System.Windows.Forms.Button();
            this.lbl_Value = new System.Windows.Forms.Label();
            this.lbl_type = new System.Windows.Forms.Label();
            this.txt_period = new System.Windows.Forms.TextBox();
            this.chk_Period = new System.Windows.Forms.CheckBox();
            this.cmb_endDate = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cmb_stDate = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.cmb_calcType = new System.Windows.Forms.ComboBox();
            this.lbl_calctype = new System.Windows.Forms.Label();
            this.cmb_Series = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Modern No. 20", 9.75F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(12, 20);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(99, 15);
            this.label2.TabIndex = 1;
            this.label2.Text = "CAGR Calculator";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btn_Cancel);
            this.groupBox1.Controls.Add(this.btn_Caption);
            this.groupBox1.Controls.Add(this.lbl_Value);
            this.groupBox1.Controls.Add(this.lbl_type);
            this.groupBox1.Controls.Add(this.txt_period);
            this.groupBox1.Controls.Add(this.chk_Period);
            this.groupBox1.Controls.Add(this.cmb_endDate);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.cmb_stDate);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.cmb_calcType);
            this.groupBox1.Controls.Add(this.lbl_calctype);
            this.groupBox1.Controls.Add(this.cmb_Series);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(12, 36);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(348, 243);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            // 
            // btn_Cancel
            // 
            this.btn_Cancel.Location = new System.Drawing.Point(237, 205);
            this.btn_Cancel.Name = "btn_Cancel";
            this.btn_Cancel.Size = new System.Drawing.Size(65, 23);
            this.btn_Cancel.TabIndex = 9;
            this.btn_Cancel.Text = "Cancel";
            this.btn_Cancel.UseVisualStyleBackColor = true;
            this.btn_Cancel.Click += new System.EventHandler(this.btn_Cancel_Click);
            // 
            // btn_Caption
            // 
            this.btn_Caption.Location = new System.Drawing.Point(166, 205);
            this.btn_Caption.Name = "btn_Caption";
            this.btn_Caption.Size = new System.Drawing.Size(65, 23);
            this.btn_Caption.TabIndex = 9;
            this.btn_Caption.Text = "Caption";
            this.btn_Caption.UseVisualStyleBackColor = true;
            this.btn_Caption.Click += new System.EventHandler(this.btn_Caption_Click);
            // 
            // lbl_Value
            // 
            this.lbl_Value.AutoSize = true;
            this.lbl_Value.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_Value.Location = new System.Drawing.Point(126, 184);
            this.lbl_Value.Name = "lbl_Value";
            this.lbl_Value.Size = new System.Drawing.Size(24, 20);
            this.lbl_Value.TabIndex = 8;
            this.lbl_Value.Text = "%";
            // 
            // lbl_type
            // 
            this.lbl_type.AutoSize = true;
            this.lbl_type.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_type.Location = new System.Drawing.Point(40, 184);
            this.lbl_type.Name = "lbl_type";
            this.lbl_type.Size = new System.Drawing.Size(80, 20);
            this.lbl_type.TabIndex = 8;
            this.lbl_type.Text = "CAGR  =";
            // 
            // txt_period
            // 
            this.txt_period.Location = new System.Drawing.Point(111, 146);
            this.txt_period.Name = "txt_period";
            this.txt_period.Size = new System.Drawing.Size(191, 20);
            this.txt_period.TabIndex = 7;
            this.txt_period.TextChanged += new System.EventHandler(this.txt_period_TextChanged);
            // 
            // chk_Period
            // 
            this.chk_Period.AutoSize = true;
            this.chk_Period.Location = new System.Drawing.Point(20, 149);
            this.chk_Period.Name = "chk_Period";
            this.chk_Period.Size = new System.Drawing.Size(66, 17);
            this.chk_Period.TabIndex = 6;
            this.chk_Period.Text = "# Period";
            this.chk_Period.UseVisualStyleBackColor = true;
            this.chk_Period.CheckedChanged += new System.EventHandler(this.chk_Period_CheckedChanged);
            // 
            // cmb_endDate
            // 
            this.cmb_endDate.FormattingEnabled = true;
            this.cmb_endDate.Location = new System.Drawing.Point(111, 108);
            this.cmb_endDate.Name = "cmb_endDate";
            this.cmb_endDate.Size = new System.Drawing.Size(191, 21);
            this.cmb_endDate.TabIndex = 5;
            this.cmb_endDate.SelectedIndexChanged += new System.EventHandler(this.cmb_endDate_SelectedIndexChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(17, 116);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(59, 13);
            this.label4.TabIndex = 4;
            this.label4.Text = "End Period";
            // 
            // cmb_stDate
            // 
            this.cmb_stDate.FormattingEnabled = true;
            this.cmb_stDate.Location = new System.Drawing.Point(111, 74);
            this.cmb_stDate.Name = "cmb_stDate";
            this.cmb_stDate.Size = new System.Drawing.Size(191, 21);
            this.cmb_stDate.TabIndex = 3;
            this.cmb_stDate.SelectedIndexChanged += new System.EventHandler(this.cmb_stDate_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(17, 82);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(62, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "Start Period";
            // 
            // cmb_calcType
            // 
            this.cmb_calcType.FormattingEnabled = true;
            this.cmb_calcType.Location = new System.Drawing.Point(111, 11);
            this.cmb_calcType.Name = "cmb_calcType";
            this.cmb_calcType.Size = new System.Drawing.Size(191, 21);
            this.cmb_calcType.TabIndex = 1;
            this.cmb_calcType.SelectedIndexChanged += new System.EventHandler(this.cmb_calcType_SelectedIndexChanged);
            // 
            // lbl_calctype
            // 
            this.lbl_calctype.AutoSize = true;
            this.lbl_calctype.Location = new System.Drawing.Point(17, 19);
            this.lbl_calctype.Name = "lbl_calctype";
            this.lbl_calctype.Size = new System.Drawing.Size(72, 13);
            this.lbl_calctype.TabIndex = 0;
            this.lbl_calctype.Text = "CAGR/AAGR";
            // 
            // cmb_Series
            // 
            this.cmb_Series.FormattingEnabled = true;
            this.cmb_Series.Location = new System.Drawing.Point(111, 43);
            this.cmb_Series.Name = "cmb_Series";
            this.cmb_Series.Size = new System.Drawing.Size(191, 21);
            this.cmb_Series.TabIndex = 1;
            this.cmb_Series.SelectedIndexChanged += new System.EventHandler(this.cmb_Series_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(17, 51);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(36, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Series";
            // 
            // frmChartcalc
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(365, 287);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label2);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmChartcalc";
            this.Text = "Growth Rates";
            this.Load += new System.EventHandler(this.frmChartcalc_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btn_Cancel;
        private System.Windows.Forms.Button btn_Caption;
        private System.Windows.Forms.Label lbl_Value;
        private System.Windows.Forms.Label lbl_type;
        private System.Windows.Forms.TextBox txt_period;
        private System.Windows.Forms.CheckBox chk_Period;
        private System.Windows.Forms.ComboBox cmb_endDate;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cmb_stDate;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cmb_Series;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cmb_calcType;
        private System.Windows.Forms.Label lbl_calctype;
    }
}