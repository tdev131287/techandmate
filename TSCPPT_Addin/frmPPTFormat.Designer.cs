namespace TSCPPT_Addin
{
    partial class frmPPTFormat
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
            this.rbselectedslide = new System.Windows.Forms.RadioButton();
            this.rbselectedppt = new System.Windows.Forms.RadioButton();
            this.rbtn_selectfolder = new System.Windows.Forms.RadioButton();
            this.btn_Browse = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btn_Cancel = new System.Windows.Forms.Button();
            this.btn_Submit = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.fileGridView = new System.Windows.Forms.DataGridView();
            this.pptname = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.pptStatus = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.pptReview = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.fileGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // rbselectedslide
            // 
            this.rbselectedslide.AutoSize = true;
            this.rbselectedslide.Location = new System.Drawing.Point(9, 19);
            this.rbselectedslide.Name = "rbselectedslide";
            this.rbselectedslide.Size = new System.Drawing.Size(117, 17);
            this.rbselectedslide.TabIndex = 0;
            this.rbselectedslide.TabStop = true;
            this.rbselectedslide.Text = "Only Selected Slide";
            this.rbselectedslide.UseVisualStyleBackColor = true;
            this.rbselectedslide.CheckedChanged += new System.EventHandler(this.rbselectedslide_CheckedChanged);
            // 
            // rbselectedppt
            // 
            this.rbselectedppt.AutoSize = true;
            this.rbselectedppt.Location = new System.Drawing.Point(9, 43);
            this.rbselectedppt.Name = "rbselectedppt";
            this.rbselectedppt.Size = new System.Drawing.Size(119, 17);
            this.rbselectedppt.TabIndex = 0;
            this.rbselectedppt.TabStop = true;
            this.rbselectedppt.Text = "Current Power Point";
            this.rbselectedppt.UseVisualStyleBackColor = true;
            this.rbselectedppt.CheckedChanged += new System.EventHandler(this.rbselectedppt_CheckedChanged);
            // 
            // rbtn_selectfolder
            // 
            this.rbtn_selectfolder.AutoSize = true;
            this.rbtn_selectfolder.Location = new System.Drawing.Point(9, 67);
            this.rbtn_selectfolder.Name = "rbtn_selectfolder";
            this.rbtn_selectfolder.Size = new System.Drawing.Size(111, 17);
            this.rbtn_selectfolder.TabIndex = 0;
            this.rbtn_selectfolder.TabStop = true;
            this.rbtn_selectfolder.Text = "Select PPT Folder";
            this.rbtn_selectfolder.UseVisualStyleBackColor = true;
            this.rbtn_selectfolder.CheckedChanged += new System.EventHandler(this.rbtn_selectfolder_CheckedChanged);
            // 
            // btn_Browse
            // 
            this.btn_Browse.Location = new System.Drawing.Point(274, 61);
            this.btn_Browse.Name = "btn_Browse";
            this.btn_Browse.Size = new System.Drawing.Size(90, 25);
            this.btn_Browse.TabIndex = 1;
            this.btn_Browse.Text = "Browse";
            this.btn_Browse.UseVisualStyleBackColor = true;
            this.btn_Browse.Click += new System.EventHandler(this.btn_Browse_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btn_Cancel);
            this.groupBox2.Controls.Add(this.btn_Submit);
            this.groupBox2.Location = new System.Drawing.Point(187, 9);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(177, 46);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            // 
            // btn_Cancel
            // 
            this.btn_Cancel.Location = new System.Drawing.Point(91, 12);
            this.btn_Cancel.Name = "btn_Cancel";
            this.btn_Cancel.Size = new System.Drawing.Size(75, 27);
            this.btn_Cancel.TabIndex = 0;
            this.btn_Cancel.Text = "Cancel";
            this.btn_Cancel.UseVisualStyleBackColor = true;
            this.btn_Cancel.Click += new System.EventHandler(this.btn_Cancel_Click);
            // 
            // btn_Submit
            // 
            this.btn_Submit.Location = new System.Drawing.Point(10, 12);
            this.btn_Submit.Name = "btn_Submit";
            this.btn_Submit.Size = new System.Drawing.Size(75, 27);
            this.btn_Submit.TabIndex = 0;
            this.btn_Submit.Text = "Submit";
            this.btn_Submit.UseVisualStyleBackColor = true;
            this.btn_Submit.Click += new System.EventHandler(this.btn_Submit_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.groupBox2);
            this.groupBox1.Controls.Add(this.btn_Browse);
            this.groupBox1.Controls.Add(this.rbtn_selectfolder);
            this.groupBox1.Controls.Add(this.rbselectedppt);
            this.groupBox1.Controls.Add(this.rbselectedslide);
            this.groupBox1.Location = new System.Drawing.Point(3, 3);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(373, 102);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // fileGridView
            // 
            this.fileGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.fileGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.pptname,
            this.pptStatus,
            this.pptReview});
            this.fileGridView.Location = new System.Drawing.Point(3, 112);
            this.fileGridView.Name = "fileGridView";
            this.fileGridView.Size = new System.Drawing.Size(373, 276);
            this.fileGridView.TabIndex = 1;
            // 
            // pptname
            // 
            this.pptname.HeaderText = "PPT Name";
            this.pptname.MinimumWidth = 100;
            this.pptname.Name = "pptname";
            this.pptname.ReadOnly = true;
            this.pptname.Width = 230;
            // 
            // pptStatus
            // 
            this.pptStatus.HeaderText = "Review";
            this.pptStatus.Name = "pptStatus";
            this.pptStatus.ReadOnly = true;
            this.pptStatus.Width = 50;
            // 
            // pptReview
            // 
            this.pptReview.HeaderText = "Select";
            this.pptReview.Name = "pptReview";
            this.pptReview.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.pptReview.Width = 50;
            // 
            // frmPPTFormat
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(373, 397);
            this.Controls.Add(this.fileGridView);
            this.Controls.Add(this.groupBox1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmPPTFormat";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Format Slide ";
            this.Load += new System.EventHandler(this.frmPPTFormat_Load);
            this.groupBox2.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.fileGridView)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.RadioButton rbselectedslide;
        private System.Windows.Forms.RadioButton rbselectedppt;
        private System.Windows.Forms.RadioButton rbtn_selectfolder;
        private System.Windows.Forms.Button btn_Browse;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btn_Cancel;
        private System.Windows.Forms.Button btn_Submit;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.DataGridView fileGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn pptname;
        private System.Windows.Forms.DataGridViewTextBoxColumn pptStatus;
        private System.Windows.Forms.DataGridViewCheckBoxColumn pptReview;
    }
}