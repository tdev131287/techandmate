namespace TSCPPT_Addin
{
    partial class frmReviewformat
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
            this.label1 = new System.Windows.Forms.Label();
            this.cmb_RType = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btn_Review = new System.Windows.Forms.Button();
            this.btn_Cancel = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.rb_CorrectAll = new System.Windows.Forms.RadioButton();
            this.rb_ReviewCorrect = new System.Windows.Forms.RadioButton();
            this.rb_Review = new System.Windows.Forms.RadioButton();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arial", 9.75F, ((System.Drawing.FontStyle)(((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic) 
                | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(12, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(124, 16);
            this.label2.TabIndex = 1;
            this.label2.Text = "Formatting Review";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(17, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(119, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Review following slides:";
            // 
            // cmb_RType
            // 
            this.cmb_RType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmb_RType.FormattingEnabled = true;
            this.cmb_RType.Location = new System.Drawing.Point(151, 19);
            this.cmb_RType.Name = "cmb_RType";
            this.cmb_RType.Size = new System.Drawing.Size(212, 21);
            this.cmb_RType.TabIndex = 1;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Arial", 9F, ((System.Drawing.FontStyle)(((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic) 
                | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(6, 58);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(149, 15);
            this.label3.TabIndex = 3;
            this.label3.Text = "Select a Review Method:";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btn_Review);
            this.groupBox1.Controls.Add(this.btn_Cancel);
            this.groupBox1.Controls.Add(this.groupBox2);
            this.groupBox1.Controls.Add(this.cmb_RType);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(12, 29);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(435, 274);
            this.groupBox1.TabIndex = 4;
            this.groupBox1.TabStop = false;
            // 
            // btn_Review
            // 
            this.btn_Review.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Review.Location = new System.Drawing.Point(182, 238);
            this.btn_Review.Name = "btn_Review";
            this.btn_Review.Size = new System.Drawing.Size(113, 28);
            this.btn_Review.TabIndex = 5;
            this.btn_Review.Text = "Start Review";
            this.btn_Review.UseVisualStyleBackColor = true;
            this.btn_Review.Click += new System.EventHandler(this.btn_Review_Click);
            // 
            // btn_Cancel
            // 
            this.btn_Cancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Cancel.Location = new System.Drawing.Point(318, 238);
            this.btn_Cancel.Name = "btn_Cancel";
            this.btn_Cancel.Size = new System.Drawing.Size(100, 28);
            this.btn_Cancel.TabIndex = 6;
            this.btn_Cancel.Text = "Cancel";
            this.btn_Cancel.UseVisualStyleBackColor = true;
            this.btn_Cancel.Click += new System.EventHandler(this.btn_Cancel_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.rb_CorrectAll);
            this.groupBox2.Controls.Add(this.rb_ReviewCorrect);
            this.groupBox2.Controls.Add(this.rb_Review);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Location = new System.Drawing.Point(9, 74);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(409, 158);
            this.groupBox2.TabIndex = 4;
            this.groupBox2.TabStop = false;
            // 
            // rb_CorrectAll
            // 
            this.rb_CorrectAll.AutoSize = true;
            this.rb_CorrectAll.Location = new System.Drawing.Point(6, 116);
            this.rb_CorrectAll.Name = "rb_CorrectAll";
            this.rb_CorrectAll.Size = new System.Drawing.Size(73, 17);
            this.rb_CorrectAll.TabIndex = 4;
            this.rb_CorrectAll.TabStop = true;
            this.rb_CorrectAll.Text = "Correct All";
            this.rb_CorrectAll.UseVisualStyleBackColor = true;
            // 
            // rb_ReviewCorrect
            // 
            this.rb_ReviewCorrect.AutoSize = true;
            this.rb_ReviewCorrect.Location = new System.Drawing.Point(6, 66);
            this.rb_ReviewCorrect.Name = "rb_ReviewCorrect";
            this.rb_ReviewCorrect.Size = new System.Drawing.Size(119, 17);
            this.rb_ReviewCorrect.TabIndex = 3;
            this.rb_ReviewCorrect.TabStop = true;
            this.rb_ReviewCorrect.Text = "Review and Correct";
            this.rb_ReviewCorrect.UseVisualStyleBackColor = true;
            // 
            // rb_Review
            // 
            this.rb_Review.AutoSize = true;
            this.rb_Review.Location = new System.Drawing.Point(6, 19);
            this.rb_Review.Name = "rb_Review";
            this.rb_Review.Size = new System.Drawing.Size(85, 17);
            this.rb_Review.TabIndex = 2;
            this.rb_Review.TabStop = true;
            this.rb_Review.Text = "Only Review";
            this.rb_Review.UseVisualStyleBackColor = true;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(26, 136);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(347, 13);
            this.label6.TabIndex = 1;
            this.label6.Text = "Reviews for formatting errors and corrects all errors without user concent";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(26, 86);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(357, 13);
            this.label5.TabIndex = 1;
            this.label5.Text = "Reviews for formatting errors and seeks user concent to correct each error";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(26, 39);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(306, 13);
            this.label4.TabIndex = 1;
            this.label4.Text = "Reviews for formatting errors and puts a comment for each error";
            // 
            // frmReviewformat
            // 
            this.AcceptButton = this.btn_Review;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(455, 305);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label2);
            this.ImeMode = System.Windows.Forms.ImeMode.On;
            this.KeyPreview = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmReviewformat";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Review format";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmReviewformat_FormClosed);
            this.Load += new System.EventHandler(this.frmReviewformat_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmReviewformat_KeyDown);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cmb_RType;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btn_Review;
        private System.Windows.Forms.Button btn_Cancel;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.RadioButton rb_CorrectAll;
        private System.Windows.Forms.RadioButton rb_ReviewCorrect;
        private System.Windows.Forms.RadioButton rb_Review;
    }
}