namespace TSCPPT_Addin
{
    partial class frmReplacewith
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
            this.lbltype = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.button3 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.btn_Correct = new System.Windows.Forms.Button();
            this.cmbReplace = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtError = new System.Windows.Forms.TextBox();
            this.lblFound = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // lbltype
            // 
            this.lbltype.AutoSize = true;
            this.lbltype.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Underline))), System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbltype.Location = new System.Drawing.Point(12, 9);
            this.lbltype.Name = "lbltype";
            this.lbltype.Size = new System.Drawing.Size(41, 13);
            this.lbltype.TabIndex = 1;
            this.lbltype.Text = "label2";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.button3);
            this.groupBox1.Controls.Add(this.button2);
            this.groupBox1.Controls.Add(this.btn_Correct);
            this.groupBox1.Controls.Add(this.cmbReplace);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.txtError);
            this.groupBox1.Controls.Add(this.lblFound);
            this.groupBox1.Location = new System.Drawing.Point(12, 25);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(350, 133);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            // 
            // button3
            // 
            this.button3.Font = new System.Drawing.Font("Modern No. 20", 8.999999F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button3.Location = new System.Drawing.Point(254, 89);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(76, 24);
            this.button3.TabIndex = 3;
            this.button3.Text = "Exit";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button2
            // 
            this.button2.Font = new System.Drawing.Font("Modern No. 20", 8.999999F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button2.Location = new System.Drawing.Point(172, 89);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(76, 24);
            this.button2.TabIndex = 3;
            this.button2.Text = "Discard";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // btn_Correct
            // 
            this.btn_Correct.Font = new System.Drawing.Font("Modern No. 20", 8.999999F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Correct.Location = new System.Drawing.Point(90, 89);
            this.btn_Correct.Name = "btn_Correct";
            this.btn_Correct.Size = new System.Drawing.Size(76, 24);
            this.btn_Correct.TabIndex = 3;
            this.btn_Correct.Text = "Correct";
            this.btn_Correct.UseVisualStyleBackColor = true;
            this.btn_Correct.Click += new System.EventHandler(this.btn_Correct_Click);
            // 
            // cmbReplace
            // 
            this.cmbReplace.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbReplace.FormattingEnabled = true;
            this.cmbReplace.Location = new System.Drawing.Point(118, 49);
            this.cmbReplace.Name = "cmbReplace";
            this.cmbReplace.Size = new System.Drawing.Size(198, 21);
            this.cmbReplace.TabIndex = 2;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(6, 55);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(102, 15);
            this.label1.TabIndex = 0;
            this.label1.Text = "String Replace";
            // 
            // txtError
            // 
            this.txtError.Location = new System.Drawing.Point(118, 16);
            this.txtError.Name = "txtError";
            this.txtError.Size = new System.Drawing.Size(198, 20);
            this.txtError.TabIndex = 1;
            // 
            // lblFound
            // 
            this.lblFound.AutoSize = true;
            this.lblFound.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFound.Location = new System.Drawing.Point(6, 19);
            this.lblFound.Name = "lblFound";
            this.lblFound.Size = new System.Drawing.Size(89, 15);
            this.lblFound.TabIndex = 0;
            this.lblFound.Text = "String Found";
            // 
            // frmReplacewith
            // 
            this.AcceptButton = this.btn_Correct;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(382, 165);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.lbltype);
            this.Name = "frmReplacewith";
            this.Text = "Replace With";
            this.Load += new System.EventHandler(this.frmReplacewith_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lbltype;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button btn_Correct;
        private System.Windows.Forms.ComboBox cmbReplace;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtError;
        private System.Windows.Forms.Label lblFound;
    }
}