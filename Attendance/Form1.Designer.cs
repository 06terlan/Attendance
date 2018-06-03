namespace Attendance
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
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.btn_file = new System.Windows.Forms.Button();
            this.btn_report = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // btn_file
            // 
            this.btn_file.Location = new System.Drawing.Point(12, 12);
            this.btn_file.Name = "btn_file";
            this.btn_file.Size = new System.Drawing.Size(174, 48);
            this.btn_file.TabIndex = 0;
            this.btn_file.Text = "Choose file";
            this.btn_file.UseVisualStyleBackColor = true;
            this.btn_file.Click += new System.EventHandler(this.btn_file_Click);
            // 
            // btn_report
            // 
            this.btn_report.Location = new System.Drawing.Point(192, 12);
            this.btn_report.Name = "btn_report";
            this.btn_report.Size = new System.Drawing.Size(174, 48);
            this.btn_report.TabIndex = 1;
            this.btn_report.Text = "Report";
            this.btn_report.UseVisualStyleBackColor = true;
            this.btn_report.Click += new System.EventHandler(this.btn_report_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.ClientSize = new System.Drawing.Size(524, 70);
            this.Controls.Add(this.btn_report);
            this.Controls.Add(this.btn_file);
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "Report";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button btn_file;
        private System.Windows.Forms.Button btn_report;
    }
}

