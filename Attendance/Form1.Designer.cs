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
            this.button1 = new System.Windows.Forms.Button();
            this.btn_report_extended = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // btn_file
            // 
            this.btn_file.Location = new System.Drawing.Point(18, 16);
            this.btn_file.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btn_file.Name = "btn_file";
            this.btn_file.Size = new System.Drawing.Size(261, 63);
            this.btn_file.TabIndex = 0;
            this.btn_file.Text = "Choose file";
            this.btn_file.UseVisualStyleBackColor = true;
            this.btn_file.Click += new System.EventHandler(this.btn_file_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(288, 16);
            this.button1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(261, 63);
            this.button1.TabIndex = 1;
            this.button1.Text = "Report";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.btn_report_Click);
            // 
            // btn_report_extended
            // 
            this.btn_report_extended.Location = new System.Drawing.Point(558, 16);
            this.btn_report_extended.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btn_report_extended.Name = "btn_report_extended";
            this.btn_report_extended.Size = new System.Drawing.Size(261, 63);
            this.btn_report_extended.TabIndex = 1;
            this.btn_report_extended.Text = "Report Extended";
            this.btn_report_extended.UseVisualStyleBackColor = true;
            this.btn_report_extended.Click += new System.EventHandler(this.btn_report_extended_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.ClientSize = new System.Drawing.Size(834, 92);
            this.Controls.Add(this.btn_report_extended);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btn_file);
            this.Font = new System.Drawing.Font("Modern No. 20", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "Attendance Report";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button btn_file;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button btn_report_extended;
    }
}

