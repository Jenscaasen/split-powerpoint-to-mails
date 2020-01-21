namespace SendSheetsViaMail
{
    partial class CSVInput
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txtSubject = new System.Windows.Forms.TextBox();
            this.txtBody = new System.Windows.Forms.TextBox();
            this.txtMapping = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(68, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Mail Subject:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 38);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Mail Body:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 128);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(133, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "Slide-To-Receipt Mapping:";
            // 
            // txtSubject
            // 
            this.txtSubject.Location = new System.Drawing.Point(88, 13);
            this.txtSubject.Name = "txtSubject";
            this.txtSubject.Size = new System.Drawing.Size(239, 20);
            this.txtSubject.TabIndex = 3;
            this.txtSubject.Text = "Ihre Daten";
            // 
            // txtBody
            // 
            this.txtBody.Location = new System.Drawing.Point(88, 40);
            this.txtBody.Multiline = true;
            this.txtBody.Name = "txtBody";
            this.txtBody.Size = new System.Drawing.Size(239, 72);
            this.txtBody.TabIndex = 4;
            this.txtBody.Text = "Sehr geehrter Empfänger,\r\nanbei Ihre Auswertung";
            // 
            // txtMapping
            // 
            this.txtMapping.Location = new System.Drawing.Point(15, 154);
            this.txtMapping.Multiline = true;
            this.txtMapping.Name = "txtMapping";
            this.txtMapping.Size = new System.Drawing.Size(312, 321);
            this.txtMapping.TabIndex = 5;
            this.txtMapping.Text = "1,2;Buchhaltung;test@test.de\r\n2;HR;test2@test.de";
            // 
            // CSVInput
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(353, 487);
            this.Controls.Add(this.txtMapping);
            this.Controls.Add(this.txtBody);
            this.Controls.Add(this.txtSubject);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "CSVInput";
            this.Text = "Slide to Mail Configuration";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        public System.Windows.Forms.TextBox txtMapping;
        public System.Windows.Forms.TextBox txtSubject;
        public System.Windows.Forms.TextBox txtBody;
    }
}