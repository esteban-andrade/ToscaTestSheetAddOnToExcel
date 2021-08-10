namespace TestSheetAddOn
{
    partial class frmTDIName
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
            this.txtNameFormat = new System.Windows.Forms.TextBox();
            this.lblNameFormat = new System.Windows.Forms.Label();
            this.btnOK = new System.Windows.Forms.Button();
            this.lblHelpText = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtNameFormat
            // 
            this.txtNameFormat.Location = new System.Drawing.Point(132, 16);
            this.txtNameFormat.Name = "txtNameFormat";
            this.txtNameFormat.Size = new System.Drawing.Size(509, 20);
            this.txtNameFormat.TabIndex = 0;
            // 
            // lblNameFormat
            // 
            this.lblNameFormat.AutoSize = true;
            this.lblNameFormat.Location = new System.Drawing.Point(12, 19);
            this.lblNameFormat.Name = "lblNameFormat";
            this.lblNameFormat.Size = new System.Drawing.Size(114, 13);
            this.lblNameFormat.TabIndex = 1;
            this.lblNameFormat.Text = "Instance Name Format";
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(479, 70);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 2;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // lblHelpText
            // 
            this.lblHelpText.AutoSize = true;
            this.lblHelpText.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblHelpText.Location = new System.Drawing.Point(129, 39);
            this.lblHelpText.Name = "lblHelpText";
            this.lblHelpText.Size = new System.Drawing.Size(398, 17);
            this.lblHelpText.TabIndex = 3;
            this.lblHelpText.Text = "eg: Establish Vehicles_[Attribute1]_[Attribute2.SubAttribute2a]";
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(560, 70);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.TabIndex = 4;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // frmTDIName
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(662, 105);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.lblHelpText);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.lblNameFormat);
            this.Controls.Add(this.txtNameFormat);
            this.Name = "frmTDIName";
            this.Text = "TestSheet Instance Name";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtNameFormat;
        private System.Windows.Forms.Label lblNameFormat;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Label lblHelpText;
        private System.Windows.Forms.Button btnCancel;
    }
}