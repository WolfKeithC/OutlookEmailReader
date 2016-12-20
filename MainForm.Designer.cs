namespace MRMEmailReader
{
    partial class MainForm
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
            this.btnRead = new System.Windows.Forms.Button();
            this.lstMsg = new System.Windows.Forms.ListView();
            this.cmbOutlookFolders = new System.Windows.Forms.ComboBox();
            this.btnLoadMessages = new System.Windows.Forms.Button();
            this.lblEmailCount = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnRead
            // 
            this.btnRead.Location = new System.Drawing.Point(12, 13);
            this.btnRead.Name = "btnRead";
            this.btnRead.Size = new System.Drawing.Size(246, 36);
            this.btnRead.TabIndex = 1;
            this.btnRead.Text = "Load Folders";
            this.btnRead.UseVisualStyleBackColor = true;
            this.btnRead.Click += new System.EventHandler(this.btnRead_Click);
            // 
            // lstMsg
            // 
            this.lstMsg.Location = new System.Drawing.Point(12, 124);
            this.lstMsg.Name = "lstMsg";
            this.lstMsg.Size = new System.Drawing.Size(938, 196);
            this.lstMsg.TabIndex = 5;
            this.lstMsg.UseCompatibleStateImageBehavior = false;
            // 
            // cmbOutlookFolders
            // 
            this.cmbOutlookFolders.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbOutlookFolders.FormattingEnabled = true;
            this.cmbOutlookFolders.Location = new System.Drawing.Point(12, 55);
            this.cmbOutlookFolders.Name = "cmbOutlookFolders";
            this.cmbOutlookFolders.Size = new System.Drawing.Size(246, 21);
            this.cmbOutlookFolders.TabIndex = 13;
            // 
            // btnLoadMessages
            // 
            this.btnLoadMessages.Location = new System.Drawing.Point(12, 82);
            this.btnLoadMessages.Name = "btnLoadMessages";
            this.btnLoadMessages.Size = new System.Drawing.Size(246, 36);
            this.btnLoadMessages.TabIndex = 14;
            this.btnLoadMessages.Text = "Load Messages";
            this.btnLoadMessages.UseVisualStyleBackColor = true;
            this.btnLoadMessages.Click += new System.EventHandler(this.btnLoadMessages_Click);
            // 
            // lblEmailCount
            // 
            this.lblEmailCount.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblEmailCount.Location = new System.Drawing.Point(12, 323);
            this.lblEmailCount.Name = "lblEmailCount";
            this.lblEmailCount.Size = new System.Drawing.Size(938, 23);
            this.lblEmailCount.TabIndex = 16;
            this.lblEmailCount.Text = "label1";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(973, 360);
            this.Controls.Add(this.lblEmailCount);
            this.Controls.Add(this.btnLoadMessages);
            this.Controls.Add(this.cmbOutlookFolders);
            this.Controls.Add(this.lstMsg);
            this.Controls.Add(this.btnRead);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "MRM Email Reader";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button btnRead;
        private System.Windows.Forms.ListView lstMsg;
        private System.Windows.Forms.ComboBox cmbOutlookFolders;
        private System.Windows.Forms.Button btnLoadMessages;
        private System.Windows.Forms.Label lblEmailCount;
    }
}

