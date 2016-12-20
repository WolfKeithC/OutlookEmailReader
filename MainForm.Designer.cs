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
            this.btnLoad = new System.Windows.Forms.Button();
            this.lblAttach = new System.Windows.Forms.Label();
            this.lstMsg = new System.Windows.Forms.ListView();
            this.btnAccessEmail = new System.Windows.Forms.Button();
            this.lblSubject = new System.Windows.Forms.Label();
            this.lblAttachmentName = new System.Windows.Forms.Label();
            this.lblSenderName = new System.Windows.Forms.Label();
            this.lblSenderEmail = new System.Windows.Forms.Label();
            this.lblCreationdate = new System.Windows.Forms.Label();
            this.txtBody = new System.Windows.Forms.TextBox();
            this.cmbOutlookFolders = new System.Windows.Forms.ComboBox();
            this.btnLoadMessages = new System.Windows.Forms.Button();
            this.pbMessages = new System.Windows.Forms.ProgressBar();
            this.lblEmailCount = new System.Windows.Forms.Label();
            this.btnPizza = new System.Windows.Forms.Button();
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
            // btnLoad
            // 
            this.btnLoad.Location = new System.Drawing.Point(38, 401);
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.Size = new System.Drawing.Size(140, 36);
            this.btnLoad.TabIndex = 3;
            this.btnLoad.Text = "Load Attachment";
            this.btnLoad.UseVisualStyleBackColor = true;
            this.btnLoad.Click += new System.EventHandler(this.btnLoad_Click);
            // 
            // lblAttach
            // 
            this.lblAttach.AutoSize = true;
            this.lblAttach.Location = new System.Drawing.Point(193, 411);
            this.lblAttach.Name = "lblAttach";
            this.lblAttach.Size = new System.Drawing.Size(139, 13);
            this.lblAttach.TabIndex = 4;
            this.lblAttach.Text = "No Attachment downloaded";
            // 
            // lstMsg
            // 
            this.lstMsg.Location = new System.Drawing.Point(12, 124);
            this.lstMsg.Name = "lstMsg";
            this.lstMsg.Size = new System.Drawing.Size(938, 196);
            this.lstMsg.TabIndex = 5;
            this.lstMsg.UseCompatibleStateImageBehavior = false;
            // 
            // btnAccessEmail
            // 
            this.btnAccessEmail.Location = new System.Drawing.Point(12, 668);
            this.btnAccessEmail.Name = "btnAccessEmail";
            this.btnAccessEmail.Size = new System.Drawing.Size(140, 36);
            this.btnAccessEmail.TabIndex = 6;
            this.btnAccessEmail.Text = "Access Email";
            this.btnAccessEmail.UseVisualStyleBackColor = true;
            this.btnAccessEmail.Click += new System.EventHandler(this.btnAccessEmail_Click);
            // 
            // lblSubject
            // 
            this.lblSubject.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblSubject.Location = new System.Drawing.Point(186, 575);
            this.lblSubject.Name = "lblSubject";
            this.lblSubject.Size = new System.Drawing.Size(370, 23);
            this.lblSubject.TabIndex = 7;
            this.lblSubject.Text = "label1";
            this.lblSubject.Click += new System.EventHandler(this.lblSubject_Click);
            // 
            // lblAttachmentName
            // 
            this.lblAttachmentName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblAttachmentName.Location = new System.Drawing.Point(186, 609);
            this.lblAttachmentName.Name = "lblAttachmentName";
            this.lblAttachmentName.Size = new System.Drawing.Size(370, 23);
            this.lblAttachmentName.TabIndex = 8;
            this.lblAttachmentName.Text = "label1";
            this.lblAttachmentName.Click += new System.EventHandler(this.lblAttachmentName_Click);
            // 
            // lblSenderName
            // 
            this.lblSenderName.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblSenderName.Location = new System.Drawing.Point(186, 644);
            this.lblSenderName.Name = "lblSenderName";
            this.lblSenderName.Size = new System.Drawing.Size(370, 23);
            this.lblSenderName.TabIndex = 9;
            this.lblSenderName.Text = "label1";
            this.lblSenderName.Click += new System.EventHandler(this.lblSenderName_Click);
            // 
            // lblSenderEmail
            // 
            this.lblSenderEmail.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblSenderEmail.Location = new System.Drawing.Point(186, 678);
            this.lblSenderEmail.Name = "lblSenderEmail";
            this.lblSenderEmail.Size = new System.Drawing.Size(370, 23);
            this.lblSenderEmail.TabIndex = 10;
            this.lblSenderEmail.Text = "label1";
            this.lblSenderEmail.Click += new System.EventHandler(this.lblSenderEmail_Click);
            // 
            // lblCreationdate
            // 
            this.lblCreationdate.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblCreationdate.Location = new System.Drawing.Point(186, 712);
            this.lblCreationdate.Name = "lblCreationdate";
            this.lblCreationdate.Size = new System.Drawing.Size(370, 23);
            this.lblCreationdate.TabIndex = 11;
            this.lblCreationdate.Text = "label1";
            this.lblCreationdate.Click += new System.EventHandler(this.label1_Click);
            // 
            // txtBody
            // 
            this.txtBody.Location = new System.Drawing.Point(563, 575);
            this.txtBody.Multiline = true;
            this.txtBody.Name = "txtBody";
            this.txtBody.ReadOnly = true;
            this.txtBody.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtBody.Size = new System.Drawing.Size(323, 160);
            this.txtBody.TabIndex = 12;
            this.txtBody.WordWrap = false;
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
            // pbMessages
            // 
            this.pbMessages.Location = new System.Drawing.Point(783, 401);
            this.pbMessages.Name = "pbMessages";
            this.pbMessages.Size = new System.Drawing.Size(100, 23);
            this.pbMessages.TabIndex = 15;
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
            // btnPizza
            // 
            this.btnPizza.Location = new System.Drawing.Point(12, 763);
            this.btnPizza.Name = "btnPizza";
            this.btnPizza.Size = new System.Drawing.Size(140, 36);
            this.btnPizza.TabIndex = 17;
            this.btnPizza.Text = "Pizza";
            this.btnPizza.UseVisualStyleBackColor = true;
            this.btnPizza.Click += new System.EventHandler(this.btnPizza_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(973, 363);
            this.Controls.Add(this.btnPizza);
            this.Controls.Add(this.lblEmailCount);
            this.Controls.Add(this.pbMessages);
            this.Controls.Add(this.btnLoadMessages);
            this.Controls.Add(this.cmbOutlookFolders);
            this.Controls.Add(this.txtBody);
            this.Controls.Add(this.lblCreationdate);
            this.Controls.Add(this.lblSenderEmail);
            this.Controls.Add(this.lblSenderName);
            this.Controls.Add(this.lblAttachmentName);
            this.Controls.Add(this.lblSubject);
            this.Controls.Add(this.btnAccessEmail);
            this.Controls.Add(this.lstMsg);
            this.Controls.Add(this.lblAttach);
            this.Controls.Add(this.btnLoad);
            this.Controls.Add(this.btnRead);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "MRM Email Reader";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button btnRead;
        private System.Windows.Forms.Button btnLoad;
        private System.Windows.Forms.Label lblAttach;
        private System.Windows.Forms.ListView lstMsg;
        private System.Windows.Forms.Button btnAccessEmail;
        private System.Windows.Forms.Label lblSubject;
        private System.Windows.Forms.Label lblAttachmentName;
        private System.Windows.Forms.Label lblSenderName;
        private System.Windows.Forms.Label lblSenderEmail;
        private System.Windows.Forms.Label lblCreationdate;
        private System.Windows.Forms.TextBox txtBody;
        private System.Windows.Forms.ComboBox cmbOutlookFolders;
        private System.Windows.Forms.Button btnLoadMessages;
        private System.Windows.Forms.ProgressBar pbMessages;
        private System.Windows.Forms.Label lblEmailCount;
        private System.Windows.Forms.Button btnPizza;
    }
}

