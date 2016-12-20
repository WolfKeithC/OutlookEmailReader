using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MRMEmailReader
{
    public partial class MainForm : Form
    {
        ExchangeService exchange = null; 

        public MainForm()
        {
            InitializeComponent();
            lstMsg.Clear();
            lstMsg.View = View.Details;
            lstMsg.Columns.Add("Date", 150);
            lstMsg.Columns.Add("From", 250);
            lstMsg.Columns.Add("Subject", 400);
            lstMsg.Columns.Add("Has Attachment", 50);
            lstMsg.Columns.Add("Id", 100);
            lstMsg.FullRowSelect = true;  
        }

        // Uses recursion to enumerate Outlook subfolders.
        private void EnumerateFolders(Microsoft.Office.Interop.Outlook.MAPIFolder folder)
        {
            //cmbOutlookFolders.Items.Add(folder.Name);
            Microsoft.Office.Interop.Outlook.Folders childFolders =
                folder.Folders;
            if (childFolders.Count > 0)
            {
                foreach (Microsoft.Office.Interop.Outlook.MAPIFolder childFolder in childFolders)
                {
                    try
                    {
                        // Write the folder path.
                        Debug.WriteLine(childFolder.FolderPath);
                        KeyValuePair<string, string> folderItem = new KeyValuePair<string, string>(childFolder.EntryID, childFolder.Name);// { }
                        cmbOutlookFolders.Items.Add(folderItem);
                        // Call EnumerateFolders using childFolder.
                        EnumerateFolders(childFolder);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex.Message);
                    }
                }

            }
        }

        private void btnRead_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Outlook.Application myApp = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.NameSpace mapiNameSpace = myApp.GetNamespace("MAPI");
            Microsoft.Office.Interop.Outlook.MAPIFolder myInbox = mapiNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);

            Microsoft.Office.Interop.Outlook.MAPIFolder oPublicFolder = myInbox.Parent;

            cmbOutlookFolders.Items.Clear();

            foreach (Microsoft.Office.Interop.Outlook.MAPIFolder folder in oPublicFolder.Folders)
            {
                EnumerateFolders(folder);
            }

            cmbOutlookFolders.DisplayMember = "Value";
            cmbOutlookFolders.ValueMember = "Key";
        }

        private void btnLoadMessages_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Outlook.Application myApp = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.NameSpace mapiNameSpace = myApp.GetNamespace("MAPI");

            KeyValuePair<string, string> selectedItem = (KeyValuePair<string, string>)cmbOutlookFolders.SelectedItem;

            Microsoft.Office.Interop.Outlook.MailItem mail01 = myApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem) as Microsoft.Office.Interop.Outlook.MailItem;
            Microsoft.Office.Interop.Outlook.MailItem mail02 = myApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem) as Microsoft.Office.Interop.Outlook.MailItem;

            mail01.Subject = "Move Deliverable Between TOWs Errors";
            mail01.Body = "Hello Laura," + Environment.NewLine + Environment.NewLine + 
                        "Attached are this week’s emails." + Environment.NewLine + Environment.NewLine + 
                        "Thanks, Keith";

            mail02.Subject = "WorkOrders with No Deliverable";
            mail02.Body = "Hello Laura," + Environment.NewLine + Environment.NewLine +
                        "Attached are this week’s emails." + Environment.NewLine + Environment.NewLine +
                        "Thanks, Keith";

            Microsoft.Office.Interop.Outlook.AddressEntry currentUser = myApp.Session.CurrentUser.AddressEntry;
            Microsoft.Office.Interop.Outlook.ExchangeUser manager = currentUser.GetExchangeUser().GetExchangeUserManager();
            // Add recipient using display name, alias, or smtp address
            //mail.Recipients.Add(manager.PrimarySmtpAddress);
            mail01.Recipients.ResolveAll();
            mail02.Recipients.ResolveAll();

            Microsoft.Office.Interop.Outlook.MAPIFolder myInbox = mapiNameSpace.GetFolderFromID(selectedItem.Key);
            //cmbOutlookFolders.SelectedItem.ToString());// GetDefaultFolder(Microsoft.Office.Interop.Outlook.O.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
            int total = myInbox.Items.Count, count = 0, xytCount = 0, unread = 0;

            if (total > 0)
            {
                foreach (Microsoft.Office.Interop.Outlook.MailItem item in myInbox.Items)
                {
                    count++;
                    unread++;
                    ListViewItem listitem = new ListViewItem(new[]   
                    {  
                        item.ReceivedTime.ToString(), item.SenderEmailAddress + "(" + item.Sender.Address.ToString() + ")", item.Subject, ((item.Attachments.Count > 0) ? "Yes" : "No"), item.EntryID.ToString()  
                    });

                    if (item.Body.Contains("Attempt to move a Deliverable between TOWs"))
                    {
                        listitem.BackColor = Color.Plum;
                        xytCount++;

                        mail01.Attachments.Add(item, Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                    }

                    if (item.Body.Contains("Failure Message Deliverable with Id"))
                    {
                        listitem.BackColor = Color.GreenYellow;
                        xytCount++;

                        mail02.Attachments.Add(item, Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, Type.Missing, Type.Missing);
                    }

                    lstMsg.Items.Add(listitem);
                }

                lblEmailCount.Text = string.Format("Total Emails: {0}, Emails Need Attention: {1}, Unread: {2} ", total, xytCount, unread);
            }

            mail01.Save();
            mail02.Save();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            lblEmailCount.Text = "Welcome";
        }
    }
}
