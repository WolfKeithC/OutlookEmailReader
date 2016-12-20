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
                        //cmbOutlookFolders.Items.Add(childFolder.Name);
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
                //cmbOutlookFolders.Items.Add(folder.Name);
                EnumerateFolders(folder);
            }

            cmbOutlookFolders.DisplayMember = "Value";
            cmbOutlookFolders.ValueMember = "Key";

            return;
            /*
            ConnectToExchangeServer();
            TimeSpan ts = new TimeSpan(0, -1, 0, 0);
            DateTime date = DateTime.Now.Add(ts);
            SearchFilter.IsGreaterThanOrEqualTo filter = new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeReceived, date);

            if (exchange != null && exchange.Url != null)
            {
                FindItemsResults<Item> findResults = exchange.FindItems(WellKnownFolderName.Inbox, filter, new ItemView(50));

                foreach (Item item in findResults)
                {

                    EmailMessage message = EmailMessage.Bind(exchange, item.Id);
                    ListViewItem listitem = new ListViewItem(new[]   
                    {  
                        message.DateTimeReceived.ToString(), message.From.Name.ToString() + "(" + message.From.Address.ToString() + ")", message.Subject, ((message.HasAttachments) ? "Yes" : "No"), message.Id.ToString()  
                    });
                    lstMsg.Items.Add(listitem);
                }
                if (findResults.Items.Count <= 0)
                {
                    lstMsg.Items.Add("No Messages found!!");
                }
            }
            else
            {
                lstMsg.Items.Add("Blank code objects!");
            }
            */
        }
        static bool RedirectionCallback(string url)
        {
            // Return true if the URL is an HTTPS URL.
            return url.ToLower().StartsWith("https://");
        }
        /*
        public void ConnectToExchangeServer()
        {

            lblMsg.Text = "Connecting to Exchange Server..";
            lblMsg.Refresh();
            try
            {
                exchange = new ExchangeService(ExchangeVersion.Exchange2013_SP1);
                exchange.Credentials = new WebCredentials("WOLFK036", "*****", "SWNA");

                string userEmailAddress = "keith.c.wolf@abc.com";
                // Look up the user's EWS endpoint by using Autodiscover.
                exchange.AutodiscoverUrl(userEmailAddress, RedirectionCallback);
                
                //exchange.AutodiscoverUrl("keith.c.wolf@abc.com");//, "https://autodiscover.disney.com/autodiscover/autodiscover.xml");

                lblMsg.Text = "Connected to Exchange Server : " + exchange.Url.Host;
                lblMsg.Refresh();

            }
            catch (Exception ex)
            {
                lblMsg.Text = "Error Connecting to Exchange Server!! " + ex.Message;
                lblMsg.Refresh();
            }
        }
        */

        private void btnAccessEmail_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Outlook.Application myApp = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.NameSpace mapiNameSpace = myApp.GetNamespace("MAPI");
            Microsoft.Office.Interop.Outlook.MAPIFolder myInbox = mapiNameSpace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox);
            if (myInbox.Items.Count > 0)
            {
                // Grab the Subject
                lblSubject.Text = ((Microsoft.Office.Interop.Outlook.MailItem)myInbox.Items[1]).Subject;
                //Grab the Attachment Name
                if (((Microsoft.Office.Interop.Outlook.MailItem)myInbox.Items[1]).Attachments.Count > 0)
                {
                    lblAttachmentName.Text = ((Microsoft.Office.Interop.Outlook.MailItem)myInbox.Items[1]).Attachments[1].FileName;
                }
                else
                {
                    lblAttachmentName.Text = "No Attachment";
                }
                // Grab the Body
                txtBody.Text = ((Microsoft.Office.Interop.Outlook.MailItem)myInbox.Items[1]).Body;
                // Sender Name
                lblSenderName.Text = ((Microsoft.Office.Interop.Outlook.MailItem)myInbox.Items[1]).SenderName;
                // Sender Email
                lblSenderEmail.Text = ((Microsoft.Office.Interop.Outlook.MailItem)myInbox.Items[1]).SenderEmailAddress;
                // Creation date
                lblCreationdate.Text = ((Microsoft.Office.Interop.Outlook.MailItem)myInbox.Items[1]).CreationTime.ToString();
            }
            else
            {
                MessageBox.Show("There are no emails in your Inbox.");
            }
        }

        private void btnLoad_Click(object sender, EventArgs e)   
        {  
            if (exchange != null)   
            {  
                if (lstMsg.Items.Count > 0)   
                {  
                    ListViewItem item = lstMsg.SelectedItems[0];  
  
                    if (item != null)   
                    {  
                        string msgid = item.SubItems[4].Text.ToString();  
                        EmailMessage message = EmailMessage.Bind(exchange, new ItemId(msgid));  
                        if (message.HasAttachments && message.Attachments[0] is FileAttachment)   
                        {  
                            FileAttachment fileAttachment = message.Attachments[0] as FileAttachment;  
                            //Change the below Path   
                            fileAttachment.Load(@"C:\\Users\\Admin\\Documents\\Visual Studio 2012\\Projects\\ReadMailFromExchangeServer\\ReadMailFromExchangeServer\\Attachments\\" + fileAttachment.Name);  
                            lblAttach.Text = "Attachment Downloaded : " + fileAttachment.Name;  
                        }   
                        else   
                        {  
                            MessageBox.Show("No Attachments found!!");  
                        }  
                    }   
                    else   
                    {  
                        MessageBox.Show("Please select a Message!!");  
                    }  
                }   
                else   
                {  
                    MessageBox.Show("Messages not loaded!!");  
                }  
                  
            }   
            else   
            {  
                MessageBox.Show("Not Connected to Mail Server!!");  
            }  
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
            pbMessages.Maximum = total;
            if (total > 0)
            {
                foreach (Microsoft.Office.Interop.Outlook.MailItem item in myInbox.Items)
                {
                    count++;
                    //pbMessages.Value = count;

                    //if (!item.UnRead) continue;

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

        private void OrderPizza()
        {
            Microsoft.Office.Interop.Outlook.Application myApp = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.MailItem mail = myApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
            mail.VotingOptions = "Cheese; Mushroom; Sausage; Combo; Veg Combo;";
            mail.Subject = "Pizza Order";
            mail.Display(false);
            mail.Save();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            lblEmailCount.Text = "Welcome";
        }

        private void lblAttachmentName_Click(object sender, EventArgs e)
        {

        }

        private void lblSubject_Click(object sender, EventArgs e)
        {

        }

        private void lblSenderName_Click(object sender, EventArgs e)
        {

        }

        private void lblSenderEmail_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void btnPizza_Click(object sender, EventArgs e)
        {
            OrderPizza();
        }
    }
}
