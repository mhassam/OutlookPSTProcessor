using System;
using System.Data;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Exception = System.Exception;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookPSTProcessor
{
    public partial class OutlookPSTProcessor : Form
    {
        private DataTable dataTable;

        public OutlookPSTProcessor()
        {
            InitializeComponent();
            InitializeDataTable();
        }

        private void InitializeDataTable()
        {
            dataTable = new DataTable();
            dataTable.Columns.Add("Sender");
            dataTable.Columns.Add("Receiver");
            dataTable.Columns.Add("Subject");
            dataTable.Columns.Add("Body");
            // Add more columns as needed

            dataGridView.DataSource = dataTable;
        }

        private void btnUploadOutlookPst_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Outlook PST Files|*.pst";
                openFileDialog.Title = "Select PST Files";
                openFileDialog.Multiselect = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    foreach (string selectedFilePath in openFileDialog.FileNames)
                    {
                        ProcessPst(selectedFilePath);
                    }
                }
            }
        }

        private void ProcessPst(string filePath)
        {
            try
            {
                //If Outlook is not installed on the system, won't able to use Microsoft.Office.Interop.Outlook
                //to directly process the PST file, as it relies on Outlook.
                //In this case, use a third-party library Aspose.Email to handle PST files without Outlook.

                // Uncomment it. In Case of Microsoft.Office.Interop.Outlook and Outlook Installed

                Outlook.Application outlookApp = new Outlook.Application();
                Outlook.NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");

                // Load the PST file
                Outlook.Folder folder = outlookNamespace.Folders.Add(filePath) as Outlook.Folder;

                // Process items in the folder
                foreach (Outlook.MailItem mailItem in folder.Items)
                {
                    DataRow row = dataTable.NewRow();
                    row["From"] = mailItem.SenderName;
                    row["To"] = mailItem.To;
                    row["Subject"] = mailItem.Subject;
                    row["Body"] = mailItem.Body;
                    // Populate other columns

                    dataTable.Rows.Add(row);
                }

                var emailAddresses = GetEmailAddressesFromFolder(folder);

                DisplayEmailAddresses(emailAddresses);

                // Release Outlook objects
                Marshal.ReleaseComObject(folder);
                Marshal.ReleaseComObject(outlookNamespace);
                Marshal.ReleaseComObject(outlookApp);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while processing PST file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string[] GetEmailAddressesFromFolder(Outlook.Folder folder)
        {
            var emailAddresses = new System.Collections.Generic.List<string>();

            foreach (object item in folder.Items)
            {
                if (item is Outlook.MailItem mailItem)
                {
                    var emailAddress = mailItem.SenderEmailAddress;
                    if (!string.IsNullOrEmpty(emailAddress))
                    {
                        emailAddresses.Add(emailAddress);
                    }
                }
            }

            // Recursively process subfolders
            foreach (Outlook.Folder subfolder in folder.Folders)
            {
                emailAddresses.AddRange(GetEmailAddressesFromFolder(subfolder));
            }

            return emailAddresses.ToArray();
        }

        private void DisplayEmailAddresses(string[] emailAddresses)
        {
            // Display email addresses in a ListBox or any other control
            listBoxEmailAddresses.Items.Clear();
            listBoxEmailAddresses.Items.AddRange(emailAddresses);
        }
    }
}
