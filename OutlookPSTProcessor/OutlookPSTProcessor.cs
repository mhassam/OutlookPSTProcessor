using System;
using System.Data;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Aspose.Email.Storage.Pst;
using Microsoft.Office.Interop.Outlook;
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

                //PersonalStorage pst = PersonalStorage.FromFile(filePath);

                //// Access the root folder
                //FolderInfo rootFolder = pst.RootFolder;

                //// Process items in the folder
                //foreach (MessageInfo messageInfo in rootFolder.EnumerateMessages())
                //{
                //    DataRow row = dataTable.NewRow();
                //    row["Subject"] = messageInfo.Subject;
                //    row["Sender"] = messageInfo.SenderRepresentativeName;
                //    // Populate other columns

                //    dataTable.Rows.Add(row);
                //}

                //// Release Aspose.Email objects
                //pst.Dispose();


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
    }
}
