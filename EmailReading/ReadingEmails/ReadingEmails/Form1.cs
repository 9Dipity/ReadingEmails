using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ReadingEmails
{
    public partial class QueryMagnet : Form
    {
        private DateTime startDate;
        private DateTime endDate;

        public QueryMagnet()
        {
            InitializeComponent();
        }

        private void readEmails(object sender, EventArgs e)
        {
            // Set start and end dates based on user input
            startDate = startDatePicker.Value.Date;
            endDate = endDatePicker.Value;

            // Specify the file path for the Excel workbook
            //string filePath = "C:\\Users\\Miuskaste\\Desktop\\EmailReading\\OutputFile.xlsx";
            try
            {
                // Create Excel workbook with ClosedXML
                using (var workbook = new XLWorkbook())
                {
                    // Add a worksheet to the workbook
                    var worksheet = workbook.Worksheets.Add("Statistics");

                    // Set column headers
                    worksheet.Cell(1, 1).Value = "Mailbox";
                    worksheet.Cell(1, 2).Value = "Time received";
                    worksheet.Cell(1, 3).Value = "Time replied";
                    worksheet.Cell(1, 4).Value = "Email";
                    worksheet.Cell(1, 5).Value = "Client";
                    worksheet.Cell(1, 6).Value = "Email Title";
                    worksheet.Cell(1, 7).Value = "Category";
                    worksheet.Cell(1, 8).Value = "Contents";

                    // Read emails using Outlook Interop API
                    Outlook.Application outlookApp = new Outlook.Application();
                    Outlook.NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");

                    // Get the root folder
                    Outlook.MAPIFolder rootFolder = outlookNamespace.Folders[1];
                    string folderName = emailFolder.Text;

                    // Navigate through the folders to find the desired folder
                    Outlook.MAPIFolder inbox = GetFolderByName(rootFolder, folderName);

                    int row = 2; // Start from row 2, leaving the first row for headers
                    int emailCounter = 0;

                    foreach (object item in inbox.Items)
                    {
                        if (item is Outlook.MailItem)
                        {
                            Outlook.MailItem email = (Outlook.MailItem)item;

                            //statusTextBox.AppendText($"");
                            statusTextBox.Text = $"Processing email {emailCounter + 1}...";

                            LogEmailDetailsToFile(email); //Logging ALL emails it processes with some information into a txt file

                            // Check if email is within the specified date range
                            if (email.CreationTime >= startDate && email.ReceivedTime <= endDate)
                            {
                                // Extract information from the email
                                string mailbox = emailFolder.Text;
                                DateTime timeReceived = email.ReceivedTime;
                                string[] parts = email.SenderEmailAddress.Split('@');
                                string domain = parts.Length > 1 ? parts[1] : "";
                                string client = domain.Split('.')[0];

                                string contents = email.Body;

                                string conversationID = email.ConversationID;

                                // Write to Excel
                                worksheet.Cell(row, 1).Value = mailbox;
                                worksheet.Cell(row, 2).Value = timeReceived.ToString();
                                worksheet.Cell(row, 4).Value = email.SenderEmailAddress;
                                worksheet.Cell(row, 5).Value = client;
                                worksheet.Cell(row, 6).Value = email.ConversationTopic;
                                worksheet.Cell(row, 8).Value = email.Body;

                                // Search for related emails in SentItems folder
                                Outlook.MAPIFolder sentItemsFolder = outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail);

                                foreach (object sentItem in sentItemsFolder.Items)
                                {
                                    if (sentItem is Outlook.MailItem)
                                    {
                                        Outlook.MailItem sentEmail = (Outlook.MailItem)sentItem;

                                        // Check if the conversation ID matches
                                        if (sentEmail.ConversationID == conversationID)
                                        {
                                            string category = GetCategoryFromKeywords(sentEmail.Body);
                                            DateTime timeReplied = sentEmail.ReceivedTime; // Assuming the reply time is the same as received time

                                            // Extract information from the sent email
                                            worksheet.Cell(row, 3).Value = timeReplied;
                                            worksheet.Cell(row, 7).Value = category;

                                            break; // Exit the inner loop once a match is found
                                        }
                                    }
                                }

                                row++;
                            }
                            emailCounter++;

                        }
                    }
                    statusTextBox.Text = $"Finished processing. Total emails: {emailCounter}";

                    using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                    {
                        saveFileDialog.Filter = "Excel Files|*.xlsx;*.xls|All Files|*.*";
                        saveFileDialog.Title = "Save Excel File";
                        saveFileDialog.FileName = "OutputFile"; // Default file name

                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            // Get the selected file path
                            string filePath = saveFileDialog.FileName;

                            // Save the workbook to the chosen file path
                            workbook.SaveAs(filePath);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                if (ex is ArgumentException argEx && argEx.Message.Contains("Empty extension is not supported"))
                {
                    LogExceptionToFile(ex);
                    MessageBox.Show($"Please select a valid KeyWords excel file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    LogExceptionToFile(ex);
                    MessageBox.Show($"An error occurred: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private void LogExceptionToFile(System.Exception ex)
        {
            // Get the directory of the executable
            string directoryPath = AppDomain.CurrentDomain.BaseDirectory;

            // Append the exception details to the log file
            string logFilePath = Path.Combine(directoryPath, "ErrorLog.txt");

            using (StreamWriter writer = new StreamWriter(logFilePath, true))
            {
                writer.WriteLine($"Timestamp: {DateTime.Now}");
                writer.WriteLine($"Exception Message: {ex.Message}");
                writer.WriteLine($"StackTrace: {ex.StackTrace}");
                writer.WriteLine("--------------------------------------------------");
            }
        }

        private void LogEmailDetailsToFile(Outlook.MailItem email)
        {
            // Get the directory of the executable
            string directoryPath = AppDomain.CurrentDomain.BaseDirectory;

            // Append the email details to the log file
            string logFilePath = Path.Combine(directoryPath, "EmailDetailsLog.txt");

            using (StreamWriter writer = new StreamWriter(logFilePath, true))
            {
                writer.WriteLine($"Subject: {email.Subject}");
                writer.WriteLine($"Received Time: {email.ReceivedTime}");
                writer.WriteLine($"Creation Time: {email.CreationTime}");
                writer.WriteLine($"Sender: {email.SenderEmailAddress}");
                // Add any other details you want to log

                writer.WriteLine("--------------------------------------------------");
            }
        }


        private Outlook.MAPIFolder GetFolderByName(Outlook.MAPIFolder parentFolder, string folderName)
        {
            foreach (Outlook.MAPIFolder folder in parentFolder.Folders)
            {
                if (folder.Name == folderName)
                {
                    return folder; // Found the desired folder
                }
                else
                {
                    // Recursively search subfolders
                    Outlook.MAPIFolder subFolder = GetFolderByName(folder, folderName);
                    if (subFolder != null)
                        return subFolder;
                }
            }
            return null; // Folder not found
        }

        private Cell CreateCell(string text)
        {
            return new Cell(new InlineString(new Text(text)));
        }

        private string GetCategoryFromKeywords(string emailBody)
        {
            // Specify the path to your Keywords Excel file
            string keywordsFilePath = keywordsFilePathTextBox.Text;

            using (var workbook = new XLWorkbook(keywordsFilePath))
            {
                var worksheet = workbook.Worksheet(1);

                // Loop through the rows in the Keywords Excel file
                foreach (var row in worksheet.RowsUsed().Skip(1))
                {
                    string keyword = row.Cell(1).Value.ToString();
                    string category = row.Cell(2).Value.ToString();

                    // Check if the keyword is present in the email body
                    if (emailBody.Contains("Keyword I would add: " + keyword, StringComparison.OrdinalIgnoreCase))
                    {
                        return category; // Return the category when a match is found
                    }
                }
            }

            return "Uncategorized"; // Return default if no match is found
        }

        private void browseKeywordsButton_Click_1(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xlsx;*.xls|All Files|*.*";
                openFileDialog.Title = "Select the Keywords Excel File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Set the selected file path to the TextBox
                    keywordsFilePathTextBox.Text = openFileDialog.FileName;
                }
            }
        }
    }
}