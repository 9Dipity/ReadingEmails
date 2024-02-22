using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Excel = Microsoft.Office.Interop.Excel;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
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
            startDate = startDatePicker.Value;
            endDate = endDatePicker.Value;

            // Create Excel workbook
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create("C:\\Users\\Miuskaste\\Desktop\\EmailReading\\OutputFile.xlsx", SpreadsheetDocumentType.Workbook))
            {
                // Add a WorkbookPart to the document
                WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
                workbookpart.Workbook = new Workbook();

                // Add a WorksheetPart to the WorkbookPart
                WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // Get the sheetData from the WorksheetPart
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                // Set column headers
                DocumentFormat.OpenXml.Spreadsheet.Row headerRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                headerRow.Append(CreateCell("Mailbox"));
                headerRow.Append(CreateCell("Time received"));
                headerRow.Append(CreateCell("Time replied"));
                headerRow.Append(CreateCell("Client"));
                headerRow.Append(CreateCell("Category"));
                sheetData.AppendChild(headerRow);

                // Read emails using Outlook Interop API
                Outlook.Application outlookApp = new Outlook.Application();
                Outlook.NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");
                Outlook.MAPIFolder inbox = outlookNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

                foreach (object item in inbox.Items)
                {
                    if (item is Outlook.MailItem)
                    {
                        Outlook.MailItem email = (Outlook.MailItem)item;

                        // Check if email is within the specified date range
                        if (email.ReceivedTime >= startDate && email.ReceivedTime <= endDate)
                        {
                            // Extract information from the email
                            string mailbox = "Inbox";
                            DateTime timeReceived = email.ReceivedTime;
                            DateTime timeReplied = timeReceived; // Assuming the reply time is the same as received time
                            string client = email.SenderEmailAddress.Split('@')[0];

                            // Retrieve category based on keyword matching
                            string category = GetCategoryFromKeywords(email.Body);

                            // Write to Excel
                            DocumentFormat.OpenXml.Spreadsheet.Row dataRow = new DocumentFormat.OpenXml.Spreadsheet.Row();
                            dataRow.Append(CreateCell(mailbox));
                            dataRow.Append(CreateCell(timeReceived.ToString()));
                            dataRow.Append(CreateCell(timeReplied.ToString()));
                            dataRow.Append(CreateCell(client));
                            dataRow.Append(CreateCell(category));
                            sheetData.AppendChild(dataRow);
                        }
                    }
                }
            }
        }
        private Cell CreateCell(string text)
        {
            return new Cell(new InlineString(new Text(text)));
        }

        private string GetCategoryFromKeywords(string emailBody)
        {
            // Specify the path to your Keywords Excel file
            string keywordsFilePath = "C:\\Users\\Miuskaste\\Desktop\\EmailReading\\Keywords.xlsx";

            using (var workbook = new XLWorkbook(keywordsFilePath))
            {
                var worksheet = workbook.Worksheet(1);

                // Loop through the rows in the Keywords Excel file
                foreach (var row in worksheet.RowsUsed().Skip(1))
                {
                    string keyword = row.Cell(1).Value.ToString();
                    string category = row.Cell(2).Value.ToString();

                    // Check if the keyword is present in the email body
                    if (emailBody.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                    {
                        return category; // Return the category when a match is found
                    }
                }
            }

            return "Uncategorized"; // Return default if no match is found
        }

        private void startDatePicker_ValueChanged(object sender, EventArgs e)
        {

        }

        private void endDatePicker_ValueChanged(object sender, EventArgs e)
        {

        }
    }
}
