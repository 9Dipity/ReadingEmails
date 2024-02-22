namespace ReadingEmails
{
    partial class QueryMagnet
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            startButton = new Button();
            startDatePicker = new DateTimePicker();
            endDatePicker = new DateTimePicker();
            startDateTextBox = new TextBox();
            endDateTextBox = new TextBox();
            keywordsFilePathTextBox = new TextBox();
            browseKeywordsButton = new Button();
            statusTextBox = new TextBox();
            emailFolder = new TextBox();
            emailFolderDescription = new TextBox();
            SuspendLayout();
            // 
            // startButton
            // 
            startButton.Location = new Point(12, 146);
            startButton.Name = "startButton";
            startButton.Size = new Size(330, 28);
            startButton.TabIndex = 0;
            startButton.Text = "Read Emails";
            startButton.UseVisualStyleBackColor = true;
            startButton.Click += readEmails;
            // 
            // startDatePicker
            // 
            startDatePicker.Location = new Point(142, 12);
            startDatePicker.Name = "startDatePicker";
            startDatePicker.Size = new Size(200, 23);
            startDatePicker.TabIndex = 1;
            // 
            // endDatePicker
            // 
            endDatePicker.Location = new Point(142, 41);
            endDatePicker.Name = "endDatePicker";
            endDatePicker.Size = new Size(200, 23);
            endDatePicker.TabIndex = 2;
            // 
            // startDateTextBox
            // 
            startDateTextBox.Location = new Point(12, 12);
            startDateTextBox.Name = "startDateTextBox";
            startDateTextBox.ReadOnly = true;
            startDateTextBox.Size = new Size(124, 23);
            startDateTextBox.TabIndex = 3;
            startDateTextBox.Text = "Start Date";
            // 
            // endDateTextBox
            // 
            endDateTextBox.Location = new Point(12, 41);
            endDateTextBox.Name = "endDateTextBox";
            endDateTextBox.ReadOnly = true;
            endDateTextBox.Size = new Size(124, 23);
            endDateTextBox.TabIndex = 4;
            endDateTextBox.Text = "End Date";
            // 
            // keywordsFilePathTextBox
            // 
            keywordsFilePathTextBox.Location = new Point(12, 70);
            keywordsFilePathTextBox.Name = "keywordsFilePathTextBox";
            keywordsFilePathTextBox.ReadOnly = true;
            keywordsFilePathTextBox.Size = new Size(124, 23);
            keywordsFilePathTextBox.TabIndex = 7;
            keywordsFilePathTextBox.Text = "KeyWords File";
            // 
            // browseKeywordsButton
            // 
            browseKeywordsButton.Location = new Point(142, 70);
            browseKeywordsButton.Name = "browseKeywordsButton";
            browseKeywordsButton.Size = new Size(200, 23);
            browseKeywordsButton.TabIndex = 8;
            browseKeywordsButton.Text = "Browse KeyWord File";
            browseKeywordsButton.UseVisualStyleBackColor = true;
            browseKeywordsButton.Click += browseKeywordsButton_Click_1;
            // 
            // statusTextBox
            // 
            statusTextBox.Location = new Point(12, 180);
            statusTextBox.Name = "statusTextBox";
            statusTextBox.ReadOnly = true;
            statusTextBox.Size = new Size(330, 23);
            statusTextBox.TabIndex = 9;
            statusTextBox.TextAlign = HorizontalAlignment.Center;
            // 
            // emailFolder
            // 
            emailFolder.Location = new Point(142, 99);
            emailFolder.Name = "emailFolder";
            emailFolder.Size = new Size(200, 23);
            emailFolder.TabIndex = 10;
            emailFolder.Text = "Inbox";
            // 
            // emailFolderDescription
            // 
            emailFolderDescription.Location = new Point(12, 99);
            emailFolderDescription.Name = "emailFolderDescription";
            emailFolderDescription.ReadOnly = true;
            emailFolderDescription.Size = new Size(124, 23);
            emailFolderDescription.TabIndex = 11;
            emailFolderDescription.Text = "Email folder";
            // 
            // QueryMagnet
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            AutoSize = true;
            ClientSize = new Size(354, 205);
            Controls.Add(emailFolderDescription);
            Controls.Add(emailFolder);
            Controls.Add(statusTextBox);
            Controls.Add(browseKeywordsButton);
            Controls.Add(keywordsFilePathTextBox);
            Controls.Add(endDateTextBox);
            Controls.Add(startDateTextBox);
            Controls.Add(endDatePicker);
            Controls.Add(startDatePicker);
            Controls.Add(startButton);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            Name = "QueryMagnet";
            Text = "QueryMagnet";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button startButton;
        private DateTimePicker startDatePicker;
        private DateTimePicker endDatePicker;
        private TextBox startDateTextBox;
        private TextBox endDateTextBox;
        private TextBox keywordsFilePathTextBox;
        private Button browseKeywordsButton;
        private TextBox statusTextBox;
        private TextBox emailFolder;
        private TextBox emailFolderDescription;
    }
}