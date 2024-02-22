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
            SuspendLayout();
            // 
            // startButton
            // 
            startButton.Location = new Point(12, 167);
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
            startDatePicker.ValueChanged += startDatePicker_ValueChanged;
            // 
            // endDatePicker
            // 
            endDatePicker.Location = new Point(142, 41);
            endDatePicker.Name = "endDatePicker";
            endDatePicker.Size = new Size(200, 23);
            endDatePicker.TabIndex = 2;
            endDatePicker.ValueChanged += endDatePicker_ValueChanged;
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
            // QueryMagnet
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            AutoSize = true;
            ClientSize = new Size(354, 205);
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
    }
}