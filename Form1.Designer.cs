namespace OutlookArchivierungApp
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
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
            mainTabControl = new System.Windows.Forms.TabControl();
            mainTabPage = new System.Windows.Forms.TabPage();
            groupBox2 = new System.Windows.Forms.GroupBox();
            outlookFolderComboBox = new System.Windows.Forms.ComboBox();
            loadEmailsButton = new System.Windows.Forms.Button();
            groupBox3 = new System.Windows.Forms.GroupBox();
            emailsDataGridView = new System.Windows.Forms.DataGridView();
            groupBox5 = new System.Windows.Forms.GroupBox();
            dateFromPicker = new System.Windows.Forms.DateTimePicker();
            dateToPicker = new System.Windows.Forms.DateTimePicker();
            label4 = new System.Windows.Forms.Label();
            label5 = new System.Windows.Forms.Label();
            statusFilterComboBox = new System.Windows.Forms.ComboBox();
            applyFilterButton = new System.Windows.Forms.Button();
            clearFilterButton = new System.Windows.Forms.Button();
            progressBar = new System.Windows.Forms.ProgressBar();
            statusLabel = new System.Windows.Forms.Label();
            exportSelectedButton = new System.Windows.Forms.Button();
            exportAllButton = new System.Windows.Forms.Button();
            settingsTabPage = new System.Windows.Forms.TabPage();
            groupBox4 = new System.Windows.Forms.GroupBox();
            btnFolderBetreff = new System.Windows.Forms.Button();
            btnFolderAbsender = new System.Windows.Forms.Button();
            btnFolderMM = new System.Windows.Forms.Button();
            btnFolderYYYY = new System.Windows.Forms.Button();
            folderPatternTextBox = new System.Windows.Forms.TextBox();
            label3 = new System.Windows.Forms.Label();
            btnSender = new System.Windows.Forms.Button();
            btnSubject = new System.Windows.Forms.Button();
            btnDate = new System.Windows.Forms.Button();
            filenamePatternTextBox = new System.Windows.Forms.TextBox();
            label2 = new System.Windows.Forms.Label();
            includeAttachmentsCheckBox = new System.Windows.Forms.CheckBox();
            includeCcBccCheckBox = new System.Windows.Forms.CheckBox();
            createSubfoldersCheckBox = new System.Windows.Forms.CheckBox();
            createLogFileCheckBox = new System.Windows.Forms.CheckBox();
            groupBox1 = new System.Windows.Forms.GroupBox();
            label1 = new System.Windows.Forms.Label();
            outputFolderTextBox = new System.Windows.Forms.TextBox();
            browseButton = new System.Windows.Forms.Button();
            folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            menuStrip1 = new System.Windows.Forms.MenuStrip();
            dateiToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            einstellungenToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            beendenToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            hilfeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            überToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            exportFormatComboBox = new System.Windows.Forms.ComboBox();
            preserveFormattingCheckBox = new System.Windows.Forms.CheckBox();
            embedImagesCheckBox = new System.Windows.Forms.CheckBox();
            toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            mainTabControl.SuspendLayout();
            mainTabPage.SuspendLayout();
            groupBox2.SuspendLayout();
            groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)emailsDataGridView).BeginInit();
            groupBox5.SuspendLayout();
            settingsTabPage.SuspendLayout();
            groupBox4.SuspendLayout();
            groupBox1.SuspendLayout();
            menuStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // mainTabControl
            // 
            mainTabControl.Controls.Add(mainTabPage);
            mainTabControl.Controls.Add(settingsTabPage);
            mainTabControl.Dock = System.Windows.Forms.DockStyle.Fill;
            mainTabControl.Location = new System.Drawing.Point(0, 24);
            mainTabControl.Name = "mainTabControl";
            mainTabControl.SelectedIndex = 0;
            mainTabControl.Size = new System.Drawing.Size(1915, 794);
            mainTabControl.TabIndex = 0;
            // 
            // mainTabPage
            // 
            mainTabPage.Controls.Add(groupBox2);
            mainTabPage.Controls.Add(groupBox3);
            mainTabPage.Controls.Add(groupBox5);
            mainTabPage.Controls.Add(progressBar);
            mainTabPage.Controls.Add(statusLabel);
            mainTabPage.Controls.Add(exportSelectedButton);
            mainTabPage.Controls.Add(exportAllButton);
            mainTabPage.Location = new System.Drawing.Point(4, 24);
            mainTabPage.Name = "mainTabPage";
            mainTabPage.Size = new System.Drawing.Size(1907, 766);
            mainTabPage.TabIndex = 0;
            mainTabPage.Text = "E-Mail Archivierung";
            mainTabPage.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            groupBox2.Controls.Add(outlookFolderComboBox);
            groupBox2.Controls.Add(loadEmailsButton);
            groupBox2.Location = new System.Drawing.Point(18, 318);
            groupBox2.Name = "groupBox2";
            groupBox2.Size = new System.Drawing.Size(580, 80);
            groupBox2.TabIndex = 1;
            groupBox2.TabStop = false;
            groupBox2.Text = "Outlook-Ordner";
            // 
            // outlookFolderComboBox
            // 
            outlookFolderComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            outlookFolderComboBox.Location = new System.Drawing.Point(10, 25);
            outlookFolderComboBox.Name = "outlookFolderComboBox";
            outlookFolderComboBox.Size = new System.Drawing.Size(450, 23);
            outlookFolderComboBox.TabIndex = 0;
            // 
            // loadEmailsButton
            // 
            loadEmailsButton.Location = new System.Drawing.Point(480, 25);
            loadEmailsButton.Name = "loadEmailsButton";
            loadEmailsButton.Size = new System.Drawing.Size(80, 23);
            loadEmailsButton.TabIndex = 1;
            loadEmailsButton.Text = "E-Mails laden";
            loadEmailsButton.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            groupBox3.Controls.Add(emailsDataGridView);
            groupBox3.Location = new System.Drawing.Point(8, 12);
            groupBox3.Name = "groupBox3";
            groupBox3.Size = new System.Drawing.Size(989, 300);
            groupBox3.TabIndex = 2;
            groupBox3.TabStop = false;
            groupBox3.Text = "E-Mail-Liste";
            // 
            // emailsDataGridView
            // 
            emailsDataGridView.AllowUserToAddRows = false;
            emailsDataGridView.AllowUserToDeleteRows = false;
            emailsDataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            emailsDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            emailsDataGridView.Location = new System.Drawing.Point(10, 25);
            emailsDataGridView.Name = "emailsDataGridView";
            emailsDataGridView.ReadOnly = true;
            emailsDataGridView.Size = new System.Drawing.Size(957, 240);
            emailsDataGridView.TabIndex = 0;
            // 
            // groupBox5
            // 
            groupBox5.Controls.Add(dateFromPicker);
            groupBox5.Controls.Add(dateToPicker);
            groupBox5.Controls.Add(label4);
            groupBox5.Controls.Add(label5);
            groupBox5.Controls.Add(statusFilterComboBox);
            groupBox5.Controls.Add(applyFilterButton);
            groupBox5.Controls.Add(clearFilterButton);
            groupBox5.Location = new System.Drawing.Point(18, 415);
            groupBox5.Name = "groupBox5";
            groupBox5.Size = new System.Drawing.Size(580, 120);
            groupBox5.TabIndex = 4;
            groupBox5.TabStop = false;
            groupBox5.Text = "Filter";
            // 
            // dateFromPicker
            // 
            dateFromPicker.Location = new System.Drawing.Point(100, 22);
            dateFromPicker.Name = "dateFromPicker";
            dateFromPicker.Size = new System.Drawing.Size(120, 23);
            dateFromPicker.TabIndex = 0;
            // 
            // dateToPicker
            // 
            dateToPicker.Location = new System.Drawing.Point(330, 22);
            dateToPicker.Name = "dateToPicker";
            dateToPicker.Size = new System.Drawing.Size(120, 23);
            dateToPicker.TabIndex = 1;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new System.Drawing.Point(10, 25);
            label4.Name = "label4";
            label4.Size = new System.Drawing.Size(30, 15);
            label4.TabIndex = 2;
            label4.Text = "Von:";
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new System.Drawing.Point(240, 25);
            label5.Name = "label5";
            label5.Size = new System.Drawing.Size(25, 15);
            label5.TabIndex = 3;
            label5.Text = "Bis:";
            // 
            // statusFilterComboBox
            // 
            statusFilterComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            statusFilterComboBox.Location = new System.Drawing.Point(10, 55);
            statusFilterComboBox.Name = "statusFilterComboBox";
            statusFilterComboBox.Size = new System.Drawing.Size(150, 23);
            statusFilterComboBox.TabIndex = 4;
            // 
            // applyFilterButton
            // 
            applyFilterButton.Location = new System.Drawing.Point(480, 55);
            applyFilterButton.Name = "applyFilterButton";
            applyFilterButton.Size = new System.Drawing.Size(80, 23);
            applyFilterButton.TabIndex = 5;
            applyFilterButton.Text = "Anwenden";
            applyFilterButton.UseVisualStyleBackColor = true;
            // 
            // clearFilterButton
            // 
            clearFilterButton.Location = new System.Drawing.Point(480, 85);
            clearFilterButton.Name = "clearFilterButton";
            clearFilterButton.Size = new System.Drawing.Size(80, 23);
            clearFilterButton.TabIndex = 6;
            clearFilterButton.Text = "Zurücksetzen";
            clearFilterButton.UseVisualStyleBackColor = true;
            // 
            // progressBar
            // 
            progressBar.Location = new System.Drawing.Point(18, 564);
            progressBar.Name = "progressBar";
            progressBar.Size = new System.Drawing.Size(580, 20);
            progressBar.TabIndex = 5;
            // 
            // statusLabel
            // 
            statusLabel.AutoSize = true;
            statusLabel.Location = new System.Drawing.Point(18, 596);
            statusLabel.Name = "statusLabel";
            statusLabel.Size = new System.Drawing.Size(37, 15);
            statusLabel.TabIndex = 6;
            statusLabel.Text = "Bereit";
            // 
            // exportSelectedButton
            // 
            exportSelectedButton.Location = new System.Drawing.Point(348, 609);
            exportSelectedButton.Name = "exportSelectedButton";
            exportSelectedButton.Size = new System.Drawing.Size(120, 43);
            exportSelectedButton.TabIndex = 7;
            exportSelectedButton.Text = "Ausgewählte exportieren";
            exportSelectedButton.UseVisualStyleBackColor = true;
            // 
            // exportAllButton
            // 
            exportAllButton.Location = new System.Drawing.Point(478, 609);
            exportAllButton.Name = "exportAllButton";
            exportAllButton.Size = new System.Drawing.Size(120, 43);
            exportAllButton.TabIndex = 8;
            exportAllButton.Text = "Alle exportieren";
            exportAllButton.UseVisualStyleBackColor = true;
            // 
            // settingsTabPage
            // 
            settingsTabPage.Controls.Add(groupBox4);
            settingsTabPage.Controls.Add(groupBox1);
            settingsTabPage.Location = new System.Drawing.Point(4, 24);
            settingsTabPage.Name = "settingsTabPage";
            settingsTabPage.Size = new System.Drawing.Size(1907, 766);
            settingsTabPage.TabIndex = 1;
            settingsTabPage.Text = "Einstellungen";
            settingsTabPage.UseVisualStyleBackColor = true;
            // 
            // groupBox4
            // 
            groupBox4.Controls.Add(btnFolderBetreff);
            groupBox4.Controls.Add(btnFolderAbsender);
            groupBox4.Controls.Add(btnFolderMM);
            groupBox4.Controls.Add(btnFolderYYYY);
            groupBox4.Controls.Add(folderPatternTextBox);
            groupBox4.Controls.Add(label3);
            groupBox4.Controls.Add(btnSender);
            groupBox4.Controls.Add(btnSubject);
            groupBox4.Controls.Add(btnDate);
            groupBox4.Controls.Add(filenamePatternTextBox);
            groupBox4.Controls.Add(label2);
            groupBox4.Controls.Add(includeAttachmentsCheckBox);
            groupBox4.Controls.Add(includeCcBccCheckBox);
            groupBox4.Controls.Add(createSubfoldersCheckBox);
            groupBox4.Controls.Add(createLogFileCheckBox);
            groupBox4.Location = new System.Drawing.Point(20, 20);
            groupBox4.Name = "groupBox4";
            groupBox4.Size = new System.Drawing.Size(674, 417);
            groupBox4.TabIndex = 3;
            groupBox4.TabStop = false;
            groupBox4.Text = "Export-Einstellungen";
            // 
            // btnFolderBetreff
            // 
            btnFolderBetreff.Location = new System.Drawing.Point(396, 243);
            btnFolderBetreff.Name = "btnFolderBetreff";
            btnFolderBetreff.Size = new System.Drawing.Size(75, 23);
            btnFolderBetreff.TabIndex = 14;
            btnFolderBetreff.Text = "Betreff";
            btnFolderBetreff.UseVisualStyleBackColor = true;
            btnFolderBetreff.Click += btnFolderBetreff_Click;
            // 
            // btnFolderAbsender
            // 
            btnFolderAbsender.Location = new System.Drawing.Point(315, 243);
            btnFolderAbsender.Name = "btnFolderAbsender";
            btnFolderAbsender.Size = new System.Drawing.Size(75, 23);
            btnFolderAbsender.TabIndex = 13;
            btnFolderAbsender.Text = "Absender";
            btnFolderAbsender.UseVisualStyleBackColor = true;
            btnFolderAbsender.Click += btnFolderAbsender_Click;
            // 
            // btnFolderMM
            // 
            btnFolderMM.Location = new System.Drawing.Point(234, 243);
            btnFolderMM.Name = "btnFolderMM";
            btnFolderMM.Size = new System.Drawing.Size(75, 23);
            btnFolderMM.TabIndex = 12;
            btnFolderMM.Text = "Monat";
            btnFolderMM.UseVisualStyleBackColor = true;
            btnFolderMM.Click += btnFolderMM_Click;
            // 
            // btnFolderYYYY
            // 
            btnFolderYYYY.Location = new System.Drawing.Point(152, 243);
            btnFolderYYYY.Name = "btnFolderYYYY";
            btnFolderYYYY.Size = new System.Drawing.Size(75, 23);
            btnFolderYYYY.TabIndex = 11;
            btnFolderYYYY.Text = "Jahr";
            btnFolderYYYY.UseVisualStyleBackColor = true;
            btnFolderYYYY.Click += btnFolderYYYY_Click;
            // 
            // folderPatternTextBox
            // 
            folderPatternTextBox.Location = new System.Drawing.Point(134, 195);
            folderPatternTextBox.Name = "folderPatternTextBox";
            folderPatternTextBox.Size = new System.Drawing.Size(336, 23);
            folderPatternTextBox.TabIndex = 10;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new System.Drawing.Point(15, 203);
            label3.Name = "label3";
            label3.Size = new System.Drawing.Size(113, 15);
            label3.TabIndex = 9;
            label3.Text = "Unterordner-Muster";
            // 
            // btnSender
            // 
            btnSender.Location = new System.Drawing.Point(318, 153);
            btnSender.Name = "btnSender";
            btnSender.Size = new System.Drawing.Size(75, 23);
            btnSender.TabIndex = 8;
            btnSender.Text = "Absender";
            btnSender.UseVisualStyleBackColor = true;
            btnSender.Click += btnSender_Click;
            // 
            // btnSubject
            // 
            btnSubject.Location = new System.Drawing.Point(233, 153);
            btnSubject.Name = "btnSubject";
            btnSubject.Size = new System.Drawing.Size(75, 23);
            btnSubject.TabIndex = 7;
            btnSubject.Text = "Betreff";
            btnSubject.UseVisualStyleBackColor = true;
            btnSubject.Click += btnSubject_Click;
            // 
            // btnDate
            // 
            btnDate.Location = new System.Drawing.Point(152, 153);
            btnDate.Name = "btnDate";
            btnDate.Size = new System.Drawing.Size(75, 23);
            btnDate.TabIndex = 6;
            btnDate.Text = "Datum";
            btnDate.UseVisualStyleBackColor = true;
            btnDate.Click += btnDate_Click;
            // 
            // filenamePatternTextBox
            // 
            filenamePatternTextBox.Location = new System.Drawing.Point(134, 124);
            filenamePatternTextBox.Name = "filenamePatternTextBox";
            filenamePatternTextBox.Size = new System.Drawing.Size(336, 23);
            filenamePatternTextBox.TabIndex = 5;
            filenamePatternTextBox.TextChanged += filenamePatternTextBox_TextChanged;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new System.Drawing.Point(15, 132);
            label2.Name = "label2";
            label2.Size = new System.Drawing.Size(113, 15);
            label2.TabIndex = 4;
            label2.Text = "Dateinamen-Muster";
            // 
            // includeAttachmentsCheckBox
            // 
            includeAttachmentsCheckBox.AutoSize = true;
            includeAttachmentsCheckBox.Location = new System.Drawing.Point(15, 25);
            includeAttachmentsCheckBox.Name = "includeAttachmentsCheckBox";
            includeAttachmentsCheckBox.Size = new System.Drawing.Size(143, 19);
            includeAttachmentsCheckBox.TabIndex = 0;
            includeAttachmentsCheckBox.Text = "Anhänge einschließen";
            includeAttachmentsCheckBox.UseVisualStyleBackColor = true;
            // 
            // includeCcBccCheckBox
            // 
            includeCcBccCheckBox.AutoSize = true;
            includeCcBccCheckBox.Location = new System.Drawing.Point(15, 50);
            includeCcBccCheckBox.Name = "includeCcBccCheckBox";
            includeCcBccCheckBox.Size = new System.Drawing.Size(139, 19);
            includeCcBccCheckBox.TabIndex = 1;
            includeCcBccCheckBox.Text = "CC/BCC einschließen";
            includeCcBccCheckBox.UseVisualStyleBackColor = true;
            // 
            // createSubfoldersCheckBox
            // 
            createSubfoldersCheckBox.AutoSize = true;
            createSubfoldersCheckBox.Location = new System.Drawing.Point(15, 75);
            createSubfoldersCheckBox.Name = "createSubfoldersCheckBox";
            createSubfoldersCheckBox.Size = new System.Drawing.Size(137, 19);
            createSubfoldersCheckBox.TabIndex = 2;
            createSubfoldersCheckBox.Text = "Unterordner erstellen";
            createSubfoldersCheckBox.UseVisualStyleBackColor = true;
            // 
            // createLogFileCheckBox
            // 
            createLogFileCheckBox.AutoSize = true;
            createLogFileCheckBox.Location = new System.Drawing.Point(15, 100);
            createLogFileCheckBox.Name = "createLogFileCheckBox";
            createLogFileCheckBox.Size = new System.Drawing.Size(125, 19);
            createLogFileCheckBox.TabIndex = 3;
            createLogFileCheckBox.Text = "Log-Datei erstellen";
            createLogFileCheckBox.UseVisualStyleBackColor = true;

            // startWithWindowsCheckBox
            // 
            startWithWindowsCheckBox = new CheckBox();
            startWithWindowsCheckBox.AutoSize = true;
            startWithWindowsCheckBox.Location = new System.Drawing.Point(15, 280);
            startWithWindowsCheckBox.Name = "startWithWindowsCheckBox";
            startWithWindowsCheckBox.Size = new System.Drawing.Size(181, 19);
            startWithWindowsCheckBox.TabIndex = 15;
            startWithWindowsCheckBox.Text = "Mit Windows starten";
            startWithWindowsCheckBox.UseVisualStyleBackColor = true;
            startWithWindowsCheckBox.CheckedChanged += startWithWindowsCheckBox_CheckedChanged;
            groupBox4.Controls.Add(startWithWindowsCheckBox);
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(label1);
            groupBox1.Controls.Add(outputFolderTextBox);
            groupBox1.Controls.Add(browseButton);
            groupBox1.Location = new System.Drawing.Point(20, 443);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new System.Drawing.Size(674, 80);
            groupBox1.TabIndex = 0;
            groupBox1.TabStop = false;
            groupBox1.Text = "Ausgabeordner";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new System.Drawing.Point(10, 25);
            label1.Name = "label1";
            label1.Size = new System.Drawing.Size(91, 15);
            label1.TabIndex = 0;
            label1.Text = "Ausgabeordner:";
            // 
            // outputFolderTextBox
            // 
            outputFolderTextBox.Location = new System.Drawing.Point(120, 22);
            outputFolderTextBox.Name = "outputFolderTextBox";
            outputFolderTextBox.Size = new System.Drawing.Size(350, 23);
            outputFolderTextBox.TabIndex = 1;
            // 
            // browseButton
            // 
            browseButton.Location = new System.Drawing.Point(480, 22);
            browseButton.Name = "browseButton";
            browseButton.Size = new System.Drawing.Size(80, 23);
            browseButton.TabIndex = 2;
            browseButton.Text = "Durchsuchen";
            browseButton.UseVisualStyleBackColor = true;
            // 
            // menuStrip1
            // 
            menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] { dateiToolStripMenuItem, hilfeToolStripMenuItem });
            menuStrip1.Location = new System.Drawing.Point(0, 0);
            menuStrip1.Name = "menuStrip1";
            menuStrip1.Size = new System.Drawing.Size(1915, 24);
            menuStrip1.TabIndex = 1;
            menuStrip1.Text = "menuStrip1";
            // 
            // dateiToolStripMenuItem
            // 
            dateiToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { einstellungenToolStripMenuItem, beendenToolStripMenuItem });
            dateiToolStripMenuItem.Name = "dateiToolStripMenuItem";
            dateiToolStripMenuItem.Size = new System.Drawing.Size(46, 20);
            dateiToolStripMenuItem.Text = "Datei";
            // 
            // einstellungenToolStripMenuItem
            // 
            einstellungenToolStripMenuItem.Name = "einstellungenToolStripMenuItem";
            einstellungenToolStripMenuItem.Size = new System.Drawing.Size(145, 22);
            einstellungenToolStripMenuItem.Text = "Einstellungen";
            einstellungenToolStripMenuItem.Click += settingsButton_Click;
            // 
            // beendenToolStripMenuItem
            // 
            beendenToolStripMenuItem.Name = "beendenToolStripMenuItem";
            beendenToolStripMenuItem.Size = new System.Drawing.Size(145, 22);
            beendenToolStripMenuItem.Text = "Beenden";
            // 
            // hilfeToolStripMenuItem
            // 
            hilfeToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] { überToolStripMenuItem });
            hilfeToolStripMenuItem.Name = "hilfeToolStripMenuItem";
            hilfeToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            hilfeToolStripMenuItem.Text = "Hilfe";
            // 
            // überToolStripMenuItem
            // 
            überToolStripMenuItem.Name = "überToolStripMenuItem";
            überToolStripMenuItem.Size = new System.Drawing.Size(99, 22);
            überToolStripMenuItem.Text = "Über";
            // 
            // exportFormatComboBox
            // 
            exportFormatComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            exportFormatComboBox.Enabled = false;
            exportFormatComboBox.Location = new System.Drawing.Point(150, 105);
            exportFormatComboBox.Name = "exportFormatComboBox";
            exportFormatComboBox.Size = new System.Drawing.Size(200, 23);
            exportFormatComboBox.TabIndex = 5;
            // 
            // preserveFormattingCheckBox
            // 
            preserveFormattingCheckBox.AutoSize = true;
            preserveFormattingCheckBox.Location = new System.Drawing.Point(150, 135);
            preserveFormattingCheckBox.Name = "preserveFormattingCheckBox";
            preserveFormattingCheckBox.Size = new System.Drawing.Size(150, 19);
            preserveFormattingCheckBox.TabIndex = 9;
            preserveFormattingCheckBox.Text = "Formatierung beibehalten";
            preserveFormattingCheckBox.UseVisualStyleBackColor = true;
            // 
            // embedImagesCheckBox
            // 
            embedImagesCheckBox.AutoSize = true;
            embedImagesCheckBox.Location = new System.Drawing.Point(150, 160);
            embedImagesCheckBox.Name = "embedImagesCheckBox";
            embedImagesCheckBox.Size = new System.Drawing.Size(104, 24);
            embedImagesCheckBox.TabIndex = 0;
            // 
            // toolStripMenuItem1
            // 
            toolStripMenuItem1.Name = "toolStripMenuItem1";
            toolStripMenuItem1.Size = new System.Drawing.Size(32, 19);
            // 
            // Form1
            // 
            AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(1915, 818);
            Controls.Add(mainTabControl);
            Controls.Add(menuStrip1);
            MainMenuStrip = menuStrip1;
            Text = "E-Mail Archivierung";
            mainTabControl.ResumeLayout(false);
            mainTabPage.ResumeLayout(false);
            mainTabPage.PerformLayout();
            groupBox2.ResumeLayout(false);
            groupBox3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)emailsDataGridView).EndInit();
            groupBox5.ResumeLayout(false);
            groupBox5.PerformLayout();
            settingsTabPage.ResumeLayout(false);
            groupBox4.ResumeLayout(false);
            groupBox4.PerformLayout();
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            menuStrip1.ResumeLayout(false);
            menuStrip1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;

        #endregion

        private TabControl mainTabControl;
        private TabPage mainTabPage;
        private TabPage settingsTabPage;
        private GroupBox groupBox1;
        private Label label1;
        private TextBox outputFolderTextBox;
        private Button browseButton;
        private GroupBox groupBox2;
        private ComboBox outlookFolderComboBox;
        private Button loadEmailsButton;
        private GroupBox groupBox3;
        private DataGridView emailsDataGridView;
        private System.Windows.Forms.GroupBox groupBox4;
        private CheckBox includeAttachmentsCheckBox;
        private CheckBox includeCcBccCheckBox;
        private CheckBox createSubfoldersCheckBox;
        private CheckBox createLogFileCheckBox;
        private GroupBox groupBox5;
        private DateTimePicker dateFromPicker;
        private DateTimePicker dateToPicker;
        private Label label4;
        private Label label5;
        private ComboBox statusFilterComboBox;
        private Button applyFilterButton;
        private Button clearFilterButton;
        private ProgressBar progressBar;
        private Label statusLabel;
        private Button exportSelectedButton;
        private Button exportAllButton;
        private FolderBrowserDialog folderBrowserDialog;
        private MenuStrip menuStrip1;
        private ToolStripMenuItem dateiToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem einstellungenToolStripMenuItem;
        private ToolStripMenuItem beendenToolStripMenuItem;
        private ToolStripMenuItem hilfeToolStripMenuItem;
        private ToolStripMenuItem überToolStripMenuItem;
        private ComboBox exportFormatComboBox;
        private CheckBox preserveFormattingCheckBox;
        private CheckBox embedImagesCheckBox;
        private TextBox filenamePatternTextBox;
        private Label label2;
        private Button btnSender;
        private Button btnSubject;
        private Button btnDate;
        private Button btnFolderBetreff;
        private Button btnFolderAbsender;
        private Button btnFolderMM;
        private Button btnFolderYYYY;
        private TextBox folderPatternTextBox;
        private Label label3;
        private CheckBox startWithWindowsCheckBox;
        
    }
}