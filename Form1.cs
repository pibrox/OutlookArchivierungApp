using System.ComponentModel;
using System.Data;
using System.Text;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Reflection;
using System.Runtime.InteropServices; 
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextFont = iTextSharp.text.Font; // Alias für iTextSharp Font
using HtmlDoc = HtmlAgilityPack.HtmlDocument;
using Newtonsoft.Json;
using Outlook = Microsoft.Office.Interop.Outlook;
using PuppeteerSharp;
using PuppeteerSharp.Media;  
using Point = System.Drawing.Point; // Alias für Point-Konflikt
using System.Diagnostics;

namespace OutlookArchivierungApp
{
    public partial class Form1 : Form
    {
        private Outlook.Application? outlookApp;
        private Outlook.NameSpace? nameSpace;
        private List<EmailInfo> allEmails;
        private List<EmailInfo> filteredEmails;
        private BackgroundWorker exportWorker;
        private BackgroundWorker loadEmailsWorker;
        private string settingsFilePath;

        // Felder für das Dateinamen-Muster
        private System.Windows.Forms.Label filenamePatternLabel;


        private System.Windows.Forms.CheckBox datePlaceholderCheckBox;
        private System.Windows.Forms.CheckBox subjectPlaceholderCheckBox;
        private System.Windows.Forms.CheckBox senderPlaceholderCheckBox;

        public Form1()
        {
            try
            {
                InitializeComponent();
                InitializeApplication();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Starten der Anwendung: {ex.Message}\n\nDetails: {ex.StackTrace}",
                    "Startfehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void InitializeApplication()
        {
            allEmails = new List<EmailInfo>();
            filteredEmails = new List<EmailInfo>();
            settingsFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OutlookArchivierungApp", "settings.json");

            try {
                // Event-Handler registrieren
                RegisterEventHandlers();

                // Background Workers initialisieren
                InitializeBackgroundWorkers();

                // UI initialisieren ZUERST
                InitializeUI();

                // Outlook initialisieren DANACH
                InitializeOutlook();

                // Einstellungen laden ZULETZT
                LoadSettings();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler bei der Initialisierung: {ex.Message}\n\nDetails: {ex.StackTrace}",
                    "Initialisierungsfehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RegisterEventHandlers()
        {
            browseButton.Click += BrowseButton_Click;
            loadEmailsButton.Click += LoadEmailsButton_Click;
            applyFilterButton.Click += ApplyFilterButton_Click;
            clearFilterButton.Click += ClearFilterButton_Click;
            exportSelectedButton.Click += ExportSelectedButton_Click;
            exportAllButton.Click += ExportAllButton_Click;
            beendenToolStripMenuItem.Click += BeendenToolStripMenuItem_Click;

            // Dateiname-Muster Dropdown Handler
        }

        private void InitializeBackgroundWorkers()
        {
            // Email-Lade-Worker
            loadEmailsWorker = new BackgroundWorker();
            loadEmailsWorker.DoWork += LoadEmailsWorker_DoWork;
            loadEmailsWorker.RunWorkerCompleted += LoadEmailsWorker_RunWorkerCompleted;
            loadEmailsWorker.ProgressChanged += LoadEmailsWorker_ProgressChanged;
            loadEmailsWorker.WorkerReportsProgress = true;

            // Export-Worker
            exportWorker = new BackgroundWorker();
            exportWorker.DoWork += ExportWorker_DoWork;
            exportWorker.RunWorkerCompleted += ExportWorker_RunWorkerCompleted;
            exportWorker.ProgressChanged += ExportWorker_ProgressChanged;
            exportWorker.WorkerReportsProgress = true;
            exportWorker.WorkerSupportsCancellation = true;
        }

        private void InitializeOutlook()
        {
            try
            {
                // Prüfen, ob Outlook installiert ist
                if (!IsOutlookInstalled())
                {
                    MessageBox.Show("Microsoft Outlook ist nicht installiert oder nicht verfügbar.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Prüfen, ob Outlook bereits läuft 
                if (!IsOutlookRunning())
                {
                    MessageBox.Show("Bitte starten Sie Microsoft Outlook zuerst und warten Sie, bis es vollständig geladen ist.", "Outlook nicht gestartet", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Outlook-Anwendung mit Retry-Mechanismus starten
                outlookApp = ConnectToOutlookWithRetry();

                // Null-Check vor GetNamespace
                if (outlookApp != null)
                {
                    nameSpace = outlookApp.GetNamespace("MAPI");

                    if (nameSpace != null)
                    {
                        // Versuchen, sich mit dem Standard-Profil zu verbinden (ohne UI-Prompt)
                        nameSpace.Logon(Type.Missing, Type.Missing, false, false);

                        LoadOutlookFolders();

                        statusLabel.Text = "Outlook erfolgreich initialisiert";
                    }
                    else
                    {
                        MessageBox.Show("Fehler beim Abrufen des Outlook-Namespace.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Fehler beim Starten der Outlook-Anwendung.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (COMException comEx)
            {
                string detailedError = GetCOMErrorDescription(comEx);
                MessageBox.Show($"COM-Fehler beim Initialisieren von Outlook: {detailedError}\n\nFehlercode: 0x{comEx.ErrorCode:X8}\n\nLösungsansätze:\n- Outlook als Administrator starten\n- Antivirus-Software temporär deaktivieren\n- Outlook-Profil reparieren", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                statusLabel.Text = "Outlook-Initialisierung fehlgeschlagen";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Initialisieren von Outlook: {ex.Message}\n\nMögliche Ursachen:\n- Outlook ist nicht installiert\n- Outlook ist nicht konfiguriert\n- Sicherheitseinstellungen blockieren den Zugriff", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                statusLabel.Text = "Outlook-Initialisierung fehlgeschlagen";
            }
        }

        private bool IsOutlookInstalled()
        {
            try
            {
                var outlookType = Type.GetTypeFromProgID("Outlook.Application");
                return outlookType != null;
            }
            catch
            {
                return false;
            }
        }

        private bool IsOutlookRunning()
        {
            try
            {
                // Prüfe alle möglichen Outlook-Prozesse
                var processes = System.Diagnostics.Process.GetProcesses();

                foreach (var process in processes)
                {
                    if (process.ProcessName.Contains("OUTLOOK") ||
                        process.ProcessName.Contains("HxOutlook") ||
                        process.ProcessName.Contains("Outlook"))
                    {
                        return true;
                    }
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        private Outlook.Application? ConnectToOutlookWithRetry()
        {
            for (int attempt = 1; attempt <= 3; attempt++)
            {
                try
                {
                    // Erstelle eine neue Outlook-Instanz
                    return new Outlook.Application();
                }
                catch (COMException)
                {
                    if (attempt == 3) throw;
                    System.Threading.Thread.Sleep(2000);
                }
            }
            return null;
        }

        private string GetCOMErrorDescription(COMException comEx)
        {
            switch ((uint)comEx.ErrorCode)
            {
                case 0x80040154:
                    return "Outlook COM-Komponente ist nicht registriert";
                case 0x80070005:
                    return "Zugriff verweigert - möglicherweise Sicherheitssoftware";
                case 0x8004010F:
                    return "Outlook ist nicht gestartet oder Profil nicht verfügbar";
                case 0x80040201:
                    return "MAPI-Subsystem nicht verfügbar";
                default:
                    return comEx.Message;
            }
        }

        private void LoadOutlookFolders()
        {
            if (nameSpace == null)
            {
                MessageBox.Show("Outlook ist nicht initialisiert.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            outlookFolderComboBox.Items.Clear();

            try
            {
                var folders = nameSpace.Folders;
                if (folders != null)
                {
                    foreach (Outlook.MAPIFolder folder in folders)
                    {
                        AddFolderToComboBox(folder, "");
                    }

                    if (outlookFolderComboBox.Items.Count > 0)
                    {
                        outlookFolderComboBox.SelectedIndex = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Fehler beim Laden der Outlook-Ordner: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AddFolderToComboBox(Outlook.MAPIFolder folder, string prefix)
        {
            try
            {
                if (folder.DefaultItemType == Outlook.OlItemType.olMailItem)
                {
                    outlookFolderComboBox.Items.Add(new FolderInfo { Name = prefix + folder.Name, Folder = folder });
                }

                var subFolders = folder.Folders;
                if (subFolders != null)
                {
                    foreach (Outlook.MAPIFolder subFolder in subFolders)
                    {
                        AddFolderToComboBox(subFolder, prefix + folder.Name + "\\");
                    }
                }
            }
            catch (Exception)
            {
                
            }
        }

        private void InitializeUI()
        {
            // Controls sichtbar machen
            MakeAllControlsVisible();

            // Nur noch ComboBox-Initialisierung und Standardwerte
            SetupDataGridView();

            // "Alle auswählen" Button
            Button selectAllButton = new Button();
            selectAllButton.Text = "Alle auswählen";
            selectAllButton.Size = new Size(100, 25);
            selectAllButton.Location = new Point(15, 270);
            selectAllButton.Click += (s, e) =>
            {
                foreach (var email in filteredEmails)
                    email.Selected = true;
                emailsDataGridView.Refresh();
            };
            groupBox3.Controls.Add(selectAllButton);

            // Standardwerte setzen
            dateFromPicker.Value = DateTime.Now.AddDays(-30);
            dateToPicker.Value = DateTime.Now;



            statusLabel.Text = "Bereit";
            //AddFilenamePatternControls();
        }

        private void MakeAllControlsVisible()
        {
            // Alle Controls sichtbar machen und zu GroupBoxes hinzufügen
            var allControls = new Control[]
            {
                label1, outputFolderTextBox, browseButton,
                outlookFolderComboBox, loadEmailsButton,
                emailsDataGridView,
                dateFromPicker, dateToPicker, statusFilterComboBox, applyFilterButton, clearFilterButton,
                includeAttachmentsCheckBox, includeCcBccCheckBox, createSubfoldersCheckBox,
                createLogFileCheckBox
            };

            foreach (var control in allControls)
            {
                if (control != null)
                {
                    control.Visible = true;
                    control.BringToFront();
                }
            }
        }


        private void AddAdvancedExportOptions()
        {
            // Export-Format Label und ComboBox zu groupBox4 hinzufügen
            Label exportFormatLabel = new Label();
            exportFormatLabel.Text = "Export-Format:";
            exportFormatLabel.Location = new Point(15, 130);
            exportFormatLabel.Size = new Size(100, 20);
            groupBox4.Controls.Add(exportFormatLabel);

            // exportFormatComboBox erstellen falls es null ist
            if (exportFormatComboBox == null)
            {
                exportFormatComboBox = new ComboBox();
                exportFormatComboBox.Name = "exportFormatComboBox";
                groupBox4.Controls.Add(exportFormatComboBox);
            }

            // exportFormatComboBox initialisieren
            exportFormatComboBox.Items.Clear();
            exportFormatComboBox.Items.AddRange(new[] { "PDF", "HTML", "MSG" });
            exportFormatComboBox.SelectedIndex = 0;
            exportFormatComboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            exportFormatComboBox.Location = new Point(120, 130);
            exportFormatComboBox.Size = new Size(100, 20);

            // preserveFormattingCheckBox erstellen falls es null ist
            if (preserveFormattingCheckBox == null)
            {
                preserveFormattingCheckBox = new CheckBox();
                preserveFormattingCheckBox.Name = "preserveFormattingCheckBox";
                groupBox4.Controls.Add(preserveFormattingCheckBox);
            }

            // preserveFormattingCheckBox konfigurieren
            preserveFormattingCheckBox.Text = "Formatierung beibehalten";
            preserveFormattingCheckBox.Checked = true;
            preserveFormattingCheckBox.Location = new Point(15, 155);
            preserveFormattingCheckBox.Size = new Size(200, 20);

            // embedImagesCheckBox erstellen falls es null ist
            if (embedImagesCheckBox == null)
            {
                embedImagesCheckBox = new CheckBox();
                embedImagesCheckBox.Name = "embedImagesCheckBox";
                groupBox4.Controls.Add(embedImagesCheckBox);
            }

            // embedImagesCheckBox konfigurieren
            embedImagesCheckBox.Text = "Bilder einbetten";
            embedImagesCheckBox.Checked = true;
            embedImagesCheckBox.Location = new Point(15, 180);
            embedImagesCheckBox.Size = new Size(200, 20);
        }

        private void settingsButton_Click(object sender, EventArgs e)
        {
            // Zum Einstellungen-Tab wechseln
            if (mainTabControl != null && mainTabControl.TabPages.Count > 1)
            {
                mainTabControl.SelectedIndex = 1; // Einstellungen-Tab
            }
        }

        public class EmailInfo : IComparable<EmailInfo>
        {
            public bool Selected { get; set; } = false;
            public string EntryID { get; set; } = "";
            public string Subject { get; set; } = "";
            public string SenderName { get; set; } = "";
            public string SenderEmailAddress { get; set; } = "";
            public DateTime ReceivedTime { get; set; }
            public string Body { get; set; } = "";
            public string HTMLBody { get; set; } = "";
            public string Status { get; set; } = "";
            public int AttachmentCount { get; set; }
            public long Size { get; set; }
            public string SizeFormatted { get; set; } = "";
            public string To { get; set; } = "";
            public string CC { get; set; } = "";
            public string BCC { get; set; } = "";

            public int CompareTo(EmailInfo? other)
            {
                if (other == null) return 1;
                return ReceivedTime.CompareTo(other.ReceivedTime);
            }
        }

        
private void LoadSettings()
{
    try
    {
        string settingsPath = @"settings.json";
        if (File.Exists(settingsPath))
        {
            string json = File.ReadAllText(settingsPath);
            var settings = JsonConvert.DeserializeObject<AppSettings>(json);
            
            if (settings != null)
            {
                if (outputFolderTextBox != null)
                    outputFolderTextBox.Text = settings.OutputFolder ?? "";

                if (includeAttachmentsCheckBox != null)
                    includeAttachmentsCheckBox.Checked = settings.IncludeAttachments;

                if (includeCcBccCheckBox != null)
                    includeCcBccCheckBox.Checked = settings.IncludeCcBcc;

                if (createSubfoldersCheckBox != null)
                    createSubfoldersCheckBox.Checked = settings.CreateSubfolders;

                if (filenamePatternTextBox != null)
                    filenamePatternTextBox.Text = settings.FilenamePattern ?? "";

                if (createLogFileCheckBox != null)
                    createLogFileCheckBox.Checked = settings.CreateLogFile;

                if (folderPatternTextBox != null)
                    folderPatternTextBox.Text = settings.FolderPattern ?? "";

                // Autostart-CheckBox basierend auf tatsächlichem Status setzen
                if (startWithWindowsCheckBox != null)
                    startWithWindowsCheckBox.Checked = IsAutostartEnabledViaStartupFolder();
            }
        }
        else
        {
            // Beim ersten Start den aktuellen Autostart-Status laden
            if (startWithWindowsCheckBox != null)
                startWithWindowsCheckBox.Checked = IsAutostartEnabledViaStartupFolder();
        }
    }
    catch (Exception ex)
    {
        MessageBox.Show($"Fehler beim Laden der Einstellungen: {ex.Message}", 
                       "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}

private void SaveSettings()
{
    try
    {
        var settings = new AppSettings
        {
            OutputFolder = outputFolderTextBox.Text,
            IncludeAttachments = includeAttachmentsCheckBox.Checked,
            IncludeCcBcc = includeCcBccCheckBox.Checked,
            CreateSubfolders = createSubfoldersCheckBox.Checked,
            FilenamePattern = filenamePatternTextBox.Text,
            CreateLogFile = createLogFileCheckBox.Checked,
            FolderPattern = folderPatternTextBox.Text,
            StartWithWindows = IsAutostartEnabledViaStartupFolder() // Aktuellen Status speichern
        };

        string json = JsonConvert.SerializeObject(settings, Formatting.Indented);
        File.WriteAllText(@"settings.json", json);
    }
    catch (Exception ex)
    {
        MessageBox.Show($"Fehler beim Speichern der Einstellungen: {ex.Message}", 
                       "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}

        public class AppSettings
        {
            public string? OutputFolder { get; set; }
            public bool IncludeAttachments { get; set; } = true;
            public bool IncludeCcBcc { get; set; } = false;
            public bool CreateSubfolders { get; set; } = false;
            public string? FilenamePattern { get; set; }
            public bool CreateLogFile { get; set; } = true;
            public string? FolderPattern { get; set; }
            public bool StartWithWindows { get; set; } = false;
        }

        private void BrowseButton_Click(object sender, EventArgs e)
        {
            using (var folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "Ausgabeordner auswählen";
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    outputFolderTextBox.Text = folderDialog.SelectedPath;
                }
            }
        }

        private void LoadEmailsButton_Click(object sender, EventArgs e)
        {
            if (outlookFolderComboBox.SelectedItem is FolderInfo folderInfo)
            {
                if (!loadEmailsWorker.IsBusy)
                {
                    loadEmailsWorker.RunWorkerAsync(folderInfo.Folder);
                }
            }
            else
            {
                MessageBox.Show("Bitte wählen Sie einen Outlook-Ordner aus.", "Warnung", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void ApplyFilterButton_Click(object sender, EventArgs e)
        {
            ApplyFilter();
        }

        private void ClearFilterButton_Click(object sender, EventArgs e)
        {
            dateFromPicker.Value = DateTime.Now.AddDays(-30);
            dateToPicker.Value = DateTime.Now;
            statusFilterComboBox.SelectedIndex = 0;
            ApplyFilter();
        }

        private void ExportSelectedButton_Click(object sender, EventArgs e)
        {
            var selectedEmails = filteredEmails.Where(e => e.Selected).ToList();
            if (selectedEmails.Count > 0)
            {
                StartExport(selectedEmails);
            }
            else
            {
                MessageBox.Show("Bitte wählen Sie mindestens eine E-Mail aus.", "Warnung", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void ExportAllButton_Click(object sender, EventArgs e)
        {
            if (filteredEmails.Count > 0)
            {
                StartExport(filteredEmails);
            }
            else
            {
                MessageBox.Show("Keine E-Mails zum Exportieren verfügbar.", "Warnung", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }



        private void BeendenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void LoadEmailsWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            var folder = (Outlook.MAPIFolder)e.Argument;
            var emails = new List<EmailInfo>();

            try
            {
                var items = folder.Items;
                int totalCount = items.Count;

                for (int i = 1; i <= totalCount; i++)
                {
                    if (loadEmailsWorker.CancellationPending)
                    {
                        e.Cancel = true;
                        return;
                    }

                    try
                    {
                        var item = items[i];
                        if (item is Outlook.MailItem mailItem)
                        {
                            var emailInfo = new EmailInfo
                            {
                                EntryID = mailItem.EntryID,
                                Subject = GetSafeProperty(() => mailItem.Subject, ""),
                                SenderName = GetSafeProperty(() => mailItem.SenderName, ""),
                                SenderEmailAddress = GetSafeProperty(() => mailItem.SenderEmailAddress, ""),
                                ReceivedTime = GetSafeProperty(() => mailItem.ReceivedTime, DateTime.Now),
                                Body = GetSafeProperty(() => mailItem.Body, ""),
                                HTMLBody = GetSafeProperty(() => mailItem.HTMLBody, ""),
                                AttachmentCount = GetSafeProperty(() => mailItem.Attachments.Count, 0),
                                Size = GetSafeProperty(() => (long)mailItem.Size, 0L),
                                To = GetSafeProperty(() => mailItem.To, ""),
                                CC = GetSafeProperty(() => mailItem.CC, ""),
                                BCC = GetSafeProperty(() => mailItem.BCC, "")
                            };

                            emailInfo.SizeFormatted = FormatFileSize(emailInfo.Size);
                            emails.Add(emailInfo);
                        }
                    }
                    catch (Exception ex)
                    {
                        // Einzelne E-Mail konnte nicht geladen werden
                    }

                    loadEmailsWorker.ReportProgress((i * 100) / totalCount);
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Fehler beim Laden der E-Mails: {ex.Message}", ex);
            }

            e.Result = emails;
        }

        private void LoadEmailsWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                statusLabel.Text = "E-Mail-Laden abgebrochen";
            }
            else if (e.Error != null)
            {
                MessageBox.Show($"Fehler beim Laden der E-Mails: {e.Error.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                statusLabel.Text = "Fehler beim Laden der E-Mails";
            }
            else
            {
                allEmails = (List<EmailInfo>)e.Result;
                filteredEmails = new List<EmailInfo>(allEmails);

                // Sortierung nach Empfangsdatum (neueste zuerst)
                filteredEmails.Sort((a, b) => b.ReceivedTime.CompareTo(a.ReceivedTime));

                emailsDataGridView.DataSource = filteredEmails;
                statusLabel.Text = $"{allEmails.Count} E-Mails geladen";
            }

            loadEmailsButton.Enabled = true;
            progressBar.Value = 0;
        }

        private void LoadEmailsWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;
            statusLabel.Text = $"Lade E-Mails... {e.ProgressPercentage}%";
        }

        private void ExportWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            var emailsToExport = (List<EmailInfo>)e.Argument;
            var result = new ExportResult();
            var errorList = new List<string>();

            // Synchroner Export mit iTextSharp 
            for (int i = 0; i < emailsToExport.Count; i++)
            {
                if (exportWorker.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }

                try
                {
                    string fileName = GenerateFileName(emailsToExport[i]);
                    string targetFolder = outputFolderTextBox.Text;

                    // Unterordner-Logik
                    if (createSubfoldersCheckBox.Checked)
                    {
                        string subfolderPath = GenerateFolderName(emailsToExport[i]);
                        targetFolder = Path.Combine(outputFolderTextBox.Text, subfolderPath);
                        if (!Directory.Exists(targetFolder))
                            Directory.CreateDirectory(targetFolder);
                    }

                    string htmlFilePath = Path.Combine(targetFolder, fileName + ".html");
                    string pdfFilePath = Path.Combine(targetFolder, fileName + ".pdf");

                    if (exportFormatComboBox.SelectedItem?.ToString() == "HTML")
                    {
                        SaveAsHtml(emailsToExport[i], htmlFilePath);
                        // PDF-Datei nicht erstellen
                    }
                    else
                    {
                        SaveAsHtml(emailsToExport[i], htmlFilePath);

                        // MailItem holen:
                        var mailItem = GetMailItemByEntryID(emailsToExport[i].EntryID);
                        if (mailItem != null)
                            ConvertHtmlToPdfWithAttachments(htmlFilePath, pdfFilePath, mailItem, errorList);
                        else
                            ConvertHtmlToPdf(htmlFilePath, pdfFilePath); // Fallback ohne Anhänge
                    }
                    result.SuccessCount++;
                }
                catch (Exception ex)
                {
                    result.ErrorCount++;
                    result.LogEntries.Add($"Fehler bei {emailsToExport[i].Subject}: {ex.Message}");
                }

                exportWorker.ReportProgress((i * 100) / emailsToExport.Count);
            }

            if (errorList.Count > 0)
            {
                MessageBox.Show("Folgende Fehler sind beim Einfügen von Anhängen aufgetreten:\n\n" + string.Join("\n", errorList), "Anhang-Fehler", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            e.Result = result;
        }

        private void ExportWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                statusLabel.Text = "Export abgebrochen";
            }
            else if (e.Error != null)
            {
                MessageBox.Show($"Fehler beim Export: {e.Error.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                statusLabel.Text = "Export fehlgeschlagen";
            }
            else
            {
                var result = (ExportResult)e.Result;
                MessageBox.Show($"Export abgeschlossen!\n\nErfolgreich: {result.SuccessCount}\nFehler: {result.ErrorCount}", "Export", MessageBoxButtons.OK, MessageBoxIcon.Information);
                statusLabel.Text = $"Export abgeschlossen: {result.SuccessCount} erfolgreich, {result.ErrorCount} Fehler";
            }

            exportSelectedButton.Enabled = true;
            exportAllButton.Enabled = true;
            progressBar.Value = 0;
        }

        private void ExportWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;
            statusLabel.Text = $"Exportiere... {e.ProgressPercentage}%";
        }

        private void SetupDataGridView()
        {
            emailsDataGridView.AutoGenerateColumns = false;
            emailsDataGridView.AllowUserToAddRows = false;
            emailsDataGridView.AllowUserToDeleteRows = false;
            emailsDataGridView.ReadOnly = false;
            emailsDataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            emailsDataGridView.MultiSelect = true;

            // Sortierung aktivieren
            emailsDataGridView.Columns.Clear();
            emailsDataGridView.Columns.Add(new DataGridViewCheckBoxColumn
            {
                Name = "Selected",
                HeaderText = "Ausgewählt",
                DataPropertyName = "Selected",
                Width = 80,
                SortMode = DataGridViewColumnSortMode.NotSortable
            });
            emailsDataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "Subject",
                HeaderText = "Betreff",
                DataPropertyName = "Subject",
                Width = 300,
                SortMode = DataGridViewColumnSortMode.Automatic,
                ReadOnly = true // Nur diese Spalte readonly
            });
            emailsDataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "SenderName",
                HeaderText = "Absender",
                DataPropertyName = "SenderName",
                Width = 150,
                SortMode = DataGridViewColumnSortMode.Automatic,
                ReadOnly = true
            });
            emailsDataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "ReceivedTime",
                HeaderText = "Empfangen",
                DataPropertyName = "ReceivedTime",
                Width = 120,
                SortMode = DataGridViewColumnSortMode.Automatic,
                ReadOnly = true
            });
            emailsDataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "SizeFormatted",
                HeaderText = "Größe",
                DataPropertyName = "SizeFormatted",
                Width = 80,
                SortMode = DataGridViewColumnSortMode.Automatic,
                ReadOnly = true
            });
            emailsDataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                Name = "AttachmentCount",
                HeaderText = "Anhänge",
                DataPropertyName = "AttachmentCount",
                Width = 80,
                SortMode = DataGridViewColumnSortMode.Automatic,
                ReadOnly = true
            });

            // Standard-Sortierung nach Empfangsdatum (neueste zuerst)
            emailsDataGridView.Sort(emailsDataGridView.Columns["ReceivedTime"], ListSortDirection.Descending);
        }


        private void ApplyFilter()
        {
            if (allEmails == null) return;

            filteredEmails = allEmails.Where(email =>
            {
                // Datumsfilter
                if (email.ReceivedTime < dateFromPicker.Value.Date ||
                    email.ReceivedTime > dateToPicker.Value.Date.AddDays(1).AddSeconds(-1))
                    return false;

                // Statusfilter
                if (statusFilterComboBox.SelectedIndex > 0)
                {
                    string selectedStatus = statusFilterComboBox.SelectedItem.ToString();
                    if (email.AttachmentCount == 0 && selectedStatus.Contains("mit Anhängen"))
                        return false;
                    if (email.AttachmentCount > 0 && selectedStatus.Contains("ohne Anhänge"))
                        return false;
                }

                return true;
            }).ToList();

            emailsDataGridView.DataSource = filteredEmails;
            statusLabel.Text = $"{filteredEmails.Count} von {allEmails.Count} E-Mails angezeigt";
        }

        private void StartExport(List<EmailInfo> emailsToExport)
        {
            if (string.IsNullOrEmpty(outputFolderTextBox.Text))
            {
                MessageBox.Show("Bitte wählen Sie einen Ausgabeordner.", "Warnung", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (!Directory.Exists(outputFolderTextBox.Text))
            {
                MessageBox.Show("Der Ausgabeordner existiert nicht.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!exportWorker.IsBusy)
            {
                exportSelectedButton.Enabled = false;
                exportAllButton.Enabled = false;
                progressBar.Value = 0;
                exportWorker.RunWorkerAsync(emailsToExport);
            }
        }

        private void ExportEmailSync(EmailInfo email)
        {
            try
            {
                // Dateiname generieren
                string fileName = GenerateFileName(email);
                string htmlFilePath = Path.Combine(outputFolderTextBox.Text, fileName + ".html");
                string pdfFilePath = Path.Combine(outputFolderTextBox.Text, fileName + ".pdf");

                // 1. Erst als HTML speichern (perfekte Formatierung)
                SaveAsHtml(email, htmlFilePath);

                // 2. Dann HTML zu PDF konvertieren
                try
                {
                    ConvertHtmlToPdf(htmlFilePath, pdfFilePath);
                }
                catch (Exception ex)
                {
                    // Falls PDF-Konvertierung fehlschlägt, HTML-Datei behalten
                    throw new Exception($"PDF-Konvertierung fehlgeschlagen, HTML-Datei wurde gespeichert: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Fehler beim Export von '{email.Subject}': {ex.Message}");
            }
        }

        private void SaveAsHtml(EmailInfo email, string htmlFilePath)
        {
            // Vollständiges HTML-Dokument mit Metadaten erstellen
            string htmlContent = $@"
<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <meta name='viewport' content='width=device-width, initial-scale=1.0'>
    <title>{EscapeHtml(email.Subject)}</title>
    <style>
        body {{ 
            font-family: Arial, Helvetica, sans-serif; 
            margin: 20px; 
            line-height: 1.6;
            color: #333;
        }}
        .email-header {{ 
            background-color: #f8f9fa; 
            padding: 15px; 
            border-radius: 5px; 
            margin-bottom: 20px; 
            font-size: 12px;
        }}
        .email-header .title {{ 
            font-size: 16px; 
            font-weight: bold; 
            color: #333; 
            margin-bottom: 10px; 
        }}
        .email-header .detail {{ 
            margin: 3px 0; 
            font-size: 12px;
        }}
        .email-subject {{ 
            font-size: 14px; 
            font-weight: bold; 
            margin-bottom: 15px; 
            color: #333;
        }}
        .email-body {{
            /* Keine zusätzlichen Styles - ursprüngliche Formatierung beibehalten */
        }}
    </style>
</head>
<body>
    <div class='email-header'>
        <div class='title'>E-Mail Export</div>
        <div class='detail'><strong>Von:</strong> {EscapeHtml(email.SenderName)} &lt;{EscapeHtml(email.SenderEmailAddress)}&gt;</div>
        <div class='detail'><strong>An:</strong> {EscapeHtml(email.To)}</div>
        {(string.IsNullOrEmpty(email.CC) ? "" : $"<div class='detail'><strong>CC:</strong> {EscapeHtml(email.CC)}</div>")}
        {(string.IsNullOrEmpty(email.BCC) ? "" : $"<div class='detail'><strong>BCC:</strong> {EscapeHtml(email.BCC)}</div>")}
        <div class='detail'><strong>Datum:</strong> {email.ReceivedTime:dd.MM.yyyy HH:mm}</div>
        <div class='detail'><strong>Größe:</strong> {email.SizeFormatted}</div>
        <div class='detail'><strong>Anhänge:</strong> {email.AttachmentCount}</div>
    </div>
    
    <div class='email-subject'>Betreff: {EscapeHtml(email.Subject)}</div>
    <div class='email-body'>
        {CleanHtmlContent(email.HTMLBody)}
    </div>
</body>
</html>";

            File.WriteAllText(htmlFilePath, htmlContent, Encoding.UTF8);
        }

        private void ConvertHtmlToPdf(string htmlFilePath, string pdfFilePath)
        {
            // PuppeteerSharp für HTML zu PDF Konvertierung
            var task = Task.Run(async () =>
            {
                await using var browser = await Puppeteer.LaunchAsync(new LaunchOptions
                {
                    Headless = true,
                    Args = new[] { "--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage" }
                });

                await using var page = await browser.NewPageAsync();

                // HTML-Datei laden
                await page.GoToAsync($"file:///{htmlFilePath.Replace("\\", "/")}");

                // Auf alle Ressourcen warten
                await page.WaitForNetworkIdleAsync();

                // PDF generieren
                await page.PdfAsync(pdfFilePath, new PdfOptions
                {
                    Format = PuppeteerSharp.Media.PaperFormat.A4,
                    MarginOptions = new MarginOptions
                    {
                        Top = "20mm",
                        Right = "20mm",
                        Bottom = "20mm",
                        Left = "20mm"
                    },
                    PrintBackground = true,
                    PreferCSSPageSize = true
                });
            });

            task.Wait(); // Synchron warten
        }

        private void ConvertHtmlToPdfWithAttachments(string htmlFilePath, string pdfFilePath, Outlook.MailItem mailItem, List<string> errorList)
        {
            ConvertHtmlToPdf(htmlFilePath, pdfFilePath);

            if (mailItem.Attachments.Count > 0)
            {
                string tempPdf = Path.GetTempFileName();
                File.Copy(pdfFilePath, tempPdf, true);

                using (var reader = new iTextSharp.text.pdf.PdfReader(tempPdf))
                using (var fs = new FileStream(pdfFilePath, FileMode.Create, FileAccess.Write))
                using (var stamper = new iTextSharp.text.pdf.PdfStamper(reader, fs))
                {
                    for (int a = 1; a <= mailItem.Attachments.Count; a++)
                    {
                        var att = mailItem.Attachments[a];
                        string ext = Path.GetExtension(att.FileName).ToLower();
                        string tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ext);

                        try
                        {
                            // Outlook-Objekte (z.B. eingebettete E-Mails) überspringen
                            if (att.Type != Outlook.OlAttachmentType.olByValue)
                            {
                                errorList.Add($"Anhang {att.FileName} ist kein Datei-Anhang und wird übersprungen.");
                                continue;
                            }

                            // SaveAsFile explizit absichern!
                            try
                            {
                                att.SaveAsFile(tempFile);
                            }
                            catch (Exception ex)
                            {
                                errorList.Add($"Anhang {att.FileName} konnte nicht gespeichert werden (SaveAsFile-Fehler): {ex.Message}");
                                continue;
                            }

                            if (!File.Exists(tempFile))
                            {
                                errorList.Add($"Anhang {att.FileName} konnte nicht gespeichert werden: Datei nicht gefunden.");
                                continue;
                            }

                            var fileInfo = new FileInfo(tempFile);
                            if (fileInfo.Length == 0)
                            {
                                errorList.Add($"Anhang {att.FileName} ist leer und wird übersprungen.");
                                continue;
                            }

                            if (ext == ".pdf")
                            {
                                try
                                {
                                    using (var attReader = new iTextSharp.text.pdf.PdfReader(tempFile))
                                    {
                                        int n = attReader.NumberOfPages;
                                        for (int i = 1; i <= n; i++)
                                        {
                                            stamper.InsertPage(reader.NumberOfPages + 1, reader.GetPageSize(1));
                                            var cb = stamper.GetOverContent(reader.NumberOfPages);
                                            var page = stamper.GetImportedPage(attReader, i);
                                            cb.AddTemplate(page, 0, 0);
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    errorList.Add($"PDF-Anhang {att.FileName} konnte nicht verarbeitet werden: {ex.Message}");
                                    stamper.InsertPage(reader.NumberOfPages + 1, reader.GetPageSize(1));
                                    iTextSharp.text.pdf.ColumnText.ShowTextAligned(
                                        stamper.GetOverContent(reader.NumberOfPages),
                                        iTextSharp.text.Element.ALIGN_LEFT,
                                        new iTextSharp.text.Phrase($"PDF-Anhang {att.FileName} konnte nicht eingefügt werden."),
                                        50, 800, 0
                                    );
                                }
                            }
                            else if (ext == ".jpg" || ext == ".jpeg" || ext == ".png" || ext == ".bmp" || ext == ".gif")
                            {
                                try
                                {
                                    var img = iTextSharp.text.Image.GetInstance(tempFile);
                                    var pageSize = new iTextSharp.text.Rectangle(img.Width, img.Height);
                                    stamper.InsertPage(reader.NumberOfPages + 1, pageSize);
                                    var cb = stamper.GetOverContent(reader.NumberOfPages);
                                    img.SetAbsolutePosition(0, 0);
                                    cb.AddImage(img);
                                }
                                catch (Exception ex)
                                {
                                    errorList.Add($"Bild-Anhang {att.FileName} konnte nicht eingefügt werden: {ex.Message}");
                                    stamper.InsertPage(reader.NumberOfPages + 1, reader.GetPageSize(1));
                                    iTextSharp.text.pdf.ColumnText.ShowTextAligned(
                                        stamper.GetOverContent(reader.NumberOfPages),
                                        iTextSharp.text.Element.ALIGN_LEFT,
                                        new iTextSharp.text.Phrase($"Bild-Anhang {att.FileName} konnte nicht eingefügt werden."),
                                        50, 800, 0
                                    );
                                }
                            }
                            else
                            {
                                stamper.InsertPage(reader.NumberOfPages + 1, reader.GetPageSize(1));
                                iTextSharp.text.pdf.ColumnText.ShowTextAligned(
                                    stamper.GetOverContent(reader.NumberOfPages),
                                    iTextSharp.text.Element.ALIGN_LEFT,
                                    new iTextSharp.text.Phrase($"Anhang: {att.FileName} (nicht direkt darstellbar)"),
                                    50, 800, 0
                                );
                            }
                        }
                        catch (Exception ex)
                        {
                            errorList.Add($"Fehler bei Anhang {att.FileName}: {ex.Message}");
                        }
                        finally
                        {
                            try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
                        }
                    }
                }
                try { if (File.Exists(tempPdf)) File.Delete(tempPdf); } catch { }
            }
        }

        private string GenerateFileName(EmailInfo email)
        {
            string pattern = filenamePatternTextBox.Text;
            string fileName = pattern
                .Replace("{YYYY-MM-DD}", email.ReceivedTime.ToString("yyyy-MM-dd"))
                .Replace("{DD.MM.YYYY}", email.ReceivedTime.ToString("dd.MM.yyyy"))
                .Replace("{YYYY}", email.ReceivedTime.ToString("yyyy"))
                .Replace("{MM}", email.ReceivedTime.ToString("MM"))
                .Replace("{DD}", email.ReceivedTime.ToString("dd"))
                .Replace("{Betreff}", SanitizeFileName(email.Subject))
                .Replace("{Absender}", SanitizeFileName(email.SenderName));
            // ggf. weitere Platzhalter
            return fileName;
        }

        private string GenerateFolderName(EmailInfo email)
        {
            string pattern = folderPatternTextBox.Text;
            string folderName = pattern
                .Replace("{YYYY-MM-DD}", email.ReceivedTime.ToString("yyyy-MM-dd"))
                .Replace("{DD.MM.YYYY}", email.ReceivedTime.ToString("dd.MM.yyyy"))
                .Replace("{YYYY}", email.ReceivedTime.ToString("yyyy"))
                .Replace("{MM}", email.ReceivedTime.ToString("MM"))
                .Replace("{DD}", email.ReceivedTime.ToString("dd"))
                .Replace("{Betreff}", SanitizeFileName(email.Subject))
                .Replace("{Absender}", SanitizeFileName(email.SenderName));
            return folderName;
        }

        private string SanitizeFileName(string fileName)
        {
            // Ungültige Zeichen entfernen
            char[] invalidChars = Path.GetInvalidFileNameChars();
            foreach (char c in invalidChars)
            {
                fileName = fileName.Replace(c, '_');
            }

            // Länge begrenzen
            if (fileName.Length > 100)
                fileName = fileName.Substring(0, 100);

            return fileName.Trim();
        }

        private string StripHtml(string html)
        {
            if (string.IsNullOrEmpty(html))
                return "";

            // Einfache HTML-Entfernung
            return System.Text.RegularExpressions.Regex.Replace(html, "<[^>]*>", "");
        }

        private T GetSafeProperty<T>(Func<T> propertyAccessor, T defaultValue)
        {
            try
            {
                return propertyAccessor();
            }
            catch
            {
                return defaultValue;
            }
        }

        private string FormatFileSize(long bytes)
        {
            string[] sizes = { "B", "KB", "MB", "GB" };
            double len = bytes;
            int order = 0;
            while (len >= 1024 && order < sizes.Length - 1)
            {
                order++;
                len = len / 1024;
            }
            return $"{len:0.##} {sizes[order]}";
        }

        public class FolderInfo
        {
            public string Name { get; set; } = "";
            public Outlook.MAPIFolder Folder { get; set; } = null!;

            public override string ToString()
            {
                return Name;
            }
        }

        public enum ExportFormat
        {
            PDF,
            HTML,
            MSG
        }

        public class ExportResult
        {
            public int SuccessCount { get; set; }
            public int ErrorCount { get; set; }
            public List<string> LogEntries { get; set; } = new List<string>();
        }

        public class ExportSettings
        {
            public bool IncludeAttachments { get; set; } = true;
            public bool IncludeCcBcc { get; set; } = false;
            public bool CreateSubfolders { get; set; } = false;
            public int SubfolderType { get; set; } = 0;
            public string FilenamePattern { get; set; } = "{YYYY-MM-DD} {Betreff} - {Absender}";
            public bool CreateLogFile { get; set; } = true;
            public bool PreserveFormatting { get; set; } = true;
            public bool EmbedImages { get; set; } = true;
            public ExportFormat Format { get; set; } = ExportFormat.PDF;
        }



        private string CleanHtmlContent(string htmlContent)
        {
            if (string.IsNullOrEmpty(htmlContent))
                return "<p>Kein Inhalt verfügbar</p>";

            // Nur gefährliche Script- und Style-Tags entfernen
            htmlContent = Regex.Replace(htmlContent, @"<script[^>]*>.*?</script>", "", RegexOptions.IgnoreCase | RegexOptions.Singleline);

            // Style-Tags entfernen, aber Inline-Styles beibehalten
            htmlContent = Regex.Replace(htmlContent, @"<style[^>]*>.*?</style>", "", RegexOptions.IgnoreCase | RegexOptions.Singleline);

            // Relative URLs zu absoluten URLs konvertieren
            htmlContent = Regex.Replace(htmlContent, @"src=""//", @"src=""https://");

            // Gefährliche Attribute entfernen
            htmlContent = Regex.Replace(htmlContent, @"on\w+\s*=\s*[""'][^""']*[""']", "", RegexOptions.IgnoreCase);
            htmlContent = Regex.Replace(htmlContent, @"javascript:", "", RegexOptions.IgnoreCase);

            return htmlContent;
        }

        private string EscapeHtml(string text)
        {
            if (string.IsNullOrEmpty(text))
                return "";

            return text
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("\"", "&quot;")
                .Replace("'", "&#39;");
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void InsertPlaceholder(string placeholder)
        {
            if (filenamePatternTextBox == null) return;
            int selectionIndex = filenamePatternTextBox.SelectionStart;
            filenamePatternTextBox.Text = filenamePatternTextBox.Text.Insert(selectionIndex, placeholder);
            filenamePatternTextBox.SelectionStart = selectionIndex + placeholder.Length;
        }

        private void btnDate_Click(object sender, EventArgs e)
        {
            InsertPlaceholder("{YYYY-MM-DD}");
        }

        private void btnSubject_Click(object sender, EventArgs e)
        {
            InsertPlaceholder("{Betreff}");
        }

        private void btnSender_Click(object sender, EventArgs e)
        {
            InsertPlaceholder("{Absender}");
        }

        private void filenamePatternTextBox_TextChanged(object sender, EventArgs e)
        {

        }

        private Outlook.MailItem? GetMailItemByEntryID(string entryID)
        {
            if (nameSpace == null) return null;
            try
            {
                var item = nameSpace.GetItemFromID(entryID);
                return item as Outlook.MailItem;
            }
            catch
            {
                return null;
            }
        }

        private void subfolderTypeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void InsertFolderPlaceholder(string placeholder)
        {
            if (folderPatternTextBox == null) return;
            int selectionIndex = folderPatternTextBox.SelectionStart;
            folderPatternTextBox.Text = folderPatternTextBox.Text.Insert(selectionIndex, placeholder);
            folderPatternTextBox.SelectionStart = selectionIndex + placeholder.Length;
        }

        private void btnFolderYYYY_Click(object sender, EventArgs e)
        {
            InsertFolderPlaceholder("{YYYY}");
        }

        private void btnFolderMM_Click(object sender, EventArgs e)
        {
            InsertFolderPlaceholder("{MM}");
        }

        private void btnFolderAbsender_Click(object sender, EventArgs e)
        {
            InsertFolderPlaceholder("{Absender}");
        }

        private void btnFolderBetreff_Click(object sender, EventArgs e)
        {
            InsertFolderPlaceholder("{Betreff}");
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e)
        {
            throw new System.NotImplementedException();
        }
        
        #region Autostart Funktionalität

/// <summary>
/// Aktiviert oder deaktiviert den Autostart über den Windows Startup-Ordner
/// </summary>
/// <param name="enable">True zum Aktivieren, False zum Deaktivieren</param>
private void SetAutostartViaStartupFolder(bool enable)
{
    try
    {
        string startupPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
        string shortcutPath = Path.Combine(startupPath, "OutlookArchivierungApp.lnk");

        if (enable)
        {
            CreateShortcut(shortcutPath, Application.ExecutablePath);
            MessageBox.Show("Autostart wurde aktiviert. Das Programm wird beim nächsten Windows-Start automatisch gestartet.", 
                           "Autostart aktiviert", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        else
        {
            if (File.Exists(shortcutPath))
            {
                File.Delete(shortcutPath);
                MessageBox.Show("Autostart wurde deaktiviert.", 
                               "Autostart deaktiviert", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
    catch (Exception ex)
    {
        MessageBox.Show($"Fehler beim Konfigurieren des Autostarts:\n{ex.Message}", 
                       "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}

/// <summary>
/// Prüft, ob der Autostart über den Startup-Ordner aktiviert ist
/// </summary>
/// <returns>True wenn aktiviert, False wenn nicht</returns>
private bool IsAutostartEnabledViaStartupFolder()
{
    try
    {
        string startupPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
        string shortcutPath = Path.Combine(startupPath, "OutlookArchivierungApp.lnk");
        return File.Exists(shortcutPath);
    }
    catch (Exception ex)
    {
        System.Diagnostics.Debug.WriteLine($"Fehler beim Prüfen des Autostarts: {ex.Message}");
        return false;
    }
}

/// <summary>
/// Erstellt eine Windows-Verknüpfung im angegebenen Pfad
/// </summary>
/// <param name="shortcutPath">Pfad zur Verknüpfungsdatei</param>
/// <param name="targetPath">Pfad zur Zielanwendung</param>
private void CreateShortcut(string shortcutPath, string targetPath)
{
    try
    {
        // COM-Objekt für Windows Shell verwenden
        Type shellType = Type.GetTypeFromProgID("WScript.Shell");
        if (shellType == null)
        {
            throw new Exception("Windows Shell COM-Objekt nicht verfügbar");
        }

        dynamic shell = Activator.CreateInstance(shellType);
        var shortcut = shell.CreateShortcut(shortcutPath);
        
        shortcut.TargetPath = targetPath;
        shortcut.WorkingDirectory = Path.GetDirectoryName(targetPath);
        shortcut.Description = "Outlook Archivierung App - Automatischer Start";
        shortcut.WindowStyle = 1; // Normal window
        
        shortcut.Save();
        
        // COM-Objekt freigeben
        System.Runtime.InteropServices.Marshal.ReleaseComObject(shortcut);
        System.Runtime.InteropServices.Marshal.ReleaseComObject(shell);
    }
    catch (Exception ex)
    {
        throw new Exception($"Konnte Verknüpfung nicht erstellen: {ex.Message}");
    }
}

/// <summary>
/// Event-Handler für die Autostart-CheckBox
/// </summary>
private void startWithWindowsCheckBox_CheckedChanged(object sender, EventArgs e)
{
    try
    {
        CheckBox checkBox = (CheckBox)sender;

        // Bestätigung anzeigen
        string message = checkBox.Checked
            ? "Möchten Sie das Programm wirklich beim Windows-Start automatisch starten?"
            : "Möchten Sie den automatischen Start wirklich deaktivieren?";

        DialogResult result = MessageBox.Show(message, "Autostart konfigurieren",
            MessageBoxButtons.YesNo, MessageBoxIcon.Question);

        if (result == DialogResult.Yes)
        {
            SetAutostartViaStartupFolder(checkBox.Checked);
            SaveSettings(); // Einstellungen sofort speichern
        }
        else
        {
            // CheckBox-Status zurücksetzen wenn abgebrochen
            checkBox.CheckedChanged -= startWithWindowsCheckBox_CheckedChanged;
            checkBox.Checked = !checkBox.Checked;
            checkBox.CheckedChanged += startWithWindowsCheckBox_CheckedChanged;
        }
    }
    catch (Exception ex)
    {
        MessageBox.Show($"Fehler beim Ändern der Autostart-Einstellung:\n{ex.Message}",
            "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}
/// <summary>
/// Öffnet den Windows Startup-Ordner im Explorer
/// </summary>
private void OpenStartupFolder()
{
    try
    {
        string startupPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
        Process.Start("explorer.exe", startupPath);
    }
    catch (Exception ex)
    {
        MessageBox.Show($"Konnte Startup-Ordner nicht öffnen:\n{ex.Message}", 
                       "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
    }
}

#endregion
    }
}
