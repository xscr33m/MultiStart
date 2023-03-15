namespace MultiStart
{
    public partial class MAIN : Form
    {
        static public string hiddenFolderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "MultiStart", "Data");

        public MAIN()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            if (!Directory.Exists(hiddenFolderPath))
            {
                btnOpenPath.Enabled = false;
            }
            btnMultiStart.Enabled = false;
            CheckIfReadyForMultiStart();
        }

        private void CheckIfReadyForMultiStart()
        {
            if (!string.IsNullOrEmpty(tbMultiStartName.Text) && !string.IsNullOrEmpty(tbSavePath.Text) && pictureBox1.Image != null && lstVerknuepfungen.Items.Count > 0)
            {
                btnMultiStart.Enabled = true;
            }
            else
            {
                btnMultiStart.Enabled = false;
            }
        }

        private void btnHinzufuegen_Click(object sender, EventArgs e)
        {
            // Erzeuge ein OpenFileDialog-Objekt und konfiguriere es
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Verkn�pfungen (*.lnk;*.url)|*.lnk;*.url|Anwendungen (*.exe)|*.exe";
            openFileDialog.Multiselect = true;
            openFileDialog.Title = "Verkn�pfung oder Programm ausw�hlen";

            // �ffne das OpenFileDialog und f�ge Verkn�pfungen zur ListView hinzu
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Erstelle ein ListViewItem-Objekt mit dem Namen der Verkn�pfung und dem Pfad zur Verkn�pfung
                var item = new ListViewItem();
                item.Text = System.IO.Path.GetFileNameWithoutExtension(openFileDialog.FileName);
                item.SubItems.Add(openFileDialog.FileName);

                // Extrahiere das Icon aus der Verkn�pfung und f�ge es zur ImageList hinzu
                var icon = Icon.ExtractAssociatedIcon(openFileDialog.FileName);
                if (icon != null)
                {
                    lstVerknuepfungen.SmallImageList.Images.Add(icon);
                    item.ImageIndex = lstVerknuepfungen.SmallImageList.Images.Count - 1;
                }

                // F�ge das ListViewItem zur ListView hinzu
                lstVerknuepfungen.Items.Add(item);
            }
            CheckIfReadyForMultiStart();
        }

        private void btnEntfernen_Click(object sender, EventArgs e)
        {
            // �berpr�fe, ob ein ListViewItem ausgew�hlt wurde
            if (lstVerknuepfungen.SelectedItems.Count > 0)
            {
                // Entferne jedes ausgew�hlte ListViewItem aus der ListView
                foreach (ListViewItem item in lstVerknuepfungen.SelectedItems)
                {
                    lstVerknuepfungen.Items.Remove(item);
                }
            }
        }

        private void btnMultiStart_Click(object sender, EventArgs e)
        {
            if (lstVerknuepfungen.Items.Count == 0)
            {
                MessageBox.Show("Es wurden keine Verkn�pfungen hinzugef�gt.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (string.IsNullOrWhiteSpace(tbMultiStartName.Text))
            {
                MessageBox.Show("Bitte geben Sie einen Namen f�r die Startdatei ein.", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Erstelle den Ordner, falls er noch nicht existiert
            Directory.CreateDirectory(hiddenFolderPath);

            // Erstelle die Datei am ausgew�hlten Speicherort
            string saveFilePath = Path.Combine(hiddenFolderPath, $"{tbMultiStartName.Text}.bat");

            // �ffne die Datei zum Schreiben
            using (StreamWriter sw = new StreamWriter(saveFilePath))
            {
                // Schreibe jede Verkn�pfung in die Batch-Datei
                foreach (ListViewItem item in lstVerknuepfungen.Items)
                {
                    sw.WriteLine($"start \"\" \"{item.SubItems[1].Text}\"");
                }
            }

            string batchFilePath = Path.Combine(hiddenFolderPath, $"{tbMultiStartName.Text}.bat");
            CreateShortcut(tbSavePath.Text, batchFilePath);

            MessageBox.Show("MultiStart Verkn�pfung wurde auf dem Desktop abgelegt!");
        }

        private void tbSavePath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.RootFolder = Environment.SpecialFolder.Desktop;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                tbSavePath.Text = dialog.SelectedPath;
            }
        }

        private void CreateShortcut(string saveFolderPath, string batchFilePath)
        {
            // Erstelle ein WSH-Objekt
            var wsh = new IWshRuntimeLibrary.WshShell();

            // Erstelle eine Verkn�pfung-Datei mit dem Namen der Batch-Datei
            string shortcutPath = Path.Combine(saveFolderPath, $"{Path.GetFileNameWithoutExtension(batchFilePath)}.lnk");
            var shortcut = (IWshRuntimeLibrary.IWshShortcut)wsh.CreateShortcut(shortcutPath);

            // Konfiguriere die Verkn�pfung-Eigenschaften
            shortcut.TargetPath = batchFilePath;
            shortcut.WorkingDirectory = Path.GetDirectoryName(batchFilePath);
            shortcut.WindowStyle = 1; // Normal-Fenster
            shortcut.Description = "MultiStart-Verkn�pfung";
            shortcut.IconLocation = selectedIconPath;


            // Speichere die Verkn�pfung-Datei
            shortcut.Save();
        }

        private void btnOpenPath_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("explorer.exe", hiddenFolderPath);
        }

        private string? selectedIconPath; // Klassenvariable zum Speichern des Pfads zum ausgew�hlten Icons

        private void btnIcon_Click(object sender, EventArgs e)
        {
            // Erzeuge ein OpenFileDialog-Objekt und konfiguriere es
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Symbole (*.ico)|*.ico";
            openFileDialog.Title = "Symbol ausw�hlen";

            // �ffne das OpenFileDialog und f�ge das ausgew�hlte Icon zur PictureBox hinzu
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                selectedIconPath = openFileDialog.FileName;
                pictureBox1.Image = Image.FromFile(selectedIconPath);
            }
            CheckIfReadyForMultiStart();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            // Pr�fe, ob das PictureBox-Steuerelement ein Bild enth�lt
            if (pictureBox1.Image != null)
            {
                // Entferne das Bild aus dem PictureBox-Steuerelement
                pictureBox1.Image.Dispose();
                pictureBox1.Image = null;
                pictureBox1.BorderStyle = BorderStyle.None;
                Cursor = Cursors.Default;
            }
            CheckIfReadyForMultiStart();
        }

        private void pictureBox1_MouseHover(object sender, EventArgs e)
        {
            if (pictureBox1.Image != null)
            {
                Cursor = Cursors.No;
                pictureBox1.BorderStyle = (BorderStyle)Border3DStyle.RaisedOuter;
            }
        }

        private void pictureBox1_MouseLeave(object sender, EventArgs e)
        {
            Cursor = Cursors.Default;
            pictureBox1.BorderStyle = BorderStyle.None;
        }

        private void tbMultiStartName_TextChanged(object sender, EventArgs e)
        {
            CheckIfReadyForMultiStart();
        }

        private void tbSavePath_TextChanged(object sender, EventArgs e)
        {
            CheckIfReadyForMultiStart();
        }
    }
}