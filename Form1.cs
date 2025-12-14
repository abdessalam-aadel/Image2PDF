using iText.IO.Image;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Kernel.Geom;
using System.Diagnostics;

namespace Image2PDF
{
    public partial class frmMain : Form
    {

        string[] IMGfiles;
        int fileCount = 0;
        string selected_path;
        string csv_path;
        public frmMain() => InitializeComponent();
        public static int SearchDirectoryTree(string path, out string[] IMGfiles)
        {
            string[] extensions = { ".jpg", ".jpeg", ".png", ".tiff" };

            IMGfiles = Directory
                .EnumerateFiles(path, "*.*", SearchOption.TopDirectoryOnly)
                .Where(f => extensions.Contains(
                    System.IO.Path.GetExtension(f),
                    StringComparer.OrdinalIgnoreCase))
                .ToArray();
            return IMGfiles.Length;
        }

        bool ScaleToFit = false;
        bool ScaleAbsolute = false;

        private void btnLoadFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog FD = new FolderBrowserDialog();
            if (selected_path != null)
                FD.SelectedPath = selected_path;

            if (FD.ShowDialog() == DialogResult.OK)
            {
                selected_path = FD.SelectedPath;
                txtBoxLoadIMG.Text = FD.SelectedPath;
                fileCount = SearchDirectoryTree(FD.SelectedPath, out IMGfiles);
                // Check the Empty Folder
                lblFileCount.Text = fileCount == 0 ? "Your Folder is Empty" : fileCount + " files.";
                // Clear the Alert message and success message
                lblDone.Text = "";
                lblAlert.Text = "";
                //IMGDone.Visible = false;
            }
        }

        private void btnLoadCSV_Click(object sender, EventArgs e)
        {
            OpenFileDialog OFD = new OpenFileDialog
            {
                InitialDirectory = "C:/",
                RestoreDirectory = true,
                Filter = "(*.csv) | *.csv",
                Title = "Chose your file"
            };
            DialogResult result = OFD.ShowDialog();
            if (result == DialogResult.OK)
            {
                txtBoxLoadCSV.Text = OFD.FileName;
                csv_path = OFD.FileName;
            }
        }
        List<CsvEntry> LoadCsv(string csvPath)
        {
            var list = new List<CsvEntry>();

            foreach (var line in File.ReadAllLines(csvPath).Skip(1)) // skip header
            {
                var parts = line.Split(',');

                list.Add(new CsvEntry
                {
                    ImageName = parts[0].Trim(),
                    X = parts[1].Trim(),
                    Y = parts[2].Trim()
                });
            }

            return list;
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            // Clear the Alert message and success message
            lblDone.Text = "";
            lblAlert.Text = "";

            // Check if the user has been selected a folder
            if (IMGfiles == null)
            {
                lblAlert.Text = "Please select or Drag your folder and try again!";
                return;
            }

            // Check the Empty Folder
            if (fileCount == 0)
            {
                lblAlert.Text = "Your Folder is Empty";
                return;
            }

            if(csv_path == null)
            {
                lblAlert.Text = "Please select your CSV file and try again!";
                return;
            }

            Cursor = Cursors.WaitCursor;

            try
            {
                var csvData = LoadCsv(txtBoxLoadCSV.Text);

                foreach (string imagePath in IMGfiles)
                {
                    string imageName = System.IO.Path.GetFileNameWithoutExtension(imagePath);
                    var csvRow = csvData
                                .FirstOrDefault(c => c.ImageName.Equals(imageName, StringComparison.OrdinalIgnoreCase));

                    if (csvRow == null)
                        continue; // no coordinates for this image

                    string pdfPath = System.IO.Path.Combine(
                        selected_path,
                        imageName + ".pdf"
                    );

                    PdfWriter writer = new PdfWriter(pdfPath);
                    PdfDocument pdf = new PdfDocument(writer);
                    Document document = new Document(pdf, PageSize.A4);

                    // Load image
                    ImageData imageData = ImageDataFactory.Create(imagePath);
                    iText.Layout.Element.Image image = new iText.Layout.Element.Image(imageData);

                    //image.SetFixedPosition(0, 0);

                    document.SetMargins(10, 10, 10, 10);

                    float usableWidth = PageSize.A4.GetWidth()
                                        - document.GetLeftMargin()
                                        - document.GetRightMargin();

                    float usableHeight = PageSize.A4.GetHeight()
                                        - document.GetTopMargin()
                                        - document.GetBottomMargin();
                    if (ScaleToFit)
                        image.ScaleToFit(usableWidth, usableHeight);
                    if (ScaleAbsolute)
                        image.ScaleAbsolute(usableWidth, usableHeight);

                    // Add image to document
                    document.Add(image);

                    // Add text positioned over the image
                    Paragraph text = new Paragraph(csvRow.ImageName + "\n X = " + csvRow.X + ", Y = " + csvRow.Y)
                        .SetFontSize(12)
                        .SetFontColor(iText.Kernel.Colors.ColorConstants.BLACK);

                    // Absolute positioning (X, Y)
                    text.SetFixedPosition(
                        1,      // page number
                        50,    // X position
                        750,    // Y position
                        500     // width
                    );

                    document.Add(text);

                    document.Close();
                }
            }
            catch (Exception ex)
            {
                // Message Exception
                MessageBox.Show(ex.Message, "Error Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Cursor = Cursors.Default;
                IMGfiles = null;
                txtBoxLoadIMG.Text = "Select your images folder ...";
                txtBoxLoadCSV.Text = "Select your CSV file ...";
                lblFileCount.Text = "...";
                return;
            }

            Cursor = Cursors.Default;
            lblDone.Text = "Done";
            IMGfiles = null;
            txtBoxLoadIMG.Text = "Select your images folder ...";
            txtBoxLoadCSV.Text = "Select your CSV file ...";
            lblFileCount.Text = "...";

        }

        private void comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox.Text == "ScaleToFit")
                ScaleToFit = true;
            if (comboBox.Text == "ScaleAbsolute")
                ScaleAbsolute = true;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // Go to Github repository
            string url = "https://github.com/abdessalam-aadel";

            // Open the URL in the default web browser
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true // Ensures the URL is opened in the default web browser
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}");
            }
        }
    }

    class CsvEntry
    {
        public string ImageName { get; set; }
        public string X { get; set; }
        public string Y { get; set; }
    }
}
