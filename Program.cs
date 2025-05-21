using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Threading;
using System.Drawing;

using Path = System.IO.Path;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using iText.IO.Image;
using iText.Kernel.Geom;
using iText.Kernel.Font;
using iText.IO.Font.Constants;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf.Canvas.Parser;
using iText.Kernel.Pdf.Canvas.Parser.Listener;
using iTextImage = iText.Layout.Element.Image;

namespace KRA.FiscalizeTool
{
    public partial class MainForm : Form
    {
        // Folders configuration
        private readonly string QRCodeFolder = @"C:\FBtemp\QRCodes";
        private readonly string OutputFolder = @"C:\FBtemp\Modified";

        // List to track processed invoices
        private List<ProcessedInvoice> ProcessedInvoices = new List<ProcessedInvoice>();

        public MainForm()
        {
            InitializeComponent();

            // Ensure all required folders exist
            EnsureFoldersExist();
        }

        private void EnsureFoldersExist()
        {
            try
            {
                // Create all required folders if they don't exist
                Directory.CreateDirectory(QRCodeFolder);
                Directory.CreateDirectory(OutputFolder);

                LogMessage("Verified all required folders exist");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to create required folders: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LogMessage(string message)
        {
            if (listBoxInvoices.InvokeRequired)
            {
                // If called from a different thread, invoke on the UI thread
                listBoxInvoices.Invoke(new Action<string>(LogMessage), message);
                return;
            }

            // Regular UI update (now on the UI thread)
            listBoxInvoices.Items.Add($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}");
            listBoxInvoices.SelectedIndex = listBoxInvoices.Items.Count - 1;
            Application.DoEvents();
        }

        private void InitializeComponent()
        {
            this.btnLoadInvoices = new Button();
            this.btnProcessInvoices = new Button();
            this.listBoxInvoices = new ListBox();
            this.lblStatus = new Label();
            this.progressBar = new ProgressBar();

            // 
            // btnLoadInvoices
            // 
            this.btnLoadInvoices.Location = new System.Drawing.Point(12, 12);
            this.btnLoadInvoices.Name = "btnLoadInvoices";
            this.btnLoadInvoices.Size = new System.Drawing.Size(150, 30);
            this.btnLoadInvoices.TabIndex = 0;
            this.btnLoadInvoices.Text = "1. Load Invoices";
            this.btnLoadInvoices.UseVisualStyleBackColor = true;
            this.btnLoadInvoices.Click += new System.EventHandler(this.btnLoadInvoices_Click);
            // 
            // listBoxInvoices
            // 
            this.listBoxInvoices.FormattingEnabled = true;
            this.listBoxInvoices.HorizontalScrollbar = true;
            this.listBoxInvoices.Location = new System.Drawing.Point(12, 48);
            this.listBoxInvoices.Name = "listBoxInvoices";
            this.listBoxInvoices.Size = new System.Drawing.Size(760, 225);
            this.listBoxInvoices.TabIndex = 1;
            // 
            // btnProcessInvoices
            // 
            this.btnProcessInvoices.Location = new System.Drawing.Point(12, 279);
            this.btnProcessInvoices.Name = "btnProcessInvoices";
            this.btnProcessInvoices.Size = new System.Drawing.Size(150, 30);
            this.btnProcessInvoices.TabIndex = 2;
            this.btnProcessInvoices.Text = "2. Add QR Codes to PDFs";
            this.btnProcessInvoices.UseVisualStyleBackColor = true;
            this.btnProcessInvoices.Click += new System.EventHandler(this.btnProcessInvoices_Click);
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(168, 288);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(0, 13);
            this.lblStatus.TabIndex = 3;
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(12, 315);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(760, 23);
            this.progressBar.TabIndex = 4;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 350);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.btnProcessInvoices);
            this.Controls.Add(this.listBoxInvoices);
            this.Controls.Add(this.btnLoadInvoices);
            this.Name = "MainForm";
            this.Text = "PDF QR Code Addition Tool";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private void btnLoadInvoices_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "PDF Files|*.pdf";
                openFileDialog.Multiselect = true;
                openFileDialog.Title = "Select Invoice PDFs";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    ProcessedInvoices.Clear();
                    listBoxInvoices.Items.Clear();

                    foreach (string file in openFileDialog.FileNames)
                    {
                        LogMessage($"Added invoice: {Path.GetFileName(file)}");
                        ProcessedInvoices.Add(new ProcessedInvoice { FilePath = file, Status = "Pending" });
                    }
                }
            }
        }

        private async void btnProcessInvoices_Click(object sender, EventArgs e)
        {
            if (ProcessedInvoices.Count == 0)
            {
                MessageBox.Show("Please load invoices first.", "No Invoices", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Disable buttons during processing
            btnLoadInvoices.Enabled = false;
            btnProcessInvoices.Enabled = false;

            try
            {
                progressBar.Minimum = 0;
                progressBar.Maximum = ProcessedInvoices.Count;
                progressBar.Value = 0;

                int successCount = 0;
                int failCount = 0;

                lblStatus.Text = "Processing invoices...";
                LogMessage("Starting invoice processing");

                for (int i = 0; i < ProcessedInvoices.Count; i++)
                {
                    var invoice = ProcessedInvoices[i];
                    lblStatus.Text = $"Processing {Path.GetFileName(invoice.FilePath)} ({i + 1}/{ProcessedInvoices.Count})";
                    LogMessage($"Processing invoice: {Path.GetFileName(invoice.FilePath)}");
                    Application.DoEvents();

                    try
                    {
                        // Extract invoice number from PDF
                        string invoiceNumber = ExtractInvoiceNumberFromPdf(invoice.FilePath);

                        if (string.IsNullOrEmpty(invoiceNumber))
                        {
                            LogMessage($"Could not extract invoice number from {Path.GetFileName(invoice.FilePath)}");
                            invoice.Status = "Failed";
                            invoice.Message = "Could not extract invoice number";
                            failCount++;
                            continue;
                        }

                        LogMessage($"Found invoice number: {invoiceNumber}");

                        // Find the matching QR code based on naming pattern
                        string qrCodePath = FindQRCodeForInvoice(invoiceNumber);

                        if (string.IsNullOrEmpty(qrCodePath))
                        {
                            LogMessage($"No QR code found for invoice {invoiceNumber}");
                            invoice.Status = "Failed";
                            invoice.Message = "No matching QR code found";
                            failCount++;
                            continue;
                        }

                        LogMessage($"Found QR code at: {qrCodePath}");

                        // Add QR code to PDF
                        string modifiedPdfPath = AddQRCodeToPdf(invoice.FilePath, qrCodePath, invoiceNumber);

                        invoice.Status = "Success";
                        invoice.Message = $"Modified PDF: {Path.GetFileName(modifiedPdfPath)}";
                        successCount++;
                        LogMessage($"Successfully added QR code to PDF: {Path.GetFileName(modifiedPdfPath)}");
                    }
                    catch (Exception ex)
                    {
                        invoice.Status = "Failed";
                        invoice.Message = ex.Message;
                        LogMessage($"FAILED: {Path.GetFileName(invoice.FilePath)} - {ex.Message}");
                        failCount++;
                    }

                    progressBar.Value = i + 1;
                    Application.DoEvents();
                }

                lblStatus.Text = $"Completed: {successCount} succeeded, {failCount} failed";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Processing error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "Processing error";
            }
            finally
            {
                // Re-enable buttons
                btnLoadInvoices.Enabled = true;
                btnProcessInvoices.Enabled = true;
            }
        }

        private string ExtractInvoiceNumberFromPdf(string pdfPath)
        {
            try
            {
                string pdfText = GetTextFromPdf(pdfPath);
                string invoiceNumber = ExtractInvoiceNumber(pdfText, Path.GetFileName(pdfPath));
                return invoiceNumber;
            }
            catch (Exception ex)
            {
                LogMessage($"Error extracting invoice number: {ex.Message}");
                return null;
            }
        }

        private string GetTextFromPdf(string pdfPath)
        {
            StringBuilder sb = new StringBuilder();

            try
            {
                // Use iText7 API
                using (PdfReader reader = new PdfReader(pdfPath))
                using (PdfDocument pdfDoc = new PdfDocument(reader))
                {
                    // Process each page
                    for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
                    {
                        // Create a text extraction strategy
                        LocationTextExtractionStrategy strategy = new LocationTextExtractionStrategy();

                        // Extract text from the page
                        PdfCanvasProcessor parser = new PdfCanvasProcessor(strategy);
                        parser.ProcessPageContent(pdfDoc.GetPage(i));

                        // Get the extracted text
                        string pageText = strategy.GetResultantText();
                        sb.AppendLine(pageText);
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error extracting text from PDF: {ex.Message}");
                throw;
            }

            return sb.ToString();
        }

        private string ExtractInvoiceNumber(string pdfText, string fileName)
        {
            LogMessage("Attempting to extract invoice number");

            // Try several patterns to find invoice number
            var patterns = new[]
            {
                @"Sales\s+Invoice\s+No\.\s*:\s*(KE\d{8})",
                @"Invoice\s+No\.\s*:\s*(KE\d{8})",
                @"No\.\s*:\s*(KE\d{8})",
                @"(KE\d{8})"
            };

            foreach (var pattern in patterns)
            {
                Match match = Regex.Match(pdfText, pattern);
                if (match.Success)
                {
                    LogMessage($"Found invoice number using pattern: {pattern}");
                    return match.Groups[1].Value;
                }
            }

            // If not found in text, try to extract from filename
            LogMessage("No invoice number pattern found in text, trying filename");
            Match fileMatch = Regex.Match(fileName, @"(KE\d{8})");
            if (fileMatch.Success)
            {
                LogMessage($"Extracted invoice number from filename: {fileMatch.Groups[1].Value}");
                return fileMatch.Groups[1].Value;
            }

            return null;
        }

        private string FindQRCodeForInvoice(string invoiceNumber)
        {
            LogMessage($"Looking for QR code for invoice {invoiceNumber}");

            // Pattern for QR code filenames: QR_KE00001017_20250520171712.png
            string pattern = $"QR_{invoiceNumber}_*.png";
            
            // Get all matching files
            string[] matchingFiles = Directory.GetFiles(QRCodeFolder, pattern);

            if (matchingFiles.Length == 0)
            {
                LogMessage($"No matching QR code files found for {invoiceNumber}");
                return null;
            }

            // Sort by newest (assuming timestamp in filename) and take the first
            var newestFile = matchingFiles.OrderByDescending(f => f).FirstOrDefault();
            LogMessage($"Found QR code file: {Path.GetFileName(newestFile)}");
            
            return newestFile;
        }

        private string AddQRCodeToPdf(string originalPdfPath, string qrCodePath, string invoiceNumber)
        {
            LogMessage($"Adding QR code to PDF for invoice {invoiceNumber}");

            // Define output path - use unique filename to avoid conflicts
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            string outputPdfPath = Path.Combine(OutputFolder, $"Modified_{invoiceNumber}_{timestamp}.pdf");
            string tempPdfPath = Path.Combine(Path.GetTempPath(), $"Temp_{invoiceNumber}_{Guid.NewGuid()}.pdf");

            try
            {
                // Ensure the directory exists
                Directory.CreateDirectory(Path.GetDirectoryName(outputPdfPath));
                LogMessage($"Verified output directory exists: {Path.GetDirectoryName(outputPdfPath)}");

                // Verify input files exist
                if (!File.Exists(originalPdfPath))
                    throw new FileNotFoundException($"Original PDF file not found: {originalPdfPath}");
                if (!File.Exists(qrCodePath))
                    throw new FileNotFoundException($"QR code image not found: {qrCodePath}");

                // Copy the PDF to temp location
                File.Copy(originalPdfPath, tempPdfPath, true);
                LogMessage($"Created temporary copy of PDF at {tempPdfPath}");

                // Ensure output file doesn't exist
                if (File.Exists(outputPdfPath))
                {
                    File.Delete(outputPdfPath);
                    LogMessage($"Deleted existing output file: {outputPdfPath}");
                }

                // Standard PDF processing approach
                using (PdfReader pdfReader = new PdfReader(tempPdfPath))
                using (PdfWriter pdfWriter = new PdfWriter(outputPdfPath))
                using (PdfDocument pdfDoc = new PdfDocument(pdfReader, pdfWriter))
                {
                    Document document = new Document(pdfDoc);

                    // Get the number of pages
                    int numberOfPages = pdfDoc.GetNumberOfPages();
                    LogMessage($"PDF has {numberOfPages} pages");

                    // Target the last page
                    PdfPage lastPage = pdfDoc.GetPage(numberOfPages);

                    // Load QR code image
                    ImageData imageData = ImageDataFactory.Create(qrCodePath);
                    iText.Layout.Element.Image qrCodeImage = new iText.Layout.Element.Image(imageData);

                    // Set size and position - smaller size and position in bottom left
                    qrCodeImage.ScaleToFit(50, 50);
                    qrCodeImage.SetFixedPosition(5, 20);

                    // Add the image to the document
                    document.Add(qrCodeImage);

                    // Add invoice number text next to QR code
                    PdfFont boldFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
                    Paragraph footerParagraph = new Paragraph($"Invoice: {invoiceNumber}");
                    footerParagraph.SetFont(boldFont);
                    footerParagraph.SetFontSize(6);

                    // Position and add the text
                    document.ShowTextAligned(
                        footerParagraph,
                        60, // x position
                        20, // y position
                        numberOfPages,
                        TextAlignment.LEFT,
                        VerticalAlignment.BOTTOM,
                        0  // rotation
                    );

                    // Close the document explicitly
                    document.Close();
                    LogMessage("Document closed successfully");
                }

                // Clean up the temporary file
                try
                {
                    if (File.Exists(tempPdfPath))
                    {
                        File.Delete(tempPdfPath);
                        LogMessage($"Deleted temporary file: {tempPdfPath}");
                    }
                }
                catch (Exception cleanupEx)
                {
                    LogMessage($"Warning: Failed to clean up temporary file: {cleanupEx.Message}");
                }

                LogMessage($"Successfully modified PDF: {outputPdfPath}");
                return outputPdfPath;
            }
            catch (Exception ex)
            {
                LogMessage($"Error modifying PDF: {ex.Message}");
                throw;
            }
        }

        // UI Controls
        private Button btnLoadInvoices;
        private Button btnProcessInvoices;
        private ListBox listBoxInvoices;
        private Label lblStatus;
        private ProgressBar progressBar;
    }

    public class ProcessedInvoice
    {
        public string FilePath { get; set; }
        public string Status { get; set; }
        public string Message { get; set; }
    }

    static class Program
    {
        [STAThread]
        static void Main()
        {
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainForm());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Application error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}