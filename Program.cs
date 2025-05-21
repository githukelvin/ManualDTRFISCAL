using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Threading;
using System.Drawing;
using System.Diagnostics;
using System.Data.SQLite;
using System.Data;

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
        private readonly string LogFolder = @"C:\FBtemp\Logs";
        private string LogFilePath;

        // List to track processed invoices
        private List<ProcessedInvoice> ProcessedInvoices = new List<ProcessedInvoice>();

        public MainForm()
        {
            InitializeComponent();

            // Initialize logging
            InitializeLogging();

            // Ensure all required folders exist
            EnsureFoldersExist();
            
            WriteLog("Application started", true);
            WriteLog($"iText.Kernel.Pdf version: {typeof(PdfDocument).Assembly.GetName().Version}", true);
        }

        private void InitializeLogging()
        {
            try
            {
                // Create log directory if it doesn't exist
                Directory.CreateDirectory(LogFolder);
                
                // Create a unique log file name with timestamp
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                LogFilePath = Path.Combine(LogFolder, $"QRCode_Log_{timestamp}.txt");
                
                // Write initial log entry
                File.WriteAllText(LogFilePath, $"=== PDF QR Code Tool Log Started at {DateTime.Now} ===\r\n");
                
                WriteLog("Logging initialized successfully", true);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to initialize logging: {ex.Message}", "Logging Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void WriteLog(string message, bool verbose = false)
        {
            string logEntry = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}";
            
            try
            {
                // Write to log file
                File.AppendAllText(LogFilePath, logEntry + "\r\n");
                
                // If it's a verbose message, prepend with [VERBOSE]
                if (verbose)
                {
                    logEntry = "[VERBOSE] " + logEntry;
                }
                
                // Also add to UI log if not too verbose
                if (!verbose || message.Contains("ERROR") || message.Contains("FAILED"))
                {
                    LogMessage(message);
                }
            }
            catch (Exception ex)
            {
                // If logging fails, at least try to show in the UI
                LogMessage($"Failed to write to log file: {ex.Message}");
            }
        }

        private void LogException(Exception ex, string context)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine($"ERROR in {context}: {ex.GetType().Name}: {ex.Message}");
            sb.AppendLine($"Stack trace: {ex.StackTrace}");
            
            // Log inner exception if present
            if (ex.InnerException != null)
            {
                sb.AppendLine($"Inner exception: {ex.InnerException.GetType().Name}: {ex.InnerException.Message}");
                sb.AppendLine($"Inner stack trace: {ex.InnerException.StackTrace}");
            }
            
            WriteLog(sb.ToString(), true);
        }

        private void EnsureFoldersExist()
        {
            try
            {
                // Create all required folders if they don't exist
                Directory.CreateDirectory(QRCodeFolder);
                WriteLog($"Verified QR code folder exists: {QRCodeFolder}", true);
                
                Directory.CreateDirectory(OutputFolder);
                WriteLog($"Verified output folder exists: {OutputFolder}", true);
                
                LogMessage("Verified all required folders exist");
            }
            catch (Exception ex)
            {
                LogException(ex, "EnsureFoldersExist");
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
            WriteLog("Load Invoices button clicked", true);
            
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "PDF Files|*.pdf";
                openFileDialog.Multiselect = true;
                openFileDialog.Title = "Select Invoice PDFs";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    ProcessedInvoices.Clear();
                    listBoxInvoices.Items.Clear();
                    
                    WriteLog($"User selected {openFileDialog.FileNames.Length} files", true);

                    foreach (string file in openFileDialog.FileNames)
                    {
                        LogMessage($"Added invoice: {Path.GetFileName(file)}");
                        WriteLog($"Added file: {file}", true);
                        ProcessedInvoices.Add(new ProcessedInvoice { FilePath = file, Status = "Pending" });
                    }
                }
                else
                {
                    WriteLog("User canceled file selection", true);
                }
            }
        }

        private async void btnProcessInvoices_Click(object sender, EventArgs e)
        {
            WriteLog("Process Invoices button clicked", true);
            
            if (ProcessedInvoices.Count == 0)
            {
                MessageBox.Show("Please load invoices first.", "No Invoices", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                WriteLog("No invoices loaded, processing canceled", true);
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
                WriteLog($"Starting processing of {ProcessedInvoices.Count} invoices", true);

                for (int i = 0; i < ProcessedInvoices.Count; i++)
                {
                    var invoice = ProcessedInvoices[i];
                    lblStatus.Text = $"Processing {Path.GetFileName(invoice.FilePath)} ({i + 1}/{ProcessedInvoices.Count})";
                    LogMessage($"Processing invoice: {Path.GetFileName(invoice.FilePath)}");
                    WriteLog($"Processing invoice {i+1}/{ProcessedInvoices.Count}: {invoice.FilePath}", true);
                    Application.DoEvents();

                    try
                    {
                        // Extract invoice number from PDF
                        WriteLog($"Attempting to extract invoice number from {invoice.FilePath}", true);
                        string invoiceNumber = ExtractInvoiceNumberFromPdf(invoice.FilePath);

                        if (string.IsNullOrEmpty(invoiceNumber))
                        {
                            LogMessage($"Could not extract invoice number from {Path.GetFileName(invoice.FilePath)}");
                            WriteLog($"Failed to extract invoice number from {invoice.FilePath}", true);
                            invoice.Status = "Failed";
                            invoice.Message = "Could not extract invoice number";
                            failCount++;
                            continue;
                        }

                        LogMessage($"Found invoice number: {invoiceNumber}");
                        WriteLog($"Successfully extracted invoice number: {invoiceNumber}", true);

                        // Find the matching QR code based on naming pattern
                        WriteLog($"Looking for QR code for invoice {invoiceNumber}", true);
                        string qrCodePath = FindQRCodeForInvoice(invoiceNumber);

                        if (string.IsNullOrEmpty(qrCodePath))
                        {
                            LogMessage($"No QR code found for invoice {invoiceNumber}");
                            WriteLog($"No matching QR code found for invoice {invoiceNumber}", true);
                            invoice.Status = "Failed";
                            invoice.Message = "No matching QR code found";
                            failCount++;
                            continue;
                        }

                        LogMessage($"Found QR code at: {qrCodePath}");
                        WriteLog($"Found QR code at: {qrCodePath}", true);

                        // Try multiple approaches to add QR code to PDF
                        string modifiedPdfPath = null;
                        string errorMessage = null;

                        // Try the first approach - add QR to existing PDF
                        try
                        {
                            WriteLog("Attempting Method 1: Direct PDF modification", true);
                            modifiedPdfPath = AddQRCodeToPdf_Method1(invoice.FilePath, qrCodePath, invoiceNumber);
                            WriteLog($"Method 1 succeeded, created: {modifiedPdfPath}", true);
                        }
                        catch (Exception ex)
                        {
                            LogMessage($"First method failed: {ex.Message}");
                            LogException(ex, "Method 1 - Direct PDF modification");
                            errorMessage = ex.Message;
                            
                            // Try the second approach - create a cover page
                            try
                            {
                                LogMessage("Trying alternative method...");
                                WriteLog("Attempting Method 2: Creating cover page", true);
                                modifiedPdfPath = CreateCoverPageWithQRCode(invoice.FilePath, qrCodePath, invoiceNumber);
                                WriteLog($"Method 2 succeeded, created: {modifiedPdfPath}", true);
                                errorMessage = null; // Clear error if second method succeeded
                            }
                            catch (Exception ex2)
                            {
                                LogMessage($"Second method also failed: {ex2.Message}");
                                LogException(ex2, "Method 2 - Creating cover page");
                                errorMessage += $" | Alternative method: {ex2.Message}";
                            }
                        }

                        if (!string.IsNullOrEmpty(modifiedPdfPath))
                        {
                            invoice.Status = "Success";
                            invoice.Message = $"Modified PDF: {Path.GetFileName(modifiedPdfPath)}";
                            successCount++;
                            LogMessage($"Successfully added QR code to PDF: {Path.GetFileName(modifiedPdfPath)}");
                            WriteLog($"Successfully processed invoice: {invoice.FilePath} -> {modifiedPdfPath}", true);
                        }
                        else
                        {
                            invoice.Status = "Failed";
                            invoice.Message = errorMessage ?? "Unknown error occurred";
                            failCount++;
                            LogMessage($"Failed to add QR code to PDF: {errorMessage}");
                            WriteLog($"Failed to process invoice: {invoice.FilePath}, errors: {errorMessage}", true);
                        }
                    }
                    catch (Exception ex)
                    {
                        invoice.Status = "Failed";
                        invoice.Message = ex.Message;
                        LogMessage($"FAILED: {Path.GetFileName(invoice.FilePath)} - {ex.Message}");
                        LogException(ex, $"Processing invoice {Path.GetFileName(invoice.FilePath)}");
                        failCount++;
                    }

                    progressBar.Value = i + 1;
                    Application.DoEvents();
                }

                string completionMessage = $"Completed: {successCount} succeeded, {failCount} failed";
                lblStatus.Text = completionMessage;
                WriteLog(completionMessage, true);
            }
            catch (Exception ex)
            {
                string errorMsg = $"Processing error: {ex.Message}";
                MessageBox.Show(errorMsg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "Processing error";
                LogException(ex, "Main processing loop");
            }
            finally
            {
                // Re-enable buttons
                btnLoadInvoices.Enabled = true;
                btnProcessInvoices.Enabled = true;
                WriteLog("Processing completed, UI re-enabled", true);
            }
        }

        private string ExtractInvoiceNumberFromPdf(string pdfPath)
        {
            try
            {
                WriteLog($"Starting text extraction from PDF: {pdfPath}", true);
                
                string pdfText = GetTextFromPdf(pdfPath);
                WriteLog($"Extracted {pdfText.Length} characters from PDF", true);
                
                string invoiceNumber = ExtractInvoiceNumber(pdfText, Path.GetFileName(pdfPath));
                
                if (!string.IsNullOrEmpty(invoiceNumber))
                {
                    WriteLog($"Successfully extracted invoice number: {invoiceNumber}", true);
                }
                else
                {
                    WriteLog("Failed to extract invoice number from text", true);
                }
                
                return invoiceNumber;
            }
            catch (Exception ex)
            {
                LogMessage($"Error extracting invoice number: {ex.Message}");
                LogException(ex, "ExtractInvoiceNumberFromPdf");
                return null;
            }
        }

        private string GetTextFromPdf(string pdfPath)
        {
            StringBuilder sb = new StringBuilder();

            try
            {
                WriteLog($"Opening PDF for text extraction: {pdfPath}", true);
                
                // Use iText7 API with simplified error handling
                using (PdfReader reader = new PdfReader(pdfPath))
                {
                    WriteLog("PdfReader created successfully", true);
                    
                    using (PdfDocument pdfDoc = new PdfDocument(reader))
                    {
                        WriteLog($"PdfDocument opened, page count: {pdfDoc.GetNumberOfPages()}", true);
                        
                        // Process each page
                        for (int i = 1; i <= pdfDoc.GetNumberOfPages(); i++)
                        {
                            WriteLog($"Processing page {i}/{pdfDoc.GetNumberOfPages()}", true);
                            
                            try
                            {
                                // Create a text extraction strategy
                                LocationTextExtractionStrategy strategy = new LocationTextExtractionStrategy();

                                // Extract text from the page
                                PdfCanvasProcessor parser = new PdfCanvasProcessor(strategy);
                                parser.ProcessPageContent(pdfDoc.GetPage(i));

                                // Get the extracted text
                                string pageText = strategy.GetResultantText();
                                sb.AppendLine(pageText);
                                
                                WriteLog($"Extracted {pageText.Length} characters from page {i}", true);
                            }
                            catch (Exception ex)
                            {
                                // Log the error but continue with next page
                                LogException(ex, $"Text extraction from page {i}");
                            }
                        }
                    }
                }
                
                WriteLog("PDF text extraction completed successfully", true);
            }
            catch (IOException ex)
            {
                LogException(ex, "PDF I/O Error");
                throw;
            }
            catch (Exception ex)
            {
                LogException(ex, "PDF Text Extraction");
                throw;
            }

            return sb.ToString();
        }

        private string ExtractInvoiceNumber(string pdfText, string fileName)
        {
            WriteLog("Attempting to extract invoice number from text", true);

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
                WriteLog($"Trying pattern: {pattern}", true);
                Match match = Regex.Match(pdfText, pattern);
                if (match.Success)
                {
                    LogMessage($"Found invoice number using pattern: {pattern}");
                    WriteLog($"Pattern matched: {pattern}, found: {match.Groups[1].Value}", true);
                    return match.Groups[1].Value;
                }
            }

            // If not found in text, try to extract from filename
            WriteLog("No invoice number pattern found in text, trying filename", true);
            Match fileMatch = Regex.Match(fileName, @"(KE\d{8})");
            if (fileMatch.Success)
            {
                LogMessage($"Extracted invoice number from filename: {fileMatch.Groups[1].Value}");
                WriteLog($"Found invoice number in filename: {fileMatch.Groups[1].Value}", true);
                return fileMatch.Groups[1].Value;
            }

            WriteLog("Failed to extract invoice number from both text and filename", true);
            return null;
        }

        private string FindQRCodeForInvoice(string invoiceNumber)
        {
            WriteLog($"Looking for QR code for invoice {invoiceNumber}", true);

            // Pattern for QR code filenames: QR_KE00001017_20250520171712.png
            string pattern = $"QR_{invoiceNumber}_*.png";
            
            try
            {
                WriteLog($"Searching for files matching pattern: {pattern} in {QRCodeFolder}", true);
                
                // Get all matching files
                string[] matchingFiles = Directory.GetFiles(QRCodeFolder, pattern);
                WriteLog($"Found {matchingFiles.Length} matching files", true);

                if (matchingFiles.Length == 0)
                {
                    LogMessage($"No matching QR code files found for {invoiceNumber}");
                    WriteLog("No matching QR code files found", true);
                    return null;
                }

                // Sort by newest (assuming timestamp in filename) and take the first
                var newestFile = matchingFiles.OrderByDescending(f => f).FirstOrDefault();
                LogMessage($"Found QR code file: {Path.GetFileName(newestFile)}");
                WriteLog($"Selected newest file: {newestFile}", true);
                
                // Verify file exists and is readable
                if (File.Exists(newestFile))
                {
                    WriteLog($"Verified QR code file exists: {newestFile}", true);
                    
                    // Check file can be read
                    try
                    {
                        using (var stream = File.OpenRead(newestFile))
                        {
                            WriteLog($"QR code file is readable, size: {stream.Length} bytes", true);
                        }
                    }
                    catch (Exception ex)
                    {
                        LogException(ex, "QR code file read test");
                        return null;
                    }
                    
                    return newestFile;
                }
                else
                {
                    WriteLog($"WARNING: QR code file not found: {newestFile}", true);
                    return null;
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error searching for QR codes: {ex.Message}");
                LogException(ex, "FindQRCodeForInvoice");
                return null;
            }
        }
        
        private QrCodeData GetQrCodeDataFromDatabase(string invoiceNumber)
        {
            WriteLog($"Querying database for invoice number: {invoiceNumber}", true);
            
            // Remove "KE" prefix if present to get the numeric part
            string numericPart = invoiceNumber;
            if (invoiceNumber.StartsWith("KE"))
            {
                numericPart = invoiceNumber.Substring(2);
            }
            
            // Pad with leading zeros if needed
            numericPart = numericPart.TrimStart('0');
            
            string connectionString = @"Data Source=C:\FBtemp\DB\FbTransaction.db;Version=3;";
            
            try
            {
                using (SQLiteConnection connection = new SQLiteConnection(connectionString))
                {
                    connection.Open();
                    WriteLog("Database connection opened successfully", true);
                    
                    string query = "SELECT Date, TsNum, ControlCode, SerialNumber FROM fb_transaction WHERE TsNum = @TsNum";
                    
                    using (SQLiteCommand command = new SQLiteCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@TsNum", invoiceNumber);
                        
                        using (SQLiteDataReader reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                QrCodeData data = new QrCodeData
                                {
                                    Date = reader["Date"].ToString(),
                                    TsNum = reader["TsNum"].ToString(),
                                    ControlCode = reader["ControlCode"].ToString(),
                                    SerialNumber = reader["SerialNumber"].ToString()
                                };
                                
                                WriteLog($"Found data in database: TSIN={data.TsNum}, CUIN={data.ControlCode}, CUSN={data.SerialNumber}", true);
                                return data;
                            }
                            else
                            {
                                WriteLog($"No data found in database for invoice number: {invoiceNumber}", true);
                                return null;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogException(ex, "Database Query");
                WriteLog($"Database query error: {ex.Message}", true);
                return null;
            }
        }

        // Method 1: Try to add QR code to existing PDF
        private string AddQRCodeToPdf_Method1(string originalPdfPath, string qrCodePath, string invoiceNumber)
        {
            WriteLog($"Adding QR code to PDF using Method 1 for invoice {invoiceNumber}", true);
            WriteLog($"Original PDF: {originalPdfPath}", true);
            WriteLog($"QR code image: {qrCodePath}", true);

            // Define output path - use unique filename to avoid conflicts
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            string outputPdfPath = Path.Combine(OutputFolder, $"Modified_{invoiceNumber}_{timestamp}.pdf");
            string tempPdfPath = Path.Combine(Path.GetTempPath(), $"Temp_{invoiceNumber}_{Guid.NewGuid()}.pdf");
            
            WriteLog($"Output path: {outputPdfPath}", true);
            WriteLog($"Temp path: {tempPdfPath}", true);

            try
            {
                // Ensure the directory exists
                Directory.CreateDirectory(Path.GetDirectoryName(outputPdfPath));
                WriteLog($"Verified output directory exists: {Path.GetDirectoryName(outputPdfPath)}", true);

                // Verify input files exist
                if (!File.Exists(originalPdfPath))
                {
                    string error = $"Original PDF file not found: {originalPdfPath}";
                    WriteLog(error, true);
                    throw new FileNotFoundException(error);
                }
                
                if (!File.Exists(qrCodePath))
                {
                    string error = $"QR code image not found: {qrCodePath}";
                    WriteLog(error, true);
                    throw new FileNotFoundException(error);
                }

                // Copy the PDF to temp location with retry logic
                int retries = 3;
                bool copied = false;
                
                while (!copied && retries > 0)
                {
                    try
                    {
                        WriteLog($"Copying PDF to temp location (attempt {4-retries}/3)", true);
                        File.Copy(originalPdfPath, tempPdfPath, true);
                        copied = true;
                        WriteLog($"Created temporary copy of PDF at {tempPdfPath}", true);
                    }
                    catch (IOException ex)
                    {
                        retries--;
                        WriteLog($"Copy failed: {ex.Message}, retries left: {retries}", true);
                        
                        if (retries == 0) throw;
                        
                        // Generate a new temp path
                        tempPdfPath = Path.Combine(Path.GetTempPath(), $"Temp_{invoiceNumber}_{Guid.NewGuid()}.pdf");
                        WriteLog($"New temp path: {tempPdfPath}", true);
                        Thread.Sleep(100); // Short delay before retry
                    }
                }

                // Ensure output file doesn't exist
                if (File.Exists(outputPdfPath))
                {
                    WriteLog($"Deleting existing output file: {outputPdfPath}", true);
                    File.Delete(outputPdfPath);
                }

                // PDF processing with detailed error handling
                WriteLog("Opening PDF with PdfReader", true);
                using (PdfReader pdfReader = new PdfReader(tempPdfPath))
                {
                    WriteLog("PDF opened successfully with PdfReader", true);
                    WriteLog("Creating PdfWriter", true);
                    
                    using (PdfWriter pdfWriter = new PdfWriter(outputPdfPath))
                    {
                        WriteLog("PdfWriter created successfully", true);
                        WriteLog("Creating PdfDocument", true);
                        
                        using (PdfDocument pdfDoc = new PdfDocument(pdfReader, pdfWriter))
                        {
                            WriteLog($"PdfDocument created successfully, pages: {pdfDoc.GetNumberOfPages()}", true);
                            WriteLog("Creating Document layout object", true);
                            
                            Document document = new Document(pdfDoc);
                            WriteLog("Document created successfully", true);

                            // Get the number of pages
                            int numberOfPages = pdfDoc.GetNumberOfPages();
                            WriteLog($"PDF has {numberOfPages} pages", true);

                            // Target the last page
                            PdfPage lastPage = pdfDoc.GetPage(numberOfPages);
                            WriteLog($"Retrieved last page (page {numberOfPages})", true);

                            // Load QR code image
                            WriteLog($"Loading QR code image: {qrCodePath}", true);
                            ImageData imageData = ImageDataFactory.Create(qrCodePath);
                            WriteLog("ImageData created successfully", true);
                            
                            iText.Layout.Element.Image qrCodeImage = new iText.Layout.Element.Image(imageData);
                            WriteLog("Image object created successfully", true);

                            // Set size and position - larger size and position in bottom left
                            qrCodeImage.ScaleToFit(100, 100);
                            WriteLog("Image scaled to 100x100", true);
                            
                            qrCodeImage.SetFixedPosition(20, 25);
                            WriteLog("Image positioned at (20, 20)", true);

                            // Add the image to the document
                            WriteLog("Adding image to document", true);
                            document.Add(qrCodeImage);
                            WriteLog("Image added successfully", true);

                            // Try to get additional data from database
                            QrCodeData qrData = GetQrCodeDataFromDatabase(invoiceNumber);
                            string additionalInfo = "";
                            if (qrData != null)
                            {
                                WriteLog("Adding database information to PDF", true);
                                additionalInfo = qrData.FormatQrCodeText();
                            }

                            // Add invoice number text next to QR code
                            WriteLog("Creating font for text", true);
                            PdfFont boldFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
                            WriteLog("Font created successfully", true);
                            
                            WriteLog("Creating paragraph for invoice info", true);
                            Paragraph footerParagraph;
                            if (!string.IsNullOrEmpty(additionalInfo))
                            {
                                /*
                                footerParagraph = new Paragraph($"Invoice: {invoiceNumber}\n{additionalInfo}\nVerification QR Code");
                                */
                                footerParagraph = new Paragraph(additionalInfo);
                            }
                            else
                            {
                                footerParagraph = new Paragraph($"Invoice: {invoiceNumber}\nVerification QR Code");
                            }
                            footerParagraph.SetFont(boldFont);
                            footerParagraph.SetFontSize(8);
                            WriteLog("Paragraph created and styled successfully", true);

                            // Position and add the text
                            WriteLog("Adding aligned text to document", true);
                            document.ShowTextAligned(
                                footerParagraph,
                                125, // x position
                                40,  // y position
                                numberOfPages,
                                TextAlignment.LEFT,
                                VerticalAlignment.BOTTOM,
                                0   // rotation
                            );
                            WriteLog("Text added successfully", true);

                            // Close the document explicitly
                            WriteLog("Closing document", true);
                            document.Close();
                            WriteLog("Document closed successfully using Method 1", true);
                        }
                        WriteLog("PdfDocument disposed", true);
                    }
                    WriteLog("PdfWriter disposed", true);
                }
                WriteLog("PdfReader disposed", true);

                // Clean up the temporary file
                try
                {
                    if (File.Exists(tempPdfPath))
                    {
                        WriteLog($"Deleting temporary file: {tempPdfPath}", true);
                        File.Delete(tempPdfPath);
                        WriteLog("Temporary file deleted successfully", true);
                    }
                }
                catch (Exception cleanupEx)
                {
                    WriteLog($"Warning: Failed to clean up temporary file: {cleanupEx.Message}", true);
                }

                WriteLog($"Successfully modified PDF using Method 1: {outputPdfPath}", true);
                return outputPdfPath;
            }
            catch (IOException ex)
            {
                LogMessage($"I/O Error modifying PDF: {ex.Message}");
                LogException(ex, "Method1 - I/O Error");
                throw new Exception($"File access error: {ex.Message}", ex);
            }
            catch (iText.IO.Exceptions.IOException ex)
            {
                LogMessage($"iText I/O Error: {ex.Message}");
                LogException(ex, "Method1 - iText I/O Error");
                throw new Exception($"PDF I/O error: {ex.Message}", ex);
            }
            catch (Exception ex)
            {
                LogMessage($"Error modifying PDF: {ex.GetType().Name}: {ex.Message}");
                LogException(ex, "Method1 - General Error");
                throw new Exception($"Error adding QR code to PDF: {ex.Message}", ex);
            }
        }

        // Method 2: Create a simple cover page with QR code and link to original
        private string CreateCoverPageWithQRCode(string originalPdfPath, string qrCodePath, string invoiceNumber)
        {
            WriteLog($"Creating cover page with QR code for invoice {invoiceNumber}", true);
            WriteLog($"Original PDF: {originalPdfPath}", true);
            WriteLog($"QR code image: {qrCodePath}", true);

            // Try to get additional data from database
            QrCodeData qrData = GetQrCodeDataFromDatabase(invoiceNumber);
            string additionalInfo = "";
            if (qrData != null)
            {
                WriteLog("Adding database information to cover page", true);
                additionalInfo = qrData.FormatQrCodeText();
            }

            // Define output path
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            string outputPdfPath = Path.Combine(OutputFolder, $"Cover_{invoiceNumber}_{timestamp}.pdf");
            WriteLog($"Output path: {outputPdfPath}", true);

            try
            {
                // Create a new PDF with just a cover page
                WriteLog("Creating PdfWriter for cover page", true);
                using (PdfWriter writer = new PdfWriter(outputPdfPath))
                {
                    WriteLog("PdfWriter created successfully", true);
                    WriteLog("Creating PdfDocument", true);
                    
                    using (PdfDocument pdfDoc = new PdfDocument(writer))
                    {
                        WriteLog("PdfDocument created successfully", true);
                        WriteLog("Creating Document layout object", true);
                        
                        Document document = new Document(pdfDoc);
                        WriteLog("Document created successfully", true);

                        // Add a page
                        WriteLog("Adding new page", true);
                        pdfDoc.AddNewPage();
                        WriteLog("Page added successfully", true);

                        // Add a title
                        WriteLog("Creating title paragraph", true);
                        Paragraph title = new Paragraph($"Fiscal Verification for Invoice {invoiceNumber}")
                            .SetFont(PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD))
                            .SetFontSize(16)
                            .SetTextAlignment(TextAlignment.CENTER);
                        WriteLog("Adding title to document", true);
                        document.Add(title);
                        WriteLog("Title added successfully", true);

                        document.Add(new Paragraph("\n\n"));

                        // Add QR code
                        WriteLog("Loading QR code image", true);
                        ImageData imageData = ImageDataFactory.Create(qrCodePath);
                        iText.Layout.Element.Image qrCodeImage = new iText.Layout.Element.Image(imageData);
                        qrCodeImage.ScaleToFit(200, 200);
                        WriteLog("Setting image alignment", true);
                        
                        // Use explicit iText namespace to avoid ambiguity
                        qrCodeImage.SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.CENTER);
                        WriteLog("Adding image to document", true);
                        document.Add(qrCodeImage);
                        WriteLog("QR code image added successfully", true);

                        document.Add(new Paragraph("\n"));

                        // Add verification text
                        WriteLog("Adding verification text", true);
                        Paragraph verification;
                        if (!string.IsNullOrEmpty(additionalInfo))
                        {
                            verification = new Paragraph("This QR code provides fiscal verification for the attached invoice.\n" + additionalInfo)
                                .SetFont(PdfFontFactory.CreateFont(StandardFonts.HELVETICA))
                                .SetFontSize(12)
                                .SetTextAlignment(TextAlignment.CENTER);
                        }
                        else
                        {
                            verification = new Paragraph("This QR code provides fiscal verification for the attached invoice.")
                                .SetFont(PdfFontFactory.CreateFont(StandardFonts.HELVETICA))
                                .SetFontSize(12)
                                .SetTextAlignment(TextAlignment.CENTER);
                        }
                        document.Add(verification);

                        document.Add(new Paragraph("\n"));

                        // Add original filename
                        WriteLog("Adding original filename text", true);
                        Paragraph originalFile = new Paragraph($"Original Invoice: {Path.GetFileName(originalPdfPath)}")
                            .SetFont(PdfFontFactory.CreateFont(StandardFonts.HELVETICA))
                            .SetFontSize(10)
                            .SetTextAlignment(TextAlignment.CENTER);
                        document.Add(originalFile);

                        // Add note about fallback method
                        document.Add(new Paragraph("\n\n"));
                        WriteLog("Adding note about fallback method", true);
                        Paragraph note = new Paragraph("Note: This cover page was created because the original PDF could not be modified directly.")
                            .SetFont(PdfFontFactory.CreateFont(StandardFonts.HELVETICA))
                            .SetFontSize(8)
                            .SetTextAlignment(TextAlignment.CENTER);
                        document.Add(note);

                        // Close the document
                        WriteLog("Closing document", true);
                        document.Close();
                        WriteLog("Cover page created successfully", true);
                    }
                }

                WriteLog($"Successfully created cover page: {outputPdfPath}", true);
                return outputPdfPath;
            }
            catch (Exception ex)
            {
                LogMessage($"Error creating cover page: {ex.GetType().Name}: {ex.Message}");
                LogException(ex, "CreateCoverPageWithQRCode");
                throw new Exception($"Fallback method failed: {ex.Message}", ex);
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
    
    public class QrCodeData
    {
        public string Date { get; set; }
        public string TsNum { get; set; }
        public string ControlCode { get; set; }
        public string SerialNumber { get; set; }
        
        public string FormatQrCodeText()
        {
            // Format as requested: ControlCode->CUIN, TsNum->TSIN, Date->Date, SerialNumber->CUSN
            return $"TSIN:{TsNum}\nDate:{Date}\nCUSN:{SerialNumber}\nCUIN:{ControlCode}";
        }
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