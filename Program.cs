using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Threading;
using System.Threading.Tasks;
using System.Drawing;
using System.Data.SQLite;
using Newtonsoft.Json;
using OfficeOpenXml;

using QRCoder;
using System.Globalization;
//using System.Management;

using Path = System.IO.Path;
using QRCode = QRCoder.QRCode;
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
        // Hard-coded settings from appsettings.json
        private readonly string TicketFolder = @"C:\FBtemp\Ticket";
        private readonly string SentFolder = @"C:\FBtemp\Ticket\Sent";
        private readonly string FailFolder = @"C:\FBtemp\Ticket\Fail";
        private readonly string GarbageFolder = @"C:\FBtemp\Ticket\Garbage";
        private readonly string DBFolder = @"C:\FBtemp\DB";
        private readonly string QRCodeFolder = @"C:\FBtemp\QRCodes";
        private readonly string PostingFilenamePrefix = "PC_";
        private readonly string SentFolderFilenamePrefix = "R_";
        private readonly int FileWatcherWaitMilliseconds = 40000; // 40 seconds
        private readonly bool IgnoreSAPDocTotalUseTotalLineTotal = true;
        private readonly bool BlockFiscalizingTotalNotMatching = false;
        private readonly bool ThrowWarningIfTotalsNotMatching = true;

        // HS Code mappings from Excel file
        private Dictionary<string, string> HSCodeMappings = new Dictionary<string, string>();

        // Additional mappings to handle fallbacks
        private Dictionary<string, MaterialInfo> MaterialInfoMappings = new Dictionary<string, MaterialInfo>();

        // Dictionary of words to their associated material numbers for similarity matching
        private Dictionary<string, HashSet<string>> wordToMaterialMap = new Dictionary<string, HashSet<string>>();

        // Dictionary to track file watchers for fiscalization responses
        private Dictionary<string, FileSystemWatcher> FileWatchers = new Dictionary<string, FileSystemWatcher>();
        private Dictionary<string, FileSystemWatcher> FailedFileWatchers = new Dictionary<string, FileSystemWatcher>();
        private Dictionary<string, CancellationTokenSource> WatcherCancellationTokens = new Dictionary<string, CancellationTokenSource>();
        private Dictionary<string, TaskCompletionSource<FiscalResponseData>> FiscalizationTasks = new Dictionary<string, TaskCompletionSource<FiscalResponseData>>();

        // Category-based HS codes (fallback)
        private readonly Dictionary<string, string> CategoryHSCodes = new Dictionary<string, string>
        {
            { "HERBICIDES", "38089390" },
            { "INSECTICIDES", "38089199" },
            { "FUNGICIDES", "38089290" }
        };

        // List to track processed invoices
        private List<ProcessedInvoice> ProcessedInvoices = new List<ProcessedInvoice>();

        // Track if HS code file is loaded
        private bool HSCodeFileLoaded = false;

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
                Directory.CreateDirectory(TicketFolder);
                Directory.CreateDirectory(SentFolder);
                Directory.CreateDirectory(FailFolder);
                Directory.CreateDirectory(GarbageFolder);
                Directory.CreateDirectory(QRCodeFolder);
                Directory.CreateDirectory(DBFolder);

                LogMessage($"Verified all required folders exist");
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
            this.btnLoadHSCodes = new Button();
            this.btnLoadInvoices = new Button();
            this.btnProcessInvoices = new Button();
            this.listBoxInvoices = new ListBox();
            this.lblStatus = new Label();
            this.progressBar = new ProgressBar();
            this.lblHSCodesStatus = new Label();

            // 
            // btnLoadHSCodes
            // 
            this.btnLoadHSCodes.Location = new System.Drawing.Point(12, 12);
            this.btnLoadHSCodes.Name = "btnLoadHSCodes";
            this.btnLoadHSCodes.Size = new System.Drawing.Size(150, 30);
            this.btnLoadHSCodes.TabIndex = 0;
            this.btnLoadHSCodes.Text = "1. Load HS Codes";
            this.btnLoadHSCodes.UseVisualStyleBackColor = true;
            this.btnLoadHSCodes.Click += new System.EventHandler(this.btnLoadHSCodes_Click);
            // 
            // lblHSCodesStatus
            // 
            this.lblHSCodesStatus.AutoSize = true;
            this.lblHSCodesStatus.Location = new System.Drawing.Point(168, 21);
            this.lblHSCodesStatus.Name = "lblHSCodesStatus";
            this.lblHSCodesStatus.Size = new System.Drawing.Size(120, 13);
            this.lblHSCodesStatus.TabIndex = 1;
            this.lblHSCodesStatus.Text = "HS Codes not loaded";
            this.lblHSCodesStatus.ForeColor = System.Drawing.Color.Red;
            // 
            // btnLoadInvoices
            // 
            this.btnLoadInvoices.Location = new System.Drawing.Point(12, 48);
            this.btnLoadInvoices.Name = "btnLoadInvoices";
            this.btnLoadInvoices.Size = new System.Drawing.Size(150, 30);
            this.btnLoadInvoices.TabIndex = 2;
            this.btnLoadInvoices.Text = "2. Load Invoices";
            this.btnLoadInvoices.UseVisualStyleBackColor = true;
            this.btnLoadInvoices.Click += new System.EventHandler(this.btnLoadInvoices_Click);
            // 
            // listBoxInvoices
            // 
            this.listBoxInvoices.FormattingEnabled = true;
            this.listBoxInvoices.HorizontalScrollbar = true;
            this.listBoxInvoices.Location = new System.Drawing.Point(12, 84);
            this.listBoxInvoices.Name = "listBoxInvoices";
            this.listBoxInvoices.Size = new System.Drawing.Size(760, 225);
            this.listBoxInvoices.TabIndex = 3;
            // 
            // btnProcessInvoices
            // 
            this.btnProcessInvoices.Location = new System.Drawing.Point(12, 315);
            this.btnProcessInvoices.Name = "btnProcessInvoices";
            this.btnProcessInvoices.Size = new System.Drawing.Size(150, 30);
            this.btnProcessInvoices.TabIndex = 4;
            this.btnProcessInvoices.Text = "3. Process Invoices";
            this.btnProcessInvoices.UseVisualStyleBackColor = true;
            this.btnProcessInvoices.Click += new System.EventHandler(this.btnProcessInvoices_Click);
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(168, 324);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(0, 13);
            this.lblStatus.TabIndex = 5;
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(12, 351);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(760, 23);
            this.progressBar.TabIndex = 6;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(784, 386);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.btnProcessInvoices);
            this.Controls.Add(this.listBoxInvoices);
            this.Controls.Add(this.btnLoadInvoices);
            this.Controls.Add(this.lblHSCodesStatus);
            this.Controls.Add(this.btnLoadHSCodes);
            this.Name = "MainForm";
            this.Text = "KRA TIMS Fiscalization Tool";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private decimal RoundOff(decimal value)
        {
            // Match the rounding behavior to SAP B1 integration
            return Math.Round(value, 2, MidpointRounding.AwayFromZero);
        }

        private string ExtractInvoiceNumberFromPdf(string pdfPath)
        {
            try
            {
                string pdfText = GetTextFromPdf(pdfPath);
                string invoiceNumber = ExtractInvoiceNumber(pdfText, Path.GetFileName(pdfPath));
                return invoiceNumber;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private string ExtractQuantity(string line, string[] lines, int lineIndex)
        {
            // Look for pattern: number (with possible commas) followed by "Piece"
            Match match = Regex.Match(line, @"([\d,]+\.\d{2})\s+Piece");
            if (match.Success)
            {
                return match.Groups[1].Value;
            }

            // Look in the next line
            if (lineIndex + 1 < lines.Length)
            {
                Match nextLineMatch = Regex.Match(lines[lineIndex + 1], @"([\d,]+\.\d{2})\s+Piece");
                if (nextLineMatch.Success)
                {
                    return nextLineMatch.Groups[1].Value;
                }
            }

            // Try alternative formats - look for any number with decimal point
            Match altMatch = Regex.Match(line, @"([\d,]+\.\d{2})");
            if (altMatch.Success)
            {
                return altMatch.Groups[1].Value;
            }

            return "1.00"; // Default
        }

        private string ExtractUnitPrice(string line, string[] lines, int lineIndex)
        {
            // Try to find unit price after "Piece"
            Match match = Regex.Match(line, @"Piece\s+([\d,]+\.\d{2})");
            if (match.Success)
            {
                return match.Groups[1].Value;
            }

            // Try alternative pattern
            Match altMatch = Regex.Match(line, @"([\d,]+\.\d{2})\s+Piece\s+([\d,]+\.\d{2})");
            if (altMatch.Success)
            {
                return altMatch.Groups[2].Value;
            }

            // Look in the next line
            if (lineIndex + 1 < lines.Length)
            {
                Match nextLineMatch = Regex.Match(lines[lineIndex + 1], @"Piece\s+([\d,]+\.\d{2})");
                if (nextLineMatch.Success)
                {
                    return nextLineMatch.Groups[1].Value;
                }
            }

            // Try to find any price-like number after the word "Price"
            Match priceMatch = Regex.Match(line, @"Price\s+([\d,]+\.\d{2})");
            if (priceMatch.Success)
            {
                return priceMatch.Groups[1].Value;
            }

            return "0.00"; // Default
        }

        private string ExtractDiscountPercent(string line, string[] lines, int lineIndex)
        {
            // Look for discount percentage with % symbol
            Match match = Regex.Match(line, @"([\d,]+\.\d{2})%");
            if (match.Success)
            {
                return match.Groups[1].Value;
            }

            // Look for discount pattern in structured format
            Match altMatch = Regex.Match(line, @"([\d,]+\.\d{2})\s+Piece\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})%");
            if (altMatch.Success)
            {
                return altMatch.Groups[3].Value;
            }

            // Look for discount in column format with just a number and % sign
            Match columnMatch = Regex.Match(line, @"(\d+(?:\.\d+)?)%");
            if (columnMatch.Success)
            {
                return columnMatch.Groups[1].Value;
            }

            // Look in the next line
            if (lineIndex + 1 < lines.Length)
            {
                Match nextLineMatch = Regex.Match(lines[lineIndex + 1], @"([\d,]+\.\d{2})%");
                if (nextLineMatch.Success)
                {
                    return nextLineMatch.Groups[1].Value;
                }
            }

            return "0.00"; // Default
        }

        private string ExtractUnitPriceAfterDiscount(string line, string[] lines, int lineIndex)
        {
            // Try to find unit price after discount percentage
            Match match = Regex.Match(line, @"([\d,]+\.\d{2})%\s+([\d,]+\.\d{2})");
            if (match.Success)
            {
                return match.Groups[2].Value;
            }

            // Try alternative pattern
            Match altMatch = Regex.Match(line, @"([\d,]+\.\d{2})\s+Piece\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})%\s+([\d,]+\.\d{2})");
            if (altMatch.Success)
            {
                return altMatch.Groups[4].Value;
            }

            // Look for column structure with heading "After discount unit price"
            Match columnMatch = Regex.Match(line, @"discount\s+unit\s+price\s+([\d,]+\.\d{2})");
            if (columnMatch.Success)
            {
                return columnMatch.Groups[1].Value;
            }

            // Look in the next line
            if (lineIndex + 1 < lines.Length)
            {
                Match nextLineMatch = Regex.Match(lines[lineIndex + 1], @"([\d,]+\.\d{2})%\s+([\d,]+\.\d{2})");
                if (nextLineMatch.Success)
                {
                    return nextLineMatch.Groups[2].Value;
                }
            }

            // If we have unit price and discount percentage, calculate
            string unitPrice = ExtractUnitPrice(line, lines, lineIndex);
            string discountPercent = ExtractDiscountPercent(line, lines, lineIndex);

            if (unitPrice != "0.00" && discountPercent != "0.00")
            {
                decimal price = decimal.Parse(unitPrice.Replace(",", ""), CultureInfo.InvariantCulture);
                decimal discount = decimal.Parse(discountPercent.Replace(",", ""), CultureInfo.InvariantCulture);
                decimal afterDiscount = price * (1 - (discount / 100));
                return RoundOff(afterDiscount).ToString("0.00", CultureInfo.InvariantCulture);
            }

            return unitPrice; // Default to unit price
        }

        private string ExtractAmount(string line, string[] lines, int lineIndex)
        {
            // Try to find amount at the end of the line
            Match match = Regex.Match(line, @"([\d,]+\.\d{2})$");
            if (match.Success)
            {
                return match.Groups[1].Value;
            }

            // Look for amount in specific format with commas
            Match commaMatch = Regex.Match(line, @"(\d{1,3}(?:,\d{3})*\.\d{2})$");
            if (commaMatch.Success)
            {
                return commaMatch.Groups[1].Value;
            }

            // Look in next line
            if (lineIndex + 1 < lines.Length)
            {
                Match nextLineMatch = Regex.Match(lines[lineIndex + 1], @"([\d,]+\.\d{2})$");
                if (nextLineMatch.Success)
                {
                    return nextLineMatch.Groups[1].Value;
                }
            }

            // If we have quantity and unit price after discount, calculate
            string quantity = ExtractQuantity(line, lines, lineIndex);
            string unitPriceAfterDiscount = ExtractUnitPriceAfterDiscount(line, lines, lineIndex);

            if (quantity != "1.00" || unitPriceAfterDiscount != "0.00")
            {
                decimal qty = decimal.Parse(quantity.Replace(",", ""), CultureInfo.InvariantCulture);
                decimal price = decimal.Parse(unitPriceAfterDiscount.Replace(",", ""), CultureInfo.InvariantCulture);
                decimal amount = qty * price;
                return RoundOff(amount).ToString("0.00", CultureInfo.InvariantCulture);
            }

            return "0.00"; // Default
        }

        private string GenerateAndSavePCFile(Invoice invoice)
        {
            LogMessage($"Generating PC file for invoice {invoice.InvoiceNumber}");

            var sb = new StringBuilder();

            // 1. Invoice header (Invoice Number)
            sb.AppendLine($"\"{invoice.InvoiceNumber}\"@1j");

            // 2. Federal Tax ID (if available)
            if (!string.IsNullOrEmpty(invoice.FederalTaxID) && invoice.FederalTaxID.Length >= 11)
            {
                sb.AppendLine($"\"{invoice.FederalTaxID}\"@39F");
            }

            // 3. Process line items
            decimal cumulativeTotalAmount = 0;

            foreach (var lineItem in invoice.LineItems)
            {
                // Format the description (max 20 chars, alphanumeric only)
                string itemDesc = Regex.Replace(
                    lineItem.Description.Substring(0, Math.Min(20, lineItem.Description.Length)),
                    "[^a-zA-Z0-9()/]", "");

                // Format quantity and unit price (2 decimal places, no decimal point)
                decimal quantity;
                if (!decimal.TryParse(lineItem.Quantity.Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out quantity))
                {
                    quantity = 1.0m;
                    LogMessage($"WARNING: Could not parse quantity '{lineItem.Quantity}' for item {lineItem.ItemCode}, using 1.0");
                }

                decimal unitPriceAfterDiscount;
                if (!decimal.TryParse(lineItem.UnitPriceAfterDiscount.Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out unitPriceAfterDiscount))
                {
                    // Try to calculate it from amount and quantity
                    decimal amount;
                    if (decimal.TryParse(lineItem.Amount.Replace(",", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out amount))
                    {
                        unitPriceAfterDiscount = amount / quantity;
                        LogMessage($"WARNING: Calculated unit price from amount for item {lineItem.ItemCode}: {unitPriceAfterDiscount}");
                    }
                    else
                    {
                        unitPriceAfterDiscount = 0;
                        LogMessage($"WARNING: Could not determine unit price for item {lineItem.ItemCode}, using 0");
                    }
                }

                // Calculate line total after VAT and discount (following SAP B1 integration methodology)
                decimal calculatedLineTotal = quantity * unitPriceAfterDiscount;
                calculatedLineTotal = RoundOff(calculatedLineTotal);

                // Format quantities and unit prices with no decimal point
                string quantityFormatted = quantity.ToString("0.00", CultureInfo.InvariantCulture).Replace(".", "");
                string unitPriceFormatted = unitPriceAfterDiscount.ToString("0.00", CultureInfo.InvariantCulture).Replace(".", "");

                // Format HS code (remove dots)
                string hsCode = lineItem.HSCode.Replace(".", "");

                // Add line item to file
                sb.AppendLine($"\"{itemDesc}\"@{quantityFormatted}*{unitPriceFormatted}H\"{hsCode}\"@P");

                // Add to cumulative total
                cumulativeTotalAmount += calculatedLineTotal;
                LogMessage($"Line item: {lineItem.ItemCode}, Qty: {quantity}, Unit Price: {unitPriceAfterDiscount}, Line Total: {calculatedLineTotal}");
            }

            // Final rounding on the cumulative total to match SAP B1 integration
            cumulativeTotalAmount = RoundOff(cumulativeTotalAmount);
            LogMessage($"Cumulative total after final rounding: {cumulativeTotalAmount}");

            // 4. Add total amount
            string totalFormatted = cumulativeTotalAmount.ToString("0.00", CultureInfo.InvariantCulture).Replace(".", "");
            sb.AppendLine($"{totalFormatted}H0T");
            sb.AppendLine($"{totalFormatted}H1T");

            // 5. Add final lines
            sb.AppendLine("S");
            sb.AppendLine("1J");

            // Create output file path
            string fileName = $"{PostingFilenamePrefix}{invoice.InvoiceNumber}.txt";
            string outputPath = Path.Combine(TicketFolder, fileName);

            try
            {
                File.WriteAllText(outputPath, sb.ToString(), new UTF8Encoding(false));
                LogMessage($"Successfully saved PC file to {outputPath}");
                return outputPath;
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR: Failed to save PC file: {ex.Message}");
                throw new Exception($"Failed to save PC file: {ex.Message}", ex);
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
            btnLoadHSCodes.Enabled = false;
            btnLoadInvoices.Enabled = false;
            btnProcessInvoices.Enabled = false;

            try
            {
                progressBar.Minimum = 0;
                progressBar.Maximum = ProcessedInvoices.Count;
                progressBar.Value = 0;

                int successCount = 0;
                int failCount = 0;
                int skipCount = 0;

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
                        // Extract just the invoice number first to check if already fiscalized
                        string invoiceNumber = ExtractInvoiceNumberFromPdf(invoice.FilePath);

                        if (string.IsNullOrEmpty(invoiceNumber))
                        {
                            LogMessage($"Could not extract invoice number from {Path.GetFileName(invoice.FilePath)}");
                            invoice.Status = "Failed";
                            invoice.Message = "Could not extract invoice number";
                            failCount++;
                            continue;
                        }

                        // Check if already fiscalized
                        bool alreadyFiscalized = await IsInvoiceAlreadyFiscalized(invoiceNumber);
                        if (alreadyFiscalized)
                        {
                            LogMessage($"Invoice {invoiceNumber} is already fiscalized. Attempting to generate QR code PDF...");

                            // Try to get fiscal data from database first
                            FiscalResponseData existingFiscalData = await TryReadFromSQLiteDatabase(invoiceNumber);

                            // If database retrieval fails, try to read from the response file in Sent folder
                            if (existingFiscalData == null)
                            {
                                LogMessage($"Database retrieval failed for invoice {invoiceNumber}. Trying response file...");
                                string sentFileName = $"{SentFolderFilenamePrefix}{PostingFilenamePrefix}{invoiceNumber}.txt";
                                string sentFilePath = Path.Combine(SentFolder, sentFileName);

                                if (File.Exists(sentFilePath))
                                {
                                    existingFiscalData = await ReadFiscalResponseFromFile(sentFilePath, invoiceNumber);
                                    LogMessage($"Read fiscal data from response file for invoice {invoiceNumber}");
                                }
                            }

                            if (existingFiscalData != null)
                            {
                                // Generate QR code using existing fiscal data
                                string qrCodePath = GenerateQRCode(existingFiscalData.FiscalSeal, invoiceNumber);
                                LogMessage($"Generated QR code for invoice {invoiceNumber}");

                                // Add QR code and fiscal footer to PDF
                                string modifiedPdfPath = AddQRCodeAndFooterToPdf(invoice.FilePath, qrCodePath, existingFiscalData, invoiceNumber);
                                LogMessage($"Created modified PDF at {modifiedPdfPath}");

                                invoice.Status = "Success";
                                invoice.Message = $"PDF recreated with fiscal data: {Path.GetFileName(modifiedPdfPath)}";
                                successCount++;
                                LogMessage($"Successfully generated PDF with existing fiscal data for invoice {invoiceNumber}");
                            }
                            else
                            {
                                LogMessage($"Could not retrieve fiscal data for already fiscalized invoice {invoiceNumber}");
                                invoice.Status = "Failed";
                                invoice.Message = "Could not retrieve fiscal data for already fiscalized invoice";
                                failCount++;
                            }
                            continue;
                        }

                        // Check if PC file already exists
                        string pcFileName = $"{PostingFilenamePrefix}{invoiceNumber}.txt";
                        string pcFilePath = Path.Combine(TicketFolder, pcFileName);

                        if (File.Exists(pcFilePath))
                        {
                            var result = MessageBox.Show(
                                $"PC file for invoice {invoiceNumber} already exists. Overwrite?",
                                "File Exists",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Question);

                            if (result == DialogResult.No)
                            {
                                LogMessage($"Skipped invoice {invoiceNumber} - PC file exists and user chose not to overwrite");
                                invoice.Status = "Skipped";
                                invoice.Message = "PC file already exists, user chose not to overwrite";
                                skipCount++;
                                continue;
                            }
                            else
                            {
                                LogMessage($"Will overwrite existing PC file for invoice {invoiceNumber}");
                            }
                        }

                        // Extract full invoice data
                        var extractedInvoice = ExtractInvoiceData(invoice.FilePath);
                        LogMessage($"Extracted invoice {extractedInvoice.InvoiceNumber} with {extractedInvoice.LineItems.Count} line items");

                        // Generate PC file
                        string pcFilePathL = GenerateAndSavePCFile(extractedInvoice);
                        LogMessage($"Generated PC file: {pcFilePath}");

                        // Wait for fiscalization response
                        var fiscalData = await WaitForFiscalizationResponse(extractedInvoice.InvoiceNumber);

                        // If fiscalization succeeded, generate QR code and modify PDF
                        if (fiscalData != null)
                        {
                            LogMessage($"Fiscalization successful for invoice {extractedInvoice.InvoiceNumber}");
                            LogMessage($"Fiscal seal: {fiscalData.FiscalSeal}");

                            // Generate QR code
                            string qrCodePath = GenerateQRCode(fiscalData.FiscalSeal, extractedInvoice.InvoiceNumber);

                            // Add QR code and fiscal footer to PDF
                            string modifiedPdfPath = AddQRCodeAndFooterToPdf(invoice.FilePath, qrCodePath, fiscalData, extractedInvoice.InvoiceNumber);

                            invoice.Status = "Success";
                            invoice.Message = $"Fiscalized and modified PDF: {Path.GetFileName(modifiedPdfPath)}";
                            successCount++;
                        }
                        else
                        {
                            LogMessage($"Fiscalization failed or timed out for invoice {extractedInvoice.InvoiceNumber}");
                            invoice.Status = "Failed";
                            invoice.Message = "Fiscalization failed or timed out";
                            failCount++;
                        }
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

                lblStatus.Text = $"Completed: {successCount} succeeded, {failCount} failed, {skipCount} skipped";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Processing error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "Processing error";
            }
            finally
            {
                // Re-enable buttons
                btnLoadHSCodes.Enabled = true;
                btnLoadInvoices.Enabled = true;
                btnProcessInvoices.Enabled = true;
            }
        }

        private async Task<FiscalResponseData> ReadFiscalResponseFromFile(string filePath, string invoiceNumber)
        {
            LogMessage($"Reading fiscal response from file: {Path.GetFileName(filePath)}");

            try
            {
                // Wait a moment to ensure the file is completely written
                await Task.Delay(1000);

                // Read all lines from the file
                string[] lines = File.ReadAllLines(filePath);

                string fiscalSeal = null;
                string tsNum = invoiceNumber;
                string controlCode = null;
                string serialNumber = null;
                DateTime transactionDate = DateTime.Now;

                // Process the file lines
                foreach (string line in lines.Reverse()) // Start from the bottom
                {
                    string finalLine = line.Trim()
                        .Replace("\u0011", "")
                        .TrimEnd('|')
                        .Trim('-')
                        .Trim();

                    if (string.IsNullOrEmpty(finalLine))
                        continue;

                    // Try to extract fiscal seal (URL)
                    if (finalLine.StartsWith("https"))
                    {
                        fiscalSeal = finalLine.Trim();
                        continue;
                    }

                    // Try to extract CUIN (control code)
                    if (finalLine.Contains("CUIN"))
                    {
                        string[] parts = finalLine.Split(':');
                        if (parts.Length > 1)
                        {
                            controlCode = parts[1].Trim();
                        }
                        continue;
                    }

                    // Try to extract TSIN (transaction number)
                    if (finalLine.Contains("TSIN"))
                    {
                        string[] parts = finalLine.Split(':');
                        if (parts.Length > 1)
                        {
                            tsNum = parts[1].Trim();
                        }
                        continue;
                    }

                    // Try to extract CUSN (serial number)
                    if (finalLine.Contains("CUSN"))
                    {
                        string[] parts = finalLine.Split(':');
                        if (parts.Length > 1)
                        {
                            serialNumber = parts[1].Trim();
                        }
                        continue;
                    }

                    // Try to extract DATE
                    if (finalLine.Contains("DATE"))
                    {
                        string[] parts = finalLine.Split(':');
                        if (parts.Length > 1)
                        {
                            string dateStr = parts[1].Trim();
                            DateTime.TryParse(dateStr, out transactionDate);
                        }
                        continue;
                    }
                }

                // Check if we found the fiscal seal
                if (string.IsNullOrEmpty(fiscalSeal))
                {
                    LogMessage($"Fiscal seal not found in file");
                    return null;
                }

                // Create the response data
                var result = new FiscalResponseData
                {
                    TransactionDate = transactionDate,
                    TsNum = tsNum,
                    ControlCode = controlCode,
                    SerialNumber = serialNumber,
                    FiscalSeal = fiscalSeal
                };

                // Build the fiscal footer
                result.FiscalFooter = BuildFiscalFooter(result);

                LogMessage($"Successfully read fiscal response from file");
                return result;
            }
            catch (Exception ex)
            {
                LogMessage($"Error reading fiscal response from file: {ex.Message}");
                return null;
            }
        }

        private string BuildFiscalFooter(FiscalResponseData data)
        {
            var sb = new StringBuilder();

            if (!string.IsNullOrEmpty(data.TsNum))
                sb.AppendLine($"TSIN :{data.TsNum}");

            sb.AppendLine($"DATE:{data.TransactionDate:yyyy-MM-dd HH:mm:ss}");

            if (!string.IsNullOrEmpty(data.SerialNumber))
                sb.AppendLine($"CUSN :{data.SerialNumber}");

            if (!string.IsNullOrEmpty(data.ControlCode))
                sb.AppendLine($"CUIN: {data.ControlCode}");

            return sb.ToString();
        }





        // Add this helper method to standardize numeric string parsing
        private decimal ParseDecimal(string value)
        {
            if (string.IsNullOrEmpty(value))
                return 0;

            // Remove commas and other potential formatting characters
            string cleanValue = value.Replace(",", "").Trim();

            decimal result;
            if (decimal.TryParse(cleanValue, NumberStyles.Any, CultureInfo.InvariantCulture, out result))
                return result;

            return 0;
        }


        // Helper method to quickly extract just the invoice number
        private async Task<string> ExtractInvoiceNumberQuickly(string pdfPath)
        {
            try
            {
                string pdfText = GetTextFromPdf(pdfPath);
                string invoiceNumber = ExtractInvoiceNumber(pdfText, Path.GetFileName(pdfPath));
                return invoiceNumber;
            }
            catch
            {
                return null;
            }
        }
       

          private void btnLoadHSCodes_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
                openFileDialog.Title = "Select HS Codes Excel File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        listBoxInvoices.Items.Clear();
                        LoadHSCodesFromExcel(openFileDialog.FileName);
                        lblHSCodesStatus.Text = $"HS Codes loaded: {HSCodeMappings.Count} items";
                        lblHSCodesStatus.ForeColor = System.Drawing.Color.Green;
                        HSCodeFileLoaded = true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error loading HS Codes: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        LogMessage($"ERROR: Failed to load HS codes: {ex.Message}");
                    }
                }
            }
        }

        private void LoadHSCodesFromExcel(string filePath)
        {
            HSCodeMappings.Clear();
            MaterialInfoMappings.Clear();

            // Set the license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            LogMessage($"Opening Excel file: {Path.GetFileName(filePath)}");

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                // Check if there are any worksheets
                if (package.Workbook.Worksheets.Count == 0)
                {
                    throw new Exception("The Excel file does not contain any worksheets.");
                }

                var worksheet = package.Workbook.Worksheets[0]; // First worksheet

                // Check if the worksheet has data
                if (worksheet.Dimension == null)
                {
                    throw new Exception("The worksheet is empty.");
                }

                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                LogMessage($"Examining Excel file with {rowCount} rows and {colCount} columns");

                // Find column indices for Material number and HS Code by header name
                int materialNumColIndex = -1;
                int hsCodeColIndex = -1;
                int materialDescColIndex = -1;
                int productCategoryColIndex = -1;

                // Search for column headers in the first row
                for (int col = 1; col <= colCount; col++)
                {
                    string headerText = worksheet.Cells[1, col].Text?.Trim() ?? "";

                    if (string.IsNullOrEmpty(headerText))
                        continue;

                    // Log all column headers to help with debugging
                    if (headerText.ToLower().Contains("material number") || headerText.ToLower().Contains("material no") || headerText.ToLower().Contains("hs code") || headerText.ToLower().Contains("tims item code"))
                    {
                        LogMessage($"Column {col} header: '{headerText}'");
                    }

                    // Look for Material number column
                    if (headerText.Equals("Material number", StringComparison.OrdinalIgnoreCase) ||
                        (headerText.Contains("Material") && headerText.Contains("number")))
                    {
                        materialNumColIndex = col;
                        LogMessage($"Found Material number column at index {col}");
                    }
                    // Look for HS Code column
                    else if (headerText.Equals("HS Code", StringComparison.OrdinalIgnoreCase) ||
                            (headerText.Contains("HS") && headerText.Contains("Code")))
                    {
                        hsCodeColIndex = col;
                        LogMessage($"Found HS Code column at index {col}");
                    }
                    // Look for Material description column
                    else if (headerText.Equals("Material description", StringComparison.OrdinalIgnoreCase) ||
                            (headerText.Contains("Material") && headerText.Contains("description")))
                    {
                        materialDescColIndex = col;
                        LogMessage($"Found Material description column at index {col}");
                    }
                    // Look for Product category description column
                    else if (headerText.Equals("Product category description", StringComparison.OrdinalIgnoreCase) ||
                            (headerText.Contains("Product") && headerText.Contains("category")))
                    {
                        productCategoryColIndex = col;
                        LogMessage($"Found Product category column at index {col}");
                    }
                }

                // If Material number column not found, check for column with 12-digit codes
                if (materialNumColIndex == -1)
                {
                    for (int col = 1; col <= colCount; col++)
                    {
                        string cellValue = worksheet.Cells[2, col].Text?.Trim() ?? "";
                        if (Regex.IsMatch(cellValue, @"^\d{12}$"))
                        {
                            materialNumColIndex = col;
                            LogMessage($"Found Material number column at index {col} (by content pattern)");
                            break;
                        }
                    }

                    // If still not found, use hardcoded column index (4 - from screenshot)
                    if (materialNumColIndex == -1)
                    {
                        materialNumColIndex = 4;
                        LogMessage($"Using hardcoded Material number column index: {materialNumColIndex}");
                    }
                }

                // If HS Code column not found, look for column with 8-digit codes
                if (hsCodeColIndex == -1)
                {
                    for (int col = colCount; col >= 1; col--)
                    {
                        string cellValue = worksheet.Cells[2, col].Text?.Trim() ?? "";
                        if (Regex.IsMatch(cellValue, @"^\d{8}$"))
                        {
                            hsCodeColIndex = col;
                            LogMessage($"Found HS Code column at index {col} (by content pattern)");
                            break;
                        }
                    }

                    // If still not found, use hardcoded value (from screenshot)
                    if (hsCodeColIndex == -1)
                    {
                        // Look at the last few columns
                        for (int col = colCount; col >= colCount - 5 && col >= 1; col--)
                        {
                            string headerText = worksheet.Cells[1, col].Text?.Trim() ?? "";
                            if (!string.IsNullOrEmpty(headerText))
                            {
                                hsCodeColIndex = col;
                                LogMessage($"Using last column with header as HS Code column: {hsCodeColIndex}");
                                break;
                            }
                        }

                        // If still not found, use a default hardcoded column
                        if (hsCodeColIndex == -1)
                        {
                            hsCodeColIndex = colCount - 1; // Assume it's the second to last column
                            LogMessage($"Using hardcoded HS Code column index: {hsCodeColIndex}");
                        }
                    }
                }

                // Use hardcoded values for material description and product category if not found
                if (materialDescColIndex == -1 && materialNumColIndex > 0)
                {
                    materialDescColIndex = materialNumColIndex + 1; // Usually next to material number
                    LogMessage($"Using inferred Material description column index: {materialDescColIndex}");
                }

                if (productCategoryColIndex == -1 && hsCodeColIndex > 0)
                {
                    productCategoryColIndex = hsCodeColIndex - 1; // Usually before HS code
                    LogMessage($"Using inferred Product category column index: {productCategoryColIndex}");
                }

                LogMessage($"Final column selection: Material number at column {materialNumColIndex}, " +
                          $"HS Code at column {hsCodeColIndex}");

                // Process ALL rows in the Excel file
                int processedRows = 0;
                int validRows = 0;
                int errorRows = 0;

                for (int row = 2; row <= rowCount; row++) // Start from row 2 (skip header)
                {
                    try
                    {
                        processedRows++;

                        // Extract cell values using multiple methods to ensure reliability
                        string materialNum = GetCellValue(worksheet, row, materialNumColIndex);
                        string hsCode = GetCellValue(worksheet, row, hsCodeColIndex);
                        string materialDesc = GetCellValue(worksheet, row, materialDescColIndex);
                        string productCategory = GetCellValue(worksheet, row, productCategoryColIndex);

                        // Skip rows with empty material number or HS code
                        if (string.IsNullOrEmpty(materialNum) || string.IsNullOrEmpty(hsCode))
                        {
                            if (processedRows <= 5 || processedRows % 100 == 0)
                                LogMessage($"Row {row}: Empty material number or HS code - skipping");
                            errorRows++;
                            continue;
                        }

                        // Skip header-like rows
                        if (materialNum.Equals("Material number", StringComparison.OrdinalIgnoreCase) ||
                            hsCode.Equals("HS Code", StringComparison.OrdinalIgnoreCase))
                        {
                            LogMessage($"Row {row}: Contains header text - skipping");
                            errorRows++;
                            continue;
                        }

                        // Validate material number format (12 digits)
                        if (!Regex.IsMatch(materialNum, @"^\d{12}$"))
                        {
                            LogMessage($"Row {row}: Invalid material number format: '{materialNum}' - skipping");
                            errorRows++;
                            continue;
                        }

                        // Normalize HS code (remove dots if present)
                        hsCode = hsCode.Replace(".", "");

                        // Determine product category
                        string category = DetermineCategory(productCategory, materialDesc);

                        // Add to mappings
                        HSCodeMappings[materialNum] = hsCode;

                        // Store comprehensive material info for fallback lookups
                        MaterialInfoMappings[materialNum] = new MaterialInfo
                        {
                            MaterialNumber = materialNum,
                            Description = materialDesc,
                            Category = category,
                            HSCode = hsCode
                        };

                        validRows++;

                        // Log periodically to avoid flooding the log
                        if (validRows <= 5 || validRows % 100 == 0 || row == rowCount)
                        {
                            LogMessage($"Mapped row {row}: Material={materialNum}, HSCode={hsCode}, Category={category}");
                        }
                    }
                    catch (Exception ex)
                    {
                        LogMessage($"Error processing row {row}: {ex.Message}");
                        errorRows++;
                    }
                }

                // Build word index for description-based similarity searching
                BuildDescriptionIndex();

                // Report final statistics
                LogMessage($"Successfully loaded {HSCodeMappings.Count} HS code mappings");
                LogMessage($"Processed {processedRows} rows, valid: {validRows}, errors: {errorRows}");

                if (HSCodeMappings.Count == 0)
                {
                    throw new Exception("No valid HS code mappings found in the Excel file.");
                }
            }
        }

        private string GetCellValue(ExcelWorksheet worksheet, int row, int col)
        {
            if (col <= 0 || worksheet.Dimension.Columns < col)
                return string.Empty;

            var cell = worksheet.Cells[row, col];

            // Try multiple methods to get cell value
            if (cell.Value != null)
            {
                return cell.Value.ToString().Trim();
            }

            if (!string.IsNullOrEmpty(cell.Text))
            {
                return cell.Text.Trim();
            }

            return string.Empty;
        }

        private void BuildDescriptionIndex()
        {
            LogMessage("Building description search index for fallback HS code lookup...");

            wordToMaterialMap.Clear();

            foreach (var material in MaterialInfoMappings.Values)
            {
                if (string.IsNullOrEmpty(material.Description))
                    continue;

                // Split description into words and index each word
                string[] words = material.Description.ToLower()
                    .Split(new[] { ' ', ',', '.', '-', '/', '+', '(', ')', '[', ']' }, StringSplitOptions.RemoveEmptyEntries);

                foreach (string word in words)
                {
                    // Skip very short words and numbers
                    if (word.Length <= 2 || word.All(char.IsDigit))
                        continue;

                    if (!wordToMaterialMap.ContainsKey(word))
                        wordToMaterialMap[word] = new HashSet<string>();

                    wordToMaterialMap[word].Add(material.MaterialNumber);
                }
            }

            LogMessage($"Description index built with {wordToMaterialMap.Count} unique terms");
        }

        private string DetermineCategory(string productCategory, string description)
        {
            // First check if the product category directly mentions one of our known categories
            if (!string.IsNullOrEmpty(productCategory))
            {
                string upperCategory = productCategory.ToUpper();
                if (upperCategory.Contains("HERBICIDE"))
                    return "HERBICIDES";
                else if (upperCategory.Contains("INSECTICIDE"))
                    return "INSECTICIDES";
                else if (upperCategory.Contains("FUNGICIDE"))
                    return "FUNGICIDES";
            }

            // If not determined from category, try to infer from description
            if (!string.IsNullOrEmpty(description))
            {
                string upperDesc = description.ToUpper();

                // Check for herbicide indicators
                if (upperDesc.Contains("GLYPHOSATE") ||
                    upperDesc.Contains("ATRAZINE") ||
                    upperDesc.Contains("DMA") ||
                    upperDesc.Contains("HERBICIDE") ||
                    upperDesc.Contains("WEED"))
                {
                    return "HERBICIDES";
                }

                // Check for insecticide indicators
                if (upperDesc.Contains("EMAMECTIN") ||
                    upperDesc.Contains("THIAMETHOXAM") ||
                    upperDesc.Contains("LAMBDA") ||
                    upperDesc.Contains("CYHALOTHRIN") ||
                    upperDesc.Contains("INSECTICIDE") ||
                    upperDesc.Contains("PEST"))
                {
                    return "INSECTICIDES";
                }

                // Check for fungicide indicators
                if (upperDesc.Contains("TEBUCONAZOLE") ||
                    upperDesc.Contains("AZOXYSTROBIN") ||
                    upperDesc.Contains("DIFENOCONAZOLE") ||
                    upperDesc.Contains("CAPTAN") ||
                    upperDesc.Contains("FUNGICIDE") ||
                    upperDesc.Contains("DISEASE"))
                {
                    return "FUNGICIDES";
                }
            }

            // Default to HERBICIDES if we can't determine
            return "HERBICIDES";
        }

        private void btnLoadInvoices_Click(object sender, EventArgs e)
        {
            if (!HSCodeFileLoaded)
            {
                MessageBox.Show("Please load HS Codes first.", "HS Codes Required", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "PDF Files|*.pdf";
                openFileDialog.Multiselect = true;
                openFileDialog.Title = "Select Invoice PDFs";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    ProcessedInvoices.Clear();

                    foreach (string file in openFileDialog.FileNames)
                    {
                        LogMessage($"Added invoice: {Path.GetFileName(file)}");
                        ProcessedInvoices.Add(new ProcessedInvoice { FilePath = file, Status = "Pending" });
                    }
                }
            }
        }

        

        private Invoice ExtractInvoiceData(string pdfPath)
        {
            LogMessage($"Extracting data from {Path.GetFileName(pdfPath)}");

            var invoice = new Invoice();
            string pdfText = "";

            // First attempt with standard PDF text extraction
            try
            {
                pdfText = GetTextFromPdf(pdfPath);
                LogMessage($"Extracted {pdfText.Length} characters from PDF");
            }
            catch (Exception ex)
            {
                LogMessage($"Error extracting text: {ex.Message}");
                throw new Exception($"Failed to extract text from PDF: {ex.Message}");
            }

            // Extract invoice number
            string invoiceNumber = ExtractInvoiceNumber(pdfText, Path.GetFileName(pdfPath));
            if (string.IsNullOrEmpty(invoiceNumber))
            {
                throw new Exception("Could not extract invoice number from PDF");
            }

            invoice.InvoiceNumber = invoiceNumber;
            LogMessage($"Found invoice number: {invoiceNumber}");

            // Extract federal tax ID
            invoice.FederalTaxID = ExtractFederalTaxID(pdfText);
            if (!string.IsNullOrEmpty(invoice.FederalTaxID))
            {
                LogMessage($"Found federal tax ID: {invoice.FederalTaxID}");
            }

            // Extract line items
            var lineItems = ExtractLineItems(pdfText);
            if (lineItems.Count == 0)
            {
                throw new Exception("No line items found in the invoice");
            }

            LogMessage($"Extracted {lineItems.Count} line items");
            invoice.LineItems = lineItems;

            return invoice;
        }

        // First, let's fix the GetTextFromPdf method to use iText7 instead of iTextSharp
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

        // Corrected method to extract federal tax ID from the Bill-to section
        private string ExtractFederalTaxID(string pdfText)
        {
            LogMessage("Extracting federal tax ID (Buyer PIN) from bill-to section");

            // Split into lines for easier processing
            string[] lines = pdfText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            // First look for the Bill TO: section with more flexible matching
            int billToIndex = -1;
            for (int i = 0; i < lines.Length; i++)
            {
                string trimmedLine = lines[i].Trim();
                if (trimmedLine.StartsWith("Bill TO:", StringComparison.OrdinalIgnoreCase) ||
                    Regex.IsMatch(trimmedLine, @"^Bill\s+TO\s*:", RegexOptions.IgnoreCase))
                {
                    billToIndex = i;
                    LogMessage($"Found Bill TO section at line {i}: '{trimmedLine}'");
                    break;
                }
            }

            if (billToIndex >= 0)
            {
                // Process the bill-to section (next 10 lines)
                for (int i = billToIndex; i < Math.Min(billToIndex + 10, lines.Length); i++)
                {
                    // Look for the buyer's tax ID pattern at the end of any line in the bill-to section
                    Match taxIdMatch = Regex.Match(lines[i], @"(P\d{9}[A-Z0-9]|P\d{8}[A-Z0-9])");
                    if (taxIdMatch.Success)
                    {
                        string taxId = taxIdMatch.Groups[1].Value;
                        // Make sure this isn't the seller's tax ID
                        if (!lines[i].Contains("TAX ID:"))
                        {
                            LogMessage($"Found buyer tax ID in bill-to section: {taxId}");
                            return taxId;
                        }
                    }
                }
            }

            // As a fallback, look for the buyer's tax ID pattern in the entire document
            // but avoid the line containing "TAX ID:" which would be the seller's
            for (int i = 0; i < lines.Length; i++)
            {
                if (!lines[i].Contains("TAX ID:"))
                {
                    Match taxIdMatch = Regex.Match(lines[i], @"(P\d{9}[A-Z0-9]|P\d{8}[A-Z0-9])");
                    if (taxIdMatch.Success)
                    {
                        string taxId = taxIdMatch.Groups[1].Value;
                        LogMessage($"Found buyer tax ID as fallback: {taxId}");
                        return taxId;
                    }
                }
            }

            LogMessage("No buyer tax ID found in document");
            return null;
        }

        // Check if an invoice is already fiscalized
        private async Task<bool> IsInvoiceAlreadyFiscalized(string invoiceNumber)
        {
            LogMessage($"Checking if invoice {invoiceNumber} is already fiscalized");

            // Check 1: Check if response file exists in Sent folder
            string sentFileName = $"{SentFolderFilenamePrefix}{PostingFilenamePrefix}{invoiceNumber}.txt";
            string sentFilePath = Path.Combine(SentFolder, sentFileName);

            if (File.Exists(sentFilePath))
            {
                LogMessage($"Found existing response file: {sentFilePath}");
                return true;
            }

            // Check 2: Check SQLite database for existing fiscalization
            try
            {
                var fiscalData = await TryReadFromSQLiteDatabase(invoiceNumber);
                if (fiscalData != null)
                {
                    LogMessage($"Found existing fiscalization data in SQLite database for invoice {invoiceNumber}");
                    return true;
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error checking SQLite database: {ex.Message}");
            }

            LogMessage($"Invoice {invoiceNumber} is not yet fiscalized");
            return false;
        }



        private string AddQRCodeAndFooterToPdf(string originalPdfPath, string qrCodePath, FiscalResponseData fiscalData, string invoiceNumber)
        {
            LogMessage($"Adding QR code and fiscal footer to PDF for invoice {invoiceNumber}");

            // Define output paths - use unique filenames to avoid conflicts
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            string outputPdfPath = Path.Combine(SentFolder, $"Modified_{invoiceNumber}_{timestamp}.pdf");
            string tempPdfPath = Path.Combine(Path.GetTempPath(), $"Temp_{invoiceNumber}_{Guid.NewGuid()}.pdf");
            string altOutputPath = Path.Combine(SentFolder, $"Fiscal_{invoiceNumber}_{timestamp}.pdf");

            try
            {
                // Ensure the directory exists with explicit error handling
                try
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(outputPdfPath));
                    LogMessage($"Verified output directory exists: {Path.GetDirectoryName(outputPdfPath)}");
                }
                catch (Exception dirEx)
                {
                    LogMessage($"Error creating output directory: {dirEx.Message}");
                    // Try to use the application directory as fallback
                    outputPdfPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"Modified_{invoiceNumber}_{timestamp}.pdf");
                    altOutputPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"Fiscal_{invoiceNumber}_{timestamp}.pdf");
                    LogMessage($"Using alternate output location: {Path.GetDirectoryName(outputPdfPath)}");
                }

                // Verify input files exist
                if (!File.Exists(originalPdfPath))
                    throw new FileNotFoundException($"Original PDF file not found: {originalPdfPath}");
                if (!File.Exists(qrCodePath))
                    throw new FileNotFoundException($"QR code image not found: {qrCodePath}");

                // Attempt primary PDF modification
                try
                {
                    // Copy the PDF to temp location with retry logic
                    int copyAttempts = 0;
                    bool copySuccess = false;

                    while (!copySuccess && copyAttempts < 3)
                    {
                        try
                        {
                            copyAttempts++;
                            // Ensure the temp file doesn't already exist
                            if (File.Exists(tempPdfPath))
                            {
                                File.Delete(tempPdfPath);
                            }

                            // Copy with explicit file stream to ensure proper handling
                            using (FileStream source = new FileStream(originalPdfPath, FileMode.Open, FileAccess.Read, FileShare.Read))
                            using (FileStream destination = new FileStream(tempPdfPath, FileMode.CreateNew))
                            {
                                source.CopyTo(destination);
                            }

                            copySuccess = true;
                            LogMessage($"Created temporary copy of PDF at {tempPdfPath}");
                        }
                        catch (IOException ioEx)
                        {
                            LogMessage($"Copy attempt {copyAttempts} failed: {ioEx.Message}");
                            tempPdfPath = Path.Combine(Path.GetTempPath(), $"Temp_{invoiceNumber}_{Guid.NewGuid()}.pdf");
                            Thread.Sleep(500); // Brief pause before retry
                        }
                    }

                    if (!copySuccess)
                    {
                        throw new IOException("Failed to create temporary copy of PDF after multiple attempts");
                    }

                    // Verify we can open the temp PDF
                    using (PdfReader reader = new PdfReader(tempPdfPath))
                    {
                        // Just check if we can open it
                        LogMessage("Successfully opened PDF for reading");
                    }

                    // Ensure output file doesn't exist to avoid access conflicts
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

                        // Set size and position - smaller size and position
                        qrCodeImage.ScaleToFit(50, 50);
                        qrCodeImage.SetFixedPosition(5, 20);

                        // Create all objects separately to avoid null reference issues
                        PdfFont boldFont = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD);
                        Paragraph footerParagraph = new Paragraph(fiscalData.FiscalFooter);
                        footerParagraph.SetFont(boldFont);
                        footerParagraph.SetFontSize(6);

                        // Add the image to the document - make sure document is not null
                        document.Add(qrCodeImage);

                        // Position and add the text - use correct parameter order for 7.2.5
                        document.ShowTextAligned(
                            footerParagraph,
                            60, // x position
                            20,  // y position
                            numberOfPages,
                            TextAlignment.LEFT,
                            VerticalAlignment.BOTTOM,
                            0   // rotation
                        );

                      

                        // Close the document explicitly
                        document.Close();
                        LogMessage("Document closed successfully");
                    }

                    LogMessage($"Successfully modified PDF: {outputPdfPath}");
                    return outputPdfPath;
                }
                catch (Exception ex)
                {
                    LogMessage($"Primary PDF modification failed: {ex.Message}");

                    // Try the alternative approach - create a completely new PDF
                    try
                    {
                        LogMessage("Attempting to create a new PDF with fiscal information");

                        // Ensure the alternative output file doesn't exist
                        if (File.Exists(altOutputPath))
                        {
                            File.Delete(altOutputPath);
                        }

                        // Create a new PDF from scratch with fiscal information
                        using (PdfWriter writer = new PdfWriter(altOutputPath))
                        using (PdfDocument pdfDoc = new PdfDocument(writer))
                        {
                            Document document = new Document(pdfDoc);

                            // Add a page
                            pdfDoc.AddNewPage();

                            // Add invoice number at the top
                            Paragraph invoiceTitle = new Paragraph($"Fiscal Information for Invoice {invoiceNumber}")
                                .SetFont(PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD))
                                .SetFontSize(14)
                                .SetTextAlignment(TextAlignment.CENTER);

                            document.Add(invoiceTitle);
                            document.Add(new Paragraph("\n"));

                            // Add QR code - use fully qualified namespace to avoid ambiguity
                            ImageData imageData = ImageDataFactory.Create(qrCodePath);
                            iText.Layout.Element.Image qrCodeImage = new iText.Layout.Element.Image(imageData);
                            qrCodeImage.ScaleToFit(100, 100); // Larger for standalone document

                            // Use fully qualified namespace for HorizontalAlignment to avoid ambiguity
                            qrCodeImage.SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.CENTER);
                            document.Add(qrCodeImage);

                            // Add fiscal information section
                            document.Add(new Paragraph("\n"));
                            document.Add(new Paragraph("Fiscal Information")
                                .SetFont(PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD))
                                .SetFontSize(12));

                            document.Add(new Paragraph(fiscalData.FiscalFooter)
                                .SetFont(PdfFontFactory.CreateFont(StandardFonts.HELVETICA))
                                .SetFontSize(10));

                            // Add fiscal URL
                            document.Add(new Paragraph("\nFiscal Verification URL:")
                                .SetFont(PdfFontFactory.CreateFont(StandardFonts.HELVETICA_BOLD))
                                .SetFontSize(10));

                            document.Add(new Paragraph(fiscalData.FiscalSeal)
                                .SetFont(PdfFontFactory.CreateFont(StandardFonts.HELVETICA))
                                .SetFontSize(8));

                            // Add disclaimer - use HELVETICA with italic style parameter since HELVETICA_ITALIC doesn't exist
                            // in this version of iText7
                            Paragraph disclaimerParagraph = new Paragraph(
                                "\n\nNote: This is an official fiscal document generated because the original invoice could not be modified.")
                                .SetFont(PdfFontFactory.CreateFont(StandardFonts.HELVETICA))
                                .SetFontSize(8)
                                .SetTextAlignment(TextAlignment.CENTER);

                            // Apply italic styling
                            disclaimerParagraph.SetItalic();

                            document.Add(disclaimerParagraph);

                            // Close the document
                            document.Close();
                        }

                        LogMessage($"Alternative fiscal document created successfully: {altOutputPath}");
                        return altOutputPath;
                    }
                    catch (Exception altEx)
                    {
                        LogMessage($"Alternative PDF creation also failed: {altEx.Message}");

                        // Last resort: Create a minimal TXT file with the fiscal information
                        try
                        {
                            string txtPath = Path.Combine(
                                Path.GetDirectoryName(altOutputPath),
                                $"Fiscal_{invoiceNumber}_{timestamp}.txt");

                            StringBuilder sb = new StringBuilder();
                            sb.AppendLine($"FISCAL INFORMATION FOR INVOICE {invoiceNumber}");
                            sb.AppendLine("=============================================");
                            sb.AppendLine();
                            sb.AppendLine("QR CODE LOCATION:");
                            sb.AppendLine(qrCodePath);
                            sb.AppendLine();
                            sb.AppendLine("FISCAL DETAILS:");
                            sb.AppendLine(fiscalData.FiscalFooter);
                            sb.AppendLine();
                            sb.AppendLine("FISCAL VERIFICATION URL:");
                            sb.AppendLine(fiscalData.FiscalSeal);

                            File.WriteAllText(txtPath, sb.ToString());

                            LogMessage($"Created text file with fiscal information: {txtPath}");
                            return txtPath;
                        }
                        catch (Exception txtEx)
                        {
                            LogMessage($"Even text file creation failed: {txtEx.Message}");
                            throw new Exception("All output attempts failed", txtEx);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Critical error in PDF processing: {ex.Message}");
                LogMessage($"Stack trace: {ex.StackTrace}");
                throw;
            }
            finally
            {
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
            }
        }

        private List<LineItem> ExtractLineItems(string pdfText)
        {
            List<LineItem> lineItems = new List<LineItem>();
            LogMessage("Extracting line items from invoice");

            // Split PDF text into lines
            string[] lines = pdfText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            // Find the item section boundaries
            int startLineIndex = -1;
            int endLineIndex = -1;

            // First try to find section based on SR.NO header
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();

                if ((line.Contains("SR.NO") || line.Contains("SR NO")) &&
                    (line.Contains("Item Code") || line.Contains("Item Description")) &&
                    line.Contains("Amount"))
                {
                    startLineIndex = i + 1;
                    LogMessage($"Found item section header at line {i}");

                    // Now look for end of section
                    for (int j = i + 1; j < lines.Length; j++)
                    {
                        string endLine = lines[j].Trim();

                        if (endLine.Contains("Bank Information") ||
                            endLine.Contains("Sub Total") ||
                            endLine.Contains("Subtotal") ||
                            endLine.StartsWith("Account Name") ||
                            endLine.Trim() == "Total :")
                        {
                            endLineIndex = j;
                            LogMessage($"Found end of item section at line {j}");
                            break;
                        }
                    }

                    break;
                }
            }

            // If section boundaries not found, try alternative method
            if (startLineIndex == -1 || endLineIndex == -1)
            {
                LogMessage("Section boundaries not found, using line pattern matching");

                // Look for lines that begin with a number followed by item code
                // Using \d+ instead of \d{1,2} to match line numbers of any length
                for (int i = 0; i < lines.Length; i++)
                {
                    string line = lines[i].Trim();
                    Match linePattern = Regex.Match(line, @"^\s*(\d+)\s+(\d{12})");

                    if (linePattern.Success)
                    {
                        if (startLineIndex == -1)
                        {
                            startLineIndex = i;
                            LogMessage($"Found first line item at line {i}");
                        }

                        // Keep updating end line as we find more matching lines
                        endLineIndex = i + 1;
                    }
                    else if (line.Contains("Bank Information") ||
                             line.Contains("Sub Total") ||
                             line.Contains("Subtotal") ||
                             line.StartsWith("Account Name") ||
                             line.Trim() == "Total :")
                    {
                        // Stop if we hit a known end marker
                        if (startLineIndex != -1 && endLineIndex == -1)
                        {
                            endLineIndex = i;
                            LogMessage($"Found end marker at line {i}");
                        }
                        break;
                    }
                }
            }

            // If we still don't have boundaries, try a desperate approach
            if (startLineIndex == -1 || endLineIndex == -1 || startLineIndex >= endLineIndex)
            {
                LogMessage("Using direct item code matching");

                // Directly look for 12-digit item codes
                for (int i = 0; i < lines.Length; i++)
                {
                    string line = lines[i].Trim();
                    Match itemCodeMatch = Regex.Match(line, @"(\d{12})");

                    if (itemCodeMatch.Success)
                    {
                        string itemCode = itemCodeMatch.Groups[1].Value;

                        // Try to find amount
                        decimal amount = 0;
                        string amountStr = "0.00";

                        // Look for amount in current line and next line
                        Match amountMatch = Regex.Match(line, @"(\d{1,3}(?:,\d{3})*\.\d{2})$");
                        if (amountMatch.Success)
                        {
                            amountStr = amountMatch.Groups[1].Value.Replace(",", "");
                        }
                        else if (i + 1 < lines.Length)
                        {
                            Match nextLineAmount = Regex.Match(lines[i + 1], @"(\d{1,3}(?:,\d{3})*\.\d{2})$");
                            if (nextLineAmount.Success)
                            {
                                amountStr = nextLineAmount.Groups[1].Value.Replace(",", "");
                            }
                        }

                        // Try to get description
                        string description = "Unknown Item";
                        if (i + 1 < lines.Length)
                        {
                            // Look for description in this line or next line
                            int itemCodePos = line.IndexOf(itemCode);
                            if (itemCodePos > 0 && line.Length > itemCodePos + itemCode.Length + 1)
                            {
                                description = line.Substring(itemCodePos + itemCode.Length).Trim();
                            }
                            else
                            {
                                string nextLine = lines[i + 1].Trim();
                                if (!nextLine.Contains(itemCode) && !Regex.IsMatch(nextLine, @"^\d+\s+\d{12}"))
                                {
                                    description = nextLine;
                                }
                            }
                        }

                        // Create line item with HS code assignment
                        var lineItem = new LineItem
                        {
                            ItemCode = itemCode,
                            Description = description,
                            Quantity = "1.00", // Default if we can't determine
                            UnitPrice = amountStr, // Default
                            DiscountPercent = "0.00", // Default
                            UnitPriceAfterDiscount = amountStr, // Default
                            Amount = amountStr
                        };

                        // Assign HS code with comprehensive fallback strategy
                        lineItem.HSCode = GetHSCodeForItem(itemCode, description);

                        // Try to extract quantity and unit price if possible
                        Match quantityMatch = Regex.Match(line, @"(\d+\.\d{2})\s+Piece");
                        if (quantityMatch.Success)
                        {
                            lineItem.Quantity = quantityMatch.Groups[1].Value;

                            // If we have quantity, try to calculate unit price
                            decimal qty = decimal.Parse(lineItem.Quantity);
                            decimal amt = decimal.Parse(amountStr);
                            if (qty > 0)
                            {
                                decimal unitPrice = amt / qty;
                                lineItem.UnitPrice = unitPrice.ToString("0.00");
                                lineItem.UnitPriceAfterDiscount = unitPrice.ToString("0.00");
                            }
                        }

                        lineItems.Add(lineItem);
                        LogMessage($"Added line item: {itemCode}, Amount: {amountStr}, HS Code: {lineItem.HSCode}");
                    }
                }

                if (lineItems.Count > 0)
                {
                    return lineItems;
                }

                throw new Exception("Could not identify line items in the invoice");
            }

            // Process the identified section
            LogMessage($"Processing lines {startLineIndex} to {endLineIndex}");

            // Track line item numbers and item codes to avoid duplicates
            HashSet<string> processedItems = new HashSet<string>();

            for (int i = startLineIndex; i < endLineIndex; i++)
            {
                string line = lines[i].Trim();

                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                // Try to extract line number and item code - updated to use \d+ for any digit length
                Match lineMatch = Regex.Match(line, @"^\s*(\d+)\s+");

                if (lineMatch.Success)
                {
                    string lineNumber = lineMatch.Groups[1].Value;
                    LogMessage($"Processing line item {lineNumber}");

                    // Look for item code in this line or next
                    string itemCode = null;
                    Match itemCodeMatch = Regex.Match(line, @"\b(\d{12})\b");

                    if (itemCodeMatch.Success)
                    {
                        itemCode = itemCodeMatch.Groups[1].Value;
                    }
                    else if (i + 1 < lines.Length)
                    {
                        // Check next line
                        Match nextLineMatch = Regex.Match(lines[i + 1], @"^\s*(\d{12})\b");
                        if (nextLineMatch.Success)
                        {
                            itemCode = nextLineMatch.Groups[1].Value;
                            i++; // Skip next line
                        }
                    }

                    if (string.IsNullOrEmpty(itemCode))
                    {
                        LogMessage($"  Could not find item code for line {lineNumber}");
                        continue;
                    }

                    // Skip if we already processed this item
                    string itemKey = $"{lineNumber}_{itemCode}";
                    if (processedItems.Contains(itemKey))
                    {
                        LogMessage($"  Skipping duplicate item: {itemKey}");
                        continue;
                    }

                    processedItems.Add(itemKey);

                    // Extract line item details
                    string description = ExtractDescription(line, lines, i, itemCode);
                    string quantity = ExtractQuantity(line, lines, i);
                    string unitPrice = ExtractUnitPrice(line, lines, i);
                    string discountPercent = ExtractDiscountPercent(line, lines, i);
                    string unitPriceAfterDiscount = ExtractUnitPriceAfterDiscount(line, lines, i);
                    string amount = ExtractAmount(line, lines, i);

                    // Create line item
                    var lineItem = new LineItem
                    {
                        ItemCode = itemCode,
                        Description = description,
                        Quantity = quantity,
                        UnitPrice = unitPrice,
                        DiscountPercent = discountPercent,
                        UnitPriceAfterDiscount = unitPriceAfterDiscount,
                        Amount = amount,
                    };

                    // Assign HS code with comprehensive fallback strategy
                    lineItem.HSCode = GetHSCodeForItem(itemCode, description);

                    lineItems.Add(lineItem);
                    LogMessage($"  Added line item: {itemCode}, HS Code: {lineItem.HSCode}");
                }
            }

            if (lineItems.Count == 0)
            {
                throw new Exception("No line items could be extracted from the invoice");
            }

            return lineItems;
        }

        // Comprehensive HS code assignment with multiple fallback strategies
        private string GetHSCodeForItem(string itemCode, string description)
        {
            // Method 1: Direct lookup by material code (most accurate)
            if (HSCodeMappings.ContainsKey(itemCode))
            {
                LogMessage($"  Found HS code for item {itemCode} using direct mapping");
                return HSCodeMappings[itemCode];
            }

            LogMessage($"  No direct HS code mapping for {itemCode}, trying description matching...");

            // Method 2: Find most similar material by description
            if (!string.IsNullOrEmpty(description))
            {
                string bestMatchMaterial = FindMostSimilarMaterial(description);
                if (!string.IsNullOrEmpty(bestMatchMaterial) && MaterialInfoMappings.ContainsKey(bestMatchMaterial))
                {
                    LogMessage($"  Found similar material {bestMatchMaterial} with description matching for item {itemCode}");
                    return MaterialInfoMappings[bestMatchMaterial].HSCode;
                }
            }

            // Method 3: Classify based on description and assign category-based HS code
            string category = DetermineProductCategory(description);
            if (CategoryHSCodes.ContainsKey(category))
            {
                LogMessage($"  Assigning HS code based on product category '{category}' for item {itemCode}");
                return CategoryHSCodes[category];
            }

            // Final fallback - should never reach here, but just in case
            LogMessage($"  Using default HS code (HERBICIDES) for item {itemCode}");
            return CategoryHSCodes["HERBICIDES"];
        }

        // Find the most similar material based on description text
        private string FindMostSimilarMaterial(string description)
        {
            if (string.IsNullOrEmpty(description))
                return null;

            // Split description into words
            string[] words = description.ToLower()
                .Split(new[] { ' ', ',', '.', '-', '/', '+', '(', ')', '[', ']' }, StringSplitOptions.RemoveEmptyEntries);

            // Dictionary to count word matches by material
            Dictionary<string, int> materialMatchCount = new Dictionary<string, int>();

            // Count matches for each word
            foreach (string word in words)
            {
                // Skip very short words
                if (word.Length <= 2)
                    continue;

                if (wordToMaterialMap.ContainsKey(word))
                {
                    foreach (string materialNum in wordToMaterialMap[word])
                    {
                        if (!materialMatchCount.ContainsKey(materialNum))
                            materialMatchCount[materialNum] = 0;

                        materialMatchCount[materialNum]++;
                    }
                }
            }

            // Find the material with the most matches
            string bestMatch = null;
            int maxMatches = 0;

            foreach (var entry in materialMatchCount)
            {
                if (entry.Value > maxMatches)
                {
                    maxMatches = entry.Value;
                    bestMatch = entry.Key;
                }
            }

            // Only return a match if we have a reasonable number of matching words
            return maxMatches >= 2 ? bestMatch : null;
        }

        // Determine product category from description
        private string DetermineProductCategory(string description)
        {
            if (string.IsNullOrEmpty(description))
                return "HERBICIDES"; // Default category

            string upperDesc = description.ToUpper();

            // Check for herbicide indicators
            if (upperDesc.Contains("GLYPHOSATE") ||
                upperDesc.Contains("ATRAZINE") ||
                upperDesc.Contains("DMA") ||
                upperDesc.Contains("HERBICIDE") ||
                upperDesc.Contains("WEED"))
            {
                return "HERBICIDES";
            }

            // Check for insecticide indicators
            if (upperDesc.Contains("EMAMECTIN") ||
                upperDesc.Contains("THIAMETHOXAM") ||
                upperDesc.Contains("LAMBDA") ||
                upperDesc.Contains("CYHALOTHRIN") ||
                upperDesc.Contains("INSECTICIDE") ||
                upperDesc.Contains("PEST"))
            {
                return "INSECTICIDES";
            }

            // Check for fungicide indicators
            if (upperDesc.Contains("TEBUCONAZOLE") ||
                upperDesc.Contains("AZOXYSTROBIN") ||
                upperDesc.Contains("DIFENOCONAZOLE") ||
                upperDesc.Contains("CAPTAN") ||
                upperDesc.Contains("FUNGICIDE") ||
                upperDesc.Contains("DISEASE"))
            {
                return "FUNGICIDES";
            }

            // Default to HERBICIDES if we can't determine
            return "HERBICIDES";
        }

        private string ExtractDescription(string line, string[] lines, int lineIndex, string itemCode)
        {
            // Try to extract description from the current line
            int itemCodePos = line.IndexOf(itemCode);
            if (itemCodePos > 0)
            {
                int startPos = line.IndexOf(" ", itemCodePos) + 1;
                if (startPos > 0 && startPos < line.Length)
                {
                    // Find the end of description (before quantity)
                    Match quantityMatch = Regex.Match(line.Substring(startPos), @"\d+\.\d{2}\s+Piece");
                    if (quantityMatch.Success)
                    {
                        int endPos = startPos + quantityMatch.Index;
                        if (endPos > startPos)
                        {
                            return line.Substring(startPos, endPos - startPos).Trim();
                        }
                    }

                    // If no quantity match, try to find where numbers start
                    Match numberMatch = Regex.Match(line.Substring(startPos), @"\d+\.\d{2}");
                    if (numberMatch.Success)
                    {
                        int endPos = startPos + numberMatch.Index;
                        if (endPos > startPos)
                        {
                            return line.Substring(startPos, endPos - startPos).Trim();
                        }
                    }
                }
            }

            // If description not found, try the next line
            if (lineIndex + 1 < lines.Length)
            {
                string nextLine = lines[lineIndex + 1].Trim();
                if (!nextLine.StartsWith(itemCode) && !Regex.IsMatch(nextLine, @"^\d{1,2}\s+"))
                {
                    return nextLine;
                }
            }

            return "Unknown Item";
        }

       
        private async Task<FiscalResponseData> WaitForFiscalizationResponse(string invoiceNumber)
        {
            LogMessage($"Waiting for fiscalization response for invoice {invoiceNumber}...");

            string pcFileName = $"{PostingFilenamePrefix}{invoiceNumber}.txt";
            string sentFileName = $"{SentFolderFilenamePrefix}{pcFileName}";
            string sentFilePath = Path.Combine(SentFolder, sentFileName);
            string failedFilePath = Path.Combine(FailFolder, pcFileName);

            // First check if response file already exists (maybe from a previous run)
            if (File.Exists(sentFilePath))
            {
                LogMessage($"Found existing response file: {sentFilePath}");
                return await ReadFiscalResponseFromFile(sentFilePath, invoiceNumber);
            }

            // Check for failure
            if (File.Exists(failedFilePath))
            {
                LogMessage($"Found failed fiscalization file: {failedFilePath}");
                return null;
            }

            // Set up a task completion source to wait for the response
            var tcs = new TaskCompletionSource<FiscalResponseData>();
            FiscalizationTasks[invoiceNumber] = tcs;

            // Set up file system watchers to detect when the response arrives
            SetupFileWatchers(invoiceNumber, pcFileName);

            // Set up a cancellation token to handle timeout
            var cts = new CancellationTokenSource();
            WatcherCancellationTokens[invoiceNumber] = cts;

            // Start a task to handle timeout
            var timeoutTask = Task.Delay(FileWatcherWaitMilliseconds, cts.Token);

            try
            {
                // Wait for either the response or timeout
                var completedTask = await Task.WhenAny(tcs.Task, timeoutTask);

                if (completedTask == tcs.Task)
                {
                    // Response received
                    cts.Cancel();
                    return await tcs.Task;
                }
                else
                {
                    // Timeout
                    LogMessage($"Timeout waiting for fiscalization response for invoice {invoiceNumber}");

                    // Try one last time to check the database
                    var dataFromDb = await TryReadFromSQLiteDatabase(invoiceNumber);
                    if (dataFromDb != null)
                    {
                        LogMessage($"Found fiscal data in SQLite database after timeout");
                        return dataFromDb;
                    }

                    return null;
                }
            }
            finally
            {
                // Clean up resources
                CleanupFileWatchers(invoiceNumber);
            }
        }

        private void SetupFileWatchers(string invoiceNumber, string pcFileName)
        {
            LogMessage($"Setting up file watchers for invoice {invoiceNumber}");

            string sentFileName = $"{SentFolderFilenamePrefix}{pcFileName}";

            // Watcher for the sent folder (success case)
            var sentWatcher = new FileSystemWatcher(SentFolder)
            {
                Filter = "*.txt",
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime
            };

            sentWatcher.Created += async (sender, e) => {
                if (e.Name.Equals(sentFileName, StringComparison.OrdinalIgnoreCase))
                {
                    LogMessage($"Detected creation of response file: {e.Name}");

                    // Wait a moment for the file to be fully written
                    await Task.Delay(1000);

                    // Try to read from SQLite database first
                    var dataFromDb = await TryReadFromSQLiteDatabase(invoiceNumber);

                    if (dataFromDb != null)
                    {
                        // Complete the task with data from the database
                        if (FiscalizationTasks.ContainsKey(invoiceNumber) && !FiscalizationTasks[invoiceNumber].Task.IsCompleted)
                        {
                            FiscalizationTasks[invoiceNumber].SetResult(dataFromDb);
                        }
                    }
                    else
                    {
                        // Fall back to reading from the file
                        var dataFromFile = await ReadFiscalResponseFromFile(e.FullPath, invoiceNumber);

                        if (dataFromFile != null && FiscalizationTasks.ContainsKey(invoiceNumber) &&
                            !FiscalizationTasks[invoiceNumber].Task.IsCompleted)
                        {
                            FiscalizationTasks[invoiceNumber].SetResult(dataFromFile);
                        }
                    }
                }
            };

            // Watcher for the fail folder (failure case)
            var failWatcher = new FileSystemWatcher(FailFolder)
            {
                Filter = "*.txt",
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.CreationTime
            };

            failWatcher.Created += (sender, e) => {
                if (e.Name.Equals(pcFileName, StringComparison.OrdinalIgnoreCase))
                {
                    LogMessage($"Detected creation of failure file: {e.Name}");

                    // Signal failure
                    if (FiscalizationTasks.ContainsKey(invoiceNumber) && !FiscalizationTasks[invoiceNumber].Task.IsCompleted)
                    {
                        FiscalizationTasks[invoiceNumber].SetResult(null);
                    }
                }
            };

            // Enable the watchers
            sentWatcher.EnableRaisingEvents = true;
            failWatcher.EnableRaisingEvents = true;

            // Store the watchers
            FileWatchers[invoiceNumber] = sentWatcher;
            FailedFileWatchers[invoiceNumber] = failWatcher;

            LogMessage($"File watchers set up for invoice {invoiceNumber}");
        }

        private void CleanupFileWatchers(string invoiceNumber)
        {
            LogMessage($"Cleaning up resources for invoice {invoiceNumber}");

            // Dispose file watchers
            if (FileWatchers.ContainsKey(invoiceNumber))
            {
                FileWatchers[invoiceNumber].EnableRaisingEvents = false;
                FileWatchers[invoiceNumber].Dispose();
                FileWatchers.Remove(invoiceNumber);
            }

            if (FailedFileWatchers.ContainsKey(invoiceNumber))
            {
                FailedFileWatchers[invoiceNumber].EnableRaisingEvents = false;
                FailedFileWatchers[invoiceNumber].Dispose();
                FailedFileWatchers.Remove(invoiceNumber);
            }

            // Dispose cancellation token
            if (WatcherCancellationTokens.ContainsKey(invoiceNumber))
            {
                if (!WatcherCancellationTokens[invoiceNumber].IsCancellationRequested)
                {
                    WatcherCancellationTokens[invoiceNumber].Cancel();
                }

                WatcherCancellationTokens[invoiceNumber].Dispose();
                WatcherCancellationTokens.Remove(invoiceNumber);
            }

            // Remove task completion source
            if (FiscalizationTasks.ContainsKey(invoiceNumber))
            {
                FiscalizationTasks.Remove(invoiceNumber);
            }
        }

        private async Task<FiscalResponseData> TryReadFromSQLiteDatabase(string invoiceNumber)
        {
            LogMessage($"Trying to read fiscal data from SQLite database for invoice {invoiceNumber}");

            try
            {
                // Construct connection string
                string dbPath = Path.Combine(DBFolder, "FbTransaction.db");
                if (!File.Exists(dbPath))
                {
                    LogMessage($"SQLite database file not found: {dbPath}");
                    return null;
                }

                string connectionString = $"Data Source={dbPath};Version=3;";

                // Try to read data from database
                using (var connection = new SQLiteConnection(connectionString))
                {
                    await connection.OpenAsync();

                    string sql = $"SELECT Date, TsNum, ControlCode, SerialNumber, QrCode FROM fb_transaction WHERE TsNum = @TsNum";

                    using (var command = new SQLiteCommand(sql, connection))
                    {
                        command.Parameters.AddWithValue("@TsNum", invoiceNumber);

                        using (var reader = await command.ExecuteReaderAsync())
                        {
                            if (await reader.ReadAsync())
                            {
                                // Extract data from the result
                                var result = new FiscalResponseData
                                {
                                    TransactionDate = reader.GetDateTime(0),
                                    TsNum = reader.GetString(1),
                                    ControlCode = reader.GetString(2),
                                    SerialNumber = reader.GetString(3),
                                    FiscalSeal = reader.GetString(4),
                                };

                                // Construct the fiscal footer
                                result.FiscalFooter = BuildFiscalFooter(result);

                                LogMessage($"Successfully read fiscal data from SQLite database");
                                return result;
                            }
                        }
                    }
                }

                LogMessage($"No data found in SQLite database for invoice {invoiceNumber}");
                return null;
            }
            catch (Exception ex)
            {
                LogMessage($"Error reading from SQLite database: {ex.Message}");
                return null;
            }
        }

       
        
        private string GenerateQRCode(string fiscalSeal, string invoiceNumber)
        {
            LogMessage($"Generating QR code for invoice {invoiceNumber}");

            try
            {
                // Create QR code generator
                using (var qrGenerator = new QRCodeGenerator())
                {
                    // Generate QR code data
                    using (var qrCodeData = qrGenerator.CreateQrCode(fiscalSeal, QRCodeGenerator.ECCLevel.Q))
                    {
                        // Create QR code
                        using (var qrCode = new QRCode(qrCodeData))
                        {
                            // Get QR code as bitmap
                            using (var qrCodeImage = qrCode.GetGraphic(10))
                            {
                                // Save to file
                                string qrCodePath = Path.Combine(QRCodeFolder, $"{invoiceNumber}.png");
                                qrCodeImage.Save(qrCodePath, System.Drawing.Imaging.ImageFormat.Png);

                                LogMessage($"QR code saved to {qrCodePath}");
                                return qrCodePath;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error generating QR code: {ex.Message}");
                throw new Exception($"Failed to generate QR code: {ex.Message}", ex);
            }
        }

         // UI Controls
        private Button btnLoadHSCodes; 
        private Button btnLoadInvoices;
        private Button btnProcessInvoices;
        private ListBox listBoxInvoices;
        private Label lblStatus;
        private ProgressBar progressBar;
        private Label lblHSCodesStatus;
    }

    public class MaterialInfo
    {
        public string MaterialNumber { get; set; }
        public string Description { get; set; }
        public string Category { get; set; }
        public string HSCode { get; set; }
    }

    public class FiscalResponseData
    {
        public DateTime TransactionDate { get; set; }
        public string TsNum { get; set; }
        public string ControlCode { get; set; }
        public string SerialNumber { get; set; }
        public string FiscalSeal { get; set; }
        public string FiscalFooter { get; set; }
    }

    public class Invoice
    {
        public string InvoiceNumber { get; set; }
        public string FederalTaxID { get; set; }
        public List<LineItem> LineItems { get; set; } = new List<LineItem>();
    }

    public class LineItem
    {
        public string ItemCode { get; set; }
        public string Description { get; set; }
        public string Quantity { get; set; }
        public string UnitPrice { get; set; }
        public string DiscountPercent { get; set; }
        public string UnitPriceAfterDiscount { get; set; }
        public string Amount { get; set; }
        public string HSCode { get; set; }
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