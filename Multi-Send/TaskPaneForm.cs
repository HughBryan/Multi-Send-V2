using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Multi_Send
{
    public partial class TaskPaneForm : UserControl
    {
        private WebView2 webView;
        private string htmlFilePath;
        private Microsoft.Office.Interop.Outlook.Application outlookApp;

        public TaskPaneForm()
        {
            InitializeComponent();
            InitializeOutlookApp();
            InitializeWebView();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // UserControl properties
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = AutoScaleMode.Font;
            this.Size = new System.Drawing.Size(400, 600);
            this.Name = "TaskPaneForm";
            this.BackColor = System.Drawing.Color.White;

            // WebView2 control
            this.webView = new WebView2()
            {
                Dock = DockStyle.Fill,
                Name = "webView"
            };

            this.Controls.Add(this.webView);
            this.ResumeLayout(false);
        }

        private void InitializeOutlookApp()
        {
            try
            {
                this.outlookApp = Globals.ThisAddIn.Application;
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Warning: Could not initialize Outlook Application: {ex.Message}");
                this.outlookApp = null;
            }
        }

        private async void InitializeWebView()
        {
            try
            {
                // Create a custom user data folder in a writable location
                string userDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "EmailDuplicator");

                // Ensure the directory exists and is writable
                if (!Directory.Exists(userDataFolder))
                {
                    Directory.CreateDirectory(userDataFolder);
                }

                // Create WebView2 environment with custom data folder
                var environment = await CoreWebView2Environment.CreateAsync(null, userDataFolder);

                // Initialize WebView2 with the custom environment
                await webView.EnsureCoreWebView2Async(environment);

                // Set up message handling (JavaScript → C#)
                webView.CoreWebView2.WebMessageReceived += WebView_MessageReceived;

                // Get the path to your HTML file
                string projectPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                htmlFilePath = Path.Combine(projectPath, "TaskPaneUI", "index.html");

                // Navigate to your HTML file
                if (File.Exists(htmlFilePath))
                {
                    webView.CoreWebView2.Navigate($"file:///{htmlFilePath.Replace('\\', '/')}");
                }
                else
                {
                    // Show error if HTML files not found
                    string errorHtml = $@"<!DOCTYPE html>
                    <html><head><title>Email Duplicator</title>
                    <style>body {{ font-family: Segoe UI; padding: 20px; }} .error {{ color: red; }}</style></head>
                    <body>
                        <h3>Email Duplicator</h3>
                        <div class='error'>
                            <p>⚠️ HTML files not found!</p>
                            <p>Expected path: {htmlFilePath}</p>
                            <p>Please make sure your TaskPaneUI folder contains index.html, taskpane.css, and taskpane.js</p>
                        </div>
                    </body></html>";

                    webView.CoreWebView2.NavigateToString(errorHtml);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"WebView2 initialization error: {ex.Message}");
            }
        }

        private void WebView_MessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                string message = "";

                // Use the working method to get the message (handles COM registration issues)
                try
                {
                    message = e.TryGetWebMessageAsString();
                }
                catch (System.Exception)
                {
                    // Fallback for COM registration issues
                    message = e.WebMessageAsJson;
                    // Remove JSON escaping if present
                    if (message.StartsWith("\"") && message.EndsWith("\""))
                    {
                        message = message.Substring(1, message.Length - 2).Replace("\\\"", "\"");
                    }
                }

                // Parse the JSON message
                if (message.StartsWith("{"))
                {
                    dynamic messageData = JsonConvert.DeserializeObject(message);
                    string action = messageData.action?.ToString() ?? "";

                    switch (action)
                    {
                        case "test":
                            SendResponseToJS("success", "🎉 Perfect! Communication working!");
                            break;

                        case "duplicateEmail":
                            _ = Task.Run(async () => await HandleDuplicateEmail(messageData.data));
                            break;

                        case "detectPlaceholder":
                            _ = Task.Run(async () => await HandleDetectPlaceholder());
                            break;

                        default:
                            SendResponseToJS("error", $"Unknown action: {action}");
                            break;
                    }
                }
                else
                {
                    SendResponseToJS("success", $"Received: {message}");
                }
            }
            catch (System.Exception ex)
            {
                SendResponseToJS("error", $"Error processing message: {ex.Message}");
            }
        }

        private void SendResponseToJS(string type, string message, object data = null)
        {
            try
            {
                if (webView?.CoreWebView2 == null) return;

                if (webView.InvokeRequired)
                {
                    webView.Invoke(new System.Action(() => SendResponseToJS(type, message, data)));
                    return;
                }

                var response = new
                {
                    type = type,
                    message = message,
                    data = data
                };

                string jsonResponse = JsonConvert.SerializeObject(response);
                webView.CoreWebView2.PostWebMessageAsString(jsonResponse);
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error sending response to JS: {ex.Message}");
            }
        }

        private void SendProgressToJS(int current, int total, string message)
        {
            try
            {
                if (webView.InvokeRequired)
                {
                    webView.Invoke(new System.Action(() => SendProgressToJS(current, total, message)));
                    return;
                }

                var response = new
                {
                    type = "progress",
                    current = current,
                    total = total,
                    message = message
                };

                string jsonResponse = JsonConvert.SerializeObject(response);
                webView.CoreWebView2.PostWebMessageAsString(jsonResponse);
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error sending progress to JS: {ex.Message}");
            }
        }

        private async Task HandleDuplicateEmail(dynamic requestData)
        {
            try
            {
                if (outlookApp == null)
                {
                    SendResponseToJS("error", "Outlook application not available. Please restart the add-in.");
                    return;
                }

                string placeholder = requestData.placeholder;
                var recipients = JsonConvert.DeserializeObject<List<Recipient>>(requestData.recipients.ToString());

                SendResponseToJS("info", $"Starting duplication for {recipients.Count} recipients...");

                // Get the currently selected email
                var selectedItem = outlookApp.ActiveExplorer().Selection;

                if (selectedItem.Count == 0)
                {
                    SendResponseToJS("error", "Please select an email to duplicate.");
                    return;
                }

                var sourceMailItem = selectedItem[1] as MailItem;
                if (sourceMailItem == null)
                {
                    SendResponseToJS("error", "Selected item is not an email.");
                    return;
                }

                // Extract email data including attachments
                var emailData = ExtractEmailData(sourceMailItem);

                int successCount = 0;
                int totalCount = recipients.Count;

                // Create duplicates for each recipient
                for (int i = 0; i < recipients.Count; i++)
                {
                    try
                    {
                        SendProgressToJS(i + 1, totalCount, $"Creating email {i + 1}/{totalCount} for {recipients[i].Name}...");

                        await CreateDuplicateEmail(emailData, placeholder, recipients[i]);
                        successCount++;

                        // Small delay to prevent overwhelming Outlook
                        await Task.Delay(100);
                    }
                    catch (System.Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Failed to create duplicate for {recipients[i].Email}: {ex.Message}");
                    }
                }

                // Send final result
                if (successCount == totalCount)
                {
                    SendResponseToJS("success", $"✅ Successfully created {successCount} duplicate emails! Check your Drafts folder.");
                }
                else
                {
                    SendResponseToJS("error", $"⚠️ Created {successCount} out of {totalCount} duplicates. Some failed - check debug output for details.");
                }

                // Clean up temp files
                CleanupTempFiles(emailData.Attachments);
            }
            catch (System.Exception ex)
            {
                SendResponseToJS("error", $"Duplication failed: {ex.Message}");
            }
        }

        private EmailData ExtractEmailData(MailItem sourceEmail)
        {
            var emailData = new EmailData
            {
                Subject = sourceEmail.Subject ?? "",
                Body = sourceEmail.Body ?? "",
                HTMLBody = sourceEmail.HTMLBody ?? "",
                Importance = sourceEmail.Importance,
                Sensitivity = sourceEmail.Sensitivity,
                Attachments = new List<AttachmentData>()
            };

            // Extract attachments
            foreach (Attachment attachment in sourceEmail.Attachments)
            {
                try
                {
                    string tempPath = Path.Combine(Path.GetTempPath(), $"EmailDup_{Guid.NewGuid()}_{attachment.FileName}");
                    attachment.SaveAsFile(tempPath);

                    emailData.Attachments.Add(new AttachmentData
                    {
                        FileName = attachment.FileName,
                        TempFilePath = tempPath,
                        Type = attachment.Type
                    });
                }
                catch (System.Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to extract attachment {attachment.FileName}: {ex.Message}");
                }
            }

            return emailData;
        }

        private async Task CreateDuplicateEmail(EmailData sourceData, string placeholder, Recipient recipient)
        {
            MailItem newMail = null;
            try
            {
                // Create new mail item
                newMail = outlookApp.CreateItem(OlItemType.olMailItem) as MailItem;

                // Replace placeholder with recipient name
                string personalizedSubject = ReplacePlaceholder(sourceData.Subject, placeholder, recipient.Name);
                string personalizedBody = ReplacePlaceholder(sourceData.Body, placeholder, recipient.Name);
                string personalizedHTMLBody = ReplacePlaceholder(sourceData.HTMLBody, placeholder, recipient.Name);

                // Set email properties
                newMail.Subject = personalizedSubject;
                newMail.Body = personalizedBody;
                newMail.HTMLBody = personalizedHTMLBody;
                newMail.Importance = sourceData.Importance;
                newMail.Sensitivity = sourceData.Sensitivity;

                // Add recipient
                newMail.Recipients.Add(recipient.Email);
                newMail.Recipients.ResolveAll();

                // Add attachments
                foreach (var attachmentData in sourceData.Attachments)
                {
                    try
                    {
                        if (File.Exists(attachmentData.TempFilePath))
                        {
                            newMail.Attachments.Add(attachmentData.TempFilePath, attachmentData.Type, 1, attachmentData.FileName);
                        }
                    }
                    catch (System.Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Failed to add attachment {attachmentData.FileName}: {ex.Message}");
                    }
                }

                // Save as draft (don't send automatically)
                newMail.Save();
            }
            catch (System.Exception ex)
            {
                newMail?.Close(OlInspectorClose.olDiscard);
                throw new System.Exception($"Failed to create duplicate email for {recipient.Email}: {ex.Message}");
            }
            finally
            {
                if (newMail != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(newMail);
                }
            }
        }

        private string ReplacePlaceholder(string text, string placeholder, string replacement)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(placeholder))
                return text;

            // Use regex for case-insensitive replacement
            string escapedPlaceholder = Regex.Escape(placeholder);
            return Regex.Replace(text, escapedPlaceholder, replacement, RegexOptions.IgnoreCase);
        }

        private async Task HandleDetectPlaceholder()
        {
            try
            {
                if (outlookApp == null)
                {
                    SendResponseToJS("error", "Outlook application not available.");
                    return;
                }

                var selectedItem = outlookApp.ActiveExplorer().Selection;

                if (selectedItem.Count == 0)
                {
                    SendResponseToJS("error", "Please select an email to detect placeholder from.");
                    return;
                }

                var mailItem = selectedItem[1] as MailItem;
                if (mailItem == null)
                {
                    SendResponseToJS("error", "Selected item is not an email.");
                    return;
                }

                // Look for common placeholder patterns
                string text = $"{mailItem.Subject} {mailItem.Body}";
                var placeholderPatterns = new[]
                {
                    @"\{\{[^}]+\}\}", // {{name}}, {{firstname}}, etc.
                    @"\[[^\]]+\]",   // [name], [firstname], etc.
                    @"\$[A-Za-z_][A-Za-z0-9_]*", // $name, $firstname, etc.
                };

                foreach (var pattern in placeholderPatterns)
                {
                    var matches = Regex.Matches(text, pattern, RegexOptions.IgnoreCase);
                    if (matches.Count > 0)
                    {
                        string detectedPlaceholder = matches[0].Value;
                        SendResponseToJS("success", $"Detected placeholder: {detectedPlaceholder}", new { placeholder = detectedPlaceholder });
                        return;
                    }
                }

                SendResponseToJS("info", "No common placeholder patterns found. Try {{name}} or [name].");
            }
            catch (System.Exception ex)
            {
                SendResponseToJS("error", $"Error detecting placeholder: {ex.Message}");
            }
        }

        private void CleanupTempFiles(List<AttachmentData> attachments)
        {
            try
            {
                foreach (var attachment in attachments)
                {
                    try
                    {
                        if (File.Exists(attachment.TempFilePath))
                        {
                            File.Delete(attachment.TempFilePath);
                        }
                    }
                    catch
                    {
                        // Ignore errors when cleaning up individual files
                    }
                }
            }
            catch
            {
                // Ignore errors during cleanup
            }
        }

        // Clean up when control is disposed
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                webView?.Dispose();
            }
            base.Dispose(disposing);
        }
    }

    // Data classes for email duplication
    public class Recipient
    {
        public string Email { get; set; }
        public string Name { get; set; }
    }

    public class EmailData
    {
        public string Subject { get; set; }
        public string Body { get; set; }
        public string HTMLBody { get; set; }
        public OlImportance Importance { get; set; }
        public OlSensitivity Sensitivity { get; set; }
        public List<AttachmentData> Attachments { get; set; }
    }

    public class AttachmentData
    {
        public string FileName { get; set; }
        public string TempFilePath { get; set; }
        public OlAttachmentType Type { get; set; }
    }
}