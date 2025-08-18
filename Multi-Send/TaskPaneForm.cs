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

            // Initialize WebView when the control is actually loaded
            this.Load += TaskPaneForm_Load;
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

        private async void TaskPaneForm_Load(object sender, EventArgs e)
        {
            await InitializeWebViewAsync();
        }

        private async Task InitializeWebViewAsync()
        {
            try
            {
                // Create a custom user data folder in a writable location
                string userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "EmailDuplicator");

                Directory.CreateDirectory(userDataFolder);

                var environment = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
                await webView.EnsureCoreWebView2Async(environment);
                webView.CoreWebView2.WebMessageReceived += WebView_MessageReceived;

                string htmlPath;

#if DEBUG
                // Always use your dev repo folder when running under F5 / Debug
                htmlPath = @"C:\Users\hughb\source\repos\Multi-Send\Multi-Send\TaskPaneUI\index.html";
#else
        // In published builds, use the add-in deployment folder
        string assemblyDir = Path.GetDirectoryName(
            System.Reflection.Assembly.GetExecutingAssembly().Location);
        System.Diagnostics.Debug.WriteLine("Add-in loaded from: " + assemblyDir);

        htmlPath = Path.Combine(assemblyDir, "TaskPaneUI", "index.html");
#endif

                if (File.Exists(htmlPath))
                {
                    htmlFilePath = htmlPath;
                    string navPath = $"file:///{htmlPath.Replace('\\', '/')}";
                    System.Diagnostics.Debug.WriteLine("[TaskPaneForm] Loading UI from: " + navPath);
                    webView.CoreWebView2.Navigate(navPath);
                }
                else
                {
                    string errorHtml = $@"
                <!DOCTYPE html>
                <html>
                <body style='font-family: Segoe UI; padding:20px;'>
                    <h3>TaskPaneUI\index.html not found</h3>
                    <p>Expected path: {htmlPath}</p>
                </body>
                </html>";
                    webView.CoreWebView2.NavigateToString(errorHtml);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"WebView2 init error: {ex.Message}");
            }
        }


        private async void WebView_MessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            try
            {
                string message = "";

                try
                {
                    message = e.TryGetWebMessageAsString();
                }
                catch
                {
                    message = e.WebMessageAsJson;
                    if (message.StartsWith("\"") && message.EndsWith("\""))
                    {
                        message = message.Substring(1, message.Length - 2).Replace("\\\"", "\"");
                    }
                }

                if (message.StartsWith("{"))
                {
                    var jObj = JsonConvert.DeserializeObject<Newtonsoft.Json.Linq.JObject>(message);
                    string actionName = jObj?["action"]?.ToString() ?? "";

                    switch (actionName)
                    {
                        case "test":
                            SendResponseToJS("success", "🎉 Perfect! Communication working!");
                            break;

                        case "duplicateEmail":
                            await HandleDuplicateEmail(jObj["data"]);
                            break;

                        case "detectPlaceholder":
                            await HandleDetectPlaceholder();
                            break;

                        default:
                            SendResponseToJS("error", $"Unknown action: {actionName}");
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

                var response = new { type, message, data };
                string jsonResponse = JsonConvert.SerializeObject(response);

                // Check if we're on the UI thread
                if (this.InvokeRequired)
                {
                    // We're on a background thread, marshal to UI thread
                    this.Invoke(new System.Action(() => {
                        try
                        {
                            if (!webView.IsDisposed && webView.IsHandleCreated && webView.CoreWebView2 != null)
                            {
                                webView.CoreWebView2.PostWebMessageAsString(jsonResponse);
                            }
                        }
                        catch (System.Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"UI thread WebView2 error: {ex.Message}");
                        }
                    }));
                }
                else
                {
                    // We're already on UI thread
                    if (!webView.IsDisposed && webView.IsHandleCreated && webView.CoreWebView2 != null)
                    {
                        webView.CoreWebView2.PostWebMessageAsString(jsonResponse);
                    }
                }
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
                if (webView?.CoreWebView2 == null) return;

                var response = new { type = "progress", current, total, message };
                string jsonResponse = JsonConvert.SerializeObject(response);

                // Use the same threading logic as SendResponseToJS
                if (this.InvokeRequired)
                {
                    // We're on a background thread, marshal to UI thread
                    this.Invoke(new System.Action(() => {
                        try
                        {
                            if (!webView.IsDisposed && webView.IsHandleCreated && webView.CoreWebView2 != null)
                            {
                                webView.CoreWebView2.PostWebMessageAsString(jsonResponse);
                            }
                        }
                        catch (System.Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"UI thread progress error: {ex.Message}");
                        }
                    }));
                }
                else
                {
                    // We're already on UI thread
                    if (!webView.IsDisposed && webView.IsHandleCreated && webView.CoreWebView2 != null)
                    {
                        webView.CoreWebView2.PostWebMessageAsString(jsonResponse);
                    }
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error sending progress to JS: {ex.Message}");
            }
        }

        private async Task HandleDuplicateEmail(Newtonsoft.Json.Linq.JToken requestData)
        {
            try
            {
                if (outlookApp == null)
                {
                    SendResponseToJS("error", "Outlook application not available. Please restart the add-in.");
                    return;
                }

                string placeholder = requestData["placeholder"]?.ToString() ?? "";
                var recipients = requestData["recipients"]?.ToObject<List<Recipient>>() ?? new List<Recipient>();

                SendResponseToJS("info", $"Starting duplication for {recipients.Count} recipients...");

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

                var emailData = ExtractEmailData(sourceMailItem);
                int successCount = 0;

                for (int i = 0; i < recipients.Count; i++)
                {
                    try
                    {
                        SendProgressToJS(i + 1, recipients.Count,
                            $"Creating email {i + 1}/{recipients.Count} for {recipients[i].Name}...");
                        await CreateDuplicateEmail(emailData, placeholder, recipients[i]);
                        successCount++;
                        await Task.Delay(100); // keep UI responsive
                    }
                    catch (System.Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Failed to create duplicate for {recipients[i].Email}: {ex.Message}");
                    }
                }

                if (successCount == recipients.Count)
                {
                    SendResponseToJS("success", $"✅ Successfully created {successCount} duplicate emails! Check Drafts.");
                }
                else
                {
                    SendResponseToJS("error", $"⚠️ Created {successCount} out of {recipients.Count}. Some failed.");
                }

                CleanupTempFiles(emailData.Attachments);

                // Release COM ref
                if (sourceMailItem != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceMailItem);
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
                newMail = outlookApp.CreateItem(OlItemType.olMailItem) as MailItem;

                newMail.Subject = ReplacePlaceholder(sourceData.Subject, placeholder, recipient.Name);
                newMail.Body = ReplacePlaceholder(sourceData.Body, placeholder, recipient.Name);
                newMail.HTMLBody = ReplacePlaceholder(sourceData.HTMLBody, placeholder, recipient.Name);
                newMail.Importance = sourceData.Importance;
                newMail.Sensitivity = sourceData.Sensitivity;

                newMail.Recipients.Add(recipient.Email);
                newMail.Recipients.ResolveAll();

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

                newMail.Save(); // Save as draft
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

                string text = $"{mailItem.Subject} {mailItem.Body}";
                var placeholderPatterns = new[]
                {
            @"\{\{[^}]+\}\}",
            @"\[[^\]]+\]",
            @"\$[A-Za-z_][A-Za-z0-9_]*",
        };

                foreach (var pattern in placeholderPatterns)
                {
                    var matches = Regex.Matches(text, pattern, RegexOptions.IgnoreCase);
                    if (matches.Count > 0)
                    {
                        string detectedPlaceholder = matches[0].Value;
                        SendResponseToJS("success", $"Detected placeholder: {detectedPlaceholder}",
                            new { placeholder = detectedPlaceholder });
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
            foreach (var attachment in attachments)
            {
                try
                {
                    if (File.Exists(attachment.TempFilePath))
                    {
                        File.Delete(attachment.TempFilePath);
                    }
                }
                catch { /* ignore */ }
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

    // Data classes
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