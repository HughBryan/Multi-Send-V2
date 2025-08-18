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
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = AutoScaleMode.Dpi; // was Font/None
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

        private void TaskPaneForm_Load(object sender, EventArgs e)
        {
            // Initialize WebView2 directly on UI thread (this event is already on UI thread)
            InitializeWebViewSafe();
        }

        private void InitializeWebViewSafe()
        {
            // Use async void for fire-and-forget from UI thread
            InitializeWebViewAsync();
        }



        private async void InitializeWebViewAsync()
        {
            string step = "Starting";
            try
            {
                step = "Creating user data folder path";
                string userDataFolder = Path.Combine(Path.GetTempPath(), "EmailDuplicator_WebView2");
                System.Diagnostics.Debug.WriteLine($"WebView2 Debug: {step} - Path: {userDataFolder}");

                step = "Creating directory";
                try
                {
                    Directory.CreateDirectory(userDataFolder);
                    System.Diagnostics.Debug.WriteLine($"WebView2 Debug: Directory created successfully");
                }
                catch (System.Exception ex)
                {
                    throw new InvalidOperationException($"Cannot create WebView2 user data folder: {ex.Message}");
                }

                step = "Creating CoreWebView2Environment";
                System.Diagnostics.Debug.WriteLine($"WebView2 Debug: {step}");
                var environment = await CoreWebView2Environment.CreateAsync(null, userDataFolder, null);
                System.Diagnostics.Debug.WriteLine($"WebView2 Debug: Environment created successfully");

                step = "Ensuring CoreWebView2";
                System.Diagnostics.Debug.WriteLine($"WebView2 Debug: {step}");
                await webView.EnsureCoreWebView2Async(environment);
                System.Diagnostics.Debug.WriteLine($"WebView2 Debug: CoreWebView2 ensured successfully");

                step = "Attaching WebMessageReceived event";
                System.Diagnostics.Debug.WriteLine($"WebView2 Debug: {step}");
                webView.CoreWebView2.WebMessageReceived += WebView_MessageReceived;
                System.Diagnostics.Debug.WriteLine($"WebView2 Debug: Event attached successfully");

                step = "Determining HTML path";
                string htmlPath;
#if DEBUG
                htmlPath = @"C:\Users\hughb\source\repos\Multi-Send\Multi-Send\TaskPaneUI\index.html";
#else
        string assemblyDir = Path.GetDirectoryName(
            System.Reflection.Assembly.GetExecutingAssembly().Location);
        htmlPath = Path.Combine(assemblyDir, "TaskPaneUI", "index.html");
#endif
                System.Diagnostics.Debug.WriteLine($"WebView2 Debug: HTML path: {htmlPath}");

                step = "Checking if HTML file exists";
                if (File.Exists(htmlPath))
                {
                    step = "Navigating to HTML file";
                    System.Diagnostics.Debug.WriteLine($"WebView2 Debug: {step}");
                    htmlFilePath = htmlPath;

                    // Use Uri for proper file URL formatting
                    string navPath = new Uri(htmlPath).ToString();
                    System.Diagnostics.Debug.WriteLine($"WebView2 Debug: Navigation URL: {navPath}");

                    webView.CoreWebView2.Navigate(navPath);
                    System.Diagnostics.Debug.WriteLine($"WebView2 Debug: Navigation initiated successfully");
                }
                else
                {
                    step = "Navigating to error HTML";
                    System.Diagnostics.Debug.WriteLine($"WebView2 Debug: {step} - File not found: {htmlPath}");
                    string errorHtml = $@"
        <!DOCTYPE html>
        <html><body style='font-family: Segoe UI; padding:20px;'>
            <h3>TaskPaneUI\index.html not found</h3>
            <p>Expected path: {htmlPath}</p>
        </body></html>";
                    webView.CoreWebView2.NavigateToString(errorHtml);
                    System.Diagnostics.Debug.WriteLine($"WebView2 Debug: Error HTML navigation completed");
                }

                System.Diagnostics.Debug.WriteLine($"WebView2 Debug: Initialization completed successfully");
            }
            catch (ArgumentException argEx)
            {
                string errorMsg = $"WebView2 ArgumentException at step '{step}':\n" +
                                 $"Message: {argEx.Message}\n" +
                                 $"Parameter: {argEx.ParamName}\n" +
                                 $"Stack Trace: {argEx.StackTrace}";

                System.Diagnostics.Debug.WriteLine($"WebView2 Error: {errorMsg}");
                MessageBox.Show(errorMsg);
            }
            catch (System.Exception ex)
            {
                string errorMsg = $"WebView2 Exception at step '{step}':\n" +
                                 $"Type: {ex.GetType().Name}\n" +
                                 $"Message: {ex.Message}\n" +
                                 $"Stack Trace: {ex.StackTrace}";

                System.Diagnostics.Debug.WriteLine($"WebView2 Error: {errorMsg}");
                MessageBox.Show(errorMsg);
            }
        }


        private void WebView_MessageReceived(object sender, CoreWebView2WebMessageReceivedEventArgs e)
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
                            System.Diagnostics.Debug.WriteLine("WebView2 Message Debug: Processing test action");
                            SendResponseToJS("success", "🎉 Perfect! Communication working!");
                            break;

                        case "duplicateEmail":
                            System.Diagnostics.Debug.WriteLine("WebView2 Message Debug: Processing duplicateEmail action");
                            HandleDuplicateEmailSafe(jObj["data"]);
                            break;

                        case "detectPlaceholder":
                            System.Diagnostics.Debug.WriteLine("WebView2 Message Debug: Processing detectPlaceholder action");
                            HandleDetectPlaceholderSafe();
                            break;

                        default:
                            System.Diagnostics.Debug.WriteLine($"WebView2 Message Debug: Unknown action: {actionName}");
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

        // Simplified thread-safe helper methods
        private void RunOnUIThread(System.Action action)
        {
            try
            {
                if (this.InvokeRequired)
                {
                    this.Invoke(action);
                }
                else
                {
                    action();
                }
            }
            catch (ObjectDisposedException)
            {
                // Control was disposed, ignore
            }
            catch (InvalidOperationException)
            {
                // Handle can be invalid, ignore
            }
        }

        private Task InvokeAsync(Func<Task> asyncAction)
        {
            var tcs = new TaskCompletionSource<bool>();

            if (this.InvokeRequired)
            {
                this.BeginInvoke(new System.Action(() =>
                {
                    ExecuteAsyncAction(asyncAction, tcs);
                }));
            }
            else
            {
                ExecuteAsyncAction(asyncAction, tcs);
            }

            return tcs.Task;
        }

        private void ExecuteAsyncAction(Func<Task> asyncAction, TaskCompletionSource<bool> tcs)
        {
            try
            {
                var task = asyncAction();
                task.ContinueWith(t =>
                {
                    if (t.IsFaulted)
                        tcs.SetException(t.Exception);
                    else
                        tcs.SetResult(true);
                });
            }
            catch (System.Exception ex)
            {
                tcs.SetException(ex);
            }
        }

        private void SendResponseToJS(string type, string message, object data = null)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"WebView2 Response Debug: Preparing response - Type: {type}, Message: {message}");

                var response = new { type, message, data };
                string jsonResponse = JsonConvert.SerializeObject(response);

                System.Diagnostics.Debug.WriteLine($"WebView2 Response Debug: JSON serialized - Length: {jsonResponse.Length}");

                RunOnUIThread(() =>
                {
                    try
                    {
                        if (webView?.CoreWebView2 != null && !webView.IsDisposed && webView.IsHandleCreated)
                        {
                            System.Diagnostics.Debug.WriteLine($"WebView2 Response Debug: Posting message to WebView");
                            webView.CoreWebView2.PostWebMessageAsString(jsonResponse);
                            System.Diagnostics.Debug.WriteLine($"WebView2 Response Debug: Message posted successfully");
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"WebView2 Response Debug: WebView not ready - disposed: {webView?.IsDisposed}, handle: {webView?.IsHandleCreated}");
                        }
                    }
                    catch (System.Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"WebView2 Response Error: {ex.Message}");
                    }
                });
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"WebView2 Response Error in SendResponseToJS: {ex.Message}");
            }
        }
        private void SendProgressToJS(int current, int total, string message)
        {
            try
            {
                var response = new { type = "progress", current, total, message };
                string jsonResponse = JsonConvert.SerializeObject(response);

                RunOnUIThread(() =>
                {
                    try
                    {
                        if (webView?.CoreWebView2 != null && !webView.IsDisposed && webView.IsHandleCreated)
                        {
                            webView.CoreWebView2.PostWebMessageAsString(jsonResponse);
                        }
                    }
                    catch (System.Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error sending progress to JS: {ex.Message}");
                    }
                });
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in SendProgressToJS: {ex.Message}");
            }
        }

        // Safe wrapper methods for background operations
        private void HandleDuplicateEmailSafe(Newtonsoft.Json.Linq.JToken requestData)
        {
            Task.Run(async () =>
            {
                try
                {
                    await HandleDuplicateEmail(requestData);
                }
                catch (System.Exception ex)
                {
                    SendResponseToJS("error", $"Duplication failed: {ex.Message}");
                }
            });
        }

        private void HandleDetectPlaceholderSafe()
        {
            Task.Run(async () =>
            {
                try
                {
                    await HandleDetectPlaceholder();
                }
                catch (System.Exception ex)
                {
                    SendResponseToJS("error", $"Error detecting placeholder: {ex.Message}");
                }
            });
        }

        // Business logic methods
        private async Task HandleDuplicateEmail(Newtonsoft.Json.Linq.JToken requestData)
        {
            if (outlookApp == null)
            {
                SendResponseToJS("error", "Outlook application not available. Please restart the add-in.");
                return;
            }

            string placeholder = requestData["placeholder"]?.ToString() ?? "";
            var recipients = requestData["recipients"]?.ToObject<List<Recipient>>() ?? new List<Recipient>();
            bool autoSend = requestData["autoSend"]?.ToObject<bool>() ?? false;
            bool forceWithoutPlaceholder = requestData["forceWithoutPlaceholder"]?.ToObject<bool>() ?? false;

            // Access Outlook safely to get the source email
            var selectedItem = GetSelectedOutlookItem();
            if (selectedItem == null)
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

            // Check if placeholder exists in the email content (unless user already confirmed)
            if (!forceWithoutPlaceholder && !string.IsNullOrEmpty(placeholder))
            {
                string emailContent = $"{sourceMailItem.Subject ?? ""} {sourceMailItem.Body ?? ""} {sourceMailItem.HTMLBody ?? ""}";

                if (!emailContent.Contains(placeholder))
                {
                    // Send a warning back to JavaScript for user confirmation
                    SendResponseToJS("placeholderWarning",
                        $"⚠️ Warning: The placeholder '{placeholder}' was not found in the selected email.\n\n" +
                        "This means the emails will be identical without personalization.\n\n" +
                        "Do you want to continue anyway?",
                        new
                        {
                            placeholder = placeholder,
                            recipients = recipients,
                            autoSend = autoSend
                        });

                    // Release COM ref and return - wait for user confirmation
                    if (sourceMailItem != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceMailItem);
                    return;
                }
            }

            // Continue with normal processing
            var emailData = ExtractEmailData(sourceMailItem);

            string actionText = autoSend ? "sending" : "creating drafts for";
            SendResponseToJS("info", $"Starting {actionText} {recipients.Count} recipients...");

            int successCount = 0;

            for (int i = 0; i < recipients.Count; i++)
            {
                try
                {
                    string progressAction = autoSend ? "Sending" : "Creating";
                    SendProgressToJS(i + 1, recipients.Count,
                        $"{progressAction} email {i + 1}/{recipients.Count} for {recipients[i].Name}...");

                    await CreateDuplicateEmail(emailData, placeholder, recipients[i], autoSend);
                    successCount++;
                    await Task.Delay(autoSend ? 500 : 100);
                }
                catch (System.Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to create duplicate for {recipients[i].Email}: {ex.Message}");
                }
            }

            string resultAction = autoSend ? "sent" : "created";
            string resultLocation = autoSend ? "Check Sent Items" : "Check Drafts";

            if (successCount == recipients.Count)
            {
                SendResponseToJS("success", $"✅ Successfully {resultAction} {successCount} emails! {resultLocation}.");
            }
            else
            {
                SendResponseToJS("error", $"⚠️ {resultAction.Substring(0, 1).ToUpper() + resultAction.Substring(1)} {successCount} out of {recipients.Count}. Some failed.");
            }

            CleanupTempFiles(emailData.Attachments);

            // Release COM ref
            if (sourceMailItem != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceMailItem);
        }

        private Selection GetSelectedOutlookItem()
        {
            try
            {
                return outlookApp.ActiveExplorer().Selection;
            }
            catch
            {
                return null;
            }
        }

        // Add these methods to your TaskPaneForm class

        private void RefreshWebViewVisuals()
        {
            try
            {
                // Method 1: Force a layout refresh
                this.SuspendLayout();
                webView.Visible = false;
                webView.Visible = true;
                this.ResumeLayout(true);

                // Method 2: Force WebView2 to recalculate its bounds
                if (webView.CoreWebView2 != null)
                {
                    // Trigger a bounds recalculation by temporarily changing size
                    var originalBounds = webView.Bounds;
                    webView.Bounds = new System.Drawing.Rectangle(
                        originalBounds.X,
                        originalBounds.Y,
                        originalBounds.Width - 1,
                        originalBounds.Height - 1
                    );

                    // Restore original bounds
                    webView.Bounds = originalBounds;

                    // Force a repaint
                    webView.Invalidate();
                    webView.Update();
                }

                System.Diagnostics.Debug.WriteLine("WebView2 visual refresh completed");
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in RefreshWebViewVisuals: {ex.Message}");
            }
        }

        // Also add this to handle DPI changes specifically

        // Handle parent changed events

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

        private async Task CreateDuplicateEmail(EmailData sourceData, string placeholder, Recipient recipient, bool autoSend = false)
        {
            await Task.Run(() =>
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

                    if (autoSend)
                    {
                        // Send the email immediately
                        newMail.Send();
                        System.Diagnostics.Debug.WriteLine($"Email sent to {recipient.Email}");
                    }
                    else
                    {
                        // Save as draft
                        newMail.Save();
                        System.Diagnostics.Debug.WriteLine($"Email saved as draft for {recipient.Email}");
                    }
                }
                catch (System.Exception ex)
                {
                    newMail?.Close(OlInspectorClose.olDiscard);
                    string action = autoSend ? "send" : "create draft";

                    // Log the error but don't throw - let the process continue
                    System.Diagnostics.Debug.WriteLine($"Failed to {action} email for {recipient.Email}: {ex.Message}");

                    // Instead of throwing, we'll let the calling method handle the failure count
                    // The HandleDuplicateEmail method already has try/catch around this call
                }
                finally
                {
                    if (newMail != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(newMail);
                    }
                }
            });
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
            if (outlookApp == null)
            {
                SendResponseToJS("error", "Outlook application not available.");
                return;
            }

            // Access Outlook safely
            var selectedItem = GetSelectedOutlookItem();
            if (selectedItem == null || selectedItem.Count == 0)
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
                try
                {
                    webView?.Dispose();
                }
                catch { /* ignore disposal errors */ }
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