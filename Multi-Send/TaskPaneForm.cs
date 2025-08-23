using Microsoft.Office.Interop.Outlook;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Multi_Send
{
    public partial class TaskPaneForm : UserControl
    {
        private WebView2 webView;
        private Microsoft.Office.Interop.Outlook.Application outlookApp;
        private Microsoft.Office.Interop.Outlook.Inspector inspector; // NEW: Track inspector context
        private bool isInspectorMode; // NEW: Flag to know if we're in compose mode

        public TaskPaneForm()
        {
            System.Diagnostics.Debug.WriteLine("DEBUG: TaskPaneForm() - Default constructor called (Explorer mode)");
            InitializeComponent();
            InitializeOutlookApp();
            this.isInspectorMode = false; // Default to Explorer mode
            this.Load += TaskPaneForm_Load;
        }

        // NEW: Constructor for Inspector mode
        public TaskPaneForm(Microsoft.Office.Interop.Outlook.Inspector inspector)
        {
            System.Diagnostics.Debug.WriteLine("DEBUG: TaskPaneForm(Inspector) - Inspector constructor called");
            InitializeComponent();
            InitializeOutlookApp();
            this.inspector = inspector;
            this.isInspectorMode = true;
            this.Load += TaskPaneForm_Load;
            System.Diagnostics.Debug.WriteLine($"DEBUG: Inspector constructor completed - isInspectorMode: {this.isInspectorMode}");
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // UserControl properties
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = AutoScaleMode.Dpi; // was Font/None
            this.Size = new System.Drawing.Size(500, 600);
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

                // Harden WebView2 BEFORE any Navigate(...)
                webView.CoreWebView2.NavigationStarting += (s, e) =>
                {
                    var uri = e.Uri ?? "";
                    // Allow file:// URLs, about:blank, and data: URLs (for NavigateToString)
                    if (!uri.StartsWith("file://", StringComparison.OrdinalIgnoreCase) &&
                        uri != "about:blank" &&
                        !uri.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
                        e.Cancel = true;
                };
                webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
                webView.CoreWebView2.Settings.AreDevToolsEnabled = false; // set true if you need devtools in DEBUG

                System.Diagnostics.Debug.WriteLine($"WebView2 Debug: CoreWebView2 ensured successfully");

                step = "Attaching WebMessageReceived event";
                System.Diagnostics.Debug.WriteLine($"WebView2 Debug: {step}");
                webView.CoreWebView2.WebMessageReceived += WebView_MessageReceived;
                System.Diagnostics.Debug.WriteLine($"WebView2 Debug: Event attached successfully");

                step = "Loading embedded TaskPaneUI resources";
                
                try
                {
                    // Get the assembly and list all embedded resources
                    var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                    var allResources = assembly.GetManifestResourceNames();
                    
                    // Get embedded resources
                    string htmlContent = GetEmbeddedResourceContent(assembly, "Multi_Send.TaskPaneUI.index.html");
                    string cssContent = GetEmbeddedResourceContent(assembly, "Multi_Send.TaskPaneUI.taskpane.css");
                    string jsContent = GetEmbeddedResourceContent(assembly, "Multi_Send.TaskPaneUI.taskpane.js");
                    
                    // Fix Content Security Policy to allow inline styles
                    string cspFixed = System.Text.RegularExpressions.Regex.Replace(
                        htmlContent,
                        @"style-src\s+'self'",
                        "style-src 'self' 'unsafe-inline'",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                        
                    // Inject CSS and JS inline
                    string finalHtml = System.Text.RegularExpressions.Regex.Replace(
                        cspFixed,
                        @"<link\s+rel=[""']stylesheet[""']\s+href=[""']taskpane\.css[""']\s*/>",
                        $"<style>{cssContent}</style>",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                        
                    finalHtml = System.Text.RegularExpressions.Regex.Replace(
                        finalHtml,
                        @"<script\s+src=[""']taskpane\.js[""']\s*></script>",
                        $"<script>{jsContent}</script>",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                    
                    webView.CoreWebView2.NavigateToString(finalHtml);
                    System.Diagnostics.Debug.WriteLine("WebView2 Debug: Loaded embedded TaskPaneUI successfully");
                }
                catch (System.Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"WebView2 Debug: Error loading embedded resources: {ex.Message}");
                    
                    // Fallback minimal HTML
                    string fallbackHtml = @"
                        <!DOCTYPE html>
                        <html><body style='font-family: Segoe UI; padding:20px;'>
                            <h3>Multi-Send TaskPane</h3>
                            <p>Error loading embedded UI resources. Please rebuild the add-in.</p>
                            <p>Error: " + ex.Message + @"</p>
                        </body></html>";
                    
                    webView.CoreWebView2.NavigateToString(fallbackHtml);
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

        private string GetEmbeddedResourceContent(System.Reflection.Assembly assembly, string resourceName)
        {
            System.Diagnostics.Debug.WriteLine($"Looking for embedded resource: {resourceName}");

            using (var stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                    throw new System.Exception($"Embedded resource '{resourceName}' not found");

                using (var reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
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

            // DEBUG: Add logging to see what mode we're in
            System.Diagnostics.Debug.WriteLine($"DEBUG: HandleDuplicateEmail - isInspectorMode: {isInspectorMode}");
            System.Diagnostics.Debug.WriteLine($"DEBUG: inspector is null: {inspector == null}");

            string placeholder = requestData["placeholder"]?.ToString() ?? "";
            var recipients = requestData["recipients"]?.ToObject<List<Recipient>>() ?? new List<Recipient>();
            bool autoSend = requestData["autoSend"]?.ToObject<bool>() ?? false;
            bool forceWithoutPlaceholder = requestData["forceWithoutPlaceholder"]?.ToObject<bool>() ?? false;

            // Get the source email based on context
            MailItem sourceMailItem = null;
            
            if (isInspectorMode && inspector != null)
            {
                System.Diagnostics.Debug.WriteLine("DEBUG: Using Inspector mode - getting CurrentItem");
                // In compose mode - use the current item
                sourceMailItem = inspector.CurrentItem as MailItem;
                if (sourceMailItem == null)
                {
                    System.Diagnostics.Debug.WriteLine("DEBUG: Inspector CurrentItem is null or not a MailItem");
                    SendResponseToJS("error", "Current item is not an email.");
                    return;
                }
                System.Diagnostics.Debug.WriteLine($"DEBUG: Got Inspector MailItem - Subject: {sourceMailItem.Subject}");
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("DEBUG: Using Explorer mode - getting Selection");
                // In Explorer mode - use selected item
                var selectedItem = GetSelectedOutlookItem() as Selection;
                if (selectedItem == null || selectedItem.Count == 0)
                {
                    SendResponseToJS("error", "Please select an email to duplicate.");
                    return;
                }

                sourceMailItem = selectedItem[1] as MailItem;
                if (sourceMailItem == null)
                {
                    SendResponseToJS("error", "Selected item is not an email.");
                    return;
                }
                System.Diagnostics.Debug.WriteLine($"DEBUG: Got Explorer MailItem - Subject: {sourceMailItem.Subject}");
            }

            // SAFETY CHECK: Ensure we're working with an unsent email (compose mode)
            try
            {
                if (sourceMailItem.Sent)
                {
                    SendResponseToJS("error", "🚫 SAFETY: This email has already been sent. Multi-Send only works with drafts you're composing.");
                    if (!isInspectorMode && sourceMailItem != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceMailItem);
                    return;
                }
                
                // Additional safety for Explorer mode
                if (!isInspectorMode && sourceMailItem.Recipients.Count > 0)
                {
                    SendResponseToJS("error", "�� SAFETY: Please use Multi-Send from compose windows, not from received emails.");
                    if (sourceMailItem != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceMailItem);
                    return;
                }
            }
            catch (System.Exception ex)
            {
                SendResponseToJS("error", $"🚫 SAFETY: Unable to verify email safety. {ex.Message}");
                if (!isInspectorMode && sourceMailItem != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceMailItem);
                return;
            }

            System.Diagnostics.Debug.WriteLine($"DEBUG: SAFETY PASSED - Email is safe to duplicate (Sent: {sourceMailItem.Sent})");

            // Continue with existing placeholder detection logic...
            if (!forceWithoutPlaceholder && !string.IsNullOrEmpty(placeholder))
            {
                string emailContent = $"{sourceMailItem.Subject ?? ""} {sourceMailItem.Body ?? ""} {sourceMailItem.HTMLBody ?? ""}";
                System.Diagnostics.Debug.WriteLine($"DEBUG: Checking for placeholder '{placeholder}' in email content (length: {emailContent.Length})");

                // Improved placeholder detection with case-insensitive search (.NET Framework compatible)
                if (emailContent.IndexOf(placeholder, StringComparison.OrdinalIgnoreCase) == -1)
                {
                    System.Diagnostics.Debug.WriteLine("DEBUG: Placeholder NOT found - showing warning");
                    // Send a warning back to JavaScript for user confirmation
                    SendResponseToJS("placeholderWarning",
                        $"⚠️ Warning: The placeholder '{placeholder}' was not found in the current email.\n\n" +
                        "This means the emails will be identical without personalization.\n\n" +
                        "Do you want to continue anyway?",
                        new
                        {
                            placeholder = placeholder,
                            recipients = recipients,
                            autoSend = autoSend
                        });

                    // Release COM ref and return - wait for user confirmation
                    if (!isInspectorMode && sourceMailItem != null)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceMailItem);
                    return;
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("DEBUG: Placeholder FOUND - proceeding");
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

            // Release COM ref only if in Explorer mode
            if (!isInspectorMode && sourceMailItem != null)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(sourceMailItem);
        }

        private object GetSelectedOutlookItem()
        {
            try
            {
                if (isInspectorMode && inspector != null)
                {
                    // In Inspector mode - work with the current item being composed/edited
                    var mailItem = inspector.CurrentItem as MailItem;
                    if (mailItem != null)
                    {
                        // Create a fake Selection-like wrapper for consistency
                        return new InspectorItemWrapper { CurrentItem = mailItem };
                    }
                    return null;
                }
                else
                {
                    // In Explorer mode - work with selected items
                    return outlookApp.ActiveExplorer().Selection;
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting selected item: {ex.Message}");
                return null;
            }
        }

        // NEW: Wrapper class to make Inspector items work like Selection
        private class InspectorItemWrapper
        {
            public MailItem CurrentItem { get; set; }
            public int Count => CurrentItem != null ? 1 : 0;
            public object this[int index] => index == 1 && CurrentItem != null ? CurrentItem : null;
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
                    var safeName = Path.GetFileName(attachment.FileName); // strips any dirs
                    string tempPath = Path.Combine(Path.GetTempPath(), $"EmailDup_{Guid.NewGuid()}_{safeName}");
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

            MailItem mailItem = null;

            if (isInspectorMode && inspector != null)
            {
                // In compose mode - use current item
                mailItem = inspector.CurrentItem as MailItem;
                if (mailItem == null)
                {
                    SendResponseToJS("error", "Current item is not an email.");
                    return;
                }
            }
            else
            {
                // In Explorer mode - use selected item
                var selectedItem = GetSelectedOutlookItem() as Selection;
                if (selectedItem == null || selectedItem.Count == 0)
                {
                    SendResponseToJS("error", "Please select an email to detect placeholder from.");
                    return;
                }

                mailItem = selectedItem[1] as MailItem;
                if (mailItem == null)
                {
                    SendResponseToJS("error", "Selected item is not an email.");
                    return;
                }
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