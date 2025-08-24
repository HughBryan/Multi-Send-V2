using Microsoft.Office.Interop.Outlook;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
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
        private Microsoft.Office.Interop.Outlook.Inspector inspector;
        private bool isInspectorMode;

        public TaskPaneForm()
        {
            System.Diagnostics.Debug.WriteLine("DEBUG: TaskPaneForm() - Default constructor called (Explorer mode)");
            InitializeComponent();
            InitializeOutlookApp();
            this.isInspectorMode = false;
            this.Load += TaskPaneForm_Load;
        }

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
            this.webView = new WebView2 { Dock = DockStyle.Fill, Name = "webView" };
            this.Controls.Add(this.webView);
            this.Size = new System.Drawing.Size(500, 600);
            this.BackColor = System.Drawing.Color.White;
        }

        private void InitializeOutlookApp()
        {
            try
            {
                this.outlookApp = Globals.ThisAddIn.Application;
                System.Diagnostics.Debug.WriteLine("DEBUG: Outlook app initialized successfully");
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Warning: Could not initialize Outlook Application: {ex.Message}");
                this.outlookApp = null;
            }
        }

        private void TaskPaneForm_Load(object sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine("DEBUG: TaskPaneForm_Load called");
            InitializeWebViewAsync();
        }

        private async void InitializeWebViewAsync()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("WebView2 Debug: Starting initialization");
                
                // Create unique user data folder
                string userDataFolder = Path.Combine(Path.GetTempPath(), $"Multi-Send_WebView2_{System.Environment.UserName}_{Guid.NewGuid():N}");
                System.Diagnostics.Debug.WriteLine($"WebView2 Debug: Creating user data folder: {userDataFolder}");
                
                Directory.CreateDirectory(userDataFolder);
                System.Diagnostics.Debug.WriteLine("WebView2 Debug: User data folder created");
                
                var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
                System.Diagnostics.Debug.WriteLine("WebView2 Debug: Environment created");
                
                await webView.EnsureCoreWebView2Async(env);
                System.Diagnostics.Debug.WriteLine("WebView2 Debug: CoreWebView2 ensured");

                // Security settings
                webView.CoreWebView2.NavigationStarting += (s, e) => {
                    var uri = e.Uri ?? "";
                    if (!uri.StartsWith("file://") && uri != "about:blank" && !uri.StartsWith("data:"))
                        e.Cancel = true;
                };
                webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
                webView.CoreWebView2.Settings.AreDevToolsEnabled = false;

                // Attach message handler
                webView.CoreWebView2.WebMessageReceived += WebView_MessageReceived;
                System.Diagnostics.Debug.WriteLine("WebView2 Debug: Message handler attached");

                // Load embedded content
                string html = GetEmbeddedResource("Multi_Send.TaskPaneUI.index.html");
                string css = GetEmbeddedResource("Multi_Send.TaskPaneUI.taskpane.css");
                string js = GetEmbeddedResource("Multi_Send.TaskPaneUI.taskpane.js");

                // Inline CSS & JS for security
                html = Regex.Replace(html, @"style-src\s+'self'", "style-src 'self' 'unsafe-inline'", RegexOptions.IgnoreCase);
                html = Regex.Replace(html, @"<link\s+rel=[""']stylesheet[""']\s+href=[""']taskpane\.css[""']\s*/>", $"<style>{css}</style>", RegexOptions.IgnoreCase);
                html = Regex.Replace(html, @"<script\s+src=[""']taskpane\.js[""']\s*></script>", $"<script>{js}</script>", RegexOptions.IgnoreCase);

                webView.CoreWebView2.NavigateToString(html);
                System.Diagnostics.Debug.WriteLine("WebView2 Debug: Content loaded successfully");
            }
            catch (System.Exception ex)
            {
                string errorMsg = $"WebView2 Error: {ex.Message}\nStack: {ex.StackTrace}";
                System.Diagnostics.Debug.WriteLine(errorMsg);
                MessageBox.Show($"WebView2 initialization failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GetEmbeddedResource(string resourceName)
        {
            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
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
                System.Diagnostics.Debug.WriteLine("DEBUG: WebView_MessageReceived started");
                
                var msg = e.TryGetWebMessageAsString();
                System.Diagnostics.Debug.WriteLine($"DEBUG: Received message: {msg}");
                
                var jobj = JObject.Parse(msg);
                string action = jobj?["action"]?.ToString();
                System.Diagnostics.Debug.WriteLine($"DEBUG: Action: {action}");
                
                switch (action)
                {
                    case "test":
                        System.Diagnostics.Debug.WriteLine("DEBUG: Processing test action");
                        SendToJS("success", "🎉 Communication working!");
                        break;

                    case "getEmailSubject":
                        GetEmailSubject();
                        break;

                    case "getAttachmentCount":
                        System.Diagnostics.Debug.WriteLine("DEBUG: Processing getAttachmentCount action");
                        RunSafe(() => HandleGetAttachmentCount(jobj["data"]));
                        break;
                    case "duplicateEmail":
                        System.Diagnostics.Debug.WriteLine("DEBUG: Processing duplicateEmail action");
                        RunSafe(() => HandleDuplicateEmail(jobj["data"]));
                        break;
                    case "detectPlaceholder":
                        System.Diagnostics.Debug.WriteLine("DEBUG: Processing detectPlaceholder action");
                        RunSafe(HandleDetectPlaceholder);
                        break;
                    default:
                        System.Diagnostics.Debug.WriteLine($"DEBUG: Unknown action: {action}");
                        SendToJS("error", "Unknown action");
                        break;
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"DEBUG: WebView_MessageReceived error: {ex.Message}");
                SendToJS("error", $"Message error: {ex.Message}");
            }
        }

        private async void RunSafe(Func<Task> func)
        {
            try 
            {
                System.Diagnostics.Debug.WriteLine("DEBUG: RunSafe started");
                await func(); 
                System.Diagnostics.Debug.WriteLine("DEBUG: RunSafe completed");
            }
            catch (System.Exception ex) 
            {
                System.Diagnostics.Debug.WriteLine($"DEBUG: RunSafe error: {ex.Message}");
                SendToJS("error", ex.Message); 
            }
        }



        private void SendToJS(string type, string message = "", object data = null)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"SendToJS: {type} - {message}");

                if (webView == null)
                {
                    System.Diagnostics.Debug.WriteLine("SendToJS: WebView is null");
                    return;
                }

                var response = new { type, message, data };
                string json = JsonConvert.SerializeObject(response, new JsonSerializerSettings
                {
                    ReferenceLoopHandling = ReferenceLoopHandling.Ignore,
                    NullValueHandling = NullValueHandling.Ignore
                });

                // FIXED: Move ALL WebView2 access inside Invoke
                if (InvokeRequired)
                {
                    this.Invoke(new System.Action(() => {
                        try
                        {
                            // Check CoreWebView2 INSIDE Invoke where it's safe
                            if (webView?.CoreWebView2 != null)
                            {
                                webView.CoreWebView2.PostWebMessageAsString(json);
                                System.Diagnostics.Debug.WriteLine("SendToJS: Message sent successfully");
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine("SendToJS: CoreWebView2 not ready");
                            }
                        }
                        catch (System.Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"SendToJS invoke error: {ex.Message}");
                        }
                    }));
                }
                else
                {
                    // We're already on UI thread
                    if (webView?.CoreWebView2 != null)
                    {
                        webView.CoreWebView2.PostWebMessageAsString(json);
                        System.Diagnostics.Debug.WriteLine("SendToJS: Message sent directly");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("SendToJS: CoreWebView2 not ready");
                    }
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"SendToJS general error: {ex.Message}");
            }
        }

        private MailItem GetActiveMailItem()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"DEBUG: GetActiveMailItem - isInspectorMode: {isInspectorMode}");
                
                if (isInspectorMode && inspector != null) 
                {
                    System.Diagnostics.Debug.WriteLine("DEBUG: Using inspector mode");
                    var currentItem = inspector.CurrentItem as MailItem;
                    System.Diagnostics.Debug.WriteLine($"DEBUG: Inspector CurrentItem: {currentItem != null}");
                    return currentItem;
                }
                
                System.Diagnostics.Debug.WriteLine("DEBUG: Using Explorer mode");
                var explorer = outlookApp.ActiveExplorer();
                if (explorer == null) return null;
                
                var sel = explorer.Selection;
                return (sel != null && sel.Count > 0) ? sel[1] as MailItem : null;
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"DEBUG: Error in GetActiveMailItem: {ex.Message}");
                return null;
            }
        }

        private async Task HandleDuplicateEmail(JToken req)
        {
            System.Diagnostics.Debug.WriteLine("DEBUG: HandleDuplicateEmail started");
            
            var source = GetActiveMailItem();
            System.Diagnostics.Debug.WriteLine($"DEBUG: GetActiveMailItem returned: {source != null}");
            
            if (source == null) { SendToJS("error", "No email selected."); return; }
            
            try
            {
                bool isSent = source.Sent;
                System.Diagnostics.Debug.WriteLine($"DEBUG: source.Sent = {isSent}");
                
                if (isSent || (!isInspectorMode && source.Recipients.Count > 0))
                {
                    SendToJS("error", "🚫 Only works on unsent drafts you're composing."); return;
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"DEBUG: Error checking email properties: {ex.Message}");
                SendToJS("error", $"Error accessing email: {ex.Message}");
                return;
            }

            string placeholder = req["placeholder"]?.ToString() ?? "";
            var recipients = req["recipients"]?.ToObject<List<Recipient>>() ?? new List<Recipient>();
            bool autoSend = req["autoSend"]?.ToObject<bool>() ?? false;
            bool force = req["forceWithoutPlaceholder"]?.ToObject<bool>() ?? false;

            string content = $"{source.Subject} {source.Body} {source.HTMLBody}";
            if (!force && !string.IsNullOrEmpty(placeholder) &&
                content.IndexOf(placeholder, StringComparison.OrdinalIgnoreCase) == -1)
            {
                SendToJS("placeholderWarning", $"⚠️ Placeholder '{placeholder}' not found.", new { placeholder, recipients, autoSend });
                return;
            }

            var emailData = ExtractEmailData(source);
            
            // Send attachment count
            try
            {
                int attachmentCount = emailData?.Attachments?.Count ?? 0;
                SendToJS("attachmentCount", "", new { count = attachmentCount });
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error sending attachment count: {ex.Message}");
                SendToJS("attachmentCount", "", new { count = 0 });
            }
            
            SendToJS("info", $"Starting {(autoSend ? "sending" : "creating drafts for")} {recipients.Count} recipients...");

            int success = 0;
            for (int i = 0; i < recipients.Count; i++)
            {
                try
                {
                    SendToJS("progress", $"{(autoSend ? "Sending" : "Creating")} email {i + 1}/{recipients.Count} for {recipients[i].Name}...", 
                        new { current = i + 1, total = recipients.Count });
                    await CreateDuplicateEmail(emailData, placeholder, recipients[i], autoSend);
                    success++;
                    await Task.Delay(autoSend ? 500 : 100);
                }
                catch (System.Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error creating email for {recipients[i].Email}: {ex.Message}");
                }
            }

            SendToJS(success == recipients.Count ? "success" : "error",
                $"✅ {success}/{recipients.Count} {(autoSend ? "sent" : "created")}.");
            CleanupTempFiles(emailData.Attachments);
        }

        private EmailData ExtractEmailData(MailItem source)
        {
            try
            {
                var data = new EmailData
                {
                    Subject = source?.Subject ?? "",
                    Body = source?.Body ?? "",
                    HTMLBody = source?.HTMLBody ?? "",
                    Importance = source?.Importance ?? OlImportance.olImportanceNormal,
                    Sensitivity = source?.Sensitivity ?? OlSensitivity.olNormal,
                    Attachments = new List<AttachmentData>()
                };

                if (source?.Attachments != null)
                {
                    foreach (Attachment a in source.Attachments)
                    {
                        try
                        {
                            string fileName = Path.GetFileName(a.FileName ?? "attachment");
                            string tmp = Path.Combine(Path.GetTempPath(), $"EmailDup_{Guid.NewGuid()}_{fileName}");
                            a.SaveAsFile(tmp);
                            data.Attachments.Add(new AttachmentData 
                            { 
                                FileName = a.FileName ?? fileName, 
                                TempFilePath = tmp, 
                                Type = a.Type 
                            });
                        }
                        catch (System.Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error extracting attachment: {ex.Message}");
                        }
                    }
                }
                
                return data;
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error extracting email data: {ex.Message}");
                return new EmailData
                {
                    Subject = "",
                    Body = "",
                    HTMLBody = "",
                    Importance = OlImportance.olImportanceNormal,
                    Sensitivity = OlSensitivity.olNormal,
                    Attachments = new List<AttachmentData>()
                };
            }
        }

        private void GetEmailSubject()
        {
            var source = GetActiveMailItem();
            if (source == null) { SendToJS("error", "No email selected."); return; }

            try
            {
                string subject = source.Subject ?? "(No Subject)";
                SendToJS("emailSubject", "", new { subject = subject });
                System.Diagnostics.Debug.WriteLine($"DEBUG: Sent email subject: {subject}");
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting email subject: {ex.Message}");
                SendToJS("emailSubject", "", new { subject = "(Unknown Subject)" });
            }
        }

        private async Task CreateDuplicateEmail(EmailData src, string ph, Recipient r, bool autoSend)
        {
            MailItem m = null;
            try
            {
                System.Diagnostics.Debug.WriteLine($"DEBUG: Creating email for {r.Email}");

                // FIXED: Remove Task.Run() - keep COM calls on UI thread
                m = outlookApp.CreateItem(OlItemType.olMailItem) as MailItem;
                if (m == null) throw new System.Exception("Failed to create MailItem");

                m.Subject = ReplacePlaceholder(src.Subject, ph, r.Name);
                m.Body = ReplacePlaceholder(src.Body, ph, r.Name);
                m.HTMLBody = ReplacePlaceholder(src.HTMLBody, ph, r.Name);
                m.Importance = src.Importance;
                m.Sensitivity = src.Sensitivity;

                m.Recipients.Add(r.Email);
                m.Recipients.ResolveAll();

                foreach (var att in src.Attachments)
                    if (File.Exists(att.TempFilePath))
                        m.Attachments.Add(att.TempFilePath, att.Type, 1, att.FileName);

                if (autoSend) m.Send(); else m.Save();

                // Add a small delay for UI responsiveness (but stay on UI thread)
                await Task.Delay(autoSend ? 100 : 50);
            }
            finally
            {
                if (m != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(m);
            }
        }

        private async Task HandleGetAttachmentCount(JToken req)
        {
            System.Diagnostics.Debug.WriteLine("DEBUG: HandleGetAttachmentCount started");

            var source = GetActiveMailItem();
            if (source == null) { SendToJS("error", "No email selected."); return; }

            try
            {
                var emailData = ExtractEmailData(source);
                int attachmentCount = emailData?.Attachments?.Count ?? 0;

                SendToJS("attachmentCount", "", new { count = attachmentCount });
                System.Diagnostics.Debug.WriteLine($"DEBUG: Sent attachment count: {attachmentCount}");
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting attachment count: {ex.Message}");
                SendToJS("attachmentCount", "", new { count = 0 });
            }
        }

        private string ReplacePlaceholder(string text, string placeholder, string repl) =>
            string.IsNullOrEmpty(text) || string.IsNullOrEmpty(placeholder)
            ? text
            : Regex.Replace(text, Regex.Escape(placeholder), repl, RegexOptions.IgnoreCase);

        private async Task HandleDetectPlaceholder()
        {
            var mail = GetActiveMailItem();
            if (mail == null) { SendToJS("error", "No email selected."); return; }
            string text = $"{mail.Subject} {mail.Body}";
            var patterns = new[] { @"\{\{[^}]+\}\}", @"\[[^\]]+\]", @"\$[A-Za-z_][A-Za-z0-9_]*" };
            foreach (var p in patterns)
            {
                var m = Regex.Matches(text, p, RegexOptions.IgnoreCase);
                if (m.Count > 0) { SendToJS("success", $"Detected: {m[0].Value}", new { placeholder = m[0].Value }); return; }
            }
            SendToJS("info", "No common placeholders found.");
        }

        private void CleanupTempFiles(List<AttachmentData> attachments)
        {
            foreach (var a in attachments)
                try { if (File.Exists(a.TempFilePath)) File.Delete(a.TempFilePath); } catch { }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) webView?.Dispose();
            base.Dispose(disposing);
        }
    }

       

    public class Recipient { public string Email { get; set; } public string Name { get; set; } }
    public class EmailData
    {
        public string Subject, Body, HTMLBody;
        public OlImportance Importance;
        public OlSensitivity Sensitivity;
        public List<AttachmentData> Attachments;
    }
    public class AttachmentData
    {
        public string FileName, TempFilePath;
        public OlAttachmentType Type;
    }
}