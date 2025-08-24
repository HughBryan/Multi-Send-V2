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
            InitializeComponent();
            InitializeOutlookApp();
            this.isInspectorMode = false;
            this.Load += TaskPaneForm_Load;
        }

        public TaskPaneForm(Microsoft.Office.Interop.Outlook.Inspector inspector)
        {
            InitializeComponent();
            InitializeOutlookApp();
            this.inspector = inspector;
            this.isInspectorMode = true;
            this.Load += TaskPaneForm_Load;
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
            }
            catch (System.Exception ex)
            {
                this.outlookApp = null;
            }
        }

        private void TaskPaneForm_Load(object sender, EventArgs e)
        {
            InitializeWebViewAsync();
        }

        private async void InitializeWebViewAsync()
        {
            try
            {
                // Create unique user data folder
                string userDataFolder = Path.Combine(Path.GetTempPath(), $"Multi-Send_WebView2_{System.Environment.UserName}_{Guid.NewGuid():N}");
                
                Directory.CreateDirectory(userDataFolder);
                
                var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder);
                
                await webView.EnsureCoreWebView2Async(env);

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

                // Load embedded content
                string html = GetEmbeddedResource("Multi_Send.TaskPaneUI.index.html");
                string css = GetEmbeddedResource("Multi_Send.TaskPaneUI.taskpane.css");
                string js = GetEmbeddedResource("Multi_Send.TaskPaneUI.taskpane.js");

                // Inline CSS & JS for security
                html = Regex.Replace(html, @"style-src\s+'self'", "style-src 'self' 'unsafe-inline'", RegexOptions.IgnoreCase);
                html = Regex.Replace(html, @"<link\s+rel=[""']stylesheet[""']\s+href=[""']taskpane\.css[""']\s*/>", $"<style>{css}</style>", RegexOptions.IgnoreCase);
                html = Regex.Replace(html, @"<script\s+src=[""']taskpane\.js[""']\s*></script>", $"<script>{js}</script>", RegexOptions.IgnoreCase);

                webView.CoreWebView2.NavigateToString(html);
            }
            catch (System.Exception ex)
            {
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
                var msg = e.TryGetWebMessageAsString();
                
                var jobj = JObject.Parse(msg);
                string action = jobj?["action"]?.ToString();
                
                switch (action)
                {
                    case "test":
                        SendToJS("success", "🎉 Communication working!");
                        break;

                    case "getEmailSubject":
                        GetEmailSubject();
                        break;

                    case "getAttachmentCount":
                        RunSafe(() => HandleGetAttachmentCount(jobj["data"]));
                        break;
                    case "duplicateEmail":
                        RunSafe(() => HandleDuplicateEmail(jobj["data"]));
                        break;
                    case "detectPlaceholder":
                        RunSafe(HandleDetectPlaceholder);
                        break;
                    default:
                        SendToJS("error", "Unknown action");
                        break;
                }
            }
            catch (System.Exception ex)
            {
                SendToJS("error", $"Message error: {ex.Message}");
            }
        }

        private async void RunSafe(Func<Task> func)
        {
            try 
            {
                await func(); 
            }
            catch (System.Exception ex) 
            {
                SendToJS("error", ex.Message); 
            }
        }



        private void SendToJS(string type, string message = "", object data = null)
        {
            try
            {
                if (webView == null)
                {
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
                            }
                        }
                        catch (System.Exception ex)
                        {
                            // Silently handle errors
                        }
                    }));
                }
                else
                {
                    // We're already on UI thread
                    if (webView?.CoreWebView2 != null)
                    {
                        webView.CoreWebView2.PostWebMessageAsString(json);
                    }
                }
            }
            catch (System.Exception ex)
            {
                // Silently handle errors
            }
        }

        private MailItem GetActiveMailItem()
        {
            try
            {
                if (isInspectorMode && inspector != null) 
                {
                    var currentItem = inspector.CurrentItem as MailItem;
                    return currentItem;
                }
                
                var explorer = outlookApp.ActiveExplorer();
                if (explorer == null) return null;
                
                var sel = explorer.Selection;
                return (sel != null && sel.Count > 0) ? sel[1] as MailItem : null;
            }
            catch (System.Exception ex)
            {
                return null;
            }
        }

        private async Task HandleDuplicateEmail(JToken req)
        {
            var source = GetActiveMailItem();
            
            if (source == null) { SendToJS("error", "No email selected."); return; }
            
            try
            {
                bool isSent = source.Sent;
                
                if (isSent || (!isInspectorMode && source.Recipients.Count > 0))
                {
                    SendToJS("error", "🚫 Only works on unsent drafts you're composing."); return;
                }
            }
            catch (System.Exception ex)
            {
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
                    // Silently handle errors for individual email creation
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
                            // Silently handle errors for individual attachment extraction
                        }
                    }
                }
                
                return data;
            }
            catch (System.Exception ex)
            {
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
            }
            catch (System.Exception ex)
            {
                SendToJS("emailSubject", "", new { subject = "(Unknown Subject)" });
            }
        }

        private async Task CreateDuplicateEmail(EmailData src, string ph, Recipient r, bool autoSend)
        {
            MailItem m = null;
            try
            {
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
            var source = GetActiveMailItem();
            if (source == null) { SendToJS("error", "No email selected."); return; }

            try
            {
                var emailData = ExtractEmailData(source);
                int attachmentCount = emailData?.Attachments?.Count ?? 0;

                SendToJS("attachmentCount", "", new { count = attachmentCount });
            }
            catch (System.Exception ex)
            {
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