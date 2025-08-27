using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Multi_Send
{
    public class OutlookService
    {
        private readonly TaskPaneForm taskPane;

        public OutlookService(TaskPaneForm taskPane)
        {
            this.taskPane = taskPane;
        }

        public string GetEmailSubject()
        {
            var source = GetActiveMailItem();
            if (source == null) return null;

            try
            {
                return source.Subject ?? "(No Subject)";
            }
            catch (System.Exception ex)
            {
                return "(Unknown Subject)";
            }
        }

        public int GetAttachmentCount()
        {
            var source = GetActiveMailItem();
            if (source == null) return 0;

            try
            {
                var emailData = ExtractEmailData(source);
                return emailData?.Attachments?.Count ?? 0;
            }
            catch (System.Exception ex)
            {
                return 0;
            }
        }

        public string GetEmailBody()
        {
            var source = GetActiveMailItem();
            if (source == null) return null;

            try
            {
                // Try HTML body first, then plain text body
                var htmlBody = source.HTMLBody ?? "";
                var textBody = source.Body ?? "";
                
                // Return the one that has content, preferring HTML
                return !string.IsNullOrWhiteSpace(htmlBody) ? htmlBody : textBody;
            }
            catch (System.Exception ex)
            {
                return null;
            }
        }

        public async Task<string> DetectPlaceholder()
        {
            var mail = GetActiveMailItem();
            if (mail == null) return null;
            
            string text = $"{mail.Subject} {mail.Body}";
            var patterns = new[] { @"\{\{[^}]+\}\}", @"\[[^\]]+\]", @"\$[A-Za-z_][A-Za-z0-9_]*" };
            
            foreach (var p in patterns)
            {
                var m = Regex.Matches(text, p, RegexOptions.IgnoreCase);
                if (m.Count > 0) return m[0].Value;
            }
            return null;
        }

        public async Task<DuplicateEmailResult> DuplicateEmails(string placeholder, List<Recipient> recipients, bool autoSend, bool force = false)
        {
            var source = GetActiveMailItem();
            
            if (source == null) 
                return new DuplicateEmailResult { Success = false, Message = "No email selected." };

            EmailData emailData = null;
            try
            {
                bool isSent = source.Sent;
                
                if (isSent || (!taskPane.IsInspectorMode && source.Recipients.Count > 0))
                {
                    return new DuplicateEmailResult { Success = false, Message = "🚫 Only works on unsent drafts you're composing." };
                }

                string content = $"{source.Subject} {source.Body} {source.HTMLBody}";
                if (!force && !string.IsNullOrEmpty(placeholder) &&
                    content.IndexOf(placeholder, StringComparison.OrdinalIgnoreCase) == -1)
                {
                    return new DuplicateEmailResult { 
                        Success = false, 
                        Message = $"⚠️ Placeholder '{placeholder}' not found.",
                        RequiresConfirmation = true,
                        Data = new { placeholder, recipients, autoSend }
                    };
                }

                emailData = ExtractEmailData(source);
                
                int success = 0;
                for (int i = 0; i < recipients.Count; i++)
                {
                    try
                    {
                        // Progress callback could be added here
                        await CreateDuplicateEmail(emailData, placeholder, recipients[i], autoSend);
                        success++;
                        await Task.Delay(autoSend ? 500 : 100);
                    }
                    catch (System.Exception ex)
                    {
                        // Silently handle errors for individual email creation
                    }
                }
                await Task.Delay(2000); // Give Outlook time to finish with files
                
                return new DuplicateEmailResult { 
                    Success = success == recipients.Count, 
                    Message = $"✅ {success}/{recipients.Count} {(autoSend ? "sent" : "created")}.",
                    AttachmentCount = emailData?.Attachments?.Count ?? 0
                };
            }
            catch (System.Exception ex)
            {
                return new DuplicateEmailResult { Success = false, Message = "Operation failed. Please try again." };
            }
            finally
            {
                // ALWAYS cleanup, even on exceptions
                if (emailData?.Attachments != null)
                {
                    CleanupTempFiles(emailData.Attachments);
                }
            }
        }

        private MailItem GetActiveMailItem()
        {
            try
            {
                if (taskPane.IsInspectorMode && taskPane.Inspector != null) 
                {
                    var currentItem = taskPane.Inspector.CurrentItem as MailItem;
                    return currentItem;
                }
                
                var explorer = taskPane.OutlookApp.ActiveExplorer();
                if (explorer == null) return null;
                
                var sel = explorer.Selection;
                return (sel != null && sel.Count > 0) ? sel[1] as MailItem : null;
            }
            catch (System.Exception ex)
            {
                return null;
            }
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
                            string fullPath = a.FileName ?? "attachment";
                            string originalFileName = Path.GetFileName(fullPath);

                            string secureDir = Path.Combine(Path.GetTempPath(), "Multi-Send", Environment.UserName, Process.GetCurrentProcess().Id.ToString());
                            Directory.CreateDirectory(secureDir);

                            // Use the ACTUAL filename instead of random GUID
                            string tmp = Path.Combine(secureDir, originalFileName);

                            // If file already exists, add a number suffix
                            int counter = 1;
                            while (File.Exists(tmp))
                            {
                                string nameWithoutExt = Path.GetFileNameWithoutExtension(originalFileName);
                                string extension = Path.GetExtension(originalFileName);
                                tmp = Path.Combine(secureDir, $"{nameWithoutExt}_{counter}{extension}");
                                counter++;
                            }

                            a.SaveAsFile(tmp);

                            data.Attachments.Add(new AttachmentData
                            {
                                FileName = originalFileName,
                                TempFilePath = tmp,
                                Type = a.Type
                            });

                            System.Diagnostics.Debug.WriteLine($"Stored filename: '{originalFileName}', Temp path: '{tmp}'");
                        }
                        catch (System.Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Attachment extraction error: {ex.Message}");
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

        private async Task CreateDuplicateEmail(EmailData src, string ph, Recipient r, bool autoSend)
        {
            MailItem m = null;
            try
            {
                System.Diagnostics.Debug.WriteLine($"=== Creating email for {r.Email} ===");
                System.Diagnostics.Debug.WriteLine($"Source has {src.Attachments?.Count ?? 0} attachments");

                m = taskPane.OutlookApp.CreateItem(OlItemType.olMailItem) as MailItem;
                if (m == null) throw new System.Exception("Failed to create MailItem");

                m.Subject = ReplacePlaceholder(src.Subject, ph, r.Name);
                m.Body = ReplacePlaceholder(src.Body, ph, r.Name);
                m.HTMLBody = ReplacePlaceholder(src.HTMLBody, ph, r.Name);
                m.Importance = src.Importance;
                m.Sensitivity = src.Sensitivity;

                if (string.IsNullOrWhiteSpace(r.Email) || !r.Email.Contains("@"))
                {
                    throw new ArgumentException("Invalid email address");
                }

                m.Recipients.Add(r.Email);
                m.Recipients.ResolveAll();

                // Add attachments with detailed debugging
                // In CreateDuplicateEmail method, replace the attachment adding section:
                // In CreateDuplicateEmail method:
                foreach (var att in src.Attachments)
                {
                    if (File.Exists(att.TempFilePath))
                    {
                        try
                        {
                            // Add attachment and explicitly set the display name
                            var attachment = m.Attachments.Add(att.TempFilePath, att.Type, 1);

                            // Explicitly set the display name after adding
                            attachment.DisplayName = att.FileName;

                            System.Diagnostics.Debug.WriteLine($"✅ Added: {att.FileName}, DisplayName set to: {attachment.DisplayName}");
                        }
                        catch (System.Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"❌ Failed to attach {att.FileName}: {ex.Message}");
                        }
                    }
                }

                System.Diagnostics.Debug.WriteLine($"Final email has {m.Attachments.Count} attachments");

                if (autoSend)
                {
                    m.Send();
                    System.Diagnostics.Debug.WriteLine("Email sent");
                }
                else
                {
                    m.Save();
                    System.Diagnostics.Debug.WriteLine("Email saved as draft");
                }

                await Task.Delay(autoSend ? 100 : 50);
            }
            finally
            {
                if (m != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(m);
            }
        }

        private string ReplacePlaceholder(string text, string placeholder, string repl) =>
            string.IsNullOrEmpty(text) || string.IsNullOrEmpty(placeholder)
            ? text
            : Regex.Replace(text, Regex.Escape(placeholder), repl, RegexOptions.IgnoreCase);

        private void CleanupTempFiles(List<AttachmentData> attachments)
        {
            foreach (var a in attachments)
            {
                try
                {
                    if (File.Exists(a.TempFilePath))
                    {
                        try
                        {
                            var random = new Random();
                            var buffer = new byte[1024];
                            using (var fs = File.OpenWrite(a.TempFilePath))
                            {
                                long fileSize = fs.Length;
                                fs.Position = 0;
                                for (long i = 0; i < fileSize; i += buffer.Length)
                                {
                                    random.NextBytes(buffer);
                                    int bytesToWrite = (int)Math.Min(buffer.Length, fileSize - i);
                                    fs.Write(buffer, 0, bytesToWrite);
                                }
                                fs.Flush();
                            }
                        }
                        catch { /* Overwrite failed, still try to delete */ }

                        File.Delete(a.TempFilePath);
                    }
                }
                catch { /* Ignore cleanup errors */ }
            }

            try
            {
                string secureDir = Path.Combine(Path.GetTempPath(), "Multi-Send", Environment.UserName, Process.GetCurrentProcess().Id.ToString());
                if (Directory.Exists(secureDir) && !Directory.EnumerateFileSystemEntries(secureDir).Any())
                {
                    Directory.Delete(secureDir);
                }
            }
            catch { /* Ignore directory cleanup errors */ }
        }
    }
}
