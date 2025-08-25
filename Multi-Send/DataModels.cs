using Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;

namespace Multi_Send
{
    public class Recipient 
    { 
        public string Email { get; set; } = "";
        public string Name { get; set; } = "";
    }

    public class EmailData
    {
        public string Subject { get; set; } = "";
        public string Body { get; set; } = "";
        public string HTMLBody { get; set; } = "";
        public OlImportance Importance { get; set; } = OlImportance.olImportanceNormal;
        public OlSensitivity Sensitivity { get; set; } = OlSensitivity.olNormal;
        public List<AttachmentData> Attachments { get; set; } = new List<AttachmentData>();
    }

    public class AttachmentData
    {
        public string FileName { get; set; } = "";
        public string TempFilePath { get; set; } = "";
        public OlAttachmentType Type { get; set; }
    }

    public class DuplicateEmailResult
    {
        public bool Success { get; set; }
        public string Message { get; set; } = "";
        public bool RequiresConfirmation { get; set; }
        public object Data { get; set; }
        public int AttachmentCount { get; set; }
    }
}
