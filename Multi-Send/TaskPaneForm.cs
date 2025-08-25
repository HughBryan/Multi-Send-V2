using Microsoft.Office.Interop.Outlook;
using System;
using System.Windows.Forms;

namespace Multi_Send
{
    public partial class TaskPaneForm : UserControl
    {
        private Microsoft.Office.Interop.Outlook.Application outlookApp;
        private Microsoft.Office.Interop.Outlook.Inspector inspector;
        private bool isInspectorMode;
        private EmailDuplicatorControl wpfControl;

        public TaskPaneForm()
        {
            InitializeWpfContent();
            InitializeOutlookApp();
            this.isInspectorMode = false;
        }

        public TaskPaneForm(Microsoft.Office.Interop.Outlook.Inspector inspector)
        {
            InitializeWpfContent();
            InitializeOutlookApp();
            this.inspector = inspector;
            this.isInspectorMode = true;
        }

        private void InitializeWpfContent()
        {
            var host = new System.Windows.Forms.Integration.ElementHost();
            host.Dock = DockStyle.Fill;
            
            wpfControl = new EmailDuplicatorControl();
            wpfControl.SetOutlookService(new OutlookService(this));
            host.Child = wpfControl;
            
            this.Controls.Add(host);
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

        // Internal access for OutlookService
        internal Microsoft.Office.Interop.Outlook.Application OutlookApp => outlookApp;
        internal Microsoft.Office.Interop.Outlook.Inspector Inspector => inspector;
        internal bool IsInspectorMode => isInspectorMode;

        protected override void Dispose(bool disposing)
        {
            if (disposing) wpfControl = null;
            base.Dispose(disposing);
        }
    }
}