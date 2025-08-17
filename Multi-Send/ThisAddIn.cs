using System;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Multi_Send
{
    public partial class ThisAddIn
    {
        private TaskPaneForm taskPane;
        private Microsoft.Office.Tools.CustomTaskPane customTaskPane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                // Create your task pane form
                taskPane = new TaskPaneForm();

                // Create a custom task pane in Outlook
                customTaskPane = this.CustomTaskPanes.Add(taskPane, "Email Duplicator");
                customTaskPane.Visible = true;
                customTaskPane.Width = 400;
                customTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;

                MessageBox.Show("Email Duplicator Add-in loaded successfully!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading add-in: {ex.Message}");
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                // Clean up
                if (customTaskPane != null)
                {
                    customTaskPane.Dispose();
                }
                if (taskPane != null)
                {
                    taskPane.Dispose();
                }
            }
            catch (Exception ex)
            {
                // Log error but don't show message box during shutdown
                System.Diagnostics.Debug.WriteLine($"Error during shutdown: {ex.Message}");
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}