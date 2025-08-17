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
                CreateTaskPane();
                MessageBox.Show("Email Duplicator Add-in loaded successfully!");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading add-in: {ex.Message}");
            }
        }

        private void CreateTaskPane()
        {
            if (customTaskPane == null)
            {
                taskPane = new TaskPaneForm();
                customTaskPane = this.CustomTaskPanes.Add(taskPane, "Email Duplicator");
                customTaskPane.Visible = true;
                customTaskPane.Width = 400;
                customTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            }
        }

        public void ToggleTaskPane()
        {
            try
            {
                // If disposed or null → rebuild
                if (customTaskPane == null ||
                    customTaskPane.Control == null ||
                    customTaskPane.Control.IsDisposed)
                {
                    RecreateTaskPane();
                }
                else
                {
                    customTaskPane.Visible = !customTaskPane.Visible;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error toggling task pane: {ex.Message}");
                RecreateTaskPane();
            }
        }

        private void RecreateTaskPane()
        {
            try
            {
                // Remove from Outlook's collection if it exists
                if (customTaskPane != null)
                {
                    try { this.CustomTaskPanes.Remove(customTaskPane); } catch { }
                    customTaskPane.Dispose();
                    customTaskPane = null;
                }

                if (taskPane != null)
                {
                    taskPane.Dispose();
                    taskPane = null;
                }

                // Build fresh one
                CreateTaskPane();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error recreating task pane: {ex.Message}");
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                if (customTaskPane != null)
                {
                    try { this.CustomTaskPanes.Remove(customTaskPane); } catch { }
                    customTaskPane.Dispose();
                    customTaskPane = null;
                }

                if (taskPane != null)
                {
                    taskPane.Dispose();
                    taskPane = null;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Shutdown error: {ex.Message}");
            }
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
