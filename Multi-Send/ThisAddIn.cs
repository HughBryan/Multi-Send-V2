using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Win32;
using System.Threading.Tasks; // Added for async/await
using System.Text.RegularExpressions; // Added for Regex
using System.IO;
using System.Linq;

namespace Multi_Send
{
    public partial class ThisAddIn
    {
        private TaskPaneForm taskPane;
        public Microsoft.Office.Tools.CustomTaskPane customTaskPane;

        // Track task panes for inspector windows (compose/reply/forward)
        private Dictionary<Outlook.Inspector, Microsoft.Office.Tools.CustomTaskPane> inspectorTaskPanes;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                // Initialize the dictionary to track inspector task panes
                inspectorTaskPanes = new Dictionary<Outlook.Inspector, Microsoft.Office.Tools.CustomTaskPane>();

                // Create task pane for main explorer window
                CreateTaskPane();

                // Hook into inspector events to create task panes for compose windows
                this.Application.Inspectors.NewInspector += Inspectors_NewInspector;

                // Set up display change detection
                SystemEvents.DisplaySettingsChanged += (_, __) => RecreateAllTaskPanesPreservingVisibility();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading add-in: {ex.Message}");
            }
        }


        private void RecreateAllTaskPanesPreservingVisibility()
        {
            try
            {
                var wantVisible = customTaskPane?.Visible ?? false;

                // Explorer pane
                RecreateTaskPane();
                if (customTaskPane != null) customTaskPane.Visible = wantVisible;

                // Inspectors
                var openInspectors = new List<Outlook.Inspector>();
                foreach (Outlook.Inspector ins in this.Application.Inspectors)
                    openInspectors.Add(ins);

                // remove old
                foreach (var kv in inspectorTaskPanes)
                {
                    try { this.CustomTaskPanes.Remove(kv.Value); } catch { }
                    kv.Value.Dispose();
                }
                inspectorTaskPanes.Clear();

                // recreate for each open inspector
                foreach (var ins in openInspectors)
                    if (ins.CurrentItem is Outlook.MailItem)
                    {
                        CreateInspectorTaskPane(ins);
                        if (inspectorTaskPanes.TryGetValue(ins, out var tp))
                            tp.Visible = wantVisible;
                    }
            }
            catch { }
        }




        private void Inspectors_NewInspector(Outlook.Inspector Inspector)
        {
            try
            {
                // Check if this is a mail item (compose, reply, forward)
                if (Inspector.CurrentItem is Outlook.MailItem)
                {
                    // Create a task pane for this inspector window
                    CreateInspectorTaskPane(Inspector);

                    // Clean up when inspector closes
                    ((Outlook.InspectorEvents_10_Event)Inspector).Close += () => Inspector_Close(Inspector);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating inspector task pane: {ex.Message}");
            }
        }

        private void CreateInspectorTaskPane(Outlook.Inspector inspector)
        {
            if (inspectorTaskPanes.ContainsKey(inspector)) return;

            // Pass the inspector context to TaskPaneForm
            var inspectorTaskPaneForm = new TaskPaneForm(inspector);
            var inspectorCustomTaskPane = this.CustomTaskPanes.Add(inspectorTaskPaneForm, "Multi-Send", inspector);
            inspectorCustomTaskPane.Width = 500;
            inspectorCustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;

            inspectorTaskPanes[inspector] = inspectorCustomTaskPane;
        }

        private void Inspector_Close(Outlook.Inspector inspector)
        {
            try
            {
                // Clean up the task pane when inspector closes
                if (inspectorTaskPanes.ContainsKey(inspector))
                {
                    var taskPane = inspectorTaskPanes[inspector];
                    try
                    {
                        this.CustomTaskPanes.Remove(taskPane);
                        taskPane.Dispose();
                    }
                    catch { }

                    inspectorTaskPanes.Remove(inspector);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error cleaning up inspector task pane: {ex.Message}");
            }
        }

        private void CreateTaskPane()
        {
            if (customTaskPane == null)
            {
                taskPane = new TaskPaneForm();
                customTaskPane = this.CustomTaskPanes.Add(taskPane, "Multi-Send");
                customTaskPane.Visible = false;
                customTaskPane.Width = 400;
                customTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            }
        }

        public void ToggleTaskPane()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("DEBUG: ToggleTaskPane started");

                // Check if we're in an Inspector context first
                var activeInspector = this.Application.ActiveInspector();
                System.Diagnostics.Debug.WriteLine($"DEBUG: ActiveInspector: {activeInspector != null}");

                if (activeInspector != null && activeInspector.CurrentItem is Outlook.MailItem)
                {
                    System.Diagnostics.Debug.WriteLine("DEBUG: In Inspector mode - creating/showing inspector task pane");
                    // We're in a compose window - create/show inspector task pane
                    if (!inspectorTaskPanes.ContainsKey(activeInspector))
                    {
                        System.Diagnostics.Debug.WriteLine("DEBUG: Creating new inspector task pane");
                        CreateInspectorTaskPane(activeInspector);
                    }

                    if (inspectorTaskPanes.ContainsKey(activeInspector))
                    {
                        System.Diagnostics.Debug.WriteLine("DEBUG: Showing inspector task pane");
                        inspectorTaskPanes[activeInspector].Visible = true;
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("DEBUG: In Explorer mode - showing main task pane");
                    // We're in main window - show main task pane
                    RecreateTaskPane();
                    if (customTaskPane != null)
                    {
                        customTaskPane.Visible = true;
                        System.Diagnostics.Debug.WriteLine("DEBUG: Main task pane made visible");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine("DEBUG: ERROR - customTaskPane is null!");
                    }
                }
                System.Diagnostics.Debug.WriteLine("DEBUG: ToggleTaskPane completed");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"DEBUG: ToggleTaskPane error: {ex.Message}");
                MessageBox.Show($"Error opening Multi-Send: {ex.Message}");
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
                System.Diagnostics.Debug.WriteLine($"Error recreating task pane: {ex.Message}");
            }
        }

        // Add ribbon support
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                // Clean up inspector task panes
                foreach (var kvp in inspectorTaskPanes)
                {
                    try
                    {
                        this.CustomTaskPanes.Remove(kvp.Value);
                        kvp.Value.Dispose();
                    }
                    catch { }
                }
                inspectorTaskPanes.Clear();

                // Clean up main task pane
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

                // Clean up any remaining temp files
                CleanupAllTempFiles();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Shutdown error: {ex.Message}");
            }
        }

        private void CleanupAllTempFiles()
        {
            try
            {
                string tempRoot = Path.Combine(Path.GetTempPath(), "Multi-Send", Environment.UserName);
                if (Directory.Exists(tempRoot))
                {
                    Directory.Delete(tempRoot, true); // Recursive delete
                }
            }
            catch
            {
                // Ignore cleanup errors on shutdown
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