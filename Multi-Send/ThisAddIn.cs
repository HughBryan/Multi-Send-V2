using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace Multi_Send
{
    public partial class ThisAddIn
    {
        private TaskPaneForm taskPane;
        public Microsoft.Office.Tools.CustomTaskPane customTaskPane;

        // Track task panes for inspector windows (compose/reply/forward)
        private Dictionary<Outlook.Inspector, Microsoft.Office.Tools.CustomTaskPane> inspectorTaskPanes;

        // Shared visibility state - this keeps all task panes in sync
        private bool isTaskPaneVisible = false;

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
                SetupDisplayChangeDetection();
                SystemEvents.DisplaySettingsChanged += (_, __) => RecreateAllTaskPanesPreservingVisibility();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading add-in: {ex.Message}");
            }
        }

        private void SetupDisplayChangeDetection()
        {
            try
            {
                // Hook into Windows display change events
                SystemEvents.DisplaySettingsChanged += SystemEvents_DisplaySettingsChanged;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error setting up display change detection: {ex.Message}");
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
        private void SystemEvents_DisplaySettingsChanged(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("Display settings changed - fixing task panes");

                // Delay the fix slightly to let Windows finish the display change
                Timer delayTimer = new Timer();
                delayTimer.Interval = 500; // 500ms delay
                delayTimer.Tick += (s, ev) =>
                {
                    delayTimer.Stop();
                    delayTimer.Dispose();
                    ForceRedockAllTaskPanes();
                };
                delayTimer.Start();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error handling display settings change: {ex.Message}");
            }
        }

        private void ForceRedockAllTaskPanes()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("Force redocking all task panes");

                // Force redock main explorer task pane
                if (customTaskPane != null && !customTaskPane.Control.IsDisposed)
                {
                    ForceRedockSingleTaskPane(customTaskPane);
                }

                // Force redock all inspector task panes
                var inspectorsToRemove = new List<Outlook.Inspector>();

                foreach (var kvp in inspectorTaskPanes)
                {
                    try
                    {
                        var inspector = kvp.Key;
                        var taskPane = kvp.Value;

                        if (taskPane == null || taskPane.Control == null || taskPane.Control.IsDisposed)
                        {
                            inspectorsToRemove.Add(inspector);
                        }
                        else
                        {
                            ForceRedockSingleTaskPane(taskPane);
                        }
                    }
                    catch
                    {
                        inspectorsToRemove.Add(kvp.Key);
                    }
                }

                // Clean up invalid inspectors
                foreach (var inspector in inspectorsToRemove)
                {
                    inspectorTaskPanes.Remove(inspector);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error force redocking task panes: {ex.Message}");
            }
        }



        private void ForceRedockSingleTaskPane(Microsoft.Office.Tools.CustomTaskPane taskPane)
        {
            try
            {
                if (taskPane != null && !taskPane.Control.IsDisposed)
                {
                    bool wasVisible = taskPane.Visible;

                    System.Diagnostics.Debug.WriteLine($"Force redocking task pane - was visible: {wasVisible}");

                    // Temporarily hide, reset properties, then restore visibility
                    taskPane.Visible = false;

                    // Force undock and redock
                    taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionFloating;
                    taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;

                    // Reset width
                    taskPane.Width = 400;

                    // Restore visibility if it was visible before
                    if (wasVisible)
                    {
                        taskPane.Visible = true;
                    }

                    System.Diagnostics.Debug.WriteLine("Task pane redocked successfully");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error force redocking single task pane: {ex.Message}");
            }
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
            try
            {
                // Don't create duplicate task panes
                if (inspectorTaskPanes.ContainsKey(inspector))
                    return;

                // Create a new task pane form for this inspector
                var inspectorTaskPaneForm = new TaskPaneForm();
                var inspectorCustomTaskPane = this.CustomTaskPanes.Add(inspectorTaskPaneForm, "Multi-Send", inspector);

                // Use the shared visibility state and ensure proper positioning
                inspectorCustomTaskPane.Visible = isTaskPaneVisible;
                inspectorCustomTaskPane.Width = 400;
                inspectorCustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;

                // Store the task pane reference
                inspectorTaskPanes[inspector] = inspectorCustomTaskPane;

                System.Diagnostics.Debug.WriteLine($"Created inspector task pane for new window");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating inspector task pane: {ex.Message}");
            }
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
                customTaskPane.Visible = isTaskPaneVisible; // Use shared state
                customTaskPane.Width = 400;
                customTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            }
        }

        public void ToggleTaskPane()
        {
            try
            {
                // Toggle the shared visibility state
                isTaskPaneVisible = !isTaskPaneVisible;

                System.Diagnostics.Debug.WriteLine($"Toggling task pane visibility to: {isTaskPaneVisible}");

                // Apply the new state to ALL task panes
                SetAllTaskPanesVisibility(isTaskPaneVisible);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error toggling task pane: {ex.Message}");
            }
        }

        private void SetAllTaskPanesVisibility(bool visible)
        {
            try
            {
                // Set main explorer task pane visibility
                if (customTaskPane == null || customTaskPane.Control == null || customTaskPane.Control.IsDisposed)
                {
                    RecreateTaskPane();
                }

                if (customTaskPane != null)
                {
                    customTaskPane.Visible = visible;
                    // Also ensure proper docking when showing
                    if (visible)
                    {
                        ForceRedockSingleTaskPane(customTaskPane);
                    }
                }

                // Set all inspector task panes visibility
                var inspectorsToRemove = new List<Outlook.Inspector>();

                foreach (var kvp in inspectorTaskPanes)
                {
                    try
                    {
                        var inspector = kvp.Key;
                        var taskPane = kvp.Value;

                        // Check if inspector is still valid
                        if (taskPane == null || taskPane.Control == null || taskPane.Control.IsDisposed)
                        {
                            inspectorsToRemove.Add(inspector);
                        }
                        else
                        {
                            taskPane.Visible = visible;
                            // Also ensure proper docking when showing
                            if (visible)
                            {
                                ForceRedockSingleTaskPane(taskPane);
                            }
                        }
                    }
                    catch
                    {
                        inspectorsToRemove.Add(kvp.Key);
                    }
                }

                // Clean up invalid inspectors
                foreach (var inspector in inspectorsToRemove)
                {
                    inspectorTaskPanes.Remove(inspector);
                }

                // Create task panes for any new inspectors that might be open
                EnsureAllInspectorsHaveTaskPanes();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error setting task pane visibility: {ex.Message}");
            }
        }

        private void EnsureAllInspectorsHaveTaskPanes()
        {
            try
            {
                // Check all open inspectors and create task panes if needed
                foreach (Outlook.Inspector inspector in this.Application.Inspectors)
                {
                    if (inspector.CurrentItem is Outlook.MailItem && !inspectorTaskPanes.ContainsKey(inspector))
                    {
                        CreateInspectorTaskPane(inspector);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error ensuring inspectors have task panes: {ex.Message}");
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
                // Unhook display change events
                SystemEvents.DisplaySettingsChanged -= SystemEvents_DisplaySettingsChanged;

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