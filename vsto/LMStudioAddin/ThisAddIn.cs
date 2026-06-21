using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Word;
using Word = Microsoft.Office.Interop.Word;

namespace LMStudioAddin
{
    public partial class ThisAddIn
    {
        private CustomTaskPane _taskPane;
        private TaskPaneControl _taskPaneControl;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // Task pane is created lazily when the ribbon button is clicked.
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            _taskPane = null;
            _taskPaneControl?.Dispose();
        }

        /// <summary>
        /// Called by the ribbon button to toggle the AI task pane.
        /// </summary>
        public void ToggleTaskPane()
        {
            if (_taskPane == null)
            {
                _taskPaneControl = new TaskPaneControl();
                _taskPane = this.CustomTaskPanes.Add(_taskPaneControl, "LMStudio AI Assistant");
                _taskPane.Width = 340;
                _taskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
                _taskPane.Visible = true;
            }
            else
            {
                _taskPane.Visible = !_taskPane.Visible;
            }
        }

        /// <summary>
        /// Returns the currently selected text in the active Word document.
        /// </summary>
        public string GetSelectedText()
        {
            try
            {
                return Application.Selection?.Text ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Replaces the current selection with <paramref name="text"/>.
        /// </summary>
        public void InsertText(string text)
        {
            try
            {
                if (Application.Selection != null)
                    Application.Selection.TypeText(text);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not insert text: {ex.Message}", "LMStudio Add-in",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
