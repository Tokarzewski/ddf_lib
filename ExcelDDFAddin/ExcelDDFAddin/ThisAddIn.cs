using System;
using System.IO;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelDDFAddin
{
    /// <summary>
    /// Main add-in class for DDF Tools.
    /// Handles add-in lifecycle and Excel events.
    /// </summary>
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Hook into Excel events
            this.Application.WorkbookOpen += OnWorkbookOpen;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Cleanup
            this.Application.WorkbookOpen -= OnWorkbookOpen;
        }

        /// <summary>
        /// Event handler for when a workbook is opened.
        /// Auto-detects .ddf files and prompts to import.
        /// </summary>
        private void OnWorkbookOpen(Excel.Workbook workbook)
        {
            try
            {
                // Check if the opened file is a .ddf file
                if (!string.IsNullOrEmpty(workbook.FullName) &&
                    Path.GetExtension(workbook.FullName).Equals(".ddf", StringComparison.OrdinalIgnoreCase))
                {
                    // Close the empty workbook that Excel created
                    string ddfPath = workbook.FullName;
                    workbook.Close(SaveChanges: false);

                    // Import the DDF file
                    DDFManager.OpenDDFFile(ddfPath);
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Error handling DDF file: {ex.Message}",
                    "DDF Tools Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error
                );
            }
        }

        /// <summary>
        /// Required by VSTO to provide custom ribbon UI.
        /// </summary>
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new DDFRibbon();
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
