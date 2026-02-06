using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using DDFLib;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelDDFAddin
{
    /// <summary>
    /// Core functionality for opening and saving DDF files in Excel.
    /// Each CDT table from the DDF becomes a separate worksheet.
    /// </summary>
    public static class DDFManager
    {
        /// <summary>
        /// Open a DDF file and load all CDT tables into a new Excel workbook.
        /// </summary>
        public static void OpenDDFFile(string ddfPath)
        {
            try
            {
                // Show progress
                Globals.ThisAddIn.Application.StatusBar = "Loading DDF file...";
                Globals.ThisAddIn.Application.ScreenUpdating = false;

                // Read DDF file using DDFLib
                var ddf = DDF.Read(ddfPath);

                // Get available CDT tables
                var availableAttributes = ddf.GetAvailableAttributes();

                if (availableAttributes.Count == 0)
                {
                    MessageBox.Show(
                        "No data found in DDF file.",
                        "DDF Import",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    return;
                }

                // Create new workbook
                var app = Globals.ThisAddIn.Application;
                var workbook = app.Workbooks.Add();

                // Delete default sheets except one
                while (workbook.Worksheets.Count > 1)
                {
                    ((Excel.Worksheet)workbook.Worksheets[workbook.Worksheets.Count]).Delete();
                }

                bool firstSheet = true;

                // Create a worksheet for each CDT table
                foreach (var attrName in availableAttributes)
                {
                    Globals.ThisAddIn.Application.StatusBar = $"Loading {attrName}...";

                    var cdt = GetCDTProperty(ddf, attrName);
                    if (cdt == null)
                        continue;

                    Excel.Worksheet worksheet;

                    if (firstSheet)
                    {
                        // Use existing first sheet
                        worksheet = (Excel.Worksheet)workbook.Worksheets[1];
                        firstSheet = false;
                    }
                    else
                    {
                        // Add new worksheet
                        worksheet = (Excel.Worksheet)workbook.Worksheets.Add(
                            After: workbook.Worksheets[workbook.Worksheets.Count]
                        );
                    }

                    // Set worksheet name
                    worksheet.Name = attrName;

                    // Write IDs to row 1
                    WriteIdsToWorksheet(cdt.Ids, worksheet);

                    // Write column headers to row 2
                    ExcelHelpers.WriteColumnHeaders(cdt.Data, worksheet, row: 2);

                    // Write data starting from row 3
                    ExcelHelpers.WriteDataTableToWorksheet(cdt.Data, worksheet, startRow: 3);

                    // Apply formatting
                    ExcelHelpers.FormatWorksheet(worksheet);
                }

                // Store DDF file path in workbook custom properties
                SetWorkbookMetadata(workbook, ddfPath);

                // Activate first worksheet
                ((Excel.Worksheet)workbook.Worksheets[1]).Activate();

                MessageBox.Show(
                    $"Successfully loaded {availableAttributes.Count} tables from DDF file.",
                    "DDF Import Complete",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error opening DDF file: {ex.Message}",
                    "DDF Import Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            finally
            {
                Globals.ThisAddIn.Application.StatusBar = false;
                Globals.ThisAddIn.Application.ScreenUpdating = true;
            }
        }

        /// <summary>
        /// Save the current Excel workbook as a DDF file.
        /// Each worksheet becomes a CDT table in the DDF.
        /// </summary>
        public static void SaveAsDDFFile(Excel.Workbook workbook, string ddfPath)
        {
            try
            {
                Globals.ThisAddIn.Application.StatusBar = "Saving as DDF file...";
                Globals.ThisAddIn.Application.ScreenUpdating = false;

                var ddf = new DDF();

                // Process each worksheet
                foreach (Excel.Worksheet worksheet in workbook.Worksheets)
                {
                    try
                    {
                        Globals.ThisAddIn.Application.StatusBar = $"Processing {worksheet.Name}...";

                        // Read IDs from row 1
                        var ids = ExcelHelpers.ReadIdsFromWorksheet(worksheet, row: 1).ToList();

                        // Read data (headers in row 2, data from row 3)
                        var dataTable = ExcelHelpers.ReadWorksheetToDataTable(worksheet, headerRow: 2);

                        // Create CDT object
                        var cdt = new CDT(ids, dataTable);

                        // Set the property on DDF object
                        SetCDTProperty(ddf, worksheet.Name, cdt);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing worksheet {worksheet.Name}: {ex.Message}");
                    }
                }

                // Save DDF file
                ddf.Save(ddfPath);

                // Update metadata
                SetWorkbookMetadata(workbook, ddfPath);

                MessageBox.Show(
                    $"Successfully saved {ddf.GetAvailableAttributes().Count} tables to DDF file.",
                    "DDF Export Complete",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error saving DDF file: {ex.Message}",
                    "DDF Export Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
            finally
            {
                Globals.ThisAddIn.Application.StatusBar = false;
                Globals.ThisAddIn.Application.ScreenUpdating = true;
            }
        }

        /// <summary>
        /// Show file picker and open selected DDF file.
        /// </summary>
        public static void ShowOpenDialog()
        {
            var openDialog = new OpenFileDialog
            {
                Filter = "DesignBuilder DDF Files (*.ddf)|*.ddf|All Files (*.*)|*.*",
                Title = "Open DDF File",
                CheckFileExists = true,
                Multiselect = false
            };

            if (openDialog.ShowDialog() == DialogResult.OK)
            {
                OpenDDFFile(openDialog.FileName);
            }
        }

        /// <summary>
        /// Show file picker and save current workbook as DDF file.
        /// </summary>
        public static void ShowSaveDialog()
        {
            var workbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            if (workbook == null)
            {
                MessageBox.Show(
                    "No active workbook to save.",
                    "DDF Export",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                return;
            }

            var saveDialog = new SaveFileDialog
            {
                Filter = "DesignBuilder DDF Files (*.ddf)|*.ddf|All Files (*.*)|*.*",
                Title = "Save as DDF File",
                OverwritePrompt = true,
                DefaultExt = "ddf"
            };

            // Use existing DDF path if available
            string existingPath = GetWorkbookMetadata(workbook);
            if (!string.IsNullOrEmpty(existingPath))
            {
                saveDialog.FileName = existingPath;
            }

            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                SaveAsDDFFile(workbook, saveDialog.FileName);
            }
        }

        #region Helper Methods

        /// <summary>
        /// Get a CDT property from DDF object by name using reflection.
        /// </summary>
        private static CDT GetCDTProperty(DDF ddf, string propertyName)
        {
            var prop = typeof(DDF).GetProperty(propertyName);
            return prop?.GetValue(ddf) as CDT;
        }

        /// <summary>
        /// Set a CDT property on DDF object by name using reflection.
        /// </summary>
        private static void SetCDTProperty(DDF ddf, string propertyName, CDT value)
        {
            var prop = typeof(DDF).GetProperty(propertyName);
            prop?.SetValue(ddf, value);
        }

        /// <summary>
        /// Write IDs to the first row of a worksheet.
        /// </summary>
        private static void WriteIdsToWorksheet(List<int> ids, Excel.Worksheet worksheet)
        {
            if (ids == null || ids.Count == 0)
                return;

            for (int i = 0; i < ids.Count; i++)
            {
                worksheet.Cells[1, i + 1] = ids[i];
            }

            // Format IDs row
            Excel.Range idsRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, ids.Count]];
            idsRange.Font.Bold = true;
            idsRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
        }

        /// <summary>
        /// Store DDF file path in workbook custom properties.
        /// </summary>
        private static void SetWorkbookMetadata(Excel.Workbook workbook, string ddfPath)
        {
            try
            {
                var customProps = workbook.CustomDocumentProperties;

                // Remove existing property if it exists
                try
                {
                    customProps["DDFFilePath"].Delete();
                }
                catch { }

                // Add new property
                customProps.Add(
                    "DDFFilePath",
                    false,
                    Microsoft.Office.Core.MsoDocProperties.msoPropertyTypeString,
                    ddfPath
                );
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error setting workbook metadata: {ex.Message}");
            }
        }

        /// <summary>
        /// Get DDF file path from workbook custom properties.
        /// </summary>
        private static string GetWorkbookMetadata(Excel.Workbook workbook)
        {
            try
            {
                var customProps = workbook.CustomDocumentProperties;
                return customProps["DDFFilePath"].Value?.ToString();
            }
            catch
            {
                return null;
            }
        }

        #endregion
    }
}
