using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace ExcelDDFAddin
{
    /// <summary>
    /// Ribbon UI handler for DDF Tools.
    /// </summary>
    [ComVisible(true)]
    public class DDFRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public DDFRibbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("ExcelDDFAddin.DDFRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        /// <summary>
        /// Handler for Open DDF button click.
        /// </summary>
        public void OnOpenDDFClick(Office.IRibbonControl control)
        {
            try
            {
                DDFManager.ShowOpenDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error: {ex.Message}",
                    "DDF Tools Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        /// <summary>
        /// Handler for Save as DDF button click.
        /// </summary>
        public void OnSaveDDFClick(Office.IRibbonControl control)
        {
            try
            {
                DDFManager.ShowSaveDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Error: {ex.Message}",
                    "DDF Tools Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
            }
        }

        /// <summary>
        /// Handler for About button click.
        /// </summary>
        public void OnAboutClick(Office.IRibbonControl control)
        {
            var version = Assembly.GetExecutingAssembly().GetName().Version;

            MessageBox.Show(
                $"DDF Tools for Excel\n\n" +
                $"Version: {version}\n\n" +
                $"This add-in enables Excel to open and save DesignBuilder DDF files.\n\n" +
                $"Each CDT table in the DDF file is loaded as a separate worksheet.\n" +
                $"You can edit the data in Excel and save back to DDF format.\n\n" +
                $"© 2026 DDF Tools",
                "About DDF Tools",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();

            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }

            return null;
        }

        #endregion
    }
}
