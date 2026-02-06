using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;

namespace DDFLib
{
    /// <summary>
    /// Represents a DDF (DesignBuilder Data File) - a ZIP archive containing multiple CDT files.
    /// Ported from Python schema.py lines 70-145
    /// </summary>
    public class DDF
    {
        // All 17 CDT properties from Python schema.py lines 72-88
        public CDT Glazing { get; set; }
        public CDT InternalBlinds { get; set; }
        public CDT Panes { get; set; }
        public CDT WindowGas { get; set; }
        public CDT Constructions { get; set; }
        public CDT Materials { get; set; }
        public CDT ActivityTemplates { get; set; }
        public CDT ConstructionTemplates { get; set; }
        public CDT DHWTemplates { get; set; }
        public CDT FacadeTemplates { get; set; }
        public CDT GlazingTemplates { get; set; }
        public CDT HourlyWeather { get; set; }
        public CDT LightingTemplates { get; set; }
        public CDT LocalShading { get; set; }
        public CDT LocationTemplates { get; set; }
        public CDT SBEMHVACSystems { get; set; }
        public CDT Schedules { get; set; }

        /// <summary>
        /// Return list of attribute names that contain data (not null).
        /// Ported from Python schema.py lines 90-94
        /// </summary>
        public List<string> GetAvailableAttributes()
        {
            var availableAttributes = new List<string>();

            foreach (var prop in GetType().GetProperties())
            {
                if (prop.PropertyType == typeof(CDT))
                {
                    var value = prop.GetValue(this) as CDT;
                    if (value != null)
                    {
                        availableAttributes.Add(prop.Name);
                    }
                }
            }

            return availableAttributes;
        }

        /// <summary>
        /// Check if a specific attribute has data (is not null).
        /// Ported from Python schema.py lines 96-98
        /// </summary>
        public bool HasData(string attributeName)
        {
            var prop = GetType().GetProperty(attributeName);
            if (prop == null || prop.PropertyType != typeof(CDT))
                return false;

            return prop.GetValue(this) != null;
        }

        /// <summary>
        /// Read and parse a DDF file.
        /// Ported from Python schema.py lines 100-126
        /// </summary>
        public static DDF Read(string ddfFilePath)
        {
            var ddf = new DDF();
            var ddfPath = new FileInfo(ddfFilePath);

            if (!ddfPath.Exists)
            {
                Console.WriteLine($"DDF file not found: {ddfFilePath}");
                return ddf;
            }

            // Create temporary directory
            var tempDir = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            Directory.CreateDirectory(tempDir);

            try
            {
                // Extract ZIP archive
                ZipFile.ExtractToDirectory(ddfFilePath, tempDir);

                // Get all defined properties of type CDT
                var definedFields = typeof(DDF)
                    .GetProperties()
                    .Where(p => p.PropertyType == typeof(CDT))
                    .Select(p => p.Name)
                    .ToHashSet();

                // Find all .cdt files in extracted directory
                var availableCdts = Directory.GetFiles(tempDir, "*.cdt")
                    .Select(f => Path.GetFileNameWithoutExtension(f))
                    .ToHashSet();

                // Warn about unknown CDT files
                var unknown = availableCdts.Except(definedFields);
                if (unknown.Any())
                {
                    Console.WriteLine($"Unknown CDT files in {ddfPath.Name}: {string.Join(", ", unknown)}");
                }

                // Load each CDT file
                foreach (var fieldName in definedFields)
                {
                    var cdtFile = Path.Combine(tempDir, $"{fieldName}.cdt");
                    var cdt = CDT.Read(cdtFile);

                    // Set property value
                    var prop = typeof(DDF).GetProperty(fieldName);
                    prop?.SetValue(ddf, cdt);
                }

                return ddf;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing DDF file {ddfPath}: {ex.Message}");
                return ddf;
            }
            finally
            {
                // Clean up temporary directory
                try
                {
                    if (Directory.Exists(tempDir))
                        Directory.Delete(tempDir, true);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error cleaning up temp directory: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Save DDF data to a new file.
        /// Ported from Python schema.py lines 128-145
        /// </summary>
        public void Save(string ddfFilePath)
        {
            var ddfPath = new FileInfo(ddfFilePath);

            // Create temporary directory
            var tempDir = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            Directory.CreateDirectory(tempDir);

            try
            {
                // Save each CDT that has data
                foreach (var fieldName in GetAvailableAttributes())
                {
                    var prop = GetType().GetProperty(fieldName);
                    var cdt = prop?.GetValue(this) as CDT;

                    if (cdt != null)
                    {
                        var cdtFile = Path.Combine(tempDir, $"{fieldName}.cdt");
                        cdt.Save(cdtFile);
                    }
                }

                // Delete existing DDF file if it exists
                if (ddfPath.Exists)
                    ddfPath.Delete();

                // Create ZIP file
                ZipFile.CreateFromDirectory(tempDir, ddfFilePath);
            }
            finally
            {
                // Clean up temporary directory
                try
                {
                    if (Directory.Exists(tempDir))
                        Directory.Delete(tempDir, true);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error cleaning up temp directory: {ex.Message}");
                }
            }
        }
    }
}
