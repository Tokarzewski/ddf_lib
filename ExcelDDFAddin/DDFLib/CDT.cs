using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace DDFLib
{
    /// <summary>
    /// Represents a CDT (Component Data Table) file from a DDF archive.
    /// CDT files contain tabular data with custom '#' separators.
    /// </summary>
    public class CDT
    {
        // CDT file format constants (from Python schema.py lines 8-11)
        private const string SEPARATOR_HEADER = " #";  // Used for IDs and column headers
        private const string SEPARATOR_DATA = "  #";   // Used for data rows
        private const string PREFIX = "#";

        public List<int> Ids { get; set; }
        public DataTable Data { get; set; }

        public CDT()
        {
            Ids = new List<int>();
            Data = new DataTable();
        }

        public CDT(List<int> ids, DataTable data)
        {
            Ids = ids ?? new List<int>();
            Data = data ?? new DataTable();
        }

        /// <summary>
        /// Parse a line by splitting on separator and stripping prefix from first element.
        /// Ported from Python schema.py lines 20-24
        /// </summary>
        private static List<string> ParseLine(string line, string separator)
        {
            var parts = line.Split(new[] { separator }, StringSplitOptions.None).ToList();

            // Strip PREFIX from first element
            if (parts.Count > 0 && parts[0].StartsWith(PREFIX))
            {
                parts[0] = parts[0].Substring(PREFIX.Length);
            }

            return parts;
        }

        /// <summary>
        /// Read and parse a CDT file.
        /// Ported from Python schema.py lines 26-49
        /// </summary>
        public static CDT Read(string cdtFilePath)
        {
            try
            {
                if (!File.Exists(cdtFilePath))
                    return null;

                var lines = File.ReadAllLines(cdtFilePath);

                if (lines.Length < 2)
                    return null;

                // Parse IDs from line 1
                var idStrings = ParseLine(lines[0], SEPARATOR_HEADER);
                var ids = idStrings.Select(s => int.TryParse(s, out int id) ? id : 0).ToList();

                // Parse column names from line 2
                var columns = ParseLine(lines[1], SEPARATOR_HEADER);

                // Parse data rows (line 3+)
                var dataTable = new DataTable();

                // Add columns to DataTable
                foreach (var col in columns)
                {
                    dataTable.Columns.Add(col);
                }

                // Add rows
                for (int i = 2; i < lines.Length; i++)
                {
                    var rowData = ParseLine(lines[i], SEPARATOR_DATA);

                    // Ensure row has enough values (pad with empty strings if needed)
                    while (rowData.Count < columns.Count)
                        rowData.Add(string.Empty);

                    dataTable.Rows.Add(rowData.ToArray());
                }

                return new CDT(ids, dataTable);
            }
            catch (FileNotFoundException)
            {
                return null;
            }
            catch (UnauthorizedAccessException ex)
            {
                Console.WriteLine($"Permission denied: {cdtFilePath} - {ex.Message}");
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error parsing {cdtFilePath}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Save CDT data to a file.
        /// Ported from Python schema.py lines 51-67
        /// </summary>
        public void Save(string cdtFilePath)
        {
            using (var writer = new StreamWriter(cdtFilePath))
            {
                // Write IDs (using header separator)
                var idsLine = string.Join(SEPARATOR_HEADER, Ids.Select(id => id.ToString()));
                writer.WriteLine($"{PREFIX}{idsLine}");

                // Write columns (using header separator)
                var colsLine = string.Join(SEPARATOR_HEADER, Data.Columns.Cast<DataColumn>().Select(c => c.ColumnName));
                writer.WriteLine($"{PREFIX}{colsLine}");

                // Write rows (using data separator)
                foreach (DataRow row in Data.Rows)
                {
                    var rowData = row.ItemArray.Select(item => item?.ToString() ?? string.Empty);
                    var rowLine = string.Join(SEPARATOR_DATA, rowData);
                    writer.WriteLine($"{PREFIX}{rowLine}");
                }
            }
        }
    }
}
