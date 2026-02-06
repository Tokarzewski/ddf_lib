using System;
using System.IO;
using System.Linq;
using DDFLib;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace DDFLib.Tests
{
    [TestClass]
    public class DDFTests
    {
        private string GetSamplePath(string filename)
        {
            // Navigate to samples folder (relative to test project)
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var projectRoot = Path.GetFullPath(Path.Combine(baseDir, @"..\..\..\.."));
            return Path.Combine(projectRoot, "samples", filename);
        }

        [TestMethod]
        public void TestDDFRead()
        {
            // Arrange
            var ddfPath = GetSamplePath("construction and materials.DDF");

            // Act
            var ddf = DDF.Read(ddfPath);

            // Assert
            Assert.IsNotNull(ddf, "DDF should not be null");
            var availableAttributes = ddf.GetAvailableAttributes();
            Assert.IsTrue(availableAttributes.Count > 0, "DDF should have at least one CDT table");

            // Should contain Materials and Constructions based on example
            Assert.IsTrue(ddf.HasData("Materials") || ddf.HasData("Constructions"),
                "DDF should contain Materials or Constructions table");
        }

        [TestMethod]
        public void TestDDFAvailableAttributes()
        {
            // Arrange
            var ddfPath = GetSamplePath("construction and materials.DDF");

            // Act
            var ddf = DDF.Read(ddfPath);
            var attributes = ddf.GetAvailableAttributes();

            // Assert
            Assert.IsNotNull(attributes);
            Assert.IsTrue(attributes.Count > 0);

            // Print available attributes for debugging
            Console.WriteLine($"Available attributes: {string.Join(", ", attributes)}");
        }

        [TestMethod]
        public void TestCDTDataTable()
        {
            // Arrange
            var ddfPath = GetSamplePath("construction and materials.DDF");

            // Act
            var ddf = DDF.Read(ddfPath);

            // Assert
            if (ddf.HasData("Materials"))
            {
                Assert.IsNotNull(ddf.Materials);
                Assert.IsNotNull(ddf.Materials.Data);
                Assert.IsNotNull(ddf.Materials.Ids);
                Assert.IsTrue(ddf.Materials.Data.Rows.Count > 0, "Materials table should have rows");
                Assert.IsTrue(ddf.Materials.Data.Columns.Count > 0, "Materials table should have columns");
                Assert.IsTrue(ddf.Materials.Ids.Count > 0, "Materials should have IDs");
            }
        }

        [TestMethod]
        public void TestDDFRoundTrip()
        {
            // Arrange
            var originalPath = GetSamplePath("construction and materials.DDF");
            var tempPath = Path.GetTempFileName();
            tempPath = Path.ChangeExtension(tempPath, ".ddf");

            try
            {
                // Act - Read original
                var originalDDF = DDF.Read(originalPath);
                var originalAttributes = originalDDF.GetAvailableAttributes();

                // Act - Save to temp file
                originalDDF.Save(tempPath);

                // Act - Read back from temp file
                var roundTripDDF = DDF.Read(tempPath);
                var roundTripAttributes = roundTripDDF.GetAvailableAttributes();

                // Assert - Same number of tables
                Assert.AreEqual(originalAttributes.Count, roundTripAttributes.Count,
                    "Round-trip should preserve number of tables");

                // Assert - Same table names
                CollectionAssert.AreEquivalent(originalAttributes, roundTripAttributes,
                    "Round-trip should preserve table names");

                // Assert - Check first table data if available
                if (originalAttributes.Count > 0)
                {
                    var tableName = originalAttributes[0];
                    var originalTable = typeof(DDF).GetProperty(tableName).GetValue(originalDDF) as CDT;
                    var roundTripTable = typeof(DDF).GetProperty(tableName).GetValue(roundTripDDF) as CDT;

                    Assert.AreEqual(originalTable.Ids.Count, roundTripTable.Ids.Count,
                        $"{tableName} should have same number of IDs");

                    Assert.AreEqual(originalTable.Data.Rows.Count, roundTripTable.Data.Rows.Count,
                        $"{tableName} should have same number of rows");

                    Assert.AreEqual(originalTable.Data.Columns.Count, roundTripTable.Data.Columns.Count,
                        $"{tableName} should have same number of columns");
                }
            }
            finally
            {
                // Cleanup
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
        }

        [TestMethod]
        public void TestDDFWithModification()
        {
            // Arrange
            var originalPath = GetSamplePath("construction and materials.DDF");
            var tempPath = Path.GetTempFileName();
            tempPath = Path.ChangeExtension(tempPath, ".ddf");

            try
            {
                // Act - Read original
                var ddf = DDF.Read(originalPath);

                // Act - Modify data (if Materials exists)
                if (ddf.HasData("Materials") && ddf.Materials.Data.Rows.Count > 0)
                {
                    var originalValue = ddf.Materials.Data.Rows[0][0];
                    ddf.Materials.Data.Rows[0][0] = "MODIFIED_TEST_VALUE";

                    // Act - Save
                    ddf.Save(tempPath);

                    // Act - Read back
                    var modifiedDDF = DDF.Read(tempPath);

                    // Assert - Modification persisted
                    Assert.AreEqual("MODIFIED_TEST_VALUE", modifiedDDF.Materials.Data.Rows[0][0],
                        "Modification should persist through save/load cycle");
                }
                else
                {
                    Assert.Inconclusive("Materials table not available for modification test");
                }
            }
            finally
            {
                // Cleanup
                if (File.Exists(tempPath))
                    File.Delete(tempPath);
            }
        }
    }
}
