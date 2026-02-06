# Quick Start Guide - DDF Tools for Excel

Get started with editing DDF files in Excel in just a few steps!

## What You Need

- **Windows 10/11**
- **Microsoft Excel 2013 or later**
- **Visual Studio 2019+** (Community Edition is free)
  - Download: https://visualstudio.microsoft.com/downloads/
  - Install with "Office/SharePoint development" workload

## Setup (5 Minutes)

### 1. Open Visual Studio

Launch Visual Studio 2019 or later

### 2. Create VSTO Project

Since VSTO requires special project templates:

1. **File** > **New** > **Project**
2. Search for: **"Excel VSTO Add-in"**
3. Name: `ExcelDDFAddin`
4. Location: Your `ExcelDDFAddin` folder's **parent directory**
5. Framework: **.NET Framework 4.7.2**
6. Click **Create**

### 3. Add DDFLib Project

1. **Right-click solution** > **Add** > **New Project**
2. Template: **"Class Library (.NET Framework)"**
3. Name: `DDFLib`
4. Framework: **.NET Framework 4.7.2**

### 4. Copy Code Files

**For DDFLib:**
- Delete the auto-generated `Class1.cs`
- Copy `DDFLib/CDT.cs` and `DDFLib/DDF.cs` from this repo
- Copy `DDFLib/DDFLib.csproj` from this repo

**For ExcelDDFAddin:**
- Delete auto-generated files
- Copy all `.cs` files from `ExcelDDFAddin/` folder
- Copy `DDFRibbon.xml`

### 5. Add Reference

1. **Right-click ExcelDDFAddin project** > **Add** > **Reference**
2. **Projects** tab
3. Check **DDFLib**
4. Click **OK**

### 6. Build

Press **Ctrl+Shift+B** to build the solution

## Using the Add-in (1 Minute)

### 1. Start Debugging

Press **F5**

Excel will launch with the add-in loaded.

### 2. Verify Installation

Look for the **"DDF Tools"** tab in Excel's ribbon.

### 3. Open a DDF File

1. Click **"DDF Tools"** tab
2. Click **"Open DDF"** button
3. Navigate to: `samples\construction and materials.DDF`
4. Click **Open**

**Result**: Each CDT table appears as a separate worksheet!

### 4. Edit Data

- Edit any cell values (except Row 1 - the IDs)
- Row 2 contains column headers
- Row 3+ contains data

### 5. Save as DDF

1. Click **"DDF Tools"** tab
2. Click **"Save as DDF"** button
3. Choose a filename
4. Click **Save**

**Done!** Your changes are now saved in DDF format.

## Verify Your Changes (Optional)

Using the Python library:

```python
from ddf_lib import DDF

# Read your modified file
ddf = DDF.read("your_modified_file.DDF")

# Check the data
print(ddf.available_attributes)
print(ddf.Materials.df)
```

## File Association (Optional)

To double-click .ddf files to open in Excel:

### Windows Registry Method

1. **Create file** `register-ddf.reg`:
   ```reg
   Windows Registry Editor Version 5.00

   [HKEY_CLASSES_ROOT\.ddf]
   @="DesignBuilder.DDF"

   [HKEY_CLASSES_ROOT\DesignBuilder.DDF]
   @="DesignBuilder DDF File"

   [HKEY_CLASSES_ROOT\DesignBuilder.DDF\shell\open\command]
   @="\"C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE\" \"%1\""
   ```

2. **Right-click** the file > **Run as Administrator**

3. **Adjust Excel path** if your Office is installed elsewhere

Now you can double-click any .ddf file to open it in Excel!

## Troubleshooting

### "DDF Tools" tab doesn't appear

**Solution**:
- Excel > File > Options > Add-ins
- Manage: **COM Add-ins** > **Go**
- Check **ExcelDDFAddin**
- Click **OK**

### Can't open DDF file

**Solution**:
- Verify it's a valid ZIP file (rename to .zip and extract to test)
- Check that .cdt files inside are text files with `#` separators

### Build errors

**Solution**:
- Build > **Clean Solution**
- Build > **Rebuild Solution**
- Ensure DDFLib builds before ExcelDDFAddin

### Excel doesn't launch when debugging

**Solution**:
- Verify Excel is installed
- Close all Excel windows
- Press F5 again

## What Next?

- **Read [README.md](README.md)** for detailed usage instructions
- **Read [DEVELOPMENT.md](DEVELOPMENT.md)** for development guide
- **Try the sample files** in `samples/` folder
- **Customize the ribbon** by editing `DDFRibbon.xml`

## Common Use Cases

### Bulk Edit Material Properties

1. Open `construction and materials.DDF`
2. Go to "Materials" worksheet
3. Use Excel's Find & Replace to update values
4. Save as DDF

### Add New Materials

1. Open existing DDF
2. Go to "Materials" worksheet
3. Add new rows (copy existing row format)
4. Update ID in Row 1 for new column if needed
5. Save as DDF

### Compare DDF Files

1. Open first DDF in Excel (save as .xlsx)
2. Open second DDF in Excel (save as .xlsx)
3. Use Excel's comparison features
4. Or use: Data > Compare Sheets

### Export to CSV

1. Open DDF in Excel
2. Select worksheet (e.g., "Materials")
3. File > Save As > CSV
4. Process in other tools

## Support

**Issues?**
- Check [README.md](README.md) Troubleshooting section
- Check [DEVELOPMENT.md](DEVELOPMENT.md) for dev issues
- Verify DDF file with Python `ddf_lib` first

**Need Help?**
- Review sample files in `samples/` folder
- Compare with Python examples in `examples/`

## Summary

You've now:
- ✅ Built the Excel DDF add-in
- ✅ Opened a DDF file in Excel
- ✅ Edited data in Excel
- ✅ Saved changes back to DDF format
- ✅ Verified round-trip compatibility

**Happy editing! 🎉**
