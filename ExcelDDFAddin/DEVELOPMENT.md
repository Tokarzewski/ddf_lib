# Development Guide - DDF Tools for Excel

This guide provides detailed instructions for developers who want to build, test, and modify the DDF Tools Excel add-in.

## Prerequisites

### Required Software

1. **Visual Studio 2019 or later**
   - Download: https://visualstudio.microsoft.com/downloads/
   - Edition: Community (free), Professional, or Enterprise
   - Workloads required:
     - "Office/SharePoint development"
     - ".NET desktop development"

2. **Microsoft Office Excel 2013 or later**
   - Must be installed on the development machine
   - Required for VSTO development and testing

3. **.NET Framework 4.7.2 or higher**
   - Usually installed with Visual Studio
   - SDK included in Visual Studio installer

4. **Visual Studio Tools for Office (VSTO) Runtime**
   - Included with Visual Studio
   - Also needed on end-user machines

### Optional Tools

- **NuGet Package Manager** (included with Visual Studio)
- **Git** for version control
- **Python 3.x** with `ddf_lib` for testing round-trips

## Project Setup

### Creating the VSTO Project (First Time Only)

Since VSTO projects require Visual Studio's project templates, you can't just open the .csproj files directly. Follow these steps:

#### Option 1: Create from Template (Recommended for new developers)

1. **Launch Visual Studio**
2. **Create New Project**:
   - File > New > Project (Ctrl+Shift+N)
   - Search for: "Excel VSTO Add-in" or "Excel Add-in"
   - Template: "Excel VSTO Add-in" (C#)
   - Name: `ExcelDDFAddin`
   - Location: Choose your `ExcelDDFAddin` folder's parent directory
   - Framework: .NET Framework 4.7.2

3. **Add DDFLib Class Library**:
   - Right-click solution > Add > New Project
   - Template: "Class Library (.NET Framework)"
   - Name: `DDFLib`
   - Framework: .NET Framework 4.7.2

4. **Replace Files**:
   - Delete the auto-generated files in `ExcelDDFAddin` project
   - Copy all files from this repository's `ExcelDDFAddin/ExcelDDFAddin/` folder
   - Delete the auto-generated files in `DDFLib` project
   - Copy all files from this repository's `ExcelDDFAddin/DDFLib/` folder

5. **Add Project Reference**:
   - Right-click `ExcelDDFAddin` project > Add > Reference
   - Projects tab > Check `DDFLib`
   - Click OK

6. **Add Test Project**:
   - Right-click solution > Add > New Project
   - Template: "MSTest Test Project (.NET Framework)"
   - Name: `DDFLib.Tests`
   - Framework: .NET Framework 4.7.2
   - Copy test files from repository

#### Option 2: Use Existing Files (For experienced developers)

If you're familiar with manually editing VSTO project files:

1. Ensure the `.csproj` files have the correct VSTO project type GUIDs
2. Add necessary NuGet packages
3. Configure build and publish settings

### Installing Dependencies

The projects use NuGet for dependencies:

```bash
# In Visual Studio Package Manager Console:
PM> Update-Package -reinstall
```

Or right-click solution > Restore NuGet Packages

## Building the Project

### Debug Build

1. **Set ExcelDDFAddin as Startup Project**:
   - Right-click `ExcelDDFAddin` project
   - Select "Set as Startup Project"

2. **Build**:
   - Build > Build Solution (Ctrl+Shift+B)
   - Or right-click solution > Build Solution

3. **Run/Debug**:
   - Press F5 or Debug > Start Debugging
   - Excel will launch with the add-in loaded
   - You can set breakpoints in your C# code

### Release Build

1. Change configuration to Release:
   - Build > Configuration Manager
   - Active solution configuration: Release

2. Build solution:
   - Build > Build Solution

3. Output will be in:
   ```
   ExcelDDFAddin/bin/Release/
   ```

## Testing

### Running Unit Tests

1. **Open Test Explorer**:
   - Test > Test Explorer (or Ctrl+E, T)

2. **Run All Tests**:
   - Click "Run All" in Test Explorer
   - Or: Test > Run > All Tests

3. **Run Specific Test**:
   - Right-click test in Test Explorer
   - Select "Run"

### Manual Testing

1. **Start Debugging** (F5):
   - Excel launches with add-in

2. **Verify Ribbon**:
   - Look for "DDF Tools" tab in Excel ribbon
   - Should have "Open DDF" and "Save as DDF" buttons

3. **Test Open**:
   - Click "Open DDF"
   - Navigate to `samples\construction and materials.DDF`
   - Open file
   - Verify worksheets are created

4. **Test Edit**:
   - Modify a cell value
   - Note the change

5. **Test Save**:
   - Click "Save as DDF"
   - Save to a new location
   - Close Excel

6. **Verify Round-Trip** (using Python):
   ```bash
   cd ..
   python -m examples.edit_cdt
   # Compare output with your changes
   ```

### Debugging Tips

1. **Attach to Excel Process**:
   - Start Excel normally
   - In Visual Studio: Debug > Attach to Process
   - Select EXCEL.EXE
   - Set breakpoints and trigger actions

2. **View Output**:
   - View > Output (Ctrl+W, O)
   - Shows Console.WriteLine() messages

3. **Immediate Window**:
   - Debug > Windows > Immediate (Ctrl+D, I)
   - Execute C# code during debugging

4. **Check Add-in Loading**:
   - Excel > File > Options > Add-ins
   - Manage: COM Add-ins > Go
   - Verify "ExcelDDFAddin" is listed and checked

## Code Structure

### DDFLib Project

**Purpose**: Core DDF/CDT parsing logic, independent of Excel

```
DDFLib/
├── CDT.cs           # CDT file parser
│   ├── Read()      # Parse .cdt file
│   ├── Save()      # Write .cdt file
│   └── ParseLine() # Internal helper
└── DDF.cs           # DDF file handler
    ├── Read()      # Extract ZIP and parse CDT files
    ├── Save()      # Create ZIP from CDT files
    ├── GetAvailableAttributes()
    └── HasData()
```

**Key Methods**:
- `CDT.Read(string path)`: Parses a .cdt file into CDT object
- `CDT.Save(string path)`: Writes CDT object to .cdt file
- `DDF.Read(string path)`: Reads .ddf ZIP and loads all CDT tables
- `DDF.Save(string path)`: Saves all CDT tables to .ddf ZIP

### ExcelDDFAddin Project

**Purpose**: Excel integration via VSTO

```
ExcelDDFAddin/
├── ThisAddIn.cs         # VSTO entry point, event handlers
├── DDFRibbon.cs/.xml    # Custom ribbon UI
├── DDFManager.cs        # Core import/export logic
└── ExcelHelpers.cs      # Excel interop utilities
```

**Flow**:
1. User clicks "Open DDF" button
2. `DDFRibbon.OnOpenDDFClick()` → `DDFManager.ShowOpenDialog()`
3. `DDFManager.OpenDDFFile(path)`:
   - Calls `DDF.Read(path)` from DDFLib
   - Creates new Excel workbook
   - For each CDT table:
     - Creates worksheet
     - Writes IDs to row 1
     - Writes headers to row 2
     - Writes data starting row 3

**Key Classes**:
- `ThisAddIn`: Add-in lifecycle, hooks WorkbookOpen event
- `DDFRibbon`: Ribbon UI callbacks
- `DDFManager`: Main business logic for open/save
- `ExcelHelpers`: Utilities for DataTable ↔ Excel conversion

## Common Development Tasks

### Adding Support for New CDT Types

If DesignBuilder adds new CDT table types:

1. **Update DDF.cs**:
   ```csharp
   public class DDF
   {
       // Existing properties...
       public CDT NewTableName { get; set; }  // Add this
   }
   ```

2. **Rebuild**: The rest happens automatically via reflection

3. **Test**: Open a DDF with the new table type

### Modifying Ribbon UI

1. **Edit DDFRibbon.xml**:
   ```xml
   <button id="NewButton"
           label="New Action"
           onAction="OnNewButtonClick"
           imageMso="FileNew"/>
   ```

2. **Add Handler in DDFRibbon.cs**:
   ```csharp
   public void OnNewButtonClick(Office.IRibbonControl control)
   {
       // Your code here
   }
   ```

3. **Rebuild and Test**

### Changing CDT File Format

If the .cdt format changes:

1. **Update CDT.cs constants**:
   ```csharp
   private const string SEPARATOR_HEADER = " #";  // Modify as needed
   private const string SEPARATOR_DATA = "  #";
   private const string PREFIX = "#";
   ```

2. **Update Read() and Save() methods** in CDT.cs

3. **Update unit tests** in DDFTests.cs

4. **Test with real DDF files**

## Deployment

### Creating Installer

#### ClickOnce Deployment

1. **Right-click ExcelDDFAddin project** > Properties
2. **Publish tab**
3. **Configure**:
   - Publishing folder location
   - Installation URL (if deploying from web)
   - Prerequisites:
     - .NET Framework 4.7.2
     - VSTO Runtime
     - Office 2013 or later

4. **Click "Publish Now"**

5. **Distribute**:
   - Share the publish folder
   - Users run `setup.exe`

#### Manual Deployment

1. **Build in Release mode**

2. **Copy these files**:
   ```
   ExcelDDFAddin.dll
   ExcelDDFAddin.dll.manifest
   ExcelDDFAddin.vsto
   DDFLib.dll
   ```

3. **On target machine**:
   - Double-click `ExcelDDFAddin.vsto`
   - Follow installation prompts

### Code Signing

For production deployment:

1. **Get a code signing certificate**
2. **Sign the manifests**:
   - Project Properties > Signing
   - Check "Sign the ClickOnce manifests"
   - Select your certificate

3. **Rebuild in Release mode**

## Troubleshooting Development Issues

### "Could not load file or assembly 'DDFLib'"

**Solution**:
- Clean solution: Build > Clean Solution
- Rebuild: Build > Rebuild Solution
- Check DDFLib is built before ExcelDDFAddin

### "The type or namespace name 'Office' could not be found"

**Solution**:
- Ensure VSTO project type is correct
- Add references:
  - Microsoft.Office.Tools.Excel
  - Microsoft.Office.Interop.Excel
  - Microsoft.Office.Core

### Excel doesn't launch when debugging

**Solution**:
- Verify Excel is installed
- Check project properties > Debug > Start Action
- Ensure VSTO runtime is installed

### Ribbon doesn't appear

**Solution**:
- Check DDFRibbon.xml is set as Embedded Resource
- Verify CreateRibbonExtensibilityObject() returns new DDFRibbon()
- Check for XML syntax errors in DDFRibbon.xml

### Changes not reflected when debugging

**Solution**:
- Stop debugging (close Excel)
- Clean solution
- Rebuild solution
- Start debugging again

## Performance Optimization

### Large DDF Files

For DDF files with many rows:

1. **Use Excel Arrays** (already implemented):
   ```csharp
   // Instead of:
   for (each cell) worksheet.Cells[r,c] = value;

   // Use:
   range.Value2 = dataArray;  // Bulk write
   ```

2. **Disable Screen Updating**:
   ```csharp
   app.ScreenUpdating = false;
   // ... do work ...
   app.ScreenUpdating = true;
   ```

3. **Suppress Events**:
   ```csharp
   app.EnableEvents = false;
   // ... do work ...
   app.EnableEvents = true;
   ```

### Memory Management

- Release COM objects when done:
  ```csharp
  Marshal.ReleaseComObject(range);
  GC.Collect();
  GC.WaitForPendingFinalizers();
  ```

## Contributing

### Code Style

- Follow C# conventions
- Use meaningful variable names
- Add XML comments for public methods
- Keep methods focused and short

### Testing

- Add unit tests for new features
- Test with real DDF files
- Verify round-trip (open → edit → save → verify)

### Pull Requests

1. Fork the repository
2. Create feature branch
3. Make changes
4. Add tests
5. Submit PR with clear description

## Resources

- [VSTO Documentation](https://docs.microsoft.com/en-us/visualstudio/vsto/)
- [Excel Interop Reference](https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel)
- [ClickOnce Deployment](https://docs.microsoft.com/en-us/visualstudio/deployment/clickonce-security-and-deployment)
