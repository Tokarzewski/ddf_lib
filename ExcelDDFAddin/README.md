# DDF Tools for Excel

Excel VSTO add-in that enables Excel to natively open and save DesignBuilder DDF (Data Definition File) files.

## Features

- **Open DDF files in Excel**: Each CDT table becomes a separate worksheet
- **Edit in Excel**: Use familiar Excel interface to edit DDF data
- **Save as DDF**: Export your changes back to DDF format
- **Custom Ribbon**: "DDF Tools" tab with easy-to-use buttons
- **Auto-detection**: Double-click .ddf files to open directly in Excel (with file association)

## Architecture

### DDFLib (C# Class Library)
Core DDF/CDT parser ported from the Python `ddf_lib` library:
- **CDT.cs**: Parses .cdt files (tab-delimited with `#` separators)
- **DDF.cs**: Handles .ddf files (ZIP archives containing .cdt files)

### ExcelDDFAddin (VSTO Add-in)
Excel integration layer:
- **DDFManager.cs**: Core open/save logic
- **ExcelHelpers.cs**: Excel interop utilities
- **DDFRibbon.cs/.xml**: Custom ribbon UI
- **ThisAddIn.cs**: Add-in lifecycle and event handlers

## DDF File Format

DDF files are ZIP archives containing multiple .cdt files. Each .cdt file represents a table:

```
construction and materials.DDF (ZIP)
├── Constructions.cdt
└── Materials.cdt
```

### CDT File Format

CDT files use custom `#` separators:

```
#123 #456 #789                    (Line 1: IDs, separated by " #")
#Name #Description #Value         (Line 2: Column headers, separated by " #")
#Item1  #Desc1  #10.5             (Line 3+: Data rows, separated by "  #")
#Item2  #Desc2  #20.0
```

## Installation

### Prerequisites

- **Windows 10/11**
- **Microsoft Excel 2013 or later**
- **Visual Studio 2019 or later** (Community Edition is fine)
- **Office Developer Tools for Visual Studio**
- **.NET Framework 4.7.2 or higher**

### Build Instructions

1. **Open in Visual Studio**:
   ```
   File > Open > Project/Solution
   Select: ExcelDDFAddin.sln (you'll need to create this)
   ```

2. **Create VSTO Project**:
   Since VSTO projects require Visual Studio templates, you need to:

   a. Create new "Excel VSTO Add-in" project in Visual Studio:
      - File > New > Project
      - Search for "Excel VSTO Add-in"
      - Name it "ExcelDDFAddin"

   b. Replace generated files with the files from this repository:
      - Copy all `.cs` files from `ExcelDDFAddin/` folder
      - Copy `DDFRibbon.xml`

   c. Add reference to DDFLib project:
      - Right-click References > Add Reference
      - Projects > DDFLib

3. **Build the Solution**:
   ```
   Build > Build Solution (Ctrl+Shift+B)
   ```

4. **Test the Add-in**:
   - Press F5 to launch Excel with the add-in loaded
   - You should see "DDF Tools" tab in the ribbon

### Deployment

#### Option 1: ClickOnce Deployment (Recommended)
1. Right-click ExcelDDFAddin project > Publish
2. Follow the wizard to create an installer
3. Share the publish folder with users

#### Option 2: Manual Installation
1. Build the project in Release mode
2. Copy the following files to user's machine:
   - `ExcelDDFAddin.dll`
   - `ExcelDDFAddin.dll.manifest`
   - `ExcelDDFAddin.vsto`
   - `DDFLib.dll`
3. Run `ExcelDDFAddin.vsto` to install

### File Association (Optional)

To enable double-clicking .ddf files to open in Excel:

1. **Create Registry Script** (`register-ddf.reg`):
   ```reg
   Windows Registry Editor Version 5.00

   [HKEY_CLASSES_ROOT\.ddf]
   @="DesignBuilder.DDF"

   [HKEY_CLASSES_ROOT\DesignBuilder.DDF]
   @="DesignBuilder DDF File"

   [HKEY_CLASSES_ROOT\DesignBuilder.DDF\DefaultIcon]
   @="C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE,0"

   [HKEY_CLASSES_ROOT\DesignBuilder.DDF\shell\open\command]
   @="\"C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE\" \"%1\""
   ```

2. **Run as Administrator** to import registry settings

3. **Adjust Excel path** if your Office installation is different

## Usage

### Opening DDF Files

**Method 1: Using Ribbon**
1. Open Excel
2. Click "DDF Tools" tab
3. Click "Open DDF" button
4. Select your .ddf file

**Method 2: File Association**
- Double-click any .ddf file in Windows Explorer
- Excel will open with data loaded

### Excel Worksheet Structure

After opening a DDF file, each worksheet follows this structure:

```
Row 1: IDs (bold, gray background)
Row 2: Column headers (bold)
Row 3+: Data rows
```

Example:
```
| 123  | 456  | 789  |         <- Row 1: IDs
| Name | Desc | Value|         <- Row 2: Headers
| Mat1 | ...  | 10.5 |         <- Row 3+: Data
| Mat2 | ...  | 20.0 |
```

### Editing Data

1. Edit cells as you would in any Excel workbook
2. **Do not modify Row 1 (IDs)** - these are structural identifiers
3. You can modify Row 2 (column headers) and Row 3+ (data)
4. You can add/remove rows and columns

### Saving as DDF

1. Click "DDF Tools" tab
2. Click "Save as DDF" button
3. Choose save location
4. The file will be saved with all worksheets as CDT tables

### Round-Trip Verification

To verify your changes:

1. Save from Excel as .ddf file
2. Use the Python library to verify:
   ```python
   from ddf_lib import DDF

   ddf = DDF.read("your_file.DDF")
   print(ddf.available_attributes)
   print(ddf.Materials.df)
   ```

## Testing

### Test with Sample Files

Located in `samples/` directory:
- `construction and materials.DDF`
- `glazing and shading.DDF`

### Test Procedure

1. **Open Test**:
   - Open `samples\construction and materials.DDF`
   - Verify worksheets: "Constructions", "Materials"
   - Check that data matches Python output

2. **Edit Test**:
   - Modify a cell value
   - Note the change

3. **Save Test**:
   - Save as `test_output.DDF`
   - Close Excel

4. **Verify Test** (using Python):
   ```python
   from ddf_lib import DDF

   original = DDF.read("samples/construction and materials.DDF")
   modified = DDF.read("test_output.DDF")

   # Verify your change persisted
   print(modified.Materials.df)
   ```

## Troubleshooting

### Add-in doesn't appear in Excel
- Check: File > Options > Add-ins > Manage: COM Add-ins > Go
- Ensure "ExcelDDFAddin" is checked
- Try: Developer > COM Add-ins > Add > Browse to .vsto file

### "File cannot be opened" error
- Verify .ddf file is a valid ZIP archive
- Try extracting manually to check contents
- Ensure .cdt files inside are properly formatted

### Data appears incorrect
- Check CDT file format (use Notepad to view)
- Verify separators: ` #` for headers, `  #` for data
- Check that all lines start with `#` prefix

### Changes not saving
- Ensure you have write permissions to the target folder
- Check if file is open in another program
- Verify worksheet names match valid CDT table names

## Development

### Project Structure

```
ExcelDDFAddin/
├── DDFLib/                          # Class library
│   ├── CDT.cs                       # CDT file parser
│   ├── DDF.cs                       # DDF file handler
│   └── DDFLib.csproj
├── ExcelDDFAddin/                   # VSTO add-in
│   ├── ThisAddIn.cs                 # Add-in entry point
│   ├── DDFRibbon.cs                 # Ribbon event handlers
│   ├── DDFRibbon.xml                # Ribbon UI definition
│   ├── DDFManager.cs                # Core open/save logic
│   ├── ExcelHelpers.cs              # Excel utilities
│   └── ExcelDDFAddin.csproj
└── README.md
```

### Adding Support for New CDT Tables

If DesignBuilder adds new CDT types:

1. Update `DDFLib/DDF.cs`:
   ```csharp
   public CDT NewTableName { get; set; }
   ```

2. Rebuild the solution

3. The add-in will automatically detect and load the new table type

## License

This project is part of the `ddf_lib` repository.

## Credits

- **Python ddf_lib**: Original DDF/CDT parsing logic
- **VSTO Add-in**: C# port and Excel integration

## Support

For issues or questions:
1. Check the Troubleshooting section above
2. Verify with Python library that DDF file is valid
3. Review Visual Studio build output for errors
4. Check Excel add-in logs
