# Installation Guide - DDF Tools for Excel

## Prerequisites

To build the VSTO Excel add-in, you need:

### Required Software

1. **Visual Studio 2019 or later (Community Edition is FREE)**
   - Download: https://visualstudio.microsoft.com/downloads/
   - Size: ~3-6 GB
   - License: Free for individual developers

2. **Required Workloads** (select during installation):
   - ✅ **Office/SharePoint development**
   - ✅ **.NET desktop development**

3. **Microsoft Excel 2013 or later**
   - Must be installed on the same machine
   - Part of Microsoft Office

---

## Installation Steps

### Step 1: Install Visual Studio Community

1. **Download Visual Studio Community 2022**:
   - Go to: https://visualstudio.microsoft.com/downloads/
   - Click "Free download" under "Community 2022"
   - Run the downloaded installer

2. **Select Workloads**:
   - In the Visual Studio Installer, select these workloads:
     - ✅ **Office/SharePoint development**
     - ✅ **.NET desktop development**
   - Click "Install" (this will take 15-30 minutes)

3. **Wait for Installation**:
   - The installer will download and install all required components
   - Restart your computer if prompted

### Step 2: Open the Project

1. **Launch Visual Studio**

2. **Open the Solution**:
   - File > Open > Project/Solution
   - Navigate to: `c:\GitHub\ddf_lib\ExcelDDFAddin\`
   - Select: `ExcelDDFAddin.sln`
   - Click "Open"

   **Note**: If the solution doesn't load properly (VSTO template issues), see "Alternative Setup" below.

### Step 3: Build the Solution

1. **Restore NuGet Packages**:
   - Right-click solution in Solution Explorer
   - Select "Restore NuGet Packages"

2. **Build**:
   - Build > Build Solution (Ctrl+Shift+B)
   - Wait for build to complete (check Output window)
   - You should see "Build succeeded"

### Step 4: Install the Add-in

#### Method A: Debug Install (For Testing)

1. **Press F5** to start debugging
2. Excel will launch with the add-in loaded
3. Look for "DDF Tools" tab in Excel ribbon
4. Test with sample DDF file

#### Method B: Permanent Install (For Regular Use)

1. **Build in Release Mode**:
   - Build > Configuration Manager
   - Active solution configuration: **Release**
   - Build > Build Solution

2. **Install the Add-in**:
   - Navigate to: `ExcelDDFAddin\bin\Release\`
   - Double-click: `ExcelDDFAddin.vsto`
   - Click "Install" when prompted
   - Trust the certificate if asked

3. **Verify Installation**:
   - Open Excel
   - Look for "DDF Tools" tab in ribbon
   - If not visible, go to File > Options > Add-ins > Manage COM Add-ins > Check "ExcelDDFAddin"

---

## Alternative Setup (If .sln doesn't work)

VSTO projects require specific Visual Studio project templates. If the solution doesn't load:

### Create New VSTO Project

1. **File > New > Project**
2. Search for: **"Excel VSTO Add-in"**
3. Name: `ExcelDDFAddin`
4. Location: `c:\GitHub\ddf_lib\` (parent of ExcelDDFAddin folder)
5. Click "Create"

### Add DDFLib Project

1. **Right-click solution** > **Add** > **New Project**
2. Template: **"Class Library (.NET Framework)"**
3. Name: `DDFLib`
4. Framework: **.NET Framework 4.7.2**

### Copy Source Files

**For DDFLib:**
1. Delete `Class1.cs`
2. Add Existing Item > Select `DDFLib\CDT.cs` and `DDFLib\DDF.cs`

**For ExcelDDFAddin:**
1. Delete auto-generated files (except `ThisAddIn.cs` - you'll replace it)
2. Add Existing Item > Select all `.cs` files from `ExcelDDFAddin\ExcelDDFAddin\`
3. Add Existing Item > Select `DDFRibbon.xml`
4. Set `DDFRibbon.xml` properties:
   - Build Action: **Embedded Resource**

### Add Reference

1. Right-click **ExcelDDFAddin** project > **Add Reference**
2. Projects tab > Check **DDFLib**
3. Click OK

### Build

- Build > Build Solution
- Should build successfully now

---

## Deployment (Share with Others)

### Create Installer

1. **Right-click ExcelDDFAddin project** > **Publish...**
2. **Configure Publishing**:
   - Publish Location: Choose a folder (e.g., `publish\`)
   - Installation folder URL: Leave blank for CD/network install
   - Prerequisites: Check:
     - ✅ .NET Framework 4.7.2
     - ✅ Visual Studio 2010 Tools for Office Runtime

3. **Click "Finish"** then **"Publish Now"**

4. **Share the Installer**:
   - Go to the publish folder
   - Share the entire folder with users
   - Users run `setup.exe`

---

## Troubleshooting

### "Office/SharePoint development workload not installed"

**Solution**:
1. Open Visual Studio Installer
2. Click "Modify" on your Visual Studio installation
3. Check "Office/SharePoint development" workload
4. Click "Modify" to install

### "Project type is not supported"

**Solution**:
- Ensure you have Office Developer Tools installed
- Use "Alternative Setup" method above
- Create new VSTO project from template

### "Could not load file or assembly 'DDFLib'"

**Solution**:
1. Build > Clean Solution
2. Build > Rebuild Solution
3. Ensure DDFLib builds successfully before ExcelDDFAddin

### "Excel doesn't launch when debugging"

**Solution**:
1. Verify Excel is installed
2. Close all Excel instances
3. Right-click ExcelDDFAddin project > Properties > Debug
4. Verify "Start external program" points to EXCEL.EXE

### "DDF Tools tab doesn't appear"

**Solution**:
1. Excel > File > Options > Add-ins
2. Bottom dropdown: **COM Add-ins** > **Go**
3. Check **ExcelDDFAddin**
4. Click OK
5. Restart Excel

### Build Errors

**Solution**:
1. Check Output window for specific errors
2. Ensure .NET Framework 4.7.2 SDK is installed
3. Restore NuGet packages
4. Clean and rebuild solution

---

## File Association (Optional)

To enable double-clicking .ddf files to open in Excel:

### Create Registry Script

1. **Create file**: `register-ddf.reg`

2. **Add content** (adjust Excel path if needed):
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

3. **Run as Administrator**:
   - Right-click `register-ddf.reg`
   - Select "Run as administrator"
   - Click "Yes" to confirm

4. **Test**:
   - Double-click any .ddf file
   - Excel should open and import the file

### Finding Your Excel Path

If Office 16 path doesn't work, find your Excel.exe:

```cmd
where excel
```

Or check common locations:
- `C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE` (Office 2016+)
- `C:\Program Files (x86)\Microsoft Office\Office15\EXCEL.EXE` (Office 2013)
- `C:\Program Files\Microsoft Office\Office15\EXCEL.EXE` (Office 2013 64-bit)

---

## Uninstall

### Uninstall Add-in

1. **Control Panel** > **Programs and Features**
2. Find: **ExcelDDFAddin**
3. Click **Uninstall**

Or:

1. Excel > File > Options > Add-ins
2. Manage: **COM Add-ins** > **Go**
3. Uncheck **ExcelDDFAddin**

### Remove File Association

1. **Create file**: `unregister-ddf.reg`
   ```reg
   Windows Registry Editor Version 5.00

   [-HKEY_CLASSES_ROOT\.ddf]
   [-HKEY_CLASSES_ROOT\DesignBuilder.DDF]
   ```

2. **Run as Administrator**

---

## Next Steps

After installation:

1. ✅ **Test with sample files**:
   - Open `samples\construction and materials.DDF`
   - Verify worksheets are created correctly

2. ✅ **Test editing**:
   - Modify a cell value
   - Save as new DDF file
   - Verify with Python `ddf_lib`

3. ✅ **Read documentation**:
   - [README.md](README.md) - Full documentation
   - [QUICKSTART.md](QUICKSTART.md) - Quick start guide
   - [DEVELOPMENT.md](DEVELOPMENT.md) - Developer guide

---

## Support

**Need Help?**
- Check Troubleshooting section above
- Review [DEVELOPMENT.md](DEVELOPMENT.md) for build issues
- Verify DDF file with Python `ddf_lib` library

**Visual Studio Installation Issues?**
- Visit: https://docs.microsoft.com/en-us/visualstudio/install/install-visual-studio
- Community support: https://developercommunity.visualstudio.com/
