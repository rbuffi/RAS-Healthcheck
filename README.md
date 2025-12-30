# Parallels RAS Settings Documentation Tool

This PowerShell script documents all settings from a Parallels RAS installation and exports them to a Word document.

## Available Scripts

1. **Export-RASSettings.ps1** - Original version that requires Microsoft Word installed (uses COM automation)
2. **Export-RASSettings-PSWriteWord.ps1** - Recommended version that does NOT require Word (uses PSWriteWord module)

## Prerequisites

1. **Parallels RAS PowerShell Module**
   - The Parallels RAS PowerShell module must be installed
   - Typically installed with Parallels RAS or available from Parallels

2. **Output Format Requirements** (choose one):
   - **For Export-RASSettings.ps1**: Microsoft Word must be installed (uses COM automation)
   - **For Export-RASSettings-PSWriteWord.ps1**: PSWriteWord module (Install-Module PSWriteWord) - **Recommended, no Word required**

3. **Permissions**
   - Appropriate permissions to access Parallels RAS configuration
   - Typically requires administrator or RAS administrator privileges

## Installation

1. Ensure the Parallels RAS PowerShell module is installed:
   ```powershell
   Get-Module -ListAvailable -Name rasadmin
   ```

2. If not installed, install it according to Parallels RAS documentation

3. **For PSWriteWord version (recommended)**, install the PSWriteWord module:
   ```powershell
   Install-Module -Name PSWriteWord -Scope CurrentUser
   ```
   The script will attempt to install it automatically if not found.

## Usage

### Basic Usage

**Recommended (no Word required):**
```powershell
.\Export-RASSettings-PSWriteWord.ps1
```

**Original version (requires Word):**
```powershell
.\Export-RASSettings.ps1
```

This will create a Word document in the current directory with a timestamped filename.

### Advanced Usage

Specify a custom output path and document name:

```powershell
.\Export-RASSettings.ps1 -OutputPath "C:\Reports" -DocumentName "RAS-Configuration-2024"
```

### Parameters

- **OutputPath** (Optional): Path where the Word document will be saved. Defaults to current directory.
- **DocumentName** (Optional): Name of the Word document (without .docx extension). Defaults to "RAS-Settings-{timestamp}".

## What Gets Documented

The script collects and documents the following information:

1. **Site Configuration** - RAS site details
2. **Server Configuration** - All RAS servers and their properties
3. **Farm Configuration** - Application and desktop farms
4. **Application Configuration** - Published applications
5. **User Configuration** - RAS users
6. **Gateway Configuration** - RAS Gateway servers
7. **License Configuration** - License information
8. **Active Sessions** - Current user sessions
9. **Detailed Configuration** - Complete property details for all objects using all available Get-RAS* cmdlets

## Output Format

The generated Word document includes:

- Title page with generation metadata
- Formatted tables for structured data
- Detailed property listings for all configuration objects
- Professional formatting with headers and sections

## Troubleshooting

### Module Not Found

If you receive an error about the Parallels RAS module not being found:

1. Verify the module is installed:
   ```powershell
   Get-Module -ListAvailable -Name rasadmin
   ```

2. Check the module path and ensure it's in your PSModulePath

### Word Not Available

**Solution: Use the PSWriteWord version instead!**

If you don't have Word installed, use:
```powershell
.\Export-RASSettings-PSWriteWord.ps1
```

This version uses the PSWriteWord module which creates Word documents without requiring Microsoft Word to be installed.

If you prefer the original version:
- Ensure Microsoft Word is installed
- Check that Word can be accessed via COM automation
- Try running PowerShell as administrator

### Permission Errors

If you receive permission errors:

- Run PowerShell as administrator
- Ensure your account has RAS administrator privileges
- Check that you can connect to the RAS management server

### Missing Data

Some cmdlets may not be available or may return no data depending on:
- Your RAS installation type
- Module version
- Permissions level
- Configuration state

The script handles missing data gracefully and will document what is available.

## Notes

- The script attempts to collect data from all available Get-RAS* cmdlets
- Some cmdlets may require specific parameters - the script uses default behavior
- Large configurations may take several minutes to document
- The Word document may be large if you have extensive configurations

## Support

For issues with:
- **This script**: Check the error messages and ensure prerequisites are met
- **Parallels RAS module**: Consult Parallels RAS documentation
- **Word automation**: Ensure Word is properly installed and accessible


