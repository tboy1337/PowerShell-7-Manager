# PowerShell 7 Manager

A comprehensive Python tool for installing PowerShell 7, setting it as the system default, and managing PowerShell configurations on Windows systems.

## Features

- ✅ **Automated Installation**: Uses WinGet to install the latest PowerShell 7
- 🔧 **System-wide Configuration**: Sets PowerShell 7 as the default system-wide
- 📁 **File Associations**: Updates .ps1 file associations to use PowerShell 7
- 🛤️ **PATH Management**: Prioritizes PowerShell 7 in system PATH
- 🖥️ **Terminal Integration**: Configures Windows Terminal to use PowerShell 7 as default
- 🔒 **PowerShell 5.1 Management**: Safely limits access to PowerShell 5.1 while preserving system integrity
- 📝 **Comprehensive Logging**: Detailed installation and configuration logs
- 🔄 **Restoration Script**: Provides option to restore PowerShell 5.1 if needed
- ⚡ **Multi-threaded**: Uses parallel processing for optimal performance
- 🛡️ **Admin Privilege Handling**: Automatically requests administrator rights when needed

## Requirements

- **Windows 10/11**
- **Python 3.6+** (Check: `python --version`)
- **WinGet** (Usually pre-installed)
- **Administrator privileges** (Auto-requested)

## 🚀 Quick Start

### Option 1: One-Click Installation (Recommended)
1. **Double-click** `install_powershell7.cmd`
2. **Allow** administrator elevation when prompted
3. **Wait** for installation to complete
4. **Restart** your terminal/command prompt

### Option 2: Manual Installation
1. Open **Command Prompt** or **PowerShell** as Administrator
2. Run: `python powershell_manager.py`
3. Follow the on-screen instructions

### Option 3: With Dependencies (Optional)
```bash
pip install -r requirements.txt
python powershell_manager.py
```

*Note: The script primarily uses built-in Python modules, so dependencies are optional for enhanced functionality.*

## ✅ What This Program Does

- ✅ Installs PowerShell 7 using: `winget install Microsoft.PowerShell`
- ✅ Sets PowerShell 7 as system-wide default
- ✅ Updates file associations (.ps1 files)
- ✅ Configures system PATH to prioritize PowerShell 7
- ✅ Limits PowerShell 5.1 access (safely renames it)
- ✅ Creates restoration script for PowerShell 5.1

## Installation Process Details

### 1. **Verifies WinGet**: Checks if WinGet package manager is available
### 2. **Installs PowerShell 7**: Runs `winget install Microsoft.PowerShell --accept-package-agreements --accept-source-agreements`
### 3. **Locates Installation**: Finds the PowerShell 7 installation path
### 4. **System Configuration**: Performs the following configurations in parallel:
   - Updates system PATH to prioritize PowerShell 7
   - Sets file associations for .ps1 files
   - Creates system-wide PowerShell alias
   - Configures Windows Terminal default profile

### PowerShell 5.1 Management

**Important Note**: PowerShell 5.1 cannot be completely uninstalled as it's an integral part of Windows. Instead, the program:

- Renames `powershell.exe` to `powershell_v51_backup.exe` to limit direct access
- Creates a restoration script for easy rollback if needed
- Preserves system integrity while encouraging PowerShell 7 usage

## 🏃‍♂️ After Installation

### Verify Installation
```powershell
# Check PowerShell version (should show 7.x.x)
$PSVersionTable.PSVersion

# Test default PowerShell
powershell -Command '$PSVersionTable.PSVersion'
```

### Generated Files
- `powershell_installation_report.txt` - Detailed installation log
- `restore_powershell51.ps1` - PowerShell 5.1 restoration script

## 🔄 Restoration

If you need to restore PowerShell 5.1 access:

```powershell
# Run as Administrator
.\restore_powershell51.ps1
```

## ⚠️ Important Notes

1. **PowerShell 5.1 cannot be uninstalled** (it's part of Windows)
2. **The program renames it** to limit access while preserving system integrity
3. **Always run as Administrator** for system-wide changes
4. **Restart terminals** after installation for changes to take effect

## Troubleshooting

### PowerShell 7 Not Working
- Restart your terminal/command prompt
- Log out and log back in
- Check PATH: `$env:PATH.Split(';') | Select-String PowerShell`

### WinGet Not Found
- Ensure you're running Windows 10 version 1809 or later
- Install App Installer from Microsoft Store
- Update Windows to the latest version

### Installation Fails
- Verify internet connection
- Run as Administrator
- Check Windows Update service is running
- Temporarily disable antivirus if it blocks the installation

### PowerShell 7 Not Found After Installation
- Check installation paths manually:
  - `C:\Program Files\PowerShell\7\`
  - `C:\Program Files (x86)\PowerShell\7\`
- Restart your terminal/command prompt
- Log out and log back in to refresh environment variables

### Permission Errors
- Ensure the script is running with Administrator privileges
- Check if User Account Control (UAC) is blocking the operation
- Verify write permissions to system directories

### Log Analysis

Check `powershell_installation_report.txt` for detailed information about:
- Each step of the installation process
- Success/failure status of operations
- Error messages and troubleshooting hints
- System configuration changes made

## Advanced Configuration

### Custom Installation Paths

The program automatically detects PowerShell 7 installation. If you have a custom installation, modify the `possible_paths` list in the `find_powershell7_path()` method.

### Registry Modifications

The program modifies the following registry areas:
- `HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment` (PATH)
- `HKEY_CLASSES_ROOT\.ps1` (File associations)
- `HKEY_CLASSES_ROOT\Microsoft.PowerShellScript.1\Shell\Open\Command` (Default program)

### Thread Safety

The program uses thread-safe logging and parallel execution for configuration tasks to optimize performance while maintaining data integrity.

## Security Considerations

- The program requires Administrator privileges for system-wide changes
- All registry modifications are logged for transparency
- PowerShell 5.1 is preserved (renamed) rather than deleted
- No external dependencies are required for core functionality

## License

This project is provided as-is for educational and practical use. Use at your own discretion and always test in a non-production environment first.

## Support

For issues or questions:
1. Check the generated installation report
2. Review the troubleshooting section
3. Verify all prerequisites are met
4. Test in a clean Windows environment if possible

---

**🎉 Success!** After installation, typing `powershell` will launch PowerShell 7 instead of 5.1!

**⚠️ Important**: Always backup your system or create a restore point before making system-wide changes. While this tool is designed to be safe and reversible, system modifications always carry inherent risks. 