#!/usr/bin/env python3
"""
PowerShell Manager - Install PowerShell 7 and set as system default
Handles installation, configuration, and system-wide default settings
"""

import subprocess
import sys
import os
import winreg
import shutil
import json
import time
from pathlib import Path
from threading import Thread, Lock
from concurrent.futures import ThreadPoolExecutor, as_completed
import ctypes
from ctypes import wintypes

class PowerShellManager:
    def __init__(self):
        self.lock = Lock()
        self.pwsh7_path = None
        self.installation_log = []
        
    def is_admin(self):
        """Check if running with administrator privileges"""
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False
    
    def run_as_admin(self):
        """Restart the script with administrator privileges"""
        if not self.is_admin():
            print("Administrator privileges required. Restarting with elevated permissions...")
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, " ".join(sys.argv), None, 1
            )
            sys.exit(0)
    
    def log_action(self, action, success=True, details=""):
        """Thread-safe logging of actions"""
        with self.lock:
            timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
            status = "SUCCESS" if success else "FAILED"
            log_entry = f"[{timestamp}] {status}: {action}"
            if details:
                log_entry += f" - {details}"
            self.installation_log.append(log_entry)
            print(log_entry)
    
    def run_command(self, command, shell=True, capture_output=True):
        """Execute command and return result"""
        try:
            result = subprocess.run(
                command, 
                shell=shell, 
                capture_output=capture_output, 
                text=True, 
                timeout=300
            )
            return result
        except subprocess.TimeoutExpired:
            self.log_action(f"Command timeout: {command}", False, "Operation timed out after 5 minutes")
            return None
        except Exception as e:
            self.log_action(f"Command error: {command}", False, str(e))
            return None
    
    def check_winget_available(self):
        """Check if winget is available and functional"""
        self.log_action("Checking winget availability")
        result = self.run_command("winget --version")
        if result and result.returncode == 0:
            self.log_action("Winget is available", True, f"Version: {result.stdout.strip()}")
            return True
        else:
            self.log_action("Winget is not available or not functioning", False)
            return False
    
    def install_powershell7(self):
        """Install PowerShell 7 using winget"""
        self.log_action("Starting PowerShell 7 installation")
        
        if not self.check_winget_available():
            self.log_action("Cannot proceed without winget", False)
            return False
        
        # Install PowerShell 7
        install_cmd = "winget install Microsoft.PowerShell --accept-package-agreements --accept-source-agreements"
        result = self.run_command(install_cmd)
        
        if result and result.returncode == 0:
            self.log_action("PowerShell 7 installation completed successfully")
            return True
        else:
            error_msg = result.stderr if result else "Unknown error"
            self.log_action("PowerShell 7 installation failed", False, error_msg)
            return False
    
    def find_powershell7_path(self):
        """Find PowerShell 7 installation path"""
        possible_paths = [
            r"C:\Program Files\PowerShell\7\pwsh.exe",
            r"C:\Program Files (x86)\PowerShell\7\pwsh.exe",
            os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\PowerShell\7\pwsh.exe")
        ]
        
        for path in possible_paths:
            if os.path.exists(path):
                self.pwsh7_path = path
                self.log_action(f"Found PowerShell 7 at: {path}")
                return path
        
        # Try to find via winget
        result = self.run_command("winget list Microsoft.PowerShell")
        if result and result.returncode == 0:
            # PowerShell 7 is installed, try common locations again
            for path in possible_paths:
                if os.path.exists(path):
                    self.pwsh7_path = path
                    self.log_action(f"Found PowerShell 7 at: {path}")
                    return path
        
        self.log_action("PowerShell 7 installation path not found", False)
        return None
    
    def update_system_path(self):
        """Add PowerShell 7 to system PATH and prioritize it"""
        if not self.pwsh7_path:
            return False
        
        pwsh7_dir = os.path.dirname(self.pwsh7_path)
        
        try:
            # Get current system PATH
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 
                              r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment",
                              0, winreg.KEY_ALL_ACCESS) as key:
                try:
                    current_path, _ = winreg.QueryValueEx(key, "PATH")
                except FileNotFoundError:
                    current_path = ""
                
                # Remove any existing PowerShell 7 entries
                path_entries = [p.strip() for p in current_path.split(';') if p.strip()]
                path_entries = [p for p in path_entries if not p.lower().startswith(pwsh7_dir.lower())]
                
                # Add PowerShell 7 at the beginning
                path_entries.insert(0, pwsh7_dir)
                new_path = ';'.join(path_entries)
                
                # Update registry
                winreg.SetValueEx(key, "PATH", 0, winreg.REG_EXPAND_SZ, new_path)
                
            self.log_action("System PATH updated successfully")
            return True
            
        except Exception as e:
            self.log_action("Failed to update system PATH", False, str(e))
            return False
    
    def set_powershell_file_associations(self):
        """Set PowerShell 7 as default for .ps1 files"""
        if not self.pwsh7_path:
            return False
        
        try:
            # Set .ps1 file association
            with winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, r"Microsoft.PowerShellScript.1\Shell\Open\Command") as key:
                command = f'"{self.pwsh7_path}" -File "%1"'
                winreg.SetValue(key, "", winreg.REG_SZ, command)
            
            # Set default program for .ps1 files
            with winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, r".ps1") as key:
                winreg.SetValue(key, "", winreg.REG_SZ, "Microsoft.PowerShellScript.1")
            
            self.log_action("PowerShell file associations updated")
            return True
            
        except Exception as e:
            self.log_action("Failed to set file associations", False, str(e))
            return False
    
    def create_powershell_alias(self):
        """Create system-wide alias for PowerShell commands"""
        if not self.pwsh7_path:
            return False
        
        try:
            # Create a batch file that redirects powershell to pwsh
            batch_content = f'''@echo off
"{self.pwsh7_path}" %*
'''
            
            system32_path = os.path.join(os.environ['SystemRoot'], 'System32')
            batch_file = os.path.join(system32_path, 'powershell.bat')
            
            with open(batch_file, 'w') as f:
                f.write(batch_content)
            
            self.log_action("PowerShell alias created successfully")
            return True
            
        except Exception as e:
            self.log_action("Failed to create PowerShell alias", False, str(e))
            return False
    
    def configure_terminal_default(self):
        """Configure Windows Terminal to use PowerShell 7 as default"""
        try:
            # Windows Terminal settings path
            terminal_settings_path = Path.home() / "AppData/Local/Packages/Microsoft.WindowsTerminal_8wekyb3d8bbwe/LocalState/settings.json"
            
            if terminal_settings_path.exists():
                with open(terminal_settings_path, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                
                # Find PowerShell 7 profile
                pwsh7_guid = None
                for profile in settings.get('profiles', {}).get('list', []):
                    if 'PowerShell' in profile.get('name', '') and '7' in profile.get('name', ''):
                        pwsh7_guid = profile.get('guid')
                        break
                
                if pwsh7_guid:
                    settings['defaultProfile'] = pwsh7_guid
                    
                    with open(terminal_settings_path, 'w', encoding='utf-8') as f:
                        json.dump(settings, f, indent=2)
                    
                    self.log_action("Windows Terminal default profile updated")
                    return True
                else:
                    self.log_action("PowerShell 7 profile not found in Windows Terminal", False)
            else:
                self.log_action("Windows Terminal settings not found")
                
        except Exception as e:
            self.log_action("Failed to configure Windows Terminal", False, str(e))
        
        return False
    
    def disable_powershell51_access(self):
        """Limit access to PowerShell 5.1 (cannot fully uninstall as it's part of Windows)"""
        try:
            # Rename powershell.exe to make it less accessible
            system32_path = os.path.join(os.environ['SystemRoot'], 'System32')
            ps51_path = os.path.join(system32_path, 'WindowsPowerShell', 'v1.0', 'powershell.exe')
            ps51_backup = os.path.join(system32_path, 'WindowsPowerShell', 'v1.0', 'powershell_v51_backup.exe')
            
            if os.path.exists(ps51_path) and not os.path.exists(ps51_backup):
                shutil.move(ps51_path, ps51_backup)
                self.log_action("PowerShell 5.1 access limited (renamed to backup)")
                return True
            else:
                self.log_action("PowerShell 5.1 already processed or not found")
                return True
                
        except Exception as e:
            self.log_action("Failed to limit PowerShell 5.1 access", False, str(e))
            return False
    
    def create_restore_script(self):
        """Create a script to restore PowerShell 5.1 if needed"""
        try:
            restore_script = """# PowerShell 5.1 Restoration Script
# Run this script as Administrator if you need to restore PowerShell 5.1

$system32 = Join-Path $env:SystemRoot 'System32'
$ps51Backup = Join-Path $system32 'WindowsPowerShell\\v1.0\\powershell_v51_backup.exe'
$ps51Original = Join-Path $system32 'WindowsPowerShell\\v1.0\\powershell.exe'

if (Test-Path $ps51Backup) {
    if (-not (Test-Path $ps51Original)) {
        Move-Item $ps51Backup $ps51Original
        Write-Host "PowerShell 5.1 restored successfully"
    } else {
        Write-Host "PowerShell 5.1 already exists"
    }
} else {
    Write-Host "PowerShell 5.1 backup not found"
}
"""
            
            with open("restore_powershell51.ps1", "w") as f:
                f.write(restore_script)
            
            self.log_action("PowerShell 5.1 restoration script created")
            return True
            
        except Exception as e:
            self.log_action("Failed to create restoration script", False, str(e))
            return False
    
    def verify_installation(self):
        """Verify PowerShell 7 installation and configuration"""
        self.log_action("Verifying PowerShell 7 installation")
        
        # Test PowerShell 7 execution
        if self.pwsh7_path and os.path.exists(self.pwsh7_path):
            result = self.run_command(f'"{self.pwsh7_path}" -Command "$PSVersionTable.PSVersion"')
            if result and result.returncode == 0:
                version = result.stdout.strip()
                self.log_action("PowerShell 7 verification successful", True, f"Version: {version}")
                return True
        
        self.log_action("PowerShell 7 verification failed", False)
        return False
    
    def run_installation_process(self):
        """Main installation process with multi-threading for performance"""
        self.log_action("Starting PowerShell 7 installation and configuration process")
        
        # Check admin privileges
        self.run_as_admin()
        
        # Step 1: Install PowerShell 7
        if not self.install_powershell7():
            return False
        
        # Step 2: Find installation path
        if not self.find_powershell7_path():
            return False
        
        # Use ThreadPoolExecutor for parallel configuration tasks
        with ThreadPoolExecutor(max_workers=4) as executor:
            # Submit configuration tasks
            tasks = {
                executor.submit(self.update_system_path): "Update System PATH",
                executor.submit(self.set_powershell_file_associations): "Set File Associations",
                executor.submit(self.create_powershell_alias): "Create PowerShell Alias",
                executor.submit(self.configure_terminal_default): "Configure Terminal Default"
            }
            
            # Process results
            for future in as_completed(tasks):
                task_name = tasks[future]
                try:
                    result = future.result()
                    if not result:
                        self.log_action(f"Configuration task failed: {task_name}", False)
                except Exception as e:
                    self.log_action(f"Configuration task error: {task_name}", False, str(e))
        
        # Step 3: Handle PowerShell 5.1 (sequential due to file system operations)
        self.disable_powershell51_access()
        self.create_restore_script()
        
        # Step 4: Verify installation
        self.verify_installation()
        
        # Broadcast environment change
        self.broadcast_environment_change()
        
        self.log_action("PowerShell 7 installation and configuration completed")
        return True
    
    def broadcast_environment_change(self):
        """Notify system of environment variable changes"""
        try:
            # Broadcast WM_SETTINGCHANGE message
            HWND_BROADCAST = 0xFFFF
            WM_SETTINGCHANGE = 0x1A
            ctypes.windll.user32.SendMessageW(HWND_BROADCAST, WM_SETTINGCHANGE, 0, "Environment")
            self.log_action("Environment change broadcasted")
        except Exception as e:
            self.log_action("Failed to broadcast environment change", False, str(e))
    
    def generate_report(self):
        """Generate installation report"""
        report_file = "powershell_installation_report.txt"
        
        with open(report_file, "w") as f:
            f.write("PowerShell 7 Installation and Configuration Report\n")
            f.write("=" * 50 + "\n\n")
            
            for log_entry in self.installation_log:
                f.write(log_entry + "\n")
            
            f.write("\n" + "=" * 50 + "\n")
            f.write("Installation Summary:\n")
            f.write(f"PowerShell 7 Path: {self.pwsh7_path}\n")
            f.write(f"Total Actions: {len(self.installation_log)}\n")
            
            success_count = len([log for log in self.installation_log if "SUCCESS" in log])
            f.write(f"Successful Actions: {success_count}\n")
            
            failed_count = len([log for log in self.installation_log if "FAILED" in log])
            f.write(f"Failed Actions: {failed_count}\n")
        
        self.log_action(f"Installation report generated: {report_file}")

def main():
    """Main execution function"""
    print("PowerShell 7 Installation and Configuration Tool")
    print("=" * 50)
    
    manager = PowerShellManager()
    
    try:
        # Run the installation process
        success = manager.run_installation_process()
        
        # Generate report
        manager.generate_report()
        
        if success:
            print("\n✅ PowerShell 7 installation and configuration completed successfully!")
            print("📝 Check 'powershell_installation_report.txt' for detailed logs")
            print("🔄 Restart your terminal or command prompt to use PowerShell 7")
            print("📋 To restore PowerShell 5.1 if needed, run 'restore_powershell51.ps1' as Administrator")
        else:
            print("\n❌ Installation encountered errors. Check the report for details.")
            
    except KeyboardInterrupt:
        print("\n⏹️ Installation cancelled by user")
    except Exception as e:
        print(f"\n💥 Unexpected error: {e}")
        manager.log_action("Unexpected error occurred", False, str(e))
        manager.generate_report()

if __name__ == "__main__":
    main() 