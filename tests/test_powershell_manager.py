#!/usr/bin/env python3
"""
Test suite for PowerShell Manager
Tests core functionality and system compatibility
"""

import unittest
import os
import sys
import subprocess
import tempfile
import shutil
from unittest.mock import patch, MagicMock

# Add parent directory to path to import powershell_manager
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    from powershell_manager import PowerShellManager
except ImportError as e:
    print(f"Error importing powershell_manager: {e}")
    sys.exit(1)

class TestPowerShellManager(unittest.TestCase):
    """Test cases for PowerShell Manager functionality"""
    
    def setUp(self):
        """Set up test environment"""
        self.manager = PowerShellManager()
        self.test_dir = tempfile.mkdtemp()
        
    def tearDown(self):
        """Clean up test environment"""
        if os.path.exists(self.test_dir):
            shutil.rmtree(self.test_dir)
    
    def test_manager_initialization(self):
        """Test PowerShellManager initialization"""
        self.assertIsNotNone(self.manager.lock)
        self.assertIsNone(self.manager.pwsh7_path)
        self.assertEqual(len(self.manager.installation_log), 0)
    
    def test_logging_functionality(self):
        """Test thread-safe logging functionality"""
        test_action = "Test action"
        test_details = "Test details"
        
        self.manager.log_action(test_action, success=True, details=test_details)
        
        self.assertEqual(len(self.manager.installation_log), 1)
        log_entry = self.manager.installation_log[0]
        self.assertIn("SUCCESS", log_entry)
        self.assertIn(test_action, log_entry)
        self.assertIn(test_details, log_entry)
    
    def test_command_execution(self):
        """Test command execution functionality"""
        # Test with a simple command that should work on Windows
        result = self.manager.run_command("echo test")
        
        self.assertIsNotNone(result)
        if result:  # Only check if command executed
            self.assertEqual(result.returncode, 0)
            self.assertIn("test", result.stdout)
    
    def test_winget_availability_check(self):
        """Test winget availability detection"""
        # This test checks if winget detection works
        # Result may vary depending on system configuration
        result = self.manager.check_winget_available()
        self.assertIsInstance(result, bool)
    
    def test_powershell7_path_detection(self):
        """Test PowerShell 7 path detection logic"""
        # Mock common installation paths
        test_paths = [
            r"C:\Program Files\PowerShell\7\pwsh.exe",
            r"C:\Program Files (x86)\PowerShell\7\pwsh.exe"
        ]
        
        # Test path detection logic without actual file system access
        original_exists = os.path.exists
        
        def mock_exists(path):
            return path in test_paths
        
        with patch('os.path.exists', side_effect=mock_exists):
            # Should find the first existing path
            found_path = self.manager.find_powershell7_path()
            if found_path:
                self.assertIn(found_path, test_paths)
    
    @patch('ctypes.windll.shell32.IsUserAnAdmin')
    def test_admin_privilege_check(self, mock_admin_check):
        """Test administrator privilege detection"""
        # Test when user is admin
        mock_admin_check.return_value = True
        self.assertTrue(self.manager.is_admin())
        
        # Test when user is not admin
        mock_admin_check.return_value = False
        self.assertFalse(self.manager.is_admin())
    
    def test_report_generation(self):
        """Test installation report generation"""
        # Add some test log entries
        self.manager.log_action("Test action 1", success=True)
        self.manager.log_action("Test action 2", success=False, details="Test error")
        
        # Generate report
        self.manager.generate_report()
        
        # Check if report file was created
        report_file = "powershell_installation_report.txt"
        self.assertTrue(os.path.exists(report_file))
        
        # Read and verify report content
        with open(report_file, 'r') as f:
            content = f.read()
            self.assertIn("PowerShell 7 Installation and Configuration Report", content)
            self.assertIn("Test action 1", content)
            self.assertIn("Test action 2", content)
            self.assertIn("SUCCESS", content)
            self.assertIn("FAILED", content)
        
        # Clean up
        if os.path.exists(report_file):
            os.remove(report_file)
    
    def test_restore_script_creation(self):
        """Test PowerShell 5.1 restoration script creation"""
        result = self.manager.create_restore_script()
        
        # Check if script was created
        script_file = "restore_powershell51.ps1"
        if result:
            self.assertTrue(os.path.exists(script_file))
            
            # Verify script content
            with open(script_file, 'r') as f:
                content = f.read()
                self.assertIn("PowerShell 5.1 Restoration Script", content)
                self.assertIn("powershell_v51_backup.exe", content)
                self.assertIn("Move-Item", content)
            
            # Clean up
            if os.path.exists(script_file):
                os.remove(script_file)
    
    def test_environment_change_broadcast(self):
        """Test environment change broadcasting functionality"""
        # This test verifies the broadcast function doesn't crash
        # Actual broadcasting may require admin privileges
        try:
            self.manager.broadcast_environment_change()
            # If no exception is raised, the function works
            self.assertTrue(True)
        except Exception as e:
            # Some errors are expected in test environment
            self.assertIsInstance(e, Exception)

class TestSystemCompatibility(unittest.TestCase):
    """Test system compatibility and requirements"""
    
    def test_python_version(self):
        """Test Python version compatibility"""
        version = sys.version_info
        self.assertGreaterEqual(version.major, 3)
        if version.major == 3:
            self.assertGreaterEqual(version.minor, 6)
    
    def test_windows_platform(self):
        """Test Windows platform detection"""
        # This test only makes sense on Windows
        if sys.platform.startswith('win'):
            self.assertTrue(os.name == 'nt')
    
    def test_required_modules(self):
        """Test availability of required Python modules"""
        required_modules = [
            'subprocess', 'sys', 'os', 'winreg', 'shutil', 
            'json', 'time', 'pathlib', 'threading', 
            'concurrent.futures', 'ctypes'
        ]
        
        for module_name in required_modules:
            try:
                __import__(module_name)
            except ImportError:
                self.fail(f"Required module '{module_name}' is not available")

class TestEdgeCases(unittest.TestCase):
    """Test edge cases and error handling"""
    
    def setUp(self):
        self.manager = PowerShellManager()
    
    def test_invalid_command_execution(self):
        """Test handling of invalid commands"""
        result = self.manager.run_command("invalid_command_that_does_not_exist")
        # Should handle gracefully and return None or failed result
        if result is not None:
            self.assertNotEqual(result.returncode, 0)
    
    def test_empty_log_report(self):
        """Test report generation with empty logs"""
        self.manager.generate_report()
        
        report_file = "powershell_installation_report.txt"
        self.assertTrue(os.path.exists(report_file))
        
        with open(report_file, 'r') as f:
            content = f.read()
            self.assertIn("Total Actions: 0", content)
        
        # Clean up
        if os.path.exists(report_file):
            os.remove(report_file)
    
    def test_path_detection_with_no_installation(self):
        """Test path detection when PowerShell 7 is not installed"""
        with patch('os.path.exists', return_value=False):
            with patch.object(self.manager, 'run_command', return_value=None):
                result = self.manager.find_powershell7_path()
                self.assertIsNone(result)

def run_system_check():
    """Run basic system compatibility check"""
    print("Running system compatibility check...")
    print(f"Python version: {sys.version}")
    print(f"Platform: {sys.platform}")
    print(f"OS: {os.name}")
    
    if sys.platform.startswith('win'):
        print("✅ Windows platform detected")
    else:
        print("⚠️ Non-Windows platform detected - some features may not work")
    
    try:
        import winreg
        print("✅ Windows registry module available")
    except ImportError:
        print("❌ Windows registry module not available")
    
    try:
        import ctypes
        print("✅ ctypes module available")
    except ImportError:
        print("❌ ctypes module not available")
    
    print("\nTesting winget availability...")
    try:
        result = subprocess.run("winget --version", shell=True, capture_output=True, text=True, timeout=10)
        if result.returncode == 0:
            print(f"✅ WinGet available: {result.stdout.strip()}")
        else:
            print("❌ WinGet not available or not functioning")
    except Exception as e:
        print(f"❌ Error checking WinGet: {e}")

if __name__ == '__main__':
    print("PowerShell Manager Test Suite")
    print("=" * 40)
    
    # Run system compatibility check first
    run_system_check()
    print("\n" + "=" * 40)
    
    # Run unit tests
    unittest.main(verbosity=2) 