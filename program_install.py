# Filename: install_programs.py
# Refactored Program Installer (Windows, PyQt5, Metadata + Heuristic Detection)
# Includes relaxed generic filter, exclusion list, color-coding, manual install option.

import os
import re
import json
import platform
import logging
import shutil
import pythoncom
import subprocess
import winreg
import sys
import fnmatch
import math
from pathlib import Path
from typing import List, Dict, Optional, Set, Callable, Tuple, Any
from dataclasses import dataclass, field
from datetime import datetime

# --- Dependency Check ---
try:
    import win32api
    import win32com.client # Needed for MSI install check
    PYWIN32_AVAILABLE = True
except ImportError:
    PYWIN32_AVAILABLE = False

try:
    from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                                 QPushButton, QTreeWidget, QTreeWidgetItem, QHeaderView,
                                 QLabel, QLineEdit, QComboBox, QProgressBar, QFileDialog,
                                 QMessageBox, QSplitter, QToolBar, QStatusBar,
                                 QStyleFactory, QStyle)
    from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSize, QSettings, QByteArray, QTimer
    from PyQt5.QtGui import QIcon, QColor, QFont
    PYQT5_AVAILABLE = True
except ImportError:
    PYQT5_AVAILABLE = False
    # Dummy classes for basic script structure integrity if PyQt5 is missing
    class QMainWindow: pass
    class QThread: pyqtSignal = lambda *args, **kwargs: (lambda func: func)
    class QWidget: pass
    class QApplication:
        @staticmethod
        def setOrganizationName(name: str): pass
        @staticmethod
        def setApplicationName(name: str): pass
        @staticmethod
        def instance(): return None
        @staticmethod
        def processEvents(): pass
    class QStyle:
        StandardPixmap = int
        SP_DialogApplyButton, SP_DialogCancelButton, SP_BrowserReload, SP_DirOpenIcon, SP_DriveNetIcon, SP_TrashIcon = 0, 0, 0, 0, 0, 0

# --- Platform & Dependency Validation ---
IS_WINDOWS = platform.system() == "Windows"
if not IS_WINDOWS: print("CRITICAL: Requires Windows.", file=sys.stderr); sys.exit(1)
if not PYWIN32_AVAILABLE: print("CRITICAL: Requires 'pywin32' (pip install pywin32). Includes win32api and win32com.", file=sys.stderr); sys.exit(1)
if not PYQT5_AVAILABLE: print("CRITICAL: Requires 'PyQt5' (pip install PyQt5).", file=sys.stderr); sys.exit(1)

# --- Logging Configuration ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)-8s - %(message)s') # Set level=logging.DEBUG for detailed scan logs
logger = logging.getLogger('ProgramInstaller')

# --- Configuration ---
# !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
# !!! NEEDS YOUR INPUT: Define your specific software details below.      !!!
# !!! The provided entries are EXAMPLES and likely need adjustments for    !!!
# !!! your exact installers (versions, names, silent switches).           !!!
# !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
PROGRAM_CONFIG: Dict[str, Dict[str, Any]] = {
    # --- Schlumberger Software Examples ---
    "petrel": {
        "display_name": "Petrel Platform",
        "target_version": "latest", # Informational
        "identity": {
            "expected_product_names": ["Petrel", "Petrel Platform", "Schlumberger Petrel"],
            "expected_descriptions": ["Petrel Setup", "Petrel Platform Installer", "Schlumberger Petrel Installation"],
            "installer_patterns": ["Petrel*.exe", "Petrel*Setup*.exe", "SLB.Petrel*.exe"],
        },
        "check_method": {
            "type": "registry",
            "keys": [
                # Standard uninstall keys (check both 32/64 bit views)
                {"path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "match_value": "DisplayName", "match_pattern": r"Petrel.*", "get_value": "DisplayVersion"},
                {"path": r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall", "match_value": "DisplayName", "match_pattern": r"Petrel.*", "get_value": "DisplayVersion"},
                # Vendor-specific keys (existence check)
                {"path": r"SOFTWARE\Schlumberger\Petrel", "check_existence": True},
                {"path": r"SOFTWARE\WOW6432Node\Schlumberger\Petrel", "check_existence": True},
                # Common installation directory pointers
                {"path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Petrel.exe", "check_existence": True},
            ],
        },
        "install_commands": {
            # Common silent switches for InstallShield (.exe) or similar wrappers
            ".exe": '{installer_path} /s /v"/qn /norestart"', # Note nested quotes for MSI properties via /v
            # Standard MSI silent install
            ".msi": 'msiexec /i "{installer_path}" /qn /norestart',
        },
    },
    "pipesim": {
        "display_name": "PIPESIM",
        "target_version": "latest",
        "identity": {
            "expected_product_names": ["Pipesim", "Schlumberger PIPESIM"],
            "expected_descriptions": ["Pipesim Setup", "PIPESIM Installer", "PIPESIM Suite"],
            "installer_patterns": ["setup.exe", "PIPESIM*.exe", "SLB.PIPESIM*.exe"], # 'setup.exe' is common but generic, use properties to differentiate
        },
        "check_method": {
            "type": "registry",
            "keys": [
                {"path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "match_value": "DisplayName", "match_pattern": r"PIPESIM .*", "get_value": "DisplayVersion"},
                {"path": r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall", "match_value": "DisplayName", "match_pattern": r"PIPESIM .*", "get_value": "DisplayVersion"},
                {"path": r"SOFTWARE\Schlumberger\PIPESIM", "check_existence": True},
                {"path": r"SOFTWARE\WOW6432Node\Schlumberger\PIPESIM", "check_existence": True},
                # Example: Check for a known MSI Product Code (replace with actual GUID if known)
                # {"path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{YOUR-PIPESIM-GUID-HERE}", "check_existence": True, "get_value": "DisplayVersion"},
            ],
        },
        "install_commands": {
            # Common switches for Nullsoft (NSIS) or Inno Setup installers
            ".exe": '{installer_path} /S /NORESTART', # Or /SILENT, /VERYSILENT depending on installer type
            ".msi": 'msiexec /i "{installer_path}" /qn /norestart',
        },
    },
     "olga": {
        "display_name": "OLGA",
        "target_version": "latest",
        "identity": {
            "expected_product_names": ["OLGA", "Schlumberger OLGA", "OLGA Multiphase Flow Simulator"],
            "expected_descriptions": ["OLGA Setup", "OLGA Installer", "OLGA Multiphase Flow Simulator Setup"],
            "installer_patterns": ["Setup-OLGA*.exe", "OLGA*Setup*.exe", "SLB.OLGA*.exe"],
        },
        "check_method": {
            "type": "registry",
            "keys": [
                {"path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "match_value": "DisplayName", "match_pattern": r".*OLGA.*", "get_value": "DisplayVersion"},
                {"path": r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall", "match_value": "DisplayName", "match_pattern": r".*OLGA.*", "get_value": "DisplayVersion"},
                {"path": r"SOFTWARE\Schlumberger\OLGA", "check_existence": True},
                {"path": r"SOFTWARE\WOW6432Node\Schlumberger\OLGA", "check_existence": True},
                {"path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OLGA.exe", "check_existence": True},
            ],
        },
        "install_commands": {
             ".exe": '{installer_path} /S /NORESTART', # Adjust switches based on actual installer type
            ".msi": 'msiexec /i "{installer_path}" /qn /norestart',
        },
    },
     "techlog": {
        "display_name": "Techlog",
        "target_version": "latest",
        "identity": {
            "expected_product_names": ["Techlog", "Schlumberger Techlog", "Techlog Wellbore Software"],
            "expected_descriptions": ["Techlog Setup", "install Techlog", "Techlog Wellbore Software Installer"],
            "installer_patterns": ["install Techlog*.exe", "Techlog*Setup*.exe", "SLB.Techlog*.exe"],
        },
        "check_method": {
            "type": "registry",
            "keys": [
                {"path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "match_value": "DisplayName", "match_pattern": r"Techlog.*", "get_value": "DisplayVersion"},
                {"path": r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall", "match_value": "DisplayName", "match_pattern": r"Techlog.*", "get_value": "DisplayVersion"},
                {"path": r"SOFTWARE\Schlumberger\Techlog", "check_existence": True},
                {"path": r"SOFTWARE\WOW6432Node\Schlumberger\Techlog", "check_existence": True},
                {"path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Techlog.exe", "check_existence": True},
            ],
        },
        "install_commands": {
            ".exe": '{installer_path} /s /v"/qn /norestart"', # Example for InstallShield wrapper
            ".msi": 'msiexec /i "{installer_path}" /qn /norestart',
        },
    },
     "eclipse": {
        "display_name": "Eclipse Reservoir Simulator",
        "target_version": "latest",
        "identity": {
            "expected_product_names": ["Eclipse", "Eclipse Simulation", "Schlumberger Eclipse"],
            "expected_descriptions": ["Eclipse Setup", "Eclipse Simulation Installer", "Schlumberger Eclipse Reservoir Simulator"],
            "installer_patterns": ["Eclipse*.exe", "Eclipse*Setup*.exe", "SLB.Eclipse*.exe"],
        },
        "check_method": {
            "type": "registry",
            "keys": [
                {"path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "match_value": "DisplayName", "match_pattern": r"Eclipse Simulation.*", "get_value": "DisplayVersion"},
                {"path": r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall", "match_value": "DisplayName", "match_pattern": r"Eclipse Simulation.*", "get_value": "DisplayVersion"},
                {"path": r"SOFTWARE\Schlumberger\Eclipse", "check_existence": True},
                {"path": r"SOFTWARE\WOW6432Node\Schlumberger\Eclipse", "check_existence": True},
                {"path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Eclipse.exe", "check_existence": True},
            ],
        },
        "install_commands": {
            ".exe": '{installer_path} /S /NORESTART', # Adjust based on actual installer
            ".msi": 'msiexec /i "{installer_path}" /qn /norestart',
        },
    },

    # --- Other Vendor Examples ---
    "kappa_workstation": {
        "display_name": "Kappa Workstation",
        "target_version": "latest",
        "identity": {
            "expected_product_names": ["Kappa Workstation", "Workstation", "KAPPA-Workstation"],
            "expected_descriptions": ["Kappa Workstation Installer", "Setup", "KAPPA-Workstation Setup"],
            "installer_patterns": ["KappaWorkstation*.exe", "Setup*.exe", "KAPPA*.exe"], # May need refinement if 'Setup.exe' is too generic
        },
        "check_method": {
            "type": "registry",
            "keys": [
                {"path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "match_value": "DisplayName", "match_pattern": r"Kappa Workstation.*", "get_value": "DisplayVersion"},
                {"path": r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall", "match_value": "DisplayName", "match_pattern": r"Kappa Workstation.*", "get_value": "DisplayVersion"},
                {"path": r"SOFTWARE\Kappa Engineering", "check_existence": True},
                {"path": r"SOFTWARE\WOW6432Node\Kappa Engineering", "check_existence": True},
                {"path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\KappaWorkstation.exe", "check_existence": True},
            ],
        },
        "install_commands": {
            ".exe": '{installer_path} /S', # Common silent switch, verify
            ".msi": 'msiexec /i "{installer_path}" /qn /norestart',
        },
    },
    "cmg": {
        "display_name": "CMG Suite",
        "target_version": "latest",
        "identity": {
            "expected_product_names": ["CMG", "Computer Modelling Group", "CMG Launcher", "CMG Suite"],
            "expected_descriptions": ["CMG Installation", "CMG Suite Setup", "CMG Launcher Installer", "CMG 20"], # Add version specifics if needed
            "installer_patterns": ["CMG*.exe", "Setup*.exe", "*2024*.exe", "CMGLauncher*.exe"], # Patterns might include year
        },
        "check_method": {
            "type": "registry",
            "keys": [
                # Check by Display Name pattern (catches different releases)
                {"path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "match_value": "DisplayName", "match_pattern": r"CMG .* Release .*", "get_value": "DisplayVersion"},
                {"path": r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall", "match_value": "DisplayName", "match_pattern": r"CMG .* Release .*", "get_value": "DisplayVersion"},
                # Broader check if the above fails
                {"path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "match_value": "DisplayName", "match_pattern": r"CMG Suite.*", "get_value": "DisplayVersion"},
                {"path": r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall", "match_value": "DisplayName", "match_pattern": r"CMG Suite.*", "get_value": "DisplayVersion"},
                # Vendor specific keys
                {"path": r"SOFTWARE\CMG", "check_existence": True},
                {"path": r"SOFTWARE\WOW6432Node\CMG", "check_existence": True},
                {"path": r"SOFTWARE\Computer Modelling Group", "check_existence": True},
                {"path": r"SOFTWARE\WOW6432Node\Computer Modelling Group", "check_existence": True},
                # Check for a known MSI Product Code (replace with actual GUID if known & stable)
                # {"path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{YOUR-CMG-GUID-HERE}", "check_existence": True, "get_value": "DisplayVersion"},
                # App Path
                {"path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\CMGLauncher.exe", "check_existence": True},
            ],
        },
        "install_commands": {
            ".exe": '{installer_path} /s /v"/qn /norestart"', # Likely InstallShield, adjust as needed
            ".msi": 'msiexec /i "{installer_path}" /qn /norestart',
        },
    },
    "harmony_enterprise": {
        "display_name": "Harmony Enterprise",
        "target_version": "latest",
        "identity": {
            "expected_product_names": ["Harmony Enterprise", "IHS Harmony", "Harmony Well Performance Software", "S&P Harmony"], # Include variants
            "expected_descriptions": ["Harmony Enterprise Setup", "Harmony Enterprise Installer", "IHS Harmony Installation"],
            "installer_patterns": ["Harmony*.exe", "Harmony*Setup*.exe", "IHS.Harmony*.exe"],
        },
        "check_method": {
            "type": "registry",
            "keys": [
                {"path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "match_value": "DisplayName", "match_pattern": r"Harmony Enterprise.*", "get_value": "DisplayVersion"},
                {"path": r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall", "match_value": "DisplayName", "match_pattern": r"Harmony Enterprise.*", "get_value": "DisplayVersion"},
                 # Check older names too
                {"path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "match_value": "DisplayName", "match_pattern": r"IHS Harmony.*", "get_value": "DisplayVersion"},
                {"path": r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall", "match_value": "DisplayName", "match_pattern": r"IHS Harmony.*", "get_value": "DisplayVersion"},
                # Vendor keys
                {"path": r"SOFTWARE\IHS", "check_existence": True}, # Check parent first
                {"path": r"SOFTWARE\WOW6432Node\IHS", "check_existence": True},
                {"path": r"SOFTWARE\IHS\Harmony", "check_existence": True},
                {"path": r"SOFTWARE\WOW6432Node\IHS\Harmony", "check_existence": True},
                # Check S&P Global paths if applicable
                {"path": r"SOFTWARE\S&P Global\Harmony", "check_existence": True},
                {"path": r"SOFTWARE\WOW6432Node\S&P Global\Harmony", "check_existence": True},
                # App Path
                {"path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Harmony.exe", "check_existence": True},
            ],
        },
        "install_commands": {
            ".exe": '{installer_path} /S', # Verify this switch
            ".msi": 'msiexec /i "{installer_path}" /qn /norestart',
        },
    },
    "ipm": {
        "display_name": "Petex IPM Suite",
        "target_version": "latest",
        "identity": {
            "expected_product_names": ["IPM", "Petroleum Experts", "IPM Suite", "Petex IPM"],
            "expected_descriptions": ["IPM Setup", "Petroleum Experts Installation", "IPM Suite Installer", "Integrated Production Modelling"],
            "installer_patterns": ["IPM*.exe", "SetupIPM*.exe", "Setup.exe", "Petex.IPM*.exe"], # 'Setup.exe' requires property matching
        },
        "check_method": {
            "type": "registry",
            "keys": [
                {"path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "match_value": "DisplayName", "match_pattern": r"Petroleum Experts IPM.*", "get_value": "DisplayVersion"},
                {"path": r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall", "match_value": "DisplayName", "match_pattern": r"Petroleum Experts IPM.*", "get_value": "DisplayVersion"},
                # Vendor keys
                {"path": r"SOFTWARE\Petroleum Experts", "check_existence": True},
                {"path": r"SOFTWARE\WOW6432Node\Petroleum Experts", "check_existence": True},
                {"path": r"SOFTWARE\Petroleum Experts\IPM", "check_existence": True},
                {"path": r"SOFTWARE\WOW6432Node\Petroleum Experts\IPM", "check_existence": True},
                # App Path (Guessing common executable names)
                {"path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\IPM.exe", "check_existence": True},
                {"path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\GAP.exe", "check_existence": True}, # Check common modules?
                {"path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\PROSPER.exe", "check_existence": True},
                {"path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MBAL.exe", "check_existence": True},
            ],
        },
        "install_commands": {
             # Check Petex docs for reliable silent switches, these are common guesses
            ".exe": '{installer_path} /S /NORESTART', # Or /SILENT
            ".msi": 'msiexec /i "{installer_path}" /qn /norestart',
        },
    },
    "tnavigator": {
        "display_name": "tNavigator",
        "target_version": "latest",
        "identity": {
            "expected_product_names": ["tNavigator", "Rock Flow Dynamics tNavigator"],
            "expected_descriptions": ["tNavigator Setup", "tNavigator Installer", "Rock Flow Dynamics tNavigator"],
            "installer_patterns": ["tNavigator*.exe", "tNav*.exe", "RFD.tNavigator*.exe"],
        },
        "check_method": {
            "type": "registry",
            "keys": [
                {"path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "match_value": "DisplayName", "match_pattern": r"tNavigator.*", "get_value": "DisplayVersion"},
                {"path": r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall", "match_value": "DisplayName", "match_pattern": r"tNavigator.*", "get_value": "DisplayVersion"},
                 # Vendor keys
                {"path": r"SOFTWARE\Rock Flow Dynamics", "check_existence": True},
                {"path": r"SOFTWARE\WOW6432Node\Rock Flow Dynamics", "check_existence": True},
                {"path": r"SOFTWARE\Rock Flow Dynamics\tNavigator", "check_existence": True},
                {"path": r"SOFTWARE\WOW6432Node\Rock Flow Dynamics\tNavigator", "check_existence": True},
                # App Path
                {"path": r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\tNavigator.exe", "check_existence": True},
            ],
        },
        "install_commands": {
            ".exe": '{installer_path} /S', # Verify this switch
            ".msi": 'msiexec /i "{installer_path}" /qn /norestart',
        },
    },

    # --- Add more program configurations here ---
    # "example_software": {
    #    "display_name": "Example Software", "target_version": "1.2.3",
    #    "identity": { "expected_product_names": ["Example Product"], "expected_descriptions": ["Example Installer"], "installer_patterns": ["ExampleSetup_*.exe"], },
    #    "check_method": { "type": "registry", "keys": [ {"path": r"SOFTWARE\ExampleVendor\ExampleApp", "get_value": "Version"} ], },
    #    "install_commands": { ".exe": '{installer_path} /VERYSILENT /SUPPRESSMSGBOXES /NORESTART', },
    # },
}

# --- Detection Tuning ---
DETECTION_SETTINGS: Dict[str, Any] = {
    # Exclude common dependencies/runtimes unlikely to be the main target application
    # Keep this list relatively small and focused on things *never* the target.
    "exclude_generic_names": ['driver', 'redist', 'runtime', 'package', 'library', 'component'],

    # Exclude files if their properties contain these substrings. Useful for filtering out
    # common libraries, utilities, or unwanted bundled software. Case-insensitive.
    "exclude_by_property_substrings": [
        '.net framework', 'visual c++', 'visual studio tools', 'vsto',
        'codemeter runtime', 'sentinel runtime', # Common licensing components
        'microsoft edge', 'webview2', 'msedge', # Browsers / Web components
        'sql server', 'sql native client', 'odbc driver', 'oledb driver', # Database components
        'java update', 'jre', 'jdk', # Java Runtimes (unless target is specifically Java)
        'directx', 'nvidia driver', 'amd driver', 'intel driver', # Graphics/System drivers
        'adobe reader', 'acrobat reader', # Common readers (unless target)
        'silverlight', 'flash player', # Deprecated runtimes
        'remote desktop', 'anydesk', 'teamviewer', # Remote access tools
        'vcredist', # Common name for Visual C++ Redistributable installers
        'report viewer', # Common reporting component
        'crystal reports', # Common reporting component
        # --- Vendor-Specific Exclusions (Examples) ---
        'schlumberger licensing', 'slb licensing', 'codemeter control center', # Licensing tools
        'software manager', 'download manager', # Utility tools
        'productmenu', 'studio manager', 'petrel workflow tools', # Specific SLB components if not primary install
        # --- Add your own specific exclusions ---
        # 'my_internal_utility',
        # 'unwanted_bundled_app',
        ],

    # Exclude files whose properties or filename suggest they are uninstallers or patches
    "exclude_uninstaller_hints": ['uninstall', 'remove', 'uninst', 'cleanup', 'fix', 'patch', 'update'],

    # Minimum file size in bytes to consider. Helps filter out small utilities or stubs.
    "min_file_size_bytes": 5 * 1024 * 1024, # 5 MB - Adjust as needed

    # Set of directory names (lowercase) to completely ignore during scanning.
    "ignore_dirs": {
        # System Folders
        '$recycle.bin', 'system volume information', 'windows', 'programdata',
        'temp', 'tmp', 'logs', 'cache',
        # Common Software/Dev Folders
        'drivers', 'fonts', 'inf', 'driverstore', 'winsxs',
        'python', 'python27', 'python3', 'python37', 'python38', 'python39', 'python310', 'python311', 'python312',
        'java', 'jre', 'jdk', 'dotnet', '.net', 'node_modules', 'ruby', 'perl',
        '.git', '.svn', '__pycache__', '.vscode', '.idea',
        # Common Application Folders (Can be too broad, use with caution)
        'common files', 'internet explorer', 'windows defender',
        'microsoft', 'google', 'mozilla firefox', 'google chrome', # Usually OS/Browser parts, not installers in target dir
        # Specific Vendor/App Subfolders often containing non-installers
        'help', 'documentation', 'docs', 'examples', 'samples', 'bin', 'lib', 'include',
        'licenses', 'thirdparty', '3rdparty', 'redistributables',
        # Schlumberger Specific Subfolder Examples (Adjust based on your structure)
        'extensions', 'plugins', 'addins', 'configuration', 'settings', 'data',
        'petrelhelp', 'studiomanager', 'gurucontentmanager', 'plug-ins', 'simulatorplugins',
        # Other potential exclusions
        'updates', 'patches', 'hotfixes',
        # Add specific project/network drive folders if they contain non-installer clutter
        # 'my_project_data', 'shared_libs',
    }
}

# --- Data Structures ---
@dataclass
class FoundInstallerInfo:
    path: Path
    file_properties: Dict[str, Any] = field(default_factory=dict)
    installer_type: str = ".unknown" # '.exe' or '.msi'

@dataclass
class ProgramStatus:
    program_key: str
    display_name: str
    config: Dict[str, Any]
    found_installer: Optional[FoundInstallerInfo] = None
    is_installed: Optional[bool] = None # True, False, or None (unchecked)
    last_checked: Optional[datetime] = None
    last_installed_version: Optional[str] = None # Version string from registry/check
    install_error: Optional[str] = None # Last error message during install attempt

# --- Windows Utilities ---
class WindowsUtils:
    """Encapsulates Windows-specific operations like reading file properties and registry."""

    @staticmethod
    def get_file_properties(file_path_str: str) -> Optional[Dict[str, Any]]:
        """Reads version information properties from an EXE or DLL file."""
        file_path = Path(file_path_str)
        if not file_path.is_file():
            logger.debug(f"Get props skipped: Not a file '{file_path_str}'")
            return None
        properties = {}
        try:
            # Get FixedFileInfo (numeric versions)
            fixed_info = win32api.GetFileVersionInfo(str(file_path), '\\')
            if fixed_info:
                ms = fixed_info['FileVersionMS']; ls = fixed_info['FileVersionLS']
                properties['FileVersion'] = f"{win32api.HIWORD(ms)}.{win32api.LOWORD(ms)}.{win32api.HIWORD(ls)}.{win32api.LOWORD(ls)}"
                ms = fixed_info['ProductVersionMS']; ls = fixed_info['ProductVersionLS']
                properties['ProductVersion'] = f"{win32api.HIWORD(ms)}.{win32api.LOWORD(ms)}.{win32api.HIWORD(ls)}.{win32api.LOWORD(ls)}"
            else:
                 logger.debug(f"No FixedFileInfo found for {file_path.name}")

            # Get StringFileInfo (text properties like ProductName)
            lang_codepages = win32api.GetFileVersionInfo(str(file_path), r'\VarFileInfo\Translation')
            if lang_codepages:
                # Use the first language/codepage found
                lang_cp = f'{lang_codepages[0][0]:04x}{lang_codepages[0][1]:04x}'
                string_info_path = f'\\StringFileInfo\\{lang_cp}\\'
                string_keys = ['CompanyName', 'FileDescription', 'InternalName', 'LegalCopyright', 'OriginalFilename', 'ProductName']
                for key in string_keys:
                    try:
                        # Read value, strip whitespace, provide empty string if error or empty
                        value = win32api.GetFileVersionInfo(str(file_path), string_info_path + key)
                        properties[key] = value.strip() if value else ""
                    except Exception:
                        properties[key] = "" # Ensure key exists even if read fails
                logger.debug(f"Read StringFileInfo for {file_path.name} (Lang/CP: {lang_cp})")
            else:
                logger.debug(f"No language/codepage info found for {file_path.name}")
                # Attempt default keys anyway, might work for some files
                string_keys = ['CompanyName', 'FileDescription', 'InternalName', 'LegalCopyright', 'OriginalFilename', 'ProductName']
                for key in string_keys:
                    try:
                         properties[key] = win32api.GetFileVersionInfo(str(file_path), f'\\StringFileInfo\\040904b0\\{key}').strip() or "" # Try common English US
                    except Exception: properties[key] = ""

            # Ensure essential keys exist even if empty
            for key in ['FileVersion', 'ProductVersion', 'CompanyName', 'FileDescription', 'OriginalFilename', 'ProductName']:
                properties.setdefault(key, "")

            return properties
        except Exception as e:
            # Log specific error type and message if possible
            logger.debug(f"Failed getting properties for {file_path.name}: {type(e).__name__} - {e}")
            return None # Return None indicates failure to read props

    @staticmethod
    def run_command(cmd: str, program_name: str, timeout: int = 900) -> Tuple[bool, int, str, str]:
        """Executes a command line process, waits for completion, and returns status."""
        logger.info(f"Executing command for '{program_name}': {cmd}")
        try:
            # Use CREATE_NO_WINDOW to hide console windows for silent installs
            # Use shell=True carefully, ensure paths in 'cmd' are quoted properly
            result = subprocess.run(cmd, shell=True, check=False, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                    text=True, encoding='utf-8', errors='ignore', timeout=timeout,
                                    creationflags=subprocess.CREATE_NO_WINDOW)

            stdout = result.stdout.strip() if result.stdout else ""
            stderr = result.stderr.strip() if result.stderr else ""

            if stdout: logger.debug(f"Cmd stdout for '{program_name}': {stdout}")
            if stderr: logger.warning(f"Cmd stderr for '{program_name}': {stderr}") # Log stderr as warning

            # Common success codes for installers (0=OK, 3010=Reboot Required, 1641=Reboot Initiated)
            success_codes = {0, 3010, 1641}
            ran_ok = result.returncode in success_codes

            logger.info(f"Command for '{program_name}' finished. Success: {ran_ok}, Return Code: {result.returncode}")
            return ran_ok, result.returncode, stdout, stderr

        except subprocess.TimeoutExpired:
            logger.error(f"Command timed out after {timeout}s for '{program_name}': {cmd}")
            return False, -1, "", "TimeoutExpired"
        except FileNotFoundError:
            logger.error(f"Command failed: Executable not found for '{program_name}': {cmd}")
            return False, -2, "", "FileNotFoundError"
        except Exception as e:
            logger.error(f"Command execution exception for '{program_name}': {e}", exc_info=True)
            return False, -3, "", str(e)

    @staticmethod
    def check_path_exists(path_str: str) -> bool:
        """Checks if a file or directory exists, expanding environment variables."""
        try:
            expanded_path = os.path.expandvars(path_str)
            return Path(expanded_path).exists()
        except Exception as e:
            logger.warning(f"Error checking path existence for '{path_str}': {e}")
            return False

    @staticmethod
    def _reg_read_string(hkey: int, key_path: str, value_name: Optional[str]) -> Optional[str]:
        """Reads a string value from the registry. Returns None if not found or error."""
        if value_name is None: return None # Handle cases where get_value is not specified
        try:
            # Open key with KEY_READ access, handling both 32/64 bit views automatically on 64bit OS
            with winreg.OpenKey(hkey, key_path, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as key:
                value, reg_type = winreg.QueryValueEx(key, value_name)
                # Return only if it's a string type (REG_SZ or REG_EXPAND_SZ)
                if reg_type in [winreg.REG_SZ, winreg.REG_EXPAND_SZ]:
                    return str(value).strip()
                else:
                    logger.debug(f"Reg value '{value_name}' at '{key_path}' is not a string type (Type: {reg_type}).")
                    return None
        except FileNotFoundError:
            # Key or value doesn't exist - this is normal, don't log as warning
            logger.debug(f"Reg value '{value_name}' not found at '{key_path}'.")
            return None
        except OSError as e:
            # Permissions error or other OS issue
            logger.warning(f"OS Error reading reg value '{value_name}' at '{key_path}': {e}")
            return None
        except Exception as e:
            # Catchall for unexpected errors
            logger.error(f"Unexpected error reading reg value '{value_name}' at '{key_path}': {e}", exc_info=True)
            return None

    @staticmethod
    def check_registry(check_config: List[Dict]) -> Tuple[bool, Optional[str]]:
        """Checks registry based on a list of rules. Returns (found, version_string)."""
        hkey_map = {'HKLM': winreg.HKEY_LOCAL_MACHINE, 'HKCU': winreg.HKEY_CURRENT_USER}
        found_globally = False
        first_found_version: Optional[str] = None

        for rule in check_config:
            key_path: Optional[str] = rule.get("path")
            base_hive_str: str = rule.get("hive", "HKLM") # Default to HKLM
            base_hive: int = hkey_map.get(base_hive_str, winreg.HKEY_LOCAL_MACHINE)

            match_value: Optional[str] = rule.get("match_value")
            match_pattern: Optional[str] = rule.get("match_pattern")
            check_existence: bool = rule.get("check_existence", False)
            get_value: Optional[str] = rule.get("get_value") # Value to retrieve if found

            if not key_path:
                logger.warning(f"Skipping invalid registry rule (no path): {rule}")
                continue

            logger.debug(f"Checking Reg Rule: Hive={base_hive_str}, Path='{key_path}', Match='{match_value}/{match_pattern}', Exist={check_existence}, Get='{get_value}'")

            try:
                # Rule Type 1: Check if a specific key exists
                if check_existence:
                    try:
                        # Attempt to open the key. If it succeeds, the key exists.
                        # Use KEY_WOW64_64KEY to ensure we check the 64-bit view primarily
                        with winreg.OpenKey(base_hive, key_path, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY):
                            logger.debug(f"Reg Check SUCCESS (Existence): Key '{key_path}' exists.")
                            found_globally = True
                            found_version = WindowsUtils._reg_read_string(base_hive, key_path, get_value)
                            if found_version and first_found_version is None:
                                first_found_version = found_version
                                logger.debug(f"  -> Retrieved version: '{found_version}'")
                            # If existence check is enough and we found it, we can potentially stop checking this rule type
                            # return True, found_version # Or continue if other rules might provide version
                    except FileNotFoundError:
                        logger.debug(f"Reg Check FAIL (Existence): Key '{key_path}' not found.")
                        continue # Try next rule

                # Rule Type 2: Iterate subkeys under a path and match a value
                elif match_value and match_pattern:
                    try:
                        with winreg.OpenKey(base_hive, key_path, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as key:
                            subkey_index = 0
                            while True: # Loop through all subkeys
                                try:
                                    subkey_name = winreg.EnumKey(key, subkey_index)
                                    subkey_full_path = f"{key_path}\\{subkey_name}"
                                    logger.debug(f"  Checking subkey: '{subkey_full_path}'")
                                    # Read the value we need to match
                                    value_to_match = WindowsUtils._reg_read_string(base_hive, subkey_full_path, match_value)

                                    if value_to_match:
                                        logger.debug(f"    Value '{match_value}' = '{value_to_match}'")
                                        # Perform regex match (case-insensitive)
                                        if re.match(match_pattern, value_to_match, re.IGNORECASE):
                                            logger.debug(f"Reg Check SUCCESS (Match): Pattern '{match_pattern}' matched '{value_to_match}' in '{subkey_full_path}'.")
                                            found_globally = True
                                            found_version = WindowsUtils._reg_read_string(base_hive, subkey_full_path, get_value)
                                            if found_version and first_found_version is None:
                                                first_found_version = found_version
                                                logger.debug(f"    -> Retrieved version: '{found_version}'")
                                            # If we found a match, we can stop checking subkeys for this rule
                                            # return True, found_version # Or continue searching other rules
                                    subkey_index += 1
                                except OSError: # Handles ERROR_NO_MORE_ITEMS
                                    break # No more subkeys
                                except Exception as subkey_e:
                                     logger.warning(f"Error processing subkey {subkey_name} under {key_path}: {subkey_e}")
                                     subkey_index += 1 # Try next subkey
                    except FileNotFoundError:
                        logger.debug(f"Reg Check: Base path '{key_path}' for subkey iteration not found.")
                        continue # Try next rule

                # Rule Type 3: Read a specific value from a specific key
                elif get_value: # Implicitly requires key_path to exist
                     found_version = WindowsUtils._reg_read_string(base_hive, key_path, get_value)
                     if found_version is not None:
                          logger.debug(f"Reg Check SUCCESS (GetValue): Found value '{get_value}' in '{key_path}'. Version: '{found_version}'")
                          found_globally = True
                          if first_found_version is None:
                              first_found_version = found_version
                          # return True, found_version # Or continue
                     else:
                          logger.debug(f"Reg Check FAIL (GetValue): Value '{get_value}' not found or not string in '{key_path}'.")
                          continue # Try next rule

            except OSError as e:
                 logger.warning(f"OS Error checking registry rule {rule}: {e}")
                 continue # Try next rule
            except Exception as e:
                logger.error(f"Unexpected error checking registry rule {rule}: {e}", exc_info=True)
                continue # Try next rule

            # If any rule succeeded in finding the app, break the loop and return
            if found_globally:
                logger.info(f"Registry check successful for rule set. Final Result: Found={found_globally}, Version='{first_found_version}'")
                return found_globally, first_found_version

        # If loop finishes without any rule succeeding
        logger.info(f"Registry check finished. No rules matched successfully. Final Result: Found={found_globally}, Version='{first_found_version}'")
        return found_globally, first_found_version

    # --- MSI Specific Methods ---
    @staticmethod
    def _get_msi_db(msi_path: str) -> Optional[Any]:
        """Helper to open MSI database safely."""
        try:
            # Ensure msilib is imported only when needed and available
            # Note: msilib is part of standard library but might have issues in some environments
            import msilib
            # ReadOnly is usually sufficient and safer
            db = msilib.OpenDatabase(msi_path, msilib.MSIDBOPEN_READONLY)
            return db
        except ImportError:
            logger.error("Python 'msilib' module not found or import failed. Cannot read MSI properties.")
            return None
        except msilib.MSIError as e:
            logger.warning(f"MSI Error opening database {msi_path}: {e}")
            return None
        except Exception as e:
            logger.error(f"Failed to open MSI database {msi_path}: {e}", exc_info=True)
            return None

    @staticmethod
    def get_msi_properties(msi_path: str) -> Dict[str, str]:
        """Extracts properties from the Property table of an MSI file."""
        properties = {}
        db = WindowsUtils._get_msi_db(msi_path)
        if not db: return {}

        try:
            # Query the Property table
            view = db.OpenView("SELECT Property, Value FROM Property")
            view.Execute(None)
            while True:
                record = view.Fetch()
                if not record: break
                try:
                    prop_name = record.GetString(1)
                    prop_value = record.GetString(2)
                    if prop_name and prop_value is not None: # Ensure we have a name and value
                         properties[prop_name] = prop_value.strip()
                except Exception as fetch_e: # Catch potential errors reading specific records
                    logger.debug(f"Error fetching MSI property record in {msi_path}: {fetch_e}")
                    continue # Skip to next record
            view.Close() # Explicitly close the view
            logger.debug(f"Successfully read {len(properties)} properties from MSI: {msi_path}")

            # Attempt to get Summary Information (less critical, might fail)
            try:
                import msilib
                si = db.GetSummaryInformation(0) # 0 = max properties
                summary_props = {
                     'Title': 2, 'Subject': 3, 'Author': 4, 'Keywords': 5,
                     'Comments': 6, 'Template': 7, 'LastSavedBy': 8, 'RevisionNumber': 9,
                     # Add more PID_ constants if needed, see msilib docs or OLE documentation
                }
                for name, pid in summary_props.items():
                     try:
                          value = si.GetProperty(pid)
                          if isinstance(value, bytes): value = value.decode('utf-8', errors='ignore')
                          properties[f"Summary_{name}"] = str(value).strip()
                     except: continue # Ignore errors fetching individual summary props
                logger.debug(f"Read summary info from MSI: {msi_path}")
            except Exception as si_e:
                 logger.debug(f"Could not read summary information from MSI {msi_path}: {si_e}")

            return properties
        except Exception as e:
            logger.warning(f"Failed to read properties view from MSI {msi_path}: {e}")
            return {} # Return empty dict on failure

    @staticmethod
    def get_msi_product_code(msi_path: str) -> Optional[str]:
        """Gets the ProductCode GUID from an MSI file."""
        db = WindowsUtils._get_msi_db(msi_path)
        if not db: return None
        try:
            view = db.OpenView("SELECT Value FROM Property WHERE Property='ProductCode'")
            view.Execute(None)
            record = view.Fetch()
            view.Close()
            if record:
                product_code = record.GetString(1)
                logger.debug(f"Found ProductCode '{product_code}' in MSI: {msi_path}")
                return product_code
            else:
                logger.debug(f"ProductCode not found in Property table for MSI: {msi_path}")
                return None
        except Exception as e:
            logger.warning(f"Failed to get ProductCode from MSI {msi_path}: {e}")
            return None

    @staticmethod
    def is_msi_product_installed(product_code: str) -> bool:
        """Checks if an MSI product with the given ProductCode is installed using COM."""
        if not product_code: return False
        try:
            installer = win32com.client.Dispatch("WindowsInstaller.Installer")
            # Query products for all users (context=7) matching the product code
            # This is generally more reliable than iterating all products
            related_products = installer.RelatedProducts(product_code)
            if related_products and len(related_products) > 0:
                 logger.debug(f"MSI check: ProductCode '{product_code}' IS installed.")
                 return True

            # Fallback: Iterate through all products if RelatedProducts fails or is empty
            # This can be slow if many products are installed.
            # products = installer.ProductsEx(product_code, "", 7) # Query specific product code directly
            # if products and len(products) > 0:
            #      logger.debug(f"MSI check via ProductsEx: ProductCode '{product_code}' IS installed.")
            #      return True

            logger.debug(f"MSI check: ProductCode '{product_code}' is NOT installed.")
            return False
        except pythoncom.com_error as e:
             # Handle common COM errors, e.g., service not running
             logger.error(f"COM Error checking MSI ProductCode '{product_code}': {e}")
             return False # Assume not installed if check fails
        except Exception as e:
            logger.error(f"Unexpected error checking MSI ProductCode '{product_code}': {e}", exc_info=True)
            return False # Assume not installed on other errors

# --- Core Installer Logic ---
class ProgramInstaller:
    """Manages program configurations, scanning, status checking, installation, and logging."""
    HEURISTIC_SCORE_THRESHOLD = 0.5 # Adjust sensitivity of heuristic detection (0.0 to 1.0)

    def __init__(self, config: Dict = PROGRAM_CONFIG, settings: Dict = DETECTION_SETTINGS):
        self.config = config
        self.settings = settings
        self.program_status: Dict[str, ProgramStatus] = self._initialize_program_status()
        self.installation_log: Dict[str, Dict] = {} # Stores info about installations done by *this* tool
        self.search_path: Optional[str] = None
        self.unidentified_installers: List[FoundInstallerInfo] = [] # Installers passing heuristics but not config
        self._log_file: Optional[Path] = self._get_log_path()
        self._load_installation_log()

    def _initialize_program_status(self) -> Dict[str, ProgramStatus]:
        """Creates the initial status dictionary from the configuration."""
        return {key: ProgramStatus(key, cfg.get("display_name", key), cfg) for key, cfg in self.config.items()}

    def set_search_path(self, path: str) -> bool:
        """Sets and validates the directory to scan for installers."""
        try:
            resolved = Path(path).resolve(strict=True) # Ensure path exists
            if resolved.is_dir():
                self.search_path = str(resolved)
                logger.info(f"Search path set to: {self.search_path}")
                # Reset found installers when path changes
                for status in self.program_status.values():
                    status.found_installer = None
                self.unidentified_installers.clear()
                return True
            else:
                logger.error(f"Invalid search path: '{path}' is not a directory.")
                self.search_path = None
                return False
        except FileNotFoundError:
             logger.error(f"Invalid search path: '{path}' does not exist.")
             self.search_path = None; return False
        except Exception as e:
            logger.error(f"Error setting search path '{path}': {e}", exc_info=True)
            self.search_path = None
            return False

    def get_current_status(self) -> Dict[str, ProgramStatus]:
        """Returns the current status of all configured programs."""
        return self.program_status

    def _score_potential_installer(self, info: FoundInstallerInfo) -> float:
        """Calculates a heuristic score (0.0-1.0) indicating likelihood of being a target installer."""
        # Base score - Start slightly positive
        score = 0.3
        filename_lower = info.path.name.lower()
        props = info.file_properties

        # --- MSI Specific Scoring ---
        if info.installer_type == '.msi':
            score += 0.3  # Higher base score for MSIs as they are usually installers
            product_name = props.get('ProductName', '').lower()
            # Penalize common non-app MSI types based on product name
            if any(p in product_name for p in ['patch', 'update', 'hotfix', 'security update']):
                score -= 0.4
            if any(p in product_name for p in ['runtime', 'redist', 'merge module', 'driver']):
                score -= 0.5
            # Boost if product name seems non-generic
            if product_name and not any(g in product_name for g in ['install', 'setup', 'package']):
                 score += 0.1

        # --- EXE Specific Scoring ---
        elif info.installer_type == '.exe':
            inst_kw = ['setup', 'install', 'installer', 'wizard', 'web', 'online'] # Keywords suggesting installer
            uninst_kw = self.settings.get('exclude_uninstaller_hints', []) # Use configured hints

            # Filename scoring
            if any(k in filename_lower for k in inst_kw): score += 0.25
            if any(k in filename_lower for k in uninst_kw): score -= 0.35 # Strong penalty for uninstall hints in name

            # Size scoring
            try:
                size_mb = info.path.stat().st_size / (1024*1024)
                if size_mb > 100: score += 0.15 # Large files more likely main installers
                elif size_mb > 10: score += 0.10
                elif size_mb < self.settings.get('min_file_size_bytes', 0) / (1024*1024): score -= 0.15 # Penalize if below min size
            except (FileNotFoundError, OSError): score -= 0.1 # Penalize if size check fails

            # Properties scoring (if available)
            if props:
                prod_name = props.get('ProductName', '').lower()
                file_desc = props.get('FileDescription', '').lower()
                comp_name = props.get('CompanyName', '').lower()

                # Boost if properties look like a specific product
                if prod_name and not any(g in prod_name for g in ['install', 'setup', 'package', 'wizard']): score += 0.15
                if file_desc and not any(g in file_desc for g in ['install', 'setup', 'package', 'wizard']): score += 0.10
                # Penalize based on properties containing uninstall hints
                if any(k in prod_name or k in file_desc for k in uninst_kw): score -= 0.30
                # Penalize common generic company names slightly
                if comp_name in ['microsoft corporation', '']: score -= 0.05

        # Normalize score to be between 0.0 and 1.0
        return max(0.0, min(score, 1.0))

    def _find_potential_installers(self, search_path_str: str) -> List[FoundInstallerInfo]:
        """Scans the search path recursively for potential installer files (.exe, .msi)."""
        potential_files: List[FoundInstallerInfo] = []
        base_path = Path(search_path_str)
        ignore_dirs_lower = {d.lower() for d in self.settings.get('ignore_dirs', set())}
        min_size = self.settings.get('min_file_size_bytes', 0)
        exclude_generics = [n.lower() for n in self.settings.get('exclude_generic_names', [])]
        exclude_by_prop = [s.lower() for s in self.settings.get('exclude_by_property_substrings', [])]
        exclude_uninstall = [n.lower() for n in self.settings.get('exclude_uninstaller_hints', [])]

        logger.info(f"Starting recursive scan in '{base_path}'...")
        logger.info(f"Ignoring directories: {ignore_dirs_lower}")
        logger.info(f"Min file size: {min_size} bytes")
        logger.info(f"Excluding generics: {exclude_generics}")
        logger.info(f"Excluding by property: {exclude_by_prop}")
        logger.info(f"Excluding uninstall hints: {exclude_uninstall}")

        # Counters for logging scan results
        file_count, dir_count, ignored_dir_count = 0, 0, 0
        prop_read_count, prop_read_fail_count = 0, 0
        size_filter_count, prop_generic_filter_count = 0, 0
        prop_exclude_filter_count, prop_uninst_filter_count = 0, 0
        ext_filter_count = 0
        msi_prop_fail_count = 0

        for root, dirs, files in os.walk(base_path, topdown=True):
            dir_count += 1
            current_path = Path(root)
            logger.debug(f"Scanning directory: {current_path}")

            # Filter directories in-place
            original_dir_count = len(dirs)
            dirs[:] = [d for d in dirs if d.lower() not in ignore_dirs_lower]
            ignored_count_this_level = original_dir_count - len(dirs)
            ignored_dir_count += ignored_count_this_level
            if ignored_count_this_level > 0:
                 logger.debug(f"  Ignored {ignored_count_this_level} subdirectories based on ignore_dirs list.")

            for filename in files:
                file_count += 1
                filename_lower = filename.lower()
                file_path = current_path / filename
                ext_lower = file_path.suffix.lower()

                # --- Filter 1: Extension ---
                if ext_lower not in ['.exe', '.msi']:
                    ext_filter_count += 1
                    continue # Skip files that are not EXE or MSI

                logger.debug(f" -> Checking file [{file_count}]: {file_path}")

                try:
                    # --- Filter 2: File Size ---
                    file_size = file_path.stat().st_size
                    if file_size < min_size:
                        logger.debug(f"    -> SKIP: Size ({file_size} bytes) below threshold ({min_size} bytes).")
                        size_filter_count += 1
                        continue

                    properties = None
                    # --- Get Properties (Different for MSI vs EXE) ---
                    if ext_lower == '.msi':
                        logger.debug(f"    -> Reading MSI properties...")
                        properties = WindowsUtils.get_msi_properties(str(file_path))
                        prop_read_count += 1
                        if not properties:
                             logger.debug(f"    -> SKIP: Failed to read MSI properties.")
                             prop_read_fail_count += 1
                             msi_prop_fail_count += 1
                             continue
                        # Add crucial MSI info if available
                        properties['MSI_ProductCode'] = WindowsUtils.get_msi_product_code(str(file_path))
                        # Use MSI ProductVersion if available, fallback needed if missing
                        properties['ProductVersion'] = properties.get('ProductVersion', '') # Standard property name
                        properties['MSI_ProductVersion'] = properties.get('ProductVersion') # Store explicit MSI version maybe?

                    elif ext_lower == '.exe':
                        logger.debug(f"    -> Reading EXE properties...")
                        properties = WindowsUtils.get_file_properties(str(file_path))
                        prop_read_count += 1
                        if not properties:
                            logger.debug(f"    -> SKIP: Failed to read EXE properties.")
                            prop_read_fail_count += 1
                            continue
                    else: # Should not happen due to extension filter, but defensive check
                         continue

                    # Log the properties found
                    log_props = {k: v for k, v in properties.items() if k in ['ProductName', 'FileDescription', 'OriginalFilename', 'FileVersion', 'ProductVersion', 'CompanyName', 'MSI_ProductCode']}
                    logger.debug(f"    -> Properties: {log_props}")

                    # --- Apply Property-Based Filters ---
                    prop_vals_lower = { str(v).lower() for k, v in properties.items() if v and k in ['ProductName', 'FileDescription', 'OriginalFilename', 'CompanyName']}

                    # Filter 3: Exclude by specific property substrings
                    excluded_by_prop = False
                    for prop_filter in exclude_by_prop:
                         if any(prop_filter in val for val in prop_vals_lower):
                              logger.debug(f"    -> SKIP: Property filter matched '{prop_filter}'.")
                              prop_exclude_filter_count += 1
                              excluded_by_prop = True
                              break
                    if excluded_by_prop: continue

                    # Filter 4: Exclude by generic names
                    excluded_by_generic = False
                    for generic_filter in exclude_generics:
                         if any(generic_filter in val for val in prop_vals_lower):
                              logger.debug(f"    -> SKIP: Generic name filter matched '{generic_filter}'.")
                              prop_generic_filter_count += 1
                              excluded_by_generic = True
                              break
                    if excluded_by_generic: continue

                    # Filter 5: Exclude by uninstaller hints (check filename too)
                    excluded_by_uninst = False
                    prop_vals_lower.add(filename_lower) # Add filename to check for hints
                    for uninst_filter in exclude_uninstall:
                         if any(uninst_filter in val for val in prop_vals_lower):
                              logger.debug(f"    -> SKIP: Uninstaller hint filter matched '{uninst_filter}'.")
                              prop_uninst_filter_count += 1
                              excluded_by_uninst = True
                              break
                    if excluded_by_uninst: continue

                    # --- Passed all filters ---
                    logger.debug(f"    -> PASSED ALL FILTERS. Adding as potential installer.")
                    potential_files.append(FoundInstallerInfo(
                        path=file_path,
                        file_properties=properties,
                        installer_type=ext_lower
                    ))

                except FileNotFoundError:
                     logger.debug(f"    -> SKIP: File not found during processing (likely deleted mid-scan): {file_path}")
                     continue
                except OSError as e:
                     logger.warning(f"    -> SKIP: OS error processing file {file_path}: {e}")
                     continue
                except Exception as e:
                    logger.error(f"    -> SKIP: Unexpected error processing file {file_path}: {e}", exc_info=True)
                    continue

        logger.info(f"--- Scan Summary ---")
        logger.info(f"Directories scanned: {dir_count} (Excluded: {ignored_dir_count})")
        logger.info(f"Files encountered: {file_count}")
        logger.info(f"Files skipped by extension: {ext_filter_count}")
        logger.info(f"Files skipped by size: {size_filter_count}")
        logger.info(f"Property Reads attempted: {prop_read_count} (Failed: {prop_read_fail_count}, MSI specific fails: {msi_prop_fail_count})")
        logger.info(f"Files skipped by property substring filter: {prop_exclude_filter_count}")
        logger.info(f"Files skipped by generic name filter: {prop_generic_filter_count}")
        logger.info(f"Files skipped by uninstaller hint filter: {prop_uninst_filter_count}")
        logger.info(f"Potential installers identified: {len(potential_files)}")
        logger.info(f"--- End Scan Summary ---")

        return potential_files

    def scan_for_installers(self) -> Tuple[List[str], List[FoundInstallerInfo]]:
        """Scans, identifies installers based on config, and finds heuristic potentials."""
        if not self.search_path:
            logger.error("Scan cannot start: Search path is not set.")
            return [], []

        logger.info(f"Starting installer identification process in: {self.search_path}")
        # Reset status before scan
        for status in self.program_status.values():
            status.found_installer = None
        self.unidentified_installers.clear()

        # 1. Find all potential files passing basic filters
        potential_files = self._find_potential_installers(self.search_path)
        if not potential_files:
            logger.info("Scan finished: No potential installer files found after initial filtering.")
            return [], []

        logger.info(f"Found {len(potential_files)} potential installers. Matching against program configurations...")

        matched_keys: List[str] = []       # Keys of configs that found a match
        processed_paths: Set[Path] = set() # Keep track of files already assigned to a config

        # 2. Match against specific program configurations
        for key, prog_config in self.config.items():
            identity = prog_config.get('identity', {})
            # Prepare matching criteria (lowercase for case-insensitive comparison)
            exp_names = [n.lower() for n in identity.get('expected_product_names', [])]
            exp_descs = [d.lower() for d in identity.get('expected_descriptions', [])]
            patterns = [p.lower() for p in identity.get('installer_patterns', [])]
            best_match_for_key: Optional[FoundInstallerInfo] = None
            best_score = -1 # Use score to prioritize better matches if multiple files fit

            logger.debug(f"--- Comparing potentials against Config Key: '{key}' ({prog_config.get('display_name', '')}) ---")
            logger.debug(f"    Criteria: Names={exp_names}, Descs={exp_descs}, Patterns={patterns}")

            for info in potential_files:
                # Skip if already matched or has no properties (shouldn't happen after _find_potential_installers)
                if info.path in processed_paths or not info.file_properties:
                    continue

                props = info.file_properties
                filename_lower = info.path.name.lower()
                # Get properties for matching (lowercase)
                prod_name = props.get('ProductName', '').lower()
                desc = props.get('FileDescription', '').lower()
                orig_name = props.get('OriginalFilename', '').lower()

                logger.debug(f"  Comparing File: '{filename_lower}' (Prod='{prod_name}', Desc='{desc}', Orig='{orig_name}')")

                # --- Matching Logic ---
                # Score based on match quality: filename pattern < description < product name
                current_score = 0
                patt_match = any(fnmatch.fnmatch(filename_lower, pat) for pat in patterns) if patterns else False
                desc_match = any(exp in desc for exp in exp_descs) if exp_descs and desc else False
                name_match = any(exp in prod_name for exp in exp_names) if exp_names and prod_name else False

                # Also check original filename if provided
                orig_name_match = any(fnmatch.fnmatch(orig_name, pat) for pat in patterns) if patterns and orig_name else False


                if patt_match or orig_name_match: current_score += 1
                if desc_match: current_score += 2
                if name_match: current_score += 3

                logger.debug(f"    -> Score: {current_score} (Pattern={patt_match or orig_name_match}, Desc={desc_match}, Name={name_match})")

                # Consider it a potential match if score is > 0 (i.e., at least one criterion met)
                if current_score > 0:
                    # If this is a better match than the previous best for this key
                    if current_score > best_score:
                        logger.debug(f"    -> New best match for '{key}' (Score: {current_score}).")
                        best_score = current_score
                        best_match_for_key = info
                    else:
                         logger.debug(f"    -> Match found for '{key}', but score {current_score} <= previous best {best_score}.")

            # After checking all potentials for the current config key
            if best_match_for_key:
                logger.info(f"  MATCH CONFIRMED: Config '{key}' -> Installer '{best_match_for_key.path.name}' (Score: {best_score})")
                self.program_status[key].found_installer = best_match_for_key
                matched_keys.append(key)
                processed_paths.add(best_match_for_key.path) # Mark this file as used
            else:
                logger.debug(f"--- No suitable match found for Config Key: '{key}' ---")

        # 3. Apply heuristics to remaining, unmatched files
        remaining_files = [info for info in potential_files if info.path not in processed_paths]
        logger.info(f"Applying heuristics to {len(remaining_files)} remaining potential installers...")

        for info in remaining_files:
            score = self._score_potential_installer(info)
            logger.debug(f"  Heuristic check: '{info.path.name}' -> Score: {score:.2f}")
            if score >= self.HEURISTIC_SCORE_THRESHOLD:
                logger.info(f"  HEURISTIC MATCH: '{info.path.name}' (Score: {score:.2f} >= {self.HEURISTIC_SCORE_THRESHOLD}). Adding to unidentified list.")
                self.unidentified_installers.append(info)
            else:
                 logger.debug(f"  Heuristic skip: '{info.path.name}' (Score: {score:.2f} < {self.HEURISTIC_SCORE_THRESHOLD}).")


        # Sort unidentified list for consistent display
        self.unidentified_installers.sort(key=lambda i: i.path.name.lower())

        logger.info(f"--- Identification Summary ---")
        logger.info(f"Matched {len(matched_keys)} configurations to specific installers.")
        logger.info(f"Found {len(self.unidentified_installers)} additional potential installers via heuristics.")
        logger.info(f"--- End Identification Summary ---")

        return matched_keys, self.unidentified_installers

    def check_installation_status(self, program_keys: Optional[List[str]] = None) -> Dict[str, bool]:
        """Checks the installation status for specified (or all) configured programs."""
        keys_to_check = program_keys if program_keys is not None else list(self.config.keys())
        results: Dict[str, bool] = {}
        now = datetime.now()
        logger.info(f"Checking installation status for programs: {keys_to_check}")

        for key in keys_to_check:
            if key not in self.program_status:
                logger.warning(f"Skipping status check for unknown program key: '{key}'")
                continue

            status = self.program_status[key]
            check_cfg = status.config.get('check_method', {})
            check_type = check_cfg.get('type')
            is_installed: Optional[bool] = None # Start as unknown
            found_version: Optional[str] = None

            logger.debug(f"Checking status for '{status.display_name}' (Key: {key}) using type: {check_type}")

            try:
                if check_type == 'registry':
                    is_installed, found_version = WindowsUtils.check_registry(check_cfg.get('keys', []))
                elif check_type == 'path':
                    # Check if *any* of the specified paths exist
                    paths_to_check = check_cfg.get('paths', [])
                    is_installed = any(WindowsUtils.check_path_exists(p) for p in paths_to_check)
                    # Path check doesn't easily provide a version
                    found_version = None
                    logger.debug(f"Path check result for '{key}': {is_installed} (Paths: {paths_to_check})")
                elif check_type == 'msi_product_code':
                    # Requires the product code to be known, e.g., from config or a previous install log
                    product_code = check_cfg.get('product_code')
                    if product_code:
                         is_installed = WindowsUtils.is_msi_product_installed(product_code)
                         # We might be able to get version via COM too, but it's complex. Rely on registry for now.
                         found_version = None # TODO: Could try getting version via COM if needed
                         logger.debug(f"MSI Product Code check result for '{key}' (Code: {product_code}): {is_installed}")
                    else:
                         logger.warning(f"MSI Product Code check type specified for '{key}', but no product_code found in config.")
                         is_installed = False
                else:
                    logger.warning(f"Unsupported check_method type '{check_type}' specified for program '{key}'. Assuming not installed.")
                    is_installed = False

                # Update status object
                status.is_installed = bool(is_installed) # Convert None->False if check failed somehow
                status.last_checked = now
                status.last_installed_version = found_version if is_installed else None # Store version only if installed
                results[key] = status.is_installed
                logger.info(f"Status check result for '{status.display_name}': Installed={status.is_installed}, Version='{status.last_installed_version or 'N/A'}'")

            except Exception as e:
                 logger.error(f"Error during installation status check for '{key}': {e}", exc_info=True)
                 # Mark as unchecked/unknown on error
                 status.is_installed = None
                 status.last_checked = now
                 status.last_installed_version = None
                 results[key] = False # Report as not installed in the results dictionary on error

        return results

    def install_program(self, program_key: str, mode: str = 'auto') -> bool:
        """Installs a configured program using the found installer and selected mode."""
        if program_key not in self.program_status:
            logger.error(f"Installation failed: Unknown program key '{program_key}'")
            return False

        status = self.program_status[program_key]
        logger.info(f"Initiating installation for '{status.display_name}' (Key: {program_key}) in '{mode}' mode.")

        if not status.found_installer:
            logger.error(f"Installation failed for '{status.display_name}': No installer file was found or matched.")
            status.install_error = "Installer not found"
            return False

        if status.is_installed:
            logger.info(f"Skipping installation for '{status.display_name}': Program is already marked as installed.")
            # Optionally: Add logic here to check if target version requires update/reinstall
            return True # Consider already installed as success

        info = status.found_installer
        commands = status.config.get('install_commands', {})
        cmd_template = commands.get(info.installer_type) # Get command based on .exe or .msi

        if not cmd_template:
            logger.error(f"Installation failed for '{status.display_name}': No install command defined in configuration for installer type '{info.installer_type}'.")
            status.install_error = f"Missing command for {info.installer_type}"
            return False

        path_str = str(info.path)
        quoted_path = f'"{path_str}"' # Ensure path is quoted

        # Determine the final command based on the installation mode
        base_cmd_for_mode = cmd_template.format(installer_path=quoted_path) # Get the base silent command
        final_cmd = base_cmd_for_mode # Default to auto/silent

        if mode == 'manual':
            # For manual, just run the installer path directly without silent switches
            final_cmd = quoted_path
            logger.info(f"Using manual mode: Executing installer directly.")
        elif mode == 'semi':
            # Attempt to convert common silent switches to passive/semi-silent ones
            # This is highly dependent on the installer technology (MSI, InstallShield, NSIS, InnoSetup etc.)
            if info.installer_type == '.msi':
                # MSI standard passive mode
                final_cmd = f'msiexec /i {quoted_path} /passive /norestart'
            elif info.installer_type == '.exe':
                # Educated guesses for EXE wrappers:
                # Replace full silent with basic silent or passive if possible
                # Convert /qn (MSI quiet) to /qb (basic UI) if passed via /v
                if '/s /v"/qn' in base_cmd_for_mode.lower(): # InstallShield MSI wrapper
                     final_cmd = base_cmd_for_mode.replace('/qn', '/qb') # Change quiet to basic UI
                     final_cmd = final_cmd.replace('/s', '') # Remove the main silent switch? Risky. Maybe keep /s?
                elif '/S' in base_cmd_for_mode or '/SILENT' in base_cmd_for_mode or '/VERYSILENT' in base_cmd_for_mode:
                     # For NSIS/InnoSetup, there isn't always a standard "passive" mode.
                     # We might just run the silent command, or remove the switch entirely (like manual).
                     # Let's default to running the silent command but log a warning.
                     final_cmd = base_cmd_for_mode # Keep silent for now
                     logger.warning(f"Semi-silent mode for EXE ({info.path.name}): No standard passive switch known, attempting configured silent command.")
                else: # If no known silent switch was in the base command, treat as manual
                     final_cmd = quoted_path
                     logger.warning(f"Semi-silent mode for EXE ({info.path.name}): No known silent switch to modify, executing manually.")
            logger.info(f"Using semi-silent mode: Attempting execution with potentially reduced UI.")

        # Use 'start /wait' to ensure script waits for installer process to exit
        # The empty "" title is necessary for start /wait with quoted paths
        run_cmd = f'start /wait "" {final_cmd}'

        logger.info(f"Executing installation command: {run_cmd}")
        ran_ok, return_code, _, stderr = WindowsUtils.run_command(run_cmd, status.display_name)

        # --- Post-Installation Check & Status Update ---
        if ran_ok:
            logger.info(f"Installation command for '{status.display_name}' completed successfully (Return Code: {return_code}).")

            # Verify installation status after command success
            logger.info(f"Verifying installation status post-run for '{status.display_name}'...")
            if info.installer_type == '.msi' and info.file_properties.get('MSI_ProductCode'):
                product_code = info.file_properties['MSI_ProductCode']
                is_really_installed = WindowsUtils.is_msi_product_installed(product_code)
                logger.info(f"Post-install MSI check for ProductCode '{product_code}': {'Installed' if is_really_installed else 'NOT Installed'}")
            else:
                # For EXE or MSI without known product code, re-run the registry check
                is_really_installed, found_version = WindowsUtils.check_registry(status.config.get('check_method', {}).get('keys', []))
                logger.info(f"Post-install registry check: {'Installed' if is_really_installed else 'NOT Installed'}, Version: {found_version}")

            if is_really_installed:
                status.is_installed = True
                status.last_checked = datetime.now()
                status.install_error = None
                # Retrieve version again if registry check was used for verification
                if not (info.installer_type == '.msi' and info.file_properties.get('MSI_ProductCode')):
                     status.last_installed_version = found_version if is_really_installed else None
                # Record successful installation
                self._record_installation(program_key, info)
                return True
            else:
                error_msg = f"Install command succeeded (Code {return_code}), but verification check failed. Program may not be installed correctly."
                logger.error(f"Installation verification failed for '{status.display_name}': {error_msg}")
                status.is_installed = False # Mark as not installed despite success code
                status.last_checked = datetime.now()
                status.install_error = "Verification failed"
                return False
        else:
            error_msg = f"Installation command failed (Return Code: {return_code}). Stderr: {stderr or 'N/A'}"
            logger.error(f"Installation failed for '{status.display_name}': {error_msg}")
            status.is_installed = False # Ensure status is marked as not installed
            status.last_checked = datetime.now()
            status.install_error = f"Failed (Code {return_code})"
            return False

    def _is_msi_installed(self, product_code: str) -> bool:
        """Convenience wrapper for the WindowsUtils MSI check."""
        return WindowsUtils.is_msi_product_installed(product_code)

    def install_unidentified_program(self, installer_info: FoundInstallerInfo, mode: str = 'auto') -> bool:
        """Attempts to install a heuristically found program using generic silent switches."""
        prog_name = installer_info.path.name
        logger.info(f"Attempting '{mode}' mode installation for heuristically identified file: '{prog_name}'")

        # Generic default silent commands - these are best guesses!
        generic_commands = {
            ".exe": '{installer_path} /S /NORESTART', # Common for NSIS/Inno Setup
            ".msi": 'msiexec /i "{installer_path}" /qn /norestart', # Standard MSI silent
        }
        cmd_template = generic_commands.get(installer_info.installer_type)

        if not cmd_template:
            logger.error(f"Installation failed for heuristic file '{prog_name}': Unknown or unsupported installer type '{installer_info.installer_type}'.")
            # We don't have a 'status' object here, so just return False
            return False

        path_str = str(installer_info.path)
        quoted_path = f'"{path_str}"'
        base_cmd = cmd_template.format(installer_path=quoted_path)
        final_cmd = base_cmd # Default to auto/silent guess

        if mode == 'manual':
            final_cmd = quoted_path
            logger.info(f"Heuristic manual mode: Executing installer directly.")
        elif mode == 'semi':
            if installer_info.installer_type == '.msi':
                final_cmd = f'msiexec /i {quoted_path} /passive /norestart'
            elif installer_info.installer_type == '.exe':
                # Similar logic as configured install, but using generic base command
                 if '/s /v"/qn' in base_cmd.lower(): # Unlikely for generic, but check
                     final_cmd = base_cmd.replace('/qn', '/qb')
                 elif '/S' in base_cmd or '/SILENT' in base_cmd or '/VERYSILENT' in base_cmd:
                     final_cmd = base_cmd # Keep generic silent
                     logger.warning(f"Heuristic semi-silent mode for EXE ({prog_name}): Attempting generic silent command.")
                 else:
                     final_cmd = quoted_path
                     logger.warning(f"Heuristic semi-silent mode for EXE ({prog_name}): No known silent switch, executing manually.")
            logger.info(f"Heuristic semi-silent mode: Attempting execution.")

        logger.warning(f"Running installation for unidentified file. Command is a guess: {final_cmd}")
        run_cmd = f'start /wait "" {final_cmd}'
        ran_ok, return_code, _, stderr = WindowsUtils.run_command(run_cmd, f"Heuristic Install {prog_name}")

        if ran_ok:
            logger.info(f"Heuristic install command for '{prog_name}' completed successfully (Code {return_code}).")
            logger.warning("Installation verification and logging are NOT automatically performed for heuristic installs. Manual check recommended.")
            # We don't update any specific program status here. User should re-scan/check status manually.
            return True
        else:
            logger.error(f"Heuristic install command for '{prog_name}' failed (Code {return_code}). Stderr: {stderr or 'N/A'}")
            return False

    def uninstall_program(self, program_key: str) -> bool:
        """Attempts to silently uninstall a program previously installed and logged by this tool."""
        if program_key not in self.installation_log:
            logger.warning(f"Cannot uninstall '{program_key}': No installation record found in the log file. Was it installed by this tool?")
            # Try checking status anyway, maybe it was installed manually but matches config
            if program_key in self.program_status:
                 status = self.program_status[program_key]
                 logger.info(f"Checking registry for potential uninstall string for '{status.display_name}' even though not in log...")
                 # This is a bit of a long shot - find uninstall string based on DisplayName match
                 uninst_info = self._get_uninstall_info_from_registry(status.display_name, "") # No installer path hint
                 cmd_orig = uninst_info.get('uninstall_string')
                 prog_name = status.display_name
                 if cmd_orig:
                      logger.info(f"Found potential uninstall string via registry: {cmd_orig}")
                 else:
                      logger.error(f"Could not find uninstall string via registry for '{status.display_name}'. Cannot proceed.")
                      return False
            else:
                 logger.error(f"Unknown program key '{program_key}'. Cannot uninstall.")
                 return False
        else:
            # Found in log file
            log_entry = self.installation_log[program_key]
            cmd_orig = log_entry.get('uninstall_string')
            prog_name = log_entry.get('name', program_key) # Use logged name

        if not cmd_orig:
            logger.error(f"Uninstallation failed for '{prog_name}': No uninstall command was found or recorded.")
            return False

        logger.info(f"Attempting uninstallation for '{prog_name}' (Key: {program_key}).")
        logger.debug(f"Original uninstall command: {cmd_orig}")

        # Check if it looks like an MSI uninstall command (often uses ProductCode GUID)
        product_code = None
        if 'msiexec' in cmd_orig.lower() and ('/x' in cmd_orig.lower() or '/uninstall' in cmd_orig.lower()):
             match = re.search(r'\{([0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12})\}', cmd_orig, re.IGNORECASE)
             if match:
                  product_code = match.group(0)
                  logger.info(f"Detected MSI uninstall via ProductCode: {product_code}")
                  # Standardize MSI uninstall command for clarity and silent flags
                  cmd_orig = f'msiexec /x {product_code}'

        # Attempt to add silent flags
        cmd_silent = self._add_silent_flags_to_command(cmd_orig)
        logger.info(f"Executing uninstallation command: start /wait \"\" {cmd_silent}")

        # Execute uninstall command (longer timeout might be needed for complex uninstalls)
        ran_ok, return_code, _, stderr = WindowsUtils.run_command(f'start /wait "" {cmd_silent}', f"Uninstall {prog_name}", timeout=600)

        if ran_ok:
            logger.info(f"Uninstallation command for '{prog_name}' completed successfully (Code {return_code}).")

            # Verify uninstallation
            verified_uninstalled = False
            if product_code:
                logger.info(f"Verifying MSI uninstallation for ProductCode '{product_code}'...")
                if not WindowsUtils.is_msi_product_installed(product_code):
                    logger.info("MSI verification successful: Product not found.")
                    verified_uninstalled = True
                else:
                    logger.error("MSI verification FAILED: ProductCode still detected after uninstall command success.")
            else:
                # For non-MSI or unknown MSI, re-run registry check
                logger.info("Verifying uninstallation via registry check...")
                if program_key in self.program_status:
                     is_still_installed, _ = WindowsUtils.check_registry(self.program_status[program_key].config.get('check_method', {}).get('keys', []))
                     if not is_still_installed:
                          logger.info("Registry verification successful: Program no longer detected.")
                          verified_uninstalled = True
                     else:
                          logger.error("Registry verification FAILED: Program still detected after uninstall command success.")
                else:
                     logger.warning("Cannot verify uninstallation via registry: Program key unknown.")
                     verified_uninstalled = True # Assume success if verification not possible

            if verified_uninstalled:
                # Remove from log ONLY if verification passed
                if program_key in self.installation_log:
                     try:
                          del self.installation_log[program_key]
                          self._save_installation_log()
                          logger.info(f"Removed '{program_key}' from installation log.")
                     except Exception as e:
                          logger.error(f"Failed to remove '{program_key}' from log file after successful uninstall: {e}")

                # Update status in memory
                if program_key in self.program_status:
                    self.program_status[program_key].is_installed = False
                    self.program_status[program_key].last_checked = datetime.now()
                    self.program_status[program_key].last_installed_version = None
                    self.program_status[program_key].install_error = None # Clear any previous install error
                return True
            else:
                 # Verification failed
                 logger.error(f"Uninstallation verification failed for '{prog_name}'. Check system manually.")
                 # Do NOT remove from log if verification failed
                 # Update status to reflect failure? Maybe keep as installed but add error?
                 if program_key in self.program_status:
                      self.program_status[program_key].install_error = "Uninstall verification failed" # Use install_error for this?
                 return False
        else:
            logger.error(f"Uninstallation command for '{prog_name}' failed (Code {return_code}). Stderr: {stderr or 'N/A'}")
            # Update status?
            if program_key in self.program_status:
                 self.program_status[program_key].install_error = f"Uninstall failed (Code {return_code})"
            return False

    def _add_silent_flags_to_command(self, command: str) -> str:
        """Attempts to append common silent/quiet flags to an uninstall command string."""
        if not command: return ""
        cmd_lower = command.lower()
        mod_cmd = command.strip() # Start with stripped original command

        # Check if it's likely an MSI command
        is_msi_uninstall = 'msiexec' in cmd_lower and ('/x' in cmd_lower or '/uninstall' in cmd_lower)

        if is_msi_uninstall:
            # Add standard MSI silent flags if not already present
            if '/qn' not in cmd_lower and '/quiet' not in cmd_lower:
                mod_cmd += ' /qn'
                logger.debug("Added '/qn' for MSI silent uninstall.")
            if '/norestart' not in cmd_lower:
                mod_cmd += ' /norestart'
                logger.debug("Added '/norestart' for MSI uninstall.")
        else:
            # Assume EXE uninstaller (InstallShield, NSIS, InnoSetup, etc.)
            # Check if common silent flags are already there
            has_silent_flag = any(flag in cmd_lower for flag in [' /s', '/silent', '/verysilent', '/q', '/quiet', '-s', '-silent', '-q']) # Check common variations
            if not has_silent_flag:
                # Add common silent flags - /S is very common for many types
                # /VERYSILENT for InnoSetup is often needed too. Add both? Risky. Start with /S.
                mod_cmd += ' /S'
                logger.debug("Added '/S' as a potential silent flag for EXE uninstaller.")
                # Optionally add /NORESTART if it seems safe
                if 'norestart' not in cmd_lower:
                     # mod_cmd += ' /NORESTART' # Be cautious adding this generically
                     # logger.debug("Considered adding /NORESTART for EXE.")
                     pass


        if mod_cmd != command.strip():
            logger.info(f"Modified uninstall command for silent execution: '{mod_cmd}' (Original: '{command.strip()}')")
            return mod_cmd
        else:
            logger.debug(f"Uninstall command already contains silent flags or no common flags identified. Using original: '{command.strip()}'")
            return command.strip() # Return original stripped command

    # --- Logging Persistence ---
    def _get_log_path(self) -> Optional[Path]:
        """Determines the path for the installation log file in AppData."""
        try:
            appdata = os.environ.get('APPDATA')
            if appdata:
                log_dir = Path(appdata) / 'ProgramInstallerApp'
                return log_dir / 'program_installer_log.json'
            else:
                logger.error("Cannot determine log file path: APPDATA environment variable not found.")
                return None
        except Exception as e:
            logger.error(f"Error getting log file path: {e}", exc_info=True)
            return None

    def _record_installation(self, program_key: str, installer_info: FoundInstallerInfo):
        """Records successful installation details into the log file."""
        if not self._log_file:
            logger.warning("Cannot record installation: Log file path is not configured.")
            return

        prog_name = self.program_status[program_key].display_name
        logger.info(f"Recording successful installation for '{prog_name}' (Key: {program_key}).")
        logger.debug(f"Attempting to find uninstall information in registry for '{prog_name}'...")

        # Try to find the corresponding Uninstall registry entry immediately after install
        # Pass installer path as a hint for matching InstallLocation
        uninst_info = self._get_uninstall_info_from_registry(prog_name, str(installer_info.path))

        if not uninst_info.get('uninstall_string'):
             logger.warning(f"Could not find UninstallString in registry for '{prog_name}' after installation. Uninstallation via this tool may fail.")
             # Special check for MSI ProductCode - construct uninstall string manually
             if installer_info.installer_type == '.msi' and installer_info.file_properties.get('MSI_ProductCode'):
                  product_code = installer_info.file_properties['MSI_ProductCode']
                  uninst_info['uninstall_string'] = f'msiexec /x {product_code}'
                  logger.info(f"Using MSI ProductCode to construct uninstall string: {uninst_info['uninstall_string']}")

        log_entry = {
            'program_key': program_key,
            'name': prog_name,
            'timestamp': datetime.now().isoformat(),
            'installer_path': str(installer_info.path),
            'installer_type': installer_info.installer_type,
            # Info obtained from registry post-install (or MSI props)
            'uninstall_string': uninst_info.get('uninstall_string'),
            'install_location': uninst_info.get('install_location'),
            'display_version': uninst_info.get('display_version'), # Version from uninstall entry
            # Include specific installer properties if useful
            'installer_product_version': installer_info.file_properties.get('ProductVersion'),
            'installer_file_version': installer_info.file_properties.get('FileVersion'),
        }

        # Add MSI-specific info if available
        if installer_info.installer_type == '.msi':
            log_entry['product_code'] = installer_info.file_properties.get('MSI_ProductCode')
            # Use MSI property for version if available, otherwise fallback
            log_entry['display_version'] = installer_info.file_properties.get('ProductVersion') or log_entry.get('display_version')


        # Filter out None values before saving
        self.installation_log[program_key] = {k: v for k, v in log_entry.items() if v is not None}
        self._save_installation_log()
        logger.info(f"Successfully recorded installation for '{program_key}'.")
        logger.debug(f"Log entry details: {self.installation_log[program_key]}")

    def _get_uninstall_info_from_registry(self, program_name_hint: str, installer_path_hint: str) -> Dict[str, Any]:
        """Scans Uninstall registry keys to find the entry best matching the installed program."""
        # Keys to check (64-bit and 32-bit views)
        uninst_key_paths = [
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        ]
        best_match_info: Dict[str, Any] = {}
        best_score = 0
        name_hint_low = program_name_hint.lower()
        # Use installer's *parent directory* as a hint for InstallLocation matching
        inst_dir_hint = ""
        if installer_path_hint:
             try: inst_dir_hint = str(Path(installer_path_hint).parent).lower()
             except: pass

        hkey = winreg.HKEY_LOCAL_MACHINE # Primarily check HKLM

        logger.debug(f"Searching registry for uninstall info. Hint Name: '{name_hint_low}', Hint Dir: '{inst_dir_hint}'")

        for key_path in uninst_key_paths:
            try:
                # Open the base Uninstall key
                with winreg.OpenKey(hkey, key_path, 0, winreg.KEY_READ | winreg.KEY_WOW64_64KEY) as base_key:
                    subkey_index = 0
                    while True: # Iterate through all subkeys (GUIDs or names)
                        try:
                            subkey_name = winreg.EnumKey(base_key, subkey_index)
                            subkey_full_path = f"{key_path}\\{subkey_name}"
                            current_info: Dict[str, Any] = {}
                            current_score = 0

                            # Try reading common values from the subkey
                            current_info['display_name'] = WindowsUtils._reg_read_string(hkey, subkey_full_path, "DisplayName")
                            current_info['uninstall_string'] = WindowsUtils._reg_read_string(hkey, subkey_full_path, "UninstallString")
                            current_info['install_location'] = WindowsUtils._reg_read_string(hkey, subkey_full_path, "InstallLocation")
                            current_info['display_version'] = WindowsUtils._reg_read_string(hkey, subkey_full_path, "DisplayVersion")
                            current_info['publisher'] = WindowsUtils._reg_read_string(hkey, subkey_full_path, "Publisher")
                            current_info['key_path'] = subkey_full_path # Store where it was found

                            # Basic validation: Must have DisplayName and UninstallString to be considered
                            if not current_info.get('display_name') or not current_info.get('uninstall_string'):
                                subkey_index += 1
                                continue

                            # --- Scoring Logic ---
                            dn_low = current_info['display_name'].lower()
                            il_low = (current_info.get('install_location') or "").lower()

                            # 1. Strong match: DisplayName contains hint name or vice-versa
                            if name_hint_low in dn_low or dn_low in name_hint_low:
                                current_score += 5
                                logger.debug(f"  Score +5: Name match ('{dn_low}' vs '{name_hint_low}') in {subkey_name}")

                            # 2. Good match: InstallLocation matches installer directory hint
                            if il_low and inst_dir_hint and il_low == inst_dir_hint:
                                current_score += 3
                                logger.debug(f"  Score +3: Install location match ('{il_low}') in {subkey_name}")

                            # 3. Bonus: Both name and location match
                            if current_score >= 8: # Matched both name and location
                                current_score += 2
                                logger.debug(f"  Score +2: Bonus for name and location match in {subkey_name}")

                            # 4. Weaker match: Partial name match (e.g., acronym) - less reliable
                            # Could add fuzzy matching here if needed

                            logger.debug(f"  Checking Subkey: {subkey_name}, DisplayName: '{current_info['display_name']}', Score: {current_score}")

                            # Update best match if current score is higher
                            if current_score > best_score:
                                logger.debug(f"  >>> New best match found: {subkey_name} (Score: {current_score})")
                                best_score = current_score
                                best_match_info = current_info

                            subkey_index += 1
                        except OSError: # ERROR_NO_MORE_ITEMS
                            break # No more subkeys under this path
                        except Exception as sub_e:
                            logger.warning(f"Error reading subkey index {subkey_index} under {key_path}: {sub_e}")
                            subkey_index += 1 # Skip to next index

            except FileNotFoundError:
                logger.debug(f"Uninstall registry path not found: {key_path}")
                continue # This path doesn't exist, try the next one
            except Exception as base_e:
                logger.error(f"Error accessing registry path {key_path}: {base_e}", exc_info=True)
                continue # Skip this path on error

        if best_score > 0:
            logger.info(f"Found best uninstall registry match for '{program_name_hint}' with score {best_score}:")
            logger.info(f"  DisplayName: {best_match_info.get('display_name')}")
            logger.info(f"  UninstallString: {best_match_info.get('uninstall_string')}")
            logger.info(f"  KeyPath: {best_match_info.get('key_path')}")

            # Basic cleanup for UninstallString (remove arguments often added by system)
            raw_uninst = best_match_info.get('uninstall_string')
            if raw_uninst:
                # Remove surrounding quotes if present
                if raw_uninst.startswith('"') and raw_uninst.endswith('"'):
                    raw_uninst = raw_uninst[1:-1].strip()

                # Try to isolate the executable path from arguments (common pattern: MsiExec.exe /X{GUID})
                # Be careful not to remove necessary arguments for some uninstallers
                if 'msiexec.exe' in raw_uninst.lower():
                     # Keep MsiExec command with its essential /X{GUID} argument
                     parts = re.split(r'(/x\{[0-9a-f-]+\})', raw_uninst, flags=re.IGNORECASE)
                     if len(parts) >= 2:
                          cleaned_uninst = parts[0].strip() + " " + parts[1].strip()
                          logger.debug(f"Cleaned MSI uninstall string: {cleaned_uninst}")
                          best_match_info['uninstall_string'] = cleaned_uninst
                     else: # Keep original if pattern doesn't match
                          best_match_info['uninstall_string'] = raw_uninst.strip()

                # For other EXEs, try splitting at common argument indicators, but this is risky
                # Example: "C:\path\uninstaller.exe" /arg1 /arg2
                # Might want to keep only "C:\path\uninstaller.exe"
                # parts = raw_uninst.split('.exe', 1)
                # if len(parts) == 2:
                #      # Keep the .exe part, potentially discard args after first space?
                #      executable = parts[0] + '.exe'
                #      args = parts[1].split(' ', 1)[0] # Keep only first arg segment?
                #      # cleaned_uninst = executable # Simplest, discard all args
                #      # best_match_info['uninstall_string'] = cleaned_uninst
                #      pass # Keep original for non-MSI for now, too risky to modify blindly
                else:
                     best_match_info['uninstall_string'] = raw_uninst.strip()


            return best_match_info
        else:
            logger.warning(f"Could not find a suitable uninstall registry entry for hint '{program_name_hint}'.")
            return {} # Return empty dict if no match found

    def _load_installation_log(self):
        """Loads the installation log from the JSON file."""
        if not self._log_file:
            logger.warning("Log file path not set, cannot load log.")
            self.installation_log = {}
            return
        if not self._log_file.exists():
            logger.info("Installation log file not found. Starting with empty log.")
            self.installation_log = {}
            return

        try:
            with open(self._log_file, 'r', encoding='utf-8') as f:
                self.installation_log = json.load(f)
            logger.info(f"Loaded {len(self.installation_log)} installation records from {self._log_file}")
        except json.JSONDecodeError as e:
             logger.error(f"Failed to decode JSON from log file {self._log_file}: {e}. Creating backup and starting fresh.", exc_info=True)
             try: shutil.move(str(self._log_file), str(self._log_file) + ".corrupt")
             except Exception as move_e: logger.error(f"Could not backup corrupt log file: {move_e}")
             self.installation_log = {}
        except Exception as e:
            logger.error(f"Failed to load installation log from {self._log_file}: {e}", exc_info=True)
            self.installation_log = {} # Start with empty log on other errors

    def _save_installation_log(self):
        """Saves the current installation log to the JSON file."""
        if not self._log_file:
            logger.error("Cannot save installation log: Log file path is not configured.")
            return

        try:
            # Ensure the directory exists
            self._log_file.parent.mkdir(parents=True, exist_ok=True)

            # Write to a temporary file first, then replace atomically
            tmp_log_file = self._log_file.with_suffix(".tmp")
            with open(tmp_log_file, 'w', encoding='utf-8') as f:
                json.dump(self.installation_log, f, indent=2, ensure_ascii=False)

            # Atomic replace (on Windows, os.replace handles this)
            os.replace(tmp_log_file, self._log_file)
            logger.info(f"Successfully saved {len(self.installation_log)} records to installation log: {self._log_file}")

        except Exception as e:
            logger.error(f"Failed to save installation log to {self._log_file}: {e}", exc_info=True)
            # Attempt to clean up temp file if it exists
            if 'tmp_log_file' in locals() and tmp_log_file.exists():
                try:
                    tmp_log_file.unlink()
                    logger.debug("Removed temporary log file after save error.")
                except OSError as unlink_e:
                    logger.warning(f"Could not remove temporary log file {tmp_log_file} after save error: {unlink_e}")


# --- GUI Components ---
class WorkerThread(QThread):
    """Handles background tasks to keep the GUI responsive."""
    task_complete = pyqtSignal(str, object)     # Emits task name and result object
    progress_update = pyqtSignal(str, str)      # Emits task name and progress message

    def __init__(self, installer: ProgramInstaller, task_name: str, *args, **kwargs):
        super().__init__()
        self.installer = installer
        self.task_name = task_name
        self.args = args
        self.kwargs = kwargs
        self._is_running = True # Flag to allow early termination request

    def run(self):
        """Executes the requested task."""
        result = None
        try:
            task_map = {
                "scan": lambda: self.installer.scan_for_installers(),
                "check_status": lambda: self.installer.check_installation_status(self.args[0] if self.args else None),
                "install": lambda: self.installer.install_program(self.args[0], self.kwargs.get('mode', 'auto')),
                "install_heuristic": lambda: self.installer.install_unidentified_program(self.args[0], self.kwargs.get('mode', 'auto')),
                "uninstall": lambda: self.installer.uninstall_program(self.args[0])
            }

            if self.task_name in task_map:
                # Identify the program being worked on for progress messages
                prog_key_or_info = self.args[0] if self.args else None
                prog_name = ""
                if self.task_name == "install":
                    status = self.installer.program_status.get(prog_key_or_info)
                    prog_name = status.display_name if status else str(prog_key_or_info)
                elif self.task_name == "install_heuristic" and isinstance(prog_key_or_info, FoundInstallerInfo):
                    prog_name = f"[Heuristic] {prog_key_or_info.path.name}"
                elif self.task_name == "uninstall":
                    log_entry = self.installer.installation_log.get(prog_key_or_info)
                    prog_name = log_entry.get('name', str(prog_key_or_info)) if log_entry else str(prog_key_or_info)

                self.progress_update.emit(self.task_name, f"Starting {self.task_name}: {prog_name}...")
                logger.info(f"WorkerThread starting task '{self.task_name}' for '{prog_name}'")

                # Execute the task
                result = task_map[self.task_name]()

                if not self._is_running: # Check if stopped during execution
                     logger.info(f"WorkerThread task '{self.task_name}' cancelled.")
                     self.progress_update.emit(self.task_name, f"Task cancelled: {prog_name}")
                     # Don't emit task_complete if cancelled? Or emit with special result?
                     # Let's emit None for cancelled tasks that ran partially.
                     self.task_complete.emit(self.task_name, None)
                     return

                self.progress_update.emit(self.task_name, f"{self.task_name.capitalize()} finished for {prog_name}.")
                logger.info(f"WorkerThread task '{self.task_name}' finished for '{prog_name}'. Result: {result}")

            else:
                errmsg = f"Unknown task requested in WorkerThread: {self.task_name}"
                logger.error(errmsg)
                self.progress_update.emit(self.task_name, errmsg)
                result = None # Indicate failure for unknown task

            # Emit completion signal if still running
            if self._is_running:
                 self.task_complete.emit(self.task_name, result)

        except Exception as e:
            logger.error(f"Error executing task '{self.task_name}' in WorkerThread: {e}", exc_info=True)
            if self._is_running:
                self.progress_update.emit(self.task_name, f"Task failed: {e}")
                self.task_complete.emit(self.task_name, None) # Emit None for result on error

        finally:
            self._is_running = False # Ensure flag is reset

    def stop(self):
        """Requests the thread to stop execution."""
        logger.info(f"Stop requested for WorkerThread task '{self.task_name}'")
        self._is_running = False
        self.progress_update.emit(self.task_name, "Cancellation requested...")
        # Note: Actual interruption depends on the task implementation checking self._is_running

class ProgramInstallerUI(QMainWindow):
    """Main application window."""
    update_tree_signal = pyqtSignal() # Signal to trigger tree refresh from main thread

    COL_PROGRAM = 0
    COL_INSTALLED = 1
    COL_FOUND = 2
    COL_PATH = 3
    COL_VERSION = 4

    def __init__(self, parent=None):
        super().__init__(parent)
        # Ensure Org/App names are set for QSettings
        QApplication.setOrganizationName("InstallerAppOrg")
        QApplication.setApplicationName("ProgramInstallerApp")

        self.installer = ProgramInstaller()
        self.active_threads: List[WorkerThread] = []
        self.pending_updates = False # Flag to coalesce tree updates

        self.setWindowTitle("Engineering Program Installer")
        self.setMinimumSize(900, 650) # Slightly larger default size
        self._setup_ui()
        self._load_settings() # Load geometry, last path etc.

        # Connect the signal to the slot in the main thread
        self.update_tree_signal.connect(self._populate_program_list)

        self._populate_program_list() # Initial population
        self._update_ui_state() # Set initial button/action states

        # Perform initial status check shortly after startup
        QTimer.singleShot(500, self._initial_status_check)

    def _setup_ui(self):
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        main_layout = QVBoxLayout(self.central_widget)

        # --- Toolbar ---
        self.toolbar = self._create_toolbar()
        self.addToolBar(Qt.TopToolBarArea, self.toolbar)

        # --- Main Splitter (Tree View / Details Panel) ---
        splitter = QSplitter(Qt.Vertical)
        main_layout.addWidget(splitter)

        # --- Program Tree View ---
        self.program_tree = QTreeWidget()
        self.program_tree.setColumnCount(5)
        self.program_tree.setHeaderLabels(["Program", "Installed?", "Installer Found?", "Detected Path", "Version (Inst/Found)"])
        hdr = self.program_tree.header()
        hdr.setSectionResizeMode(self.COL_PROGRAM, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(self.COL_INSTALLED, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(self.COL_FOUND, QHeaderView.ResizeToContents)
        hdr.setSectionResizeMode(self.COL_PATH, QHeaderView.Stretch) # Path takes remaining space
        hdr.setSectionResizeMode(self.COL_VERSION, QHeaderView.ResizeToContents)
        self.program_tree.setSelectionMode(QTreeWidget.ExtendedSelection) # Allow multi-select
        self.program_tree.setSortingEnabled(True)
        self.program_tree.sortByColumn(self.COL_PROGRAM, Qt.AscendingOrder) # Default sort
        self.program_tree.itemSelectionChanged.connect(self._update_selection_info)
        self.program_tree.setAlternatingRowColors(True)
        splitter.addWidget(self.program_tree)

        # --- Details Panel (Bottom) ---
        details_panel = QWidget()
        details_layout = QHBoxLayout(details_panel)
        details_layout.setContentsMargins(5, 5, 5, 5)

        # Left side: Information Label
        self.selected_info = QLabel("Select a program from the list above for details.")
        self.selected_info.setWordWrap(True)
        self.selected_info.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        details_layout.addWidget(self.selected_info, 1) # Stretch factor 1

        # Right side: Control Buttons and Install Mode
        control_panel_widget = QWidget() # Use a widget for background styling if needed
        control_panel_layout = QVBoxLayout(control_panel_widget)
        control_panel_layout.setAlignment(Qt.AlignTop)
        control_panel_layout.setSpacing(10)

        # Install Mode ComboBox
        mode_layout = QHBoxLayout()
        mode_label = QLabel("Install Mode:")
        self.install_mode_combo = QComboBox()
        self.install_mode_combo.addItems(["Auto (Silent)", "Semi-Silent", "Manual (Interactive)"])
        self.install_mode_combo.setToolTip("Choose how installers are executed:\n"
                                           " - Auto: Attempts fully silent install using configured switches.\n"
                                           " - Semi-Silent: Attempts install with minimal UI (e.g., progress bar).\n"
                                           " - Manual: Runs the installer interactively.")
        self.install_mode_combo.setCurrentIndex(0) # Default to Auto
        mode_layout.addWidget(mode_label)
        mode_layout.addWidget(self.install_mode_combo)
        control_panel_layout.addLayout(mode_layout)

        # Install Button
        self.install_button = QPushButton(self._get_icon(QStyle.SP_DialogApplyButton), " Install Selected")
        self.install_button.clicked.connect(self._install_selected)
        self.install_button.setToolTip("Install selected program(s) that are not yet installed and have a found installer.")
        self.install_button.setEnabled(False) # Initially disabled
        control_panel_layout.addWidget(self.install_button)

        # Uninstall Button
        self.uninstall_button = QPushButton(self._get_icon(QStyle.SP_TrashIcon), " Uninstall Selected")
        self.uninstall_button.clicked.connect(self._uninstall_selected)
        self.uninstall_button.setToolTip("Attempt to uninstall selected program(s) using logged information.\nOnly works for programs installed via this tool or with discoverable uninstall strings.")
        self.uninstall_button.setEnabled(False) # Initially disabled
        control_panel_layout.addWidget(self.uninstall_button)

        control_panel_layout.addStretch() # Push controls to the top
        details_layout.addWidget(control_panel_widget) # Add the control panel to the right

        # Add details panel to splitter
        splitter.addWidget(details_panel)
        # Adjust initial splitter sizes (adjust ratio as needed)
        splitter.setSizes([int(self.height() * 0.7), int(self.height() * 0.3)])

        # --- Status Bar ---
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready. Set scan path to begin.")

    def _get_icon(self, standard_icon: QStyle.StandardPixmap) -> QIcon:
        """Helper to get standard Qt icons safely."""
        app_instance = QApplication.instance()
        if app_instance:
            return app_instance.style().standardIcon(standard_icon)
        else:
            return QIcon() # Return empty icon if no app instance

    def _create_toolbar(self) -> QToolBar:
        """Creates the main application toolbar."""
        toolbar = QToolBar("Main Toolbar")
        toolbar.setMovable(False)
        toolbar.setToolButtonStyle(Qt.ToolButtonTextBesideIcon)
        toolbar.setIconSize(QSize(16, 16)) # Standard icon size

        # Set Scan Path Action
        setp_action = toolbar.addAction(self._get_icon(QStyle.SP_DirOpenIcon), "Set Scan Path...")
        setp_action.triggered.connect(self._set_search_path)
        setp_action.setToolTip("Select the root directory containing program installers to scan.")

        # Scan Action
        self.scan_action = toolbar.addAction(self._get_icon(QStyle.SP_BrowserReload), "Scan")
        self.scan_action.triggered.connect(self._scan_programs)
        self.scan_action.setToolTip("Scan the selected path for installers and match against configurations.")
        self.scan_action.setEnabled(False) # Disabled until path is set

        # Check Status Action
        self.check_status_action = toolbar.addAction(self._get_icon(QStyle.SP_DriveNetIcon), "Check Status")
        self.check_status_action.triggered.connect(self._check_status)
        self.check_status_action.setToolTip("Check the installation status of all configured programs via the registry.")

        toolbar.addSeparator()

        # Filter Input
        toolbar.addWidget(QLabel(" Filter List:"))
        self.filter_input = QLineEdit()
        self.filter_input.setPlaceholderText("Type to filter programs...")
        self.filter_input.textChanged.connect(self._filter_program_list)
        self.filter_input.setClearButtonEnabled(True)
        self.filter_input.setToolTip("Filter the list by program name, path, or version.")
        toolbar.addWidget(self.filter_input)

        return toolbar

    def _set_search_path(self):
        """Opens a dialog to select the installer search directory."""
        settings = QSettings()
        # Start dialog in last used path or user's home directory
        start_dir = self.installer.search_path or settings.value("last_search_path", str(Path.home()))
        path = QFileDialog.getExistingDirectory(self, "Select Installer Scan Directory", start_dir)

        if path:
            logger.info(f"User selected path: {path}")
            if self.installer.set_search_path(path):
                self.status_bar.showMessage(f"Scan path set: {path}", 5000)
                settings.setValue("last_search_path", path) # Save for next time
                # Clear the tree and update UI state after path change
                self._clear_program_list_results()
                self._update_ui_state()
                # Optionally trigger an automatic scan after setting path?
                # self._scan_programs()
            else:
                QMessageBox.warning(self, "Invalid Path", f"The selected path '{path}' could not be used. Please ensure it exists and is accessible.")
                self._update_ui_state() # Update state even on failure

    def _clear_program_list_results(self):
         """ Clears scan/status results from the tree, keeping config entries """
         self.program_tree.setSortingEnabled(False)
         for i in range(self.program_tree.topLevelItemCount()):
              item = self.program_tree.topLevelItem(i)
              data = item.data(0, Qt.UserRole)
              if data and data['type'] == 'config':
                   item.setText(self.COL_INSTALLED, "?")
                   item.setText(self.COL_FOUND, "No")
                   item.setText(self.COL_PATH, "-")
                   item.setText(self.COL_VERSION, "-")
                   item.setForeground(self.COL_INSTALLED, QColor("orange"))
                   item.setForeground(self.COL_FOUND, QColor("gray"))
                   item.setFont(self.COL_PROGRAM, QFont()) # Reset font
                   # Reset tooltips?
              elif data and data['type'] == 'heuristic':
                   item.setHidden(True) # Hide old heuristic items - they will be repopulated by scan

         self.program_tree.setSortingEnabled(True)
         self._update_selection_info()


    def _scan_programs(self):
        """Starts the background task to scan for installers."""
        if not self.installer.search_path:
            QMessageBox.warning(self, "Scan Path Not Set", "Please set the directory to scan for installers first using 'Set Scan Path...'.")
            return
        if self._is_task_running("scan"):
            self.status_bar.showMessage("Scan is already in progress.", 3000)
            return
        self._run_task("scan")

    def _check_status(self):
        """Starts the background task to check installation status."""
        if self._is_task_running("check_status"):
             self.status_bar.showMessage("Status check is already in progress.", 3000)
             return
        # Check status for all configured programs
        self._run_task("check_status", list(self.installer.program_status.keys()))

    def _initial_status_check(self):
        """Performs an initial status check on startup if path is loaded."""
        if self.installer.search_path:
            logger.info("Performing initial installation status check...")
            self._check_status()
        else:
             logger.info("Skipping initial status check as no search path is loaded.")


    def _install_selected(self):
        """Starts the background task(s) to install selected items."""
        selected_data = self._get_selected_item_data()
        tasks_to_run: List[Tuple[str, Any]] = [] # List of (task_type, data)

        for data in selected_data:
            item_type = data.get("type")
            if item_type == "config":
                key = data["key"]
                status = self.installer.program_status.get(key)
                # Check if installable: config exists, installer found, not already installed
                if status and status.found_installer and status.is_installed is False:
                    tasks_to_run.append(("install", key))
                else:
                     logger.debug(f"Skipping install for config '{key}': Not installable (Found: {status.found_installer is not None}, Installed: {status.is_installed})")
            elif item_type == "heuristic":
                info = data.get("info")
                if info:
                     tasks_to_run.append(("install_heuristic", info))

        if not tasks_to_run:
            QMessageBox.information(self, "Nothing to Install", "No installable programs were selected. Select programs that are 'Not Installed' but have 'Installer Found'.")
            return

        # Get selected installation mode
        mode_text = self.install_mode_combo.currentText()
        install_mode = 'auto' # Default
        if "Semi-Silent" in mode_text: install_mode = 'semi'
        elif "Manual" in mode_text: install_mode = 'manual'
        logger.info(f"Selected install mode: {install_mode} ('{mode_text}')")

        # Confirmation dialog
        names_to_install = []
        for task_type, data in tasks_to_run:
            if task_type == "install": names_to_install.append(self.installer.program_status[data].display_name)
            elif task_type == "install_heuristic": names_to_install.append(f"[Heuristic] {data.path.name}")

        confirm_msg = f"Are you sure you want to attempt installation for {len(tasks_to_run)} item(s) using '{install_mode}' mode?\n\n"
        confirm_msg += f"Items:\n - {', '.join(names_to_install[:5])}"
        if len(names_to_install) > 5: confirm_msg += "\n - ..."
        if any(t == "install_heuristic" for t, d in tasks_to_run):
            confirm_msg += "\n\nNote: Installation commands for heuristic items are best guesses and may fail or require interaction."

        reply = QMessageBox.question(self, "Confirm Installation", confirm_msg, QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)

        if reply == QMessageBox.Yes:
            logger.info(f"User confirmed installation for {len(tasks_to_run)} items in '{install_mode}' mode.")
            for task_type, data in tasks_to_run:
                # Pass the selected mode to the worker task
                self._run_task(task_type, data, mode=install_mode)
        else:
             logger.info("User cancelled installation.")


    def _uninstall_selected(self):
        """Starts the background task(s) to uninstall selected items."""
        selected_data = self._get_selected_item_data()
        tasks_to_run: List[str] = [] # List of program keys to uninstall
        names_to_uninstall: List[str] = []

        for data in selected_data:
            item_type = data.get("type")
            if item_type == "config":
                key = data["key"]
                # Check if uninstallable: must be in log and have an uninstall string
                log_entry = self.installer.installation_log.get(key)
                if log_entry and log_entry.get('uninstall_string'):
                    tasks_to_run.append(key)
                    names_to_uninstall.append(log_entry.get('name', key))
                elif self.installer.program_status.get(key, ProgramStatus(key,"",{})).is_installed:
                     # If installed but not in log, mention it might fail
                     logger.warning(f"Config item '{key}' selected for uninstall, but not found in log or no uninstall string. Attempting registry lookup (may fail).")
                     # Allow attempting uninstall even if not logged, _uninstall_program handles registry lookup
                     tasks_to_run.append(key)
                     names_to_uninstall.append(self.installer.program_status[key].display_name + " (Not Logged)")

            elif item_type == "heuristic":
                QMessageBox.warning(self, "Cannot Uninstall Heuristic Item", "Uninstallation is not supported for heuristically identified items as their installation wasn't tracked.")
                return # Stop if any heuristic item is selected

        if not tasks_to_run:
            QMessageBox.warning(self, "Nothing to Uninstall", "No selected programs have recorded uninstall information or are currently installed.\nUninstallation only works reliably for programs installed via this tool or with discoverable registry entries.")
            return

        # Confirmation dialog
        confirm_msg = f"Are you sure you want to attempt UNINSTALLATION for {len(tasks_to_run)} program(s)?\n\n"
        confirm_msg += f"Items:\n - {', '.join(names_to_uninstall[:5])}"
        if len(names_to_uninstall) > 5: confirm_msg += "\n - ..."
        confirm_msg += "\n\nWarning: This will attempt to run the uninstaller silently. Ensure no critical work is running."
        if any("(Not Logged)" in name for name in names_to_uninstall):
             confirm_msg += "\nSome selected items were not found in the installation log; uninstallation success is less likely."


        reply = QMessageBox.question(self, "Confirm Uninstallation", confirm_msg, QMessageBox.Yes | QMessageBox.No, QMessageBox.No) # Default No

        if reply == QMessageBox.Yes:
            logger.info(f"User confirmed uninstallation for {len(tasks_to_run)} items.")
            for key in tasks_to_run:
                self._run_task("uninstall", key)
        else:
             logger.info("User cancelled uninstallation.")


    def _populate_program_list(self):
        """Refreshes the QTreeWidget with current program status."""
        logger.debug("Populating program list tree...")
        self.program_tree.setSortingEnabled(False) # Disable sorting during update
        self.program_tree.clear() # Clear existing items

        items_to_add: List[QTreeWidgetItem] = []
        status_dict = self.installer.get_current_status()

        # Define colors and fonts for clarity
        colors = {
            'installed': QColor("darkGreen"),
            'not_installed': QColor("red"),
            'found': QColor("darkBlue"),
            'not_found': QColor("gray"),
            'unknown': QColor("orange"),
            'heuristic': QColor("purple"),
            'error': QColor("darkRed"),
        }
        font_normal = QFont()
        font_bold = QFont(); font_bold.setBold(True)
        font_italic = QFont(); font_italic.setItalic(True)

        # 1. Add Configured Programs
        for key, status in status_dict.items():
            # Determine status strings and colors
            if status.is_installed is True:
                inst_text, inst_color = "Yes", colors['installed']
            elif status.is_installed is False:
                inst_text, inst_color = "No", colors['not_installed']
            else: # None (unknown/unchecked)
                inst_text, inst_color = "?", colors['unknown']

            installer_found = status.found_installer is not None
            found_text = "Yes" if installer_found else "No"
            found_color = colors['found'] if installer_found else colors['not_found']

            path_str = str(status.found_installer.path) if installer_found else "-"
            path_tooltip = path_str

            # Determine Version String
            version_str = "-"
            version_tooltip = ""
            if status.last_installed_version:
                version_str = f"Inst: {status.last_installed_version}"
                version_tooltip = f"Installed Version (from check): {status.last_installed_version}"
            elif installer_found:
                props = status.found_installer.file_properties
                pv = props.get('ProductVersion') or props.get('FileVersion') # Prefer ProductVersion
                if pv:
                     version_str = f"Found: {pv}"
                     version_tooltip = f"Installer Version (File: {props.get('FileVersion', 'N/A')}, Prod: {props.get('ProductVersion', 'N/A')})"
                else:
                     version_str = "Found: ?"
                     version_tooltip = "Installer found, but version property missing."

            # Add install error hint
            if status.install_error:
                 inst_text += " (Error)"
                 inst_color = colors['error']


            # Create Tree Item
            item = QTreeWidgetItem([status.display_name, inst_text, found_text, path_str, version_str])
            item.setData(0, Qt.UserRole, {"type": "config", "key": key}) # Store key in item data

            # Apply Colors
            item.setForeground(self.COL_INSTALLED, inst_color)
            item.setForeground(self.COL_FOUND, found_color)

            # Apply Tooltips
            item.setToolTip(self.COL_PROGRAM, f"Configured Program: {status.display_name}\nKey: {key}")
            item.setToolTip(self.COL_INSTALLED, f"Checked: {status.last_checked.strftime('%Y-%m-%d %H:%M:%S') if status.last_checked else 'Never'}\nError: {status.install_error or 'None'}")
            item.setToolTip(self.COL_FOUND, f"Installer file found by scan.")
            item.setToolTip(self.COL_PATH, path_tooltip)
            item.setToolTip(self.COL_VERSION, version_tooltip)


            # Apply Font Styling based on state
            if status.is_installed:
                item.setForeground(self.COL_PROGRAM, colors['installed'])
                item.setFont(self.COL_PROGRAM, font_normal)
            elif installer_found and status.is_installed is False: # Installable
                item.setForeground(self.COL_PROGRAM, colors['not_installed'])
                item.setFont(self.COL_PROGRAM, font_bold) # Bold if ready to install
            else: # Not installed, installer not found OR status unknown
                item.setForeground(self.COL_PROGRAM, colors['not_found'] if status.is_installed is False else colors['unknown'])
                item.setFont(self.COL_PROGRAM, font_normal)

            items_to_add.append(item)

        # 2. Add Heuristically Found Programs
        for info in self.installer.unidentified_installers:
            props = info.file_properties
            # Construct a display name for heuristic items
            heur_name = "[SUS] " # Prefix to indicate suspicious/heuristic
            heur_name += props.get('ProductName') or props.get('FileDescription') or Path(info.path.name).stem
            if not props.get('ProductName') and not props.get('FileDescription'):
                 heur_name += " (No Name Prop)"

            inst_text, inst_color = "?", colors['unknown'] # Installation status is unknown
            found_text, found_color = "Heuristic", colors['heuristic']
            path_str = str(info.path)
            path_tooltip = path_str

            # Version from file properties
            pv = props.get('ProductVersion') or props.get('FileVersion')
            version_str = f"Found: {pv}" if pv else "Found: ?"
            version_tooltip = f"Heuristic Installer Version (File: {props.get('FileVersion', 'N/A')}, Prod: {props.get('ProductVersion', 'N/A')})"

            item = QTreeWidgetItem([heur_name, inst_text, found_text, path_str, version_str])
            item.setData(0, Qt.UserRole, {"type": "heuristic", "info": info}) # Store FoundInstallerInfo

            # Apply Colors & Font
            item.setForeground(self.COL_PROGRAM, colors['heuristic'])
            item.setForeground(self.COL_INSTALLED, inst_color)
            item.setForeground(self.COL_FOUND, found_color)
            item.setFont(self.COL_PROGRAM, font_italic) # Italic for heuristic items

             # Apply Tooltips
            score = self.installer._score_potential_installer(info)
            item.setToolTip(self.COL_PROGRAM, f"Heuristically Identified Potential Installer\nScore: {score:.2f}\nPath: {path_str}")
            item.setToolTip(self.COL_INSTALLED, "Installation status unknown for heuristic items.")
            item.setToolTip(self.COL_FOUND, f"Identified by heuristic rules (Score: {score:.2f}). Not matched to specific configuration.")
            item.setToolTip(self.COL_PATH, path_tooltip)
            item.setToolTip(self.COL_VERSION, version_tooltip)

            items_to_add.append(item)

        # Add all items to the tree
        self.program_tree.addTopLevelItems(items_to_add)

        # Resize columns after adding items (except the stretch column)
        for i in range(self.program_tree.columnCount()):
            if i != self.COL_PATH: # Don't resize the stretch column
                self.program_tree.resizeColumnToContents(i)

        self.program_tree.setSortingEnabled(True) # Re-enable sorting
        self._filter_program_list() # Apply current filter
        self._update_selection_info() # Update details panel based on selection (might be empty now)
        logger.debug("Finished populating program list tree.")

    def _filter_program_list(self):
        """Hides/shows tree items based on the filter input text."""
        filter_text = self.filter_input.text().lower().strip()
        logger.debug(f"Filtering list with text: '{filter_text}'")

        for i in range(self.program_tree.topLevelItemCount()):
            item = self.program_tree.topLevelItem(i)
            data = item.data(0, Qt.UserRole) # Get associated data
            matches_filter = True # Assume match by default

            if filter_text: # Only filter if text is entered
                # Gather text content from relevant columns for searching
                text_to_search = [
                    item.text(self.COL_PROGRAM).lower(),      # Program Name
                    item.text(self.COL_PATH).lower(),        # Detected Path
                    item.text(self.COL_VERSION).lower(),     # Version String
                ]
                # Include program key if available
                if data and data['type'] == 'config':
                     text_to_search.append(data['key'].lower())

                # Check if filter text exists in any searchable text
                matches_filter = any(filter_text in txt for txt in text_to_search)

            # Hide item if it doesn't match the filter
            item.setHidden(not matches_filter)


    def _get_selected_item_data(self) -> List[Dict]:
        """Returns the custom data dictionary stored in selected tree items."""
        return [item.data(0, Qt.UserRole) for item in self.program_tree.selectedItems() if item.data(0, Qt.UserRole)]

    def _update_selection_info(self):
        """Updates the bottom details panel based on the currently selected tree item(s)."""
        selected_data = self._get_selected_item_data()
        count = len(selected_data)
        can_install_any = False
        can_uninstall_any = False

        info_html = ""

        if count == 0:
            info_html = "<i>Select an item from the list above for details.</i>"
        elif count == 1:
            # Display detailed info for single selection
            data = selected_data[0]
            item_type = data.get("type")
            info_lines = []

            if item_type == "config":
                key = data["key"]
                status = self.installer.program_status.get(key)
                log_entry = self.installer.installation_log.get(key)

                if status:
                    info_lines.append(f"<b>{status.display_name}</b> (Config Key: {key})")
                    # Installation Status
                    if status.is_installed is True: info_lines.append(f"&nbsp;&nbsp;Status: <font color='darkGreen'>Installed</font> (Version: {status.last_installed_version or '?'})")
                    elif status.is_installed is False: info_lines.append(f"&nbsp;&nbsp;Status: <font color='red'>Not Installed</font>")
                    else: info_lines.append(f"&nbsp;&nbsp;Status: <font color='orange'>Unknown / Not Checked</font>")
                    if status.last_checked: info_lines.append(f"&nbsp;&nbsp;Last Checked: {status.last_checked.strftime('%Y-%m-%d %H:%M')}")
                    if status.install_error: info_lines.append(f"&nbsp;&nbsp;<font color='darkRed'>Last Action Error:</font> {status.install_error}")

                    # Installer Info
                    if status.found_installer:
                        info_lines.append(f"&nbsp;&nbsp;Installer: <font color='darkBlue'>Found</font>")
                        info_lines.append(f"&nbsp;&nbsp;&nbsp;&nbsp;Path: {status.found_installer.path}")
                        props = status.found_installer.file_properties
                        pv = props.get('ProductVersion') or props.get('FileVersion')
                        info_lines.append(f"&nbsp;&nbsp;&nbsp;&nbsp;Installer Ver: {pv or 'N/A'}")
                        # Check if installable
                        if status.is_installed is False: can_install_any = True
                    else:
                        info_lines.append(f"&nbsp;&nbsp;Installer: <font color='gray'>Not Found</font>")

                    # Uninstall Info
                    if log_entry and log_entry.get('uninstall_string'):
                        info_lines.append(f"&nbsp;&nbsp;Uninstall: Logged")
                        info_lines.append(f"&nbsp;&nbsp;&nbsp;&nbsp;Command: {log_entry.get('uninstall_string')}")
                        info_lines.append(f"&nbsp;&nbsp;&nbsp;&nbsp;Installed On: {log_entry.get('timestamp')}")
                        if status.is_installed: can_uninstall_any = True
                    elif status.is_installed: # Installed but not logged
                        info_lines.append(f"&nbsp;&nbsp;Uninstall: <font color='orange'>Not logged (may attempt registry lookup)</font>")
                        can_uninstall_any = True # Allow attempt even if not logged
                    else:
                        info_lines.append(f"&nbsp;&nbsp;Uninstall: Not logged / Not applicable")

            elif item_type == "heuristic":
                info: FoundInstallerInfo = data.get("info")
                if info:
                    props = info.file_properties
                    name = props.get('ProductName') or props.get('FileDescription') or Path(info.path.name).stem
                    info_lines.append(f"<b>[Heuristic] {name}</b>")
                    info_lines.append(f"&nbsp;&nbsp;Type: Potential Installer (Heuristic Match)")
                    score = self.installer._score_potential_installer(info)
                    info_lines.append(f"&nbsp;&nbsp;Confidence Score: {score:.2f}")
                    info_lines.append(f"&nbsp;&nbsp;Path: {info.path}")
                    pv = props.get('ProductVersion') or props.get('FileVersion')
                    info_lines.append(f"&nbsp;&nbsp;Version: {pv or 'N/A'}")
                    info_lines.append(f"&nbsp;&nbsp;Company: {props.get('CompanyName') or 'N/A'}")
                    # Heuristic items are always considered installable (but not uninstallable via tool)
                    can_install_any = True

            info_html = "<br>".join(info_lines)

        else: # Multiple items selected
            info_html = f"<b>{count} items selected.</b><br>"
            installable_count = 0
            uninstallable_count = 0
            # Check aggregate possibility
            for data in selected_data:
                item_type = data.get("type")
                if item_type == 'config':
                    key = data['key']
                    status = self.installer.program_status.get(key)
                    log = self.installer.installation_log.get(key)
                    if status and status.found_installer and status.is_installed is False:
                        can_install_any = True
                        installable_count += 1
                    if status and status.is_installed and (log and log.get('uninstall_string')):
                         can_uninstall_any = True
                         uninstallable_count +=1
                    elif status and status.is_installed: # Installed but not logged
                         can_uninstall_any = True # Allow attempt
                         # uninstallable_count += 1 # Don't count if not logged? Or count as potential?

                elif item_type == 'heuristic':
                    can_install_any = True
                    installable_count += 1

            info_html += f"&nbsp;&nbsp;Installable: {installable_count}<br>"
            info_html += f"&nbsp;&nbsp;Uninstallable (Logged): {uninstallable_count}"


        self.selected_info.setText(info_html)

        # Enable/disable buttons based on *overall* possibility for selection
        self.install_button.setEnabled(can_install_any)
        self.uninstall_button.setEnabled(can_uninstall_any)

    def _update_ui_state(self):
        """Enables/disables UI elements based on application state (path set, tasks running)."""
        is_idle = not any(t.isRunning() for t in self.active_threads)
        path_is_set = self.installer.search_path is not None

        # Toolbar Actions
        self.scan_action.setEnabled(is_idle and path_is_set)
        self.scan_action.setToolTip("Scan the selected path for installers." if path_is_set else "Set scan path first.")
        self.check_status_action.setEnabled(is_idle)
        self.toolbar.setEnabled(is_idle) # Disable whole toolbar while busy? Maybe too restrictive.

        # Detail Panel Buttons (depend on selection AND idle state)
        # Re-evaluate install/uninstall possibility based on current selection
        self._update_selection_info()
        # Further disable if not idle
        if not is_idle:
             self.install_button.setEnabled(False)
             self.uninstall_button.setEnabled(False)

        # Update status bar message
        if not is_idle:
            running_tasks = [t.task_name for t in self.active_threads if t.isRunning()]
            self.status_bar.showMessage(f"Busy: {', '.join(running_tasks)}...")
        elif not path_is_set:
             self.status_bar.showMessage("Ready. Set scan path to begin.")
        else:
             self.status_bar.showMessage("Ready.")


    def _is_task_running(self, task_name_prefix: str) -> bool:
        """Checks if any active thread matches the task name prefix."""
        return any(t.isRunning() and t.task_name.startswith(task_name_prefix) for t in self.active_threads)

    def _run_task(self, task_name: str, *args, **kwargs):
        """Creates and starts a WorkerThread for a given task."""
        # Prevent duplicate scans or status checks from starting if already running
        # Allow multiple install/uninstall tasks concurrently
        if task_name in ["scan", "check_status"] and self._is_task_running(task_name):
            self.status_bar.showMessage(f"{task_name.capitalize()} is already running.", 3000)
            logger.warning(f"Task '{task_name}' requested but already running.")
            return

        logger.info(f"Starting worker thread for task: '{task_name}'")
        self._set_actions_enabled(False) # Disable actions while task runs

        thread = WorkerThread(self.installer, task_name, *args, **kwargs)
        # Connect signals
        thread.task_complete.connect(self._on_task_complete)
        thread.progress_update.connect(self._on_progress_update)
        thread.finished.connect(lambda th=thread: self._thread_finished(th)) # Ensure correct thread ref

        self.active_threads.append(thread)
        thread.start()
        self._update_ui_state() # Update UI to reflect busy state

    def _set_actions_enabled(self, enabled: bool):
        """Enable/disable main actions based on whether tasks are running."""
        path_is_set = self.installer.search_path is not None
        self.scan_action.setEnabled(enabled and path_is_set)
        self.check_status_action.setEnabled(enabled)

        # Enable/disable install/uninstall based on idle state AND selection state
        if enabled:
            self._update_selection_info() # Recalculate button states based on selection
        else:
            self.install_button.setEnabled(False)
            self.uninstall_button.setEnabled(False)

        # Also enable/disable the toolbar widgets like the filter input?
        self.toolbar.setEnabled(enabled) # Simple approach: disable whole toolbar


    def _can_install_selected(self) -> bool:
        """Helper to check if any selected item is currently installable."""
        data = self._get_selected_item_data()
        return any(
            (d.get('type') == 'heuristic') or
            (d.get('type') == 'config' and (st := self.installer.program_status.get(d.get('key'))) and st.found_installer and st.is_installed is False)
            for d in data
        )

    def _can_uninstall_selected(self) -> bool:
        """Helper to check if any selected item is currently uninstallable."""
        data = self._get_selected_item_data()
        return any(
             d.get('type') == 'config' and
             (st := self.installer.program_status.get(d.get('key'))) and
             st.is_installed and # Must be installed
             # Either logged with uninstall string OR allow attempt if not logged
             ( (key := d.get('key')) in self.installer.installation_log and self.installer.installation_log[key].get('uninstall_string') or True) # Simpler: allow if installed
            for d in data
        )

    # --- Signal Handlers / Slots ---

    def _on_progress_update(self, task_name: str, message: str):
        """Updates the status bar with progress messages from worker threads."""
        self.status_bar.showMessage(f"[{task_name}] {message}", 5000) # Show for 5 seconds

    def _on_task_complete(self, task_name: str, result: Any):
        """Handles task completion signals from worker threads."""
        logger.info(f"GUI received task_complete signal: Task='{task_name}', Result type='{type(result)}'")

        # Trigger UI update for tasks that change program status or lists
        if task_name in ["scan", "check_status", "install", "install_heuristic", "uninstall"]:
            logger.debug(f"Requesting UI update after task '{task_name}' completion.")
            # Use a timer to coalesce multiple rapid updates
            if not self.pending_updates:
                 self.pending_updates = True
                 QTimer.singleShot(150, self._trigger_tree_update) # Delay slightly (150ms)

        # Show specific messages for success/failure of install/uninstall
        if task_name == "install" or task_name == "install_heuristic":
            prog_name = self.sender().args[0] # Get target from thread args
            if isinstance(prog_name, FoundInstallerInfo): prog_name = prog_name.path.name
            else: prog_name = self.installer.program_status.get(prog_name, ProgramStatus("","",{})).display_name

            if result is True:
                QMessageBox.information(self, "Installation Success", f"Installation completed successfully for '{prog_name}'.")
            elif result is False:
                QMessageBox.warning(self, "Installation Failed", f"Installation failed for '{prog_name}'.\nCheck logs for details.")
            # Handle None result (e.g., error in thread)
            elif result is None:
                 QMessageBox.critical(self, "Installation Error", f"An unexpected error occurred during installation for '{prog_name}'. Check logs.")

        elif task_name == "uninstall":
            prog_key = self.sender().args[0]
            prog_name = self.installer.installation_log.get(prog_key,{}).get('name') or prog_key
            if result is True:
                QMessageBox.information(self, "Uninstallation Success", f"Uninstallation completed successfully for '{prog_name}'.")
            elif result is False:
                QMessageBox.warning(self, "Uninstallation Failed", f"Uninstallation failed for '{prog_name}'.\nMay require manual removal. Check logs for details.")
            elif result is None:
                 QMessageBox.critical(self, "Uninstallation Error", f"An unexpected error occurred during uninstallation for '{prog_name}'. Check logs.")

    def _trigger_tree_update(self):
        """Emits the signal to update the tree (called by timer)."""
        if self.pending_updates: # Check flag again in case it was handled
             logger.debug("Timer triggered: Emitting update_tree_signal.")
             self.update_tree_signal.emit()
             self.pending_updates = False # Reset flag

    def _thread_finished(self, thread: WorkerThread):
        """Cleans up after a worker thread finishes."""
        logger.debug(f"WorkerThread finished signal received for task '{thread.task_name}'.")
        try:
            # Remove the thread from the active list
            if thread in self.active_threads:
                self.active_threads.remove(thread)
            else:
                 logger.warning(f"Finished thread for task '{thread.task_name}' was not found in active list.")
        except Exception as e:
            logger.error(f"Error removing finished thread from active list: {e}", exc_info=True)

        # Update UI state only if *all* threads are now idle
        if not any(t.isRunning() for t in self.active_threads):
            logger.info("All worker threads have finished. Updating UI state to idle.")
            self._update_ui_state()
            self.status_bar.showMessage("Ready.", 3000) # Reset status bar
        else:
            logger.debug(f"{len([t for t in self.active_threads if t.isRunning()])} threads still running. UI remains busy.")
            self._update_ui_state() # Update UI state to reflect current running tasks


    # --- Settings Persistence ---

    def _load_settings(self):
        """Loads UI settings (geometry, last path) from QSettings."""
        self.settings = QSettings() # Uses Org/App name set earlier
        logger.info(f"Loading UI settings from: {self.settings.fileName()}")

        # Restore window geometry and state
        geometry = self.settings.value("geometry")
        state = self.settings.value("windowState")
        if isinstance(geometry, QByteArray):
            self.restoreGeometry(geometry)
            logger.debug("Restored window geometry.")
        if isinstance(state, QByteArray):
            self.restoreState(state)
            logger.debug("Restored window state.")

        # Restore last search path
        last_path = self.settings.value("last_search_path", "")
        if last_path and isinstance(last_path, str):
            logger.info(f"Attempting to set last used search path: {last_path}")
            # Use the setter to validate the path
            self.installer.set_search_path(last_path)
            # No need to update UI state here, constructor does it later

    def _save_settings(self):
        """Saves UI settings to QSettings."""
        if not hasattr(self, 'settings'):
            self.settings = QSettings()

        logger.info(f"Saving UI settings to: {self.settings.fileName()}")
        # Save geometry and state
        self.settings.setValue("geometry", self.saveGeometry())
        self.settings.setValue("windowState", self.saveState())

        # Save the currently set search path
        if self.installer.search_path:
            self.settings.setValue("last_search_path", self.installer.search_path)
            logger.debug(f"Saved last search path: {self.installer.search_path}")

        self.settings.sync() # Ensure settings are written to storage
        logger.info("UI settings saved.")

    # --- Window Closing ---

    def closeEvent(self, event):
        """Handles the window close event."""
        logger.info("Close event triggered. Saving settings and checking for running tasks.")
        self._save_settings()

        # Check for running background tasks
        running_threads = [t for t in self.active_threads if t.isRunning()]
        if running_threads:
            task_names = [t.task_name for t in running_threads]
            logger.warning(f"Attempting to close window, but background tasks are running: {task_names}")
            reply = QMessageBox.question(self, "Tasks Running",
                                         f"The following tasks are still running:\n - {', '.join(task_names)}\n\n"
                                         "Closing the application now may interrupt them. Are you sure you want to quit?",
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

            if reply == QMessageBox.No:
                logger.info("User cancelled window close due to running tasks.")
                event.ignore() # Prevent window from closing
                return
            else:
                logger.warning("User confirmed closing despite running tasks. Attempting to signal stop...")
                # Try to signal threads to stop (may not be immediate)
                for thread in running_threads:
                    thread.stop()
                # Give a brief moment for signals to process?
                QApplication.processEvents()
                logger.warning("Continuing close process. Tasks may be forcefully terminated.")

        logger.info("Proceeding with window close.")
        super().closeEvent(event)

# --- Main Execution ---
if __name__ == "__main__":
    # Set AppUserModelID for proper taskbar icon grouping on Windows
    try:
        import ctypes
        myappid = 'MyOrganization.MyProduct.ProgramInstaller.1.0' # Choose a unique ID
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        logger.debug(f"AppUserModelID set to: {myappid}")
    except Exception as e:
        # This is not critical, just improves Windows integration
        print(f"Warning: Could not set AppUserModelID: {e}", file=sys.stderr)
        logger.warning(f"Could not set AppUserModelID: {e}")

    # Create the Qt Application
    app = QApplication(sys.argv)

    # Apply a modern style if available (Fusion is generally good cross-platform)
    styles = QStyleFactory.keys()
    if 'Fusion' in styles:
        app.setStyle(QStyleFactory.create('Fusion'))
        logger.debug("Applied 'Fusion' style.")
    elif 'WindowsVista' in styles: # Fallback for Windows
        app.setStyle(QStyleFactory.create('WindowsVista'))
        logger.debug("Applied 'WindowsVista' style.")
    else:
         logger.debug(f"Default style '{app.style().objectName()}' will be used. Available: {styles}")


    # Create and show the main window
    try:
        window = ProgramInstallerUI()
        window.show()
        logger.info("Application startup successful. Entering main event loop.")
        sys.exit(app.exec_()) # Start the Qt event loop
    except Exception as e:
         # Catch critical startup errors
         logger.critical(f"Application failed to start: {e}", exc_info=True)
         # Try to show a message box even if the main UI failed
         try:
              msg = QMessageBox()
              msg.setIcon(QMessageBox.Critical)
              msg.setWindowTitle("Fatal Startup Error")
              msg.setText("The application failed to initialize.")
              msg.setInformativeText(f"Error: {e}\n\nPlease check the console output or logs for more details.")
              msg.exec_()
         except Exception as msg_e:
              # Fallback to console if even QMessageBox fails
              print(f"FATAL APPLICATION ERROR: {e}\nCannot show error message box: {msg_e}", file=sys.stderr)
         sys.exit(1) # Exit with error code