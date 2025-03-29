# Windows Program Installer Tool

## Overview

This Python script provides a graphical user interface (GUI) built with PyQt5 to help automate the discovery, installation, and tracking of specific software programs on Windows, with a focus on engineering applications. It scans specified directories for installer files (.exe, .msi), attempts to identify them using predefined configurations and file metadata, checks if the corresponding program is already installed via the Windows Registry, and allows for automated (silent), semi-silent, or manual installation. It also logs successful installations to aid in potential uninstallation later.

## Features

*   **GUI:** User-friendly interface powered by PyQt5.
*   **Installer Scanning:** Scans user-defined directories for `.exe` and `.msi` files.
*   **Configurable Program Definitions:** Uses a Python dictionary (`PROGRAM_CONFIG`) to define expected properties (product names, descriptions, filename patterns) for known software.
*   **Metadata Extraction:** Reads metadata (Product Name, File Description, Version, etc.) from installer files using `pywin32`.
*   **MSI Property Reading:** Extracts detailed properties from MSI files, including ProductCode.
*   **Heuristic Detection:** Attempts to identify potential installers that don't match specific configurations based on file properties, size, and naming conventions.
*   **Installation Status Check:** Queries the Windows Registry (primarily Uninstall keys) to determine if configured programs are already installed and retrieve their version.
*   **Multiple Install Modes:**
    *   **Auto (Silent):** Attempts fully silent installation using predefined command-line switches.
    *   **Semi-Silent:** Attempts installation with minimal UI (e.g., progress bar only) using passive switches.
    *   **Manual (Interactive):** Launches the installer normally for user interaction.
*   **Installation Logging:** Records successfully installed programs (via this tool) including installer path and discovered uninstall command in a JSON file (`%APPDATA%\ProgramInstallerApp\program_installer_log.json`).
*   **Silent Uninstallation:** Uses logged information (primarily the UninstallString or MSI ProductCode) to attempt silent uninstallation.
*   **Filtering & Sorting:** Allows filtering the program list in the GUI and sorting by columns.
*   **Color-Coding:** Visually indicates program status (Installed, Not Installed, Installer Found, Heuristic) in the list.
*   **Dependency Checks:** Verifies required libraries (`PyQt5`, `pywin32`) are available on startup.
*   **Background Operations:** Uses QThreads for long-running tasks (scanning, installing, checking status) to keep the GUI responsive.
*   **UI State Persistence:** Saves and restores window size/position and the last used scan path.
*   **Customizable Filtering:** `DETECTION_SETTINGS` allows tuning the scanner by excluding common dependency names, uninstallers, small files, and specific directories.

## Requirements

*   **Operating System:** Windows (tested on Windows 10/11)
*   **Python:** Python 3.7+
*   **Python Packages:**
    *   `PyQt5`: For the graphical user interface.
    *   `pywin32`: For accessing Windows APIs (file properties, registry, MSI info).

    Install required packages using pip:
    ```bash
    pip install PyQt5 pywin32
    ```

## Configuration

The script requires configuration for the specific software you want to manage.

1.  **`PROGRAM_CONFIG` Dictionary:**
    *   Located near the top of the script.
    *   Each key is a unique identifier for a program (e.g., `"petrel"`).
    *   The value is a dictionary containing:
        *   `display_name`: How the program appears in the GUI.
        *   `target_version`: (Currently informational) Can be "latest" or a specific version string.
        *   `identity`: Dictionary with lists of expected strings/patterns:
            *   `expected_product_names`: Likely "ProductName" property values.
            *   `expected_descriptions`: Likely "FileDescription" property values.
            *   `installer_patterns`: `fnmatch` patterns for installer filenames (e.g., `Petrel*.exe`).
        *   `check_method`: How to check if the program is installed.
            *   `type`: Currently `"registry"`.
            *   `keys`: A list of registry check rules. Each rule is a dictionary specifying:
                *   `path`: Registry key path (e.g., `SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall`).
                *   `hive`: (Optional, defaults to `HKLM`) `HKLM` or `HKCU`.
                *   `match_value`: Registry value name to check within subkeys (e.g., `DisplayName`).
                *   `match_pattern`: Regex pattern to match against `match_value`.
                *   `get_value`: Registry value name to retrieve if matched (e.g., `DisplayVersion`).
                *   `check_existence`: If `True`, simply check if the `path` exists.
        *   `install_commands`: Dictionary mapping installer extension (`.exe`, `.msi`) to the command-line template for silent installation. Use `{installer_path}` as a placeholder for the full installer path (which will be quoted).

    **--> You MUST populate `PROGRAM_CONFIG` with details specific to the engineering software installers you have.** The provided entries are examples and may need significant adjustment for your exact installer versions and types. Finding the correct silent install switches often requires consulting vendor documentation or experimentation (`installer.exe /?` or `/help`).

2.  **`DETECTION_SETTINGS` Dictionary:**
    *   Controls how the script filters potential installers during the scan.
    *   `exclude_generic_names`: Lowercase strings likely found in dependency installers (e.g., 'driver', 'redist').
    *   `exclude_by_property_substrings`: Lowercase strings to filter out based on ProductName/FileDescription (e.g., '.net framework', 'visual c++', 'codemeter').
    *   `exclude_uninstaller_hints`: Lowercase strings indicating an uninstaller (e.g., 'uninstall', 'remove').
    *   `min_file_size_bytes`: Minimum size in bytes for a file to be considered.
    *   `ignore_dirs`: Set of lowercase directory names to skip during scanning (e.g., 'windows', 'temp', 'drivers'). Add project-specific subfolders here if needed (e.g. 'HelpFiles', 'Documentation').

## Usage

1.  **Configure:** Edit `PROGRAM_CONFIG` and `DETECTION_SETTINGS` in the `install_programs.py` script for your specific software and environment.
2.  **Run:** Execute the script from your terminal:
    ```bash
    python install_programs.py
    ```
3.  **Set Scan Path:** Click the "Set Scan Path..." button in the toolbar and select the root directory containing your program installers. The tool will scan recursively, respecting the `ignore_dirs` setting.
4.  **Scan:** Click the "Scan" button. The tool will search for potential installers and match them against the `PROGRAM_CONFIG`. Unmatched files passing heuristic checks will appear as "[Unk]" or "[SuS]" (Suspicious/Unidentified).
5.  **Check Status:** Click "Check Status". The tool will query the registry based on `check_method` rules to update the "Installed" status for configured programs.
6.  **Review:** Examine the list. Check the "Installed" and "Installer Found" columns. Select an item to see more details in the bottom panel. Use the filter box to narrow the list.
7.  **Select & Install/Uninstall:**
    *   Select one or more programs from the list.
    *   Choose the desired installation mode ("Auto", "Semi-Silent", "Manual") from the dropdown.
    *   Click "Install Selected" to install programs where an installer was found and the program is not already installed.
    *   Click "Uninstall Selected" to attempt silent uninstallation for configured programs that were previously installed *by this tool* (and have an uninstall string logged). Uninstallation of heuristic items is not supported.
8.  **Monitor:** Observe the status bar for progress updates and check the console output/log file for detailed information or errors.

## How it Works

*   **Scanning:** Uses `os.walk` to traverse the selected directory, applying filters (`ignore_dirs`, file size, extensions).
*   **Identification:**
    *   For `.exe` and `.msi` files passing initial filters, `win32api.GetFileVersionInfo` (for EXE) or `msilib` (for MSI) is used to extract metadata.
    *   Files are filtered further based on `DETECTION_SETTINGS` (generic names, property substrings, uninstaller hints).
    *   Remaining files are compared against `PROGRAM_CONFIG` identity rules (product name, description, filename patterns).
    *   Files not matching a config entry but passing a heuristic score threshold (`_score_potential_installer`) are listed as unidentified.
*   **Status Check:** Uses `winreg` to access the registry according to `check_method` rules in `PROGRAM_CONFIG`.
*   **Installation:** Uses `subprocess.run` with `start /wait ""` to execute the installer command (potentially with silent switches derived from `PROGRAM_CONFIG` or defaults for heuristics) based on the selected mode. Waits for the installer process to finish. Checks common success exit codes (0, 3010, 1641). For MSI installs, it verifies installation using the Windows Installer COM API if the ProductCode is known.
*   **Uninstallation:** Retrieves the uninstall command from the log file (`installation_log.json`). Attempts to add common silent flags (`/S`, `/qn`, etc.) and executes the command via `subprocess.run`. Verifies MSI uninstallation via COM API if ProductCode was available.
*   **Logging:** Uses Python's `logging` module for runtime output. Stores persistent installation records (program key, name, timestamp, installer path, uninstall string, version) in a JSON file in the user's AppData directory.

## Limitations & Considerations

*   **Silent Switches:** Finding the correct silent installation switches (`install_commands`) is crucial and often requires trial-and-error or vendor documentation. Incorrect switches may lead to interactive prompts or failed installations.
*   **Complex Installers:** Installers requiring multiple steps, prerequisites, or complex configuration might not work well with simple silent commands.
*   **Administrator Privileges:** Installing/uninstalling most software requires administrator privileges. Run the script as an administrator if necessary.
*   **Heuristic Accuracy:** Heuristic detection is not perfect and may identify non-installers or miss actual installers. Manual verification is recommended for heuristic items.
*   **Uninstallation Reliability:** Uninstallation relies on finding a valid `UninstallString` in the registry after installation or using the MSI `ProductCode`. If this fails or the uninstaller itself is flawed, silent uninstall might not work. Uninstalling programs *not* installed by this tool is not directly supported (though the registry check might find them).
*   **Error Handling:** While basic error handling is present, complex installer errors might not be caught gracefully. Check the logs for details.

## License

MIT License (or specify your preferred license)