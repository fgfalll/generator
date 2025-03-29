# Desktop Organizer v4.2 (Авто-організатор робочого столу)

![Desktop Organizer Screenshot]()
*Automated desktop file management with modular extensions, configurable timer, drive selection, settings persistence, and dynamic module loading.*

## Table of Contents

*   [Features](#features)
*   [Installation](#installation)
    *   [Requirements](#requirements)
    *   [Optional Modules](#optional-modules)
*   [Usage](#usage)
    *   [Interface Overview](#interface-overview)
    *   [Workflow](#workflow)
*   [Settings Configuration](#settings-configuration)
    *   [YAML File (`config.yaml`)](#yaml-file-configyaml)
    *   [Settings Dialog](#settings-dialog)
    *   [File Mover Filters (Configuration Only)](#file-mover-filters-configuration-only)
*   [Module Management](#module-management)
    *   [Module Concept](#module-concept)
    *   [Module Locations & Priority](#module-locations--priority)
    *   [Importing Modules via UI](#importing-modules-via-ui)
    *   [Developer Note: Module Structure](#developer-note-module-structure)
*   [Platform Behavior](#platform-behavior)
*   [Troubleshooting](#troubleshooting)
*   [Future Work / Placeholders](#future-work--placeholders)
*   [License](#license)
*   [Documentation Information](#documentation-information)

## Features

✅ **Core Functionality**
*   🧹 **Automated Desktop Cleaning:** Moves files and folders from the Desktop to organized, timestamped folders.
*   ⏱️ **Timer Control:** Configurable countdown timer for automatic operation (1-60 mins). Manual start/stop controls.
*   ⚙️ **Manual Start:** Option to trigger the cleaning process immediately.
*   🧵 **Background Processing:** File moving occurs in a separate thread (`FileMover`) to keep the UI responsive.

✅ **Configuration & Flexibility**
*   💾 **Persistent Settings:** Application settings (including module paths) saved to `config.yaml` and loaded on startup.
*   🔧 **Settings Dialog:** Dedicated window (File -> Settings...) to manage application behavior:
    *   **Timer:** Override default duration, enable/disable timer autostart on launch.
    *   **Drives:** Choose main target drive policy (Prefer D:, Auto-detect next available). Fallback to C: is automatic.
    *   **File Filters:** Configure maximum file size, extensions to skip, and filenames to skip *(Note: Filter application is not yet implemented in the move process)*.
    *   **Modules:** Specify an optional external directory to load modules from (takes priority).
*   🚗 **Dynamic Drive Selection:** Automatically determines the target drive based on settings and availability (D:, E:, Auto-detected, C: fallback). Visual indicators (🟢/🔴) show drive status. Manual override via buttons.

✅ **Modularity & Extensibility (New in v4.2)**
*   🔌 **Dynamic Module Loading:** Optional features (like License Manager, Install Programs) are loaded at runtime from `.py` files.
*   📁 **Standard Module Directory:** Looks for modules in a `modules` subdirectory located next to the main application executable (or script).
*   🌍 **External Module Path:** Configure an additional directory in Settings to load modules from. Modules here override those in the standard directory.
*   📥 **Module Import:** Easily copy compatible `.py` module files into the standard `modules` directory via "File -> Import Module...".
*   🔄 **Automatic Reloading:** Application detects newly imported modules or changes to the external module path in settings and attempts to reload available modules.
*   🚦 **Conditional UI:** Menu items for optional modules (e.g., "Manage Licenses...", "Install Programs") are enabled *only if* the corresponding module is found and loaded successfully.

✅ **Integration & UI**
*   📑 **Menu Bar:** Provides access to Settings, Module Import, License Manager (if loaded), Install Programs (if loaded), and Exit.
*   📄 **License Manager Integration:** Opens the License Manager module (v2.0+) via the "License" menu *if* `license_manager.py` is loaded.
*   ➕ **Install Programs Integration:** Opens the program installer module via the "Install Programs" menu *if* `program_install.py` (or similarly named module) is loaded.
*   📊 **Detailed Logging:** Logs actions, progress, errors, settings changes, and *module loading status* with timestamps in a dedicated panel.
*   ✨ **Modern UI:** Built with PyQt5, includes visual cues for drive status and selection.

## Installation

### Requirements
*   Python 3.8+
*   PyQt5
*   PyYAML (for settings)
*   psutil (for drive detection)

```bash
# Install core dependencies
pip install pyqt5 pyyaml psutil
```

### Optional Modules

*   The core Desktop Organizer can run without any optional modules.
*   To enable features like the **License Manager** or **Install Programs**, place the corresponding `.py` files (e.g., `license_manager.py`, `program_install.py`) into one of the recognized module locations:
    1.  The **External Module Directory** specified in the application's Settings (if configured).
    2.  The **Standard `modules` Directory**, which is a folder named `modules` located in the same directory as the main application executable (`DesktopOrganizer.exe`) or script (`v4.x.py`).
*   Modules can also be added using the "File -> Import Module..." feature.

### Platform Notes
*   The application is primarily developed and tested on **Windows** due to its reliance on drive letters (D:, E:, C:).
*   Core file moving logic, settings management, and module loading should work on **macOS** or **Linux**, but drive auto-detection and target paths might require adjustments.
*   **Optional modules themselves** might have specific platform dependencies (e.g., License Manager using Windows Registry).

## Usage

### Interface Overview

1.  **Menu Bar:**
    *   **Файл (File):** Access `Налаштування... (Settings...)`, `Import Module to Standard Folder...` (New), and `Вихід (Exit)`.
    *   **Install Programs:** Access `Install New Program...` (*Enabled only if the corresponding module is loaded*).
    *   **License:** Access `Manage Licenses...` (*Enabled only if the License Manager module is loaded*).
2.  **Timer Label:** Displays the countdown or status.
3.  **Timer Controls:** (`QComboBox`, `🚀 Старт зараз`, `⏱️ Стоп/▶️ Старт таймер`).
4.  **Drive Selection Group:** (`Диск D:`, `Диск E:`) with availability indicators (🟢/🔴).
5.  **Log Area (`QTextEdit`):** Displays timestamped messages, including module loading successes and failures.

### Workflow

```mermaid
graph TD
    A[App Start] --> B{"Load Settings (config.yaml)"}
    B --> C[Determine Module Paths (Settings + Default)]
    C --> D{Load Optional Modules}
    D -- Success/Fail --> E[Update UI (Enable/Disable Module Menus)]
    E --> F{"Determine Target Drive (Settings + Availability)"}
    F --> G[Apply Other UI Settings (Timer, etc.)]
    G --> H{"Autostart Timer? (Setting)"}
    H -- Yes --> I[Start Countdown Timer]
    H -- No --> J[Timer Disabled]

    subgraph "User Interaction"
        direction LR
        U1[Select Time] --> UTimer
        U2[Select Drive] --> UDrive
        U3[Toggle Timer] --> UTimer
        U4[Start Now] --> UManualStart
        U5[Open Settings] --> USettings
        U6[Open Loaded Module (e.g., License Mgr)] --> UModuleWin
        U7[Import Module] --> UImport
    end

    I --> T{Timer Ends}
    UTimer --> I
    UTimer --> J
    UDrive --> J
    UManualStart --> K[Start FileMover Thread]

    T --> K
    K --> L["Move Desktop Items (Skip Filters*)"]
    L --> M[Log Results]
    M --> N{"Autostart Timer? (Setting)"}
    N -- Yes --> I
    N -- No --> J

    USettings -- Apply/OK --> SaveSettings[Save Settings to config.yaml]
    SaveSettings --> R[Reload Modules & Update UI]
    R --> F

    UImport --> CopyFiles[Copy .py to ./modules]
    CopyFiles --> R
```

1.  App starts, loads settings (`config.yaml`), determines module search paths.
2.  Attempts to load optional modules from External and Standard locations. Logs success/failure.
3.  UI is updated; menu items for loaded modules are enabled.
4.  Target drive determined, timer configured based on settings.
5.  User interacts (changes timer, starts manually, selects drive, opens loaded modules).
6.  **Settings Change:** If settings (especially External Module Path) are applied, modules are reloaded, and UI is updated.
7.  **Import Module:** User selects `.py` files, they are copied to `./modules`, modules are reloaded, and UI is updated.
8.  When timer ends or "Start Now" is clicked, `FileMover` runs (UI disabled).
9.  File moving occurs (filters still limited), results logged (UI re-enabled).
10. Timer restarts if configured.

## Settings Configuration

### YAML File (`config.yaml`)

Settings are stored in YAML format at:
*   **Windows:** `C:\Users\<YourUser>\.DesktopOrganizer\config.yaml`
*   **macOS/Linux:** `~/.DesktopOrganizer/config.yaml`

**Example `config.yaml` Structure (v4.2):**

```yaml
application:
  autostart_timer_enabled: true
  external_module_path: '' # NEW: Optional path to external modules directory (e.g., C:\MyOrgModules)
timer:
  override_default_enabled: false
  default_minutes: 3
drives:
  main_drive_policy: D
file_manager:
  max_file_size_mb: 100
  allowed_extensions:
  - .lnk
  allowed_filenames: []
```

### Settings Dialog

Accessed via **File (Файл) -> Settings... (Налаштування...)**.

*   **Tabs:**
    *   **General:**
        *   `Application Behavior`: Enable/disable timer autostart. **New:** `External Modules Path` field and `Browse...` button to select a directory where modules will be searched for *first*.
        *   `Timer Settings`: Override default timer duration.
        *   `Drive Settings`: Select main drive policy.
    *   **File Filters:** Configure file size, skip extensions, skip filenames *(Filters mostly not yet active in v4.2)*.
*   **Buttons:** `OK`, `Cancel`, `Apply`. Clicking `OK` or `Apply` saves settings and triggers a reload of modules if the external path changed.

### File Mover Filters (Configuration Only)

⚠️ **Important:** While the Settings dialog allows configuring "Max File Size", "Skip Extensions", and "Skip Filenames", the actual logic to *use* these filters (except for the default `.lnk` skip) within the `FileMover` thread **is not implemented in version 4.2**. These settings are saved but currently only `.lnk` files are skipped by the move process.

## Module Management

### Module Concept

Desktop Organizer v4.2 allows extending its functionality through optional Python modules (`.py` files). Features like the License Manager are now separate modules that the main application loads dynamically at startup or when settings change. This keeps the core application lighter and allows modules to be updated independently.

### Module Locations & Priority

The application searches for required module files (defined internally, e.g., `license_manager.py`) in the following locations, in order of priority:

1.  **External Module Directory:** The path specified in `File -> Settings... -> General -> External Modules Path`. If a valid directory is set, the application looks here first.
2.  **Standard Module Directory:** A folder named `modules` located directly next to the main application executable (e.g., `DesktopOrganizer.exe`) or the main script (`v4.2.py`). This is the default location.

If a module (e.g., `license_manager.py`) is found in the *External* directory, the application loads that version and **does not** look for it in the *Standard* directory.

### Importing Modules via UI

You can easily add or update modules in the **Standard Module Directory** (`./modules`):

1.  Go to **File -> Import Module to Standard Folder...**
2.  A file dialog opens. Select one or more `.py` module files you want to import.
3.  Click "Open".
4.  If a module with the same name already exists in the `./modules` folder, you will be asked to confirm overwriting it.
5.  The selected files are copied into the `./modules` folder.
6.  The application automatically re-scans for modules and updates the UI (enabling menu items for newly available modules). Check the log panel for details.

### Developer Note: Module Structure

For a `.py` file to be loaded successfully as a module by Desktop Organizer:
*   It must contain a specific class expected by the main application (e.g., `LicenseManager` for the license module). This class name is defined internally in the main app's configuration.
*   The class's `__init__` method should accept an optional `parent` argument (`def __init__(self, parent=None):`). The main application window instance will be passed as the parent, allowing the module to interact with it (e.g., using its logging function or accessing shared resources if necessary).

## Platform Behavior

*   **Drive Selection:** Primarily designed for Windows drive letters. Auto-detect uses `psutil`.
*   **Module Loading:** Works cross-platform.
*   **Module Functionality:** Individual modules (like License Manager) may have their own platform-specific behavior (e.g., registry use on Windows).

## Troubleshooting

1.  **Target Drive Not Found / Errors Writing:** (Same as v4.1) Check drive existence, permissions.
2.  **Timer Doesn't Start Automatically:** (Same as v4.1) Check settings.
3.  **Settings Not Saving:** (Same as v4.1) Check user profile permissions, logs.
4.  **Module Menu Item Disabled (e.g., "Manage Licenses..." is greyed out):**
    *   **Module file missing:** Ensure the required `.py` file (e.g., `license_manager.py`) exists in either the configured External Module Directory or the standard `./modules` directory.
    *   **Module load error:** Check the application's log panel for error messages during startup or after importing/changing settings. Errors could be due to syntax errors in the module, missing dependencies *within* the module, or incorrect structure (e.g., missing the expected class).
    *   **Incorrect module filename:** The main application expects specific filenames for specific features.
    *   **External Path Incorrect:** If using the external path setting, ensure the path is correct and the application has permission to read from it.
5.  **Module Import Fails:**
    *   **Permissions:** Ensure the application has write permissions to the standard `./modules` directory (located next to the executable/script).
    *   Check the log panel for specific copy errors.
6.  **Files Not Skipped Based on Filters:** As noted, size/extension/filename filters configured in settings are not yet applied by the `FileMover` in v4.2 (except `.lnk`).
7.  **(Resolved) Crash after Module Import (0xC0000409):** This specific crash related to reloading should be resolved in v4.2 due to improved cleanup. If it recurs, it may indicate a severe issue within the imported module itself, especially if it uses C extensions.

## Future Work / Placeholders

*   **Implement File Mover Filters:** Add logic to `FileMover` to respect configured filters.
*   **Implement "Install Programs":** Develop the actual functionality for the `program_install` module.
*   **macOS/Linux Drive Logic:** Improve non-Windows drive/path handling.
*   **Error Handling:** More granular error reporting during file moves and module loading.
*   **Module Interface:** Define a more formal interface or base class for modules.

## License

Proprietary - © 2024 Taras

## Documentation Information

**Version**: 4.2
**Last Updated**: March 26, 2024