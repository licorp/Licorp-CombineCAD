# Licorp_CombiCAD

**Export Revit sheets to DWG with multi-layout merge capabilities.**

A Revit add-in for exporting sheets to DWG format with advanced merge features powered by AutoCAD.

## Features

- **Individual Export**: Export each sheet as a separate DWG file
- **Multi-Layout Merge**: Export sheets to individual DWGs, then merge into 1 file with multiple layouts (each sheet = 1 layout)
- **Single Layout Merge**: Combine all sheets into 1 DWG file with 1 layout
- **Model Space Export**: Export sheets into Model Space with title blocks arranged in a grid
- **Auto-Bind XRefs**: Automatically prevent/handle XREF files during export
- **Smart View Scale**: Detect primary view scale and update title block parameters
- **Layer Manager**: Export/Import DWG layer mapping settings to share across team
- **Persistent Settings**: Remembers your last export settings

## Requirements

- **Revit**: 2020-2026
- **AutoCAD**: 2020-2026 (required for merge features)

## Installation

### Automated Install

```powershell
# From project root
.\deploy\Deploy.ps1
```

### Manual Install

1. Build the solution:
   ```powershell
   .\build.ps1
   ```

2. Copy files to Revit Addins folder:
   - `bin\R2025\Release\Licorp_CombineCAD.dll` → `%APPDATA%\Autodesk\REVIT\Addins\2025\`
   - `bin\R2025\Release\Licorp_CombineCAD.addin` → `%APPDATA%\Autodesk\REVIT\Addins\2025\`

3. For AutoCAD merge features, install the plugin:
   - Copy `src.acad\Licorp_MergeSheets\PackageContents.xml` → `%PROGRAMDATA%\Autodesk\ApplicationPlugins\Licorp_MergeSheets.bundle\`

## Usage

### From Revit Ribbon

1. Open Revit and load a project with sheets
2. Go to the **Licorp** tab → **Combine CAD** panel
3. Choose export mode:
   - **Multi-Layout DWG**: Merged file with multiple layouts
   - **Individual DWG**: Separate files per sheet
   - **Single Layout DWG**: All sheets in one layout
   - **Model Space DWG**: Sheets arranged in Model Space grid

### Export Dialog

1. **Select Sheets**: Check sheets to export or use Select All/Deselect All
2. **Filter**: Search by sheet number or name
3. **Output Folder**: Choose destination folder
4. **DWG Export Setup**: Select from document's saved setups
5. **Export Mode**: Choose merge behavior
6. **Options**:
   - Auto-Bind XRefs: Prevent XREF files
   - Smart View Scale: Update title block with actual scale
   - Open after export: Auto-open result in AutoCAD
   - Progress always on top: Keep progress visible

### Layer Manager

Export or import DWG layer mapping settings to share with team members.

1. **Ribbon** → **Layer Manager**
2. Select export setup
3. Export to `.txt` or Import from `.txt`

## Project Structure

```
Licorp_CombiCAD/
├── src/
│   ├── Licorp_CombineCAD.Shared/      # Shared code
│   │   ├── Commands/                   # Revit commands
│   │   ├── Helpers/                    # Reflection utilities
│   │   ├── Models/                     # Data models
│   │   ├── Services/                   # Business logic
│   │   ├── ViewModels/                 # MVVM ViewModels
│   │   └── Views/                      # WPF dialogs
│   ├── Licorp_CombineCAD.R2020/       # .NET Framework 4.8 (Revit 2020-2024)
│   └── Licorp_CombineCAD.R2025/       # .NET 8 (Revit 2025-2026)
└── src.acad/
    └── Licorp_MergeSheets/             # AutoCAD plugin for merge
```

## Architecture

### Services

| Service | Purpose |
|---------|---------|
| `DwgExportService` | Core DWG export with linked model handling |
| `DwgMergeService` | Coordinate AutoCAD merge via AcCoreConsole |
| `DwgCleanupService` | XREF file detection and cleanup |
| `AutoCadLocatorService` | Find AutoCAD/AcCoreConsole on system |
| `SmartScaleService` | Detect and apply view scales to title blocks |
| `LayerMappingService` | Export/Import layer settings |
| `ProfileService` | Save/load user preferences |
| `SheetCollectorService` | Collect sheets with viewport analysis |

### MVVM Pattern

- **ViewModels**: Handle UI logic, commands, property changes
- **Views**: XAML-based WPF dialogs with dark theme
- **Models**: Simple data objects with INotifyPropertyChanged

## Edge Cases Handled

- AutoCAD not installed → Only Individual export available
- User cancel → Cleanup temp files
- Sheet without viewports → Skip with warning
- File already exists → Auto-rename with counter
- Output folder doesn't exist → Create automatically
- Export fails → Log error, continue with remaining sheets
- Linked models → Auto-unload before export, reload after

## Build

```powershell
.\build.ps1
```

Output:
- `bin\R2025\Release\` - Revit 2025-2026 add-in
- `bin\R2020\Release\` - Revit 2020-2024 add-in
- `bin\acad\Release\` - AutoCAD plugin

## License

Copyright © LICORP. All rights reserved.

## Credits

Built upon foundations from:
- [Nice3point Revit Extensions](https://github.com/Nice3point)
- [ricaun](https://github.com/ricaun-io)
- [chuongmep](https://github.com/chuongmep)