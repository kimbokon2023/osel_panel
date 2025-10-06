# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build and Run Commands

### Running the application
```bash
# Run the main application with Gooey GUI
python dawan_jamb.py

# For development/debugging without GUI
# Note: Main function is decorated with @Gooey, so direct python execution will show GUI
# The application requires Excel files in 'excel파일/' directory to run

# Build executable with PyInstaller
pyinstaller dawan_jamb.spec

# Build and run executable
pyinstaller dawan_jamb.spec && ./dist/dawan_jamb/dawan_jamb.exe

# Clean build artifacts
rmdir /s /q build dist
del /q *.spec.bak
```

### Dependencies Installation
```bash
pip install ezdxf openpyxl gooey requests psutil numpy matplotlib lxml
pip install pyinstaller  # For building executables
```

### Testing and Debugging
```bash
# No formal test suite exists - testing is done through:
# 1. Running with sample Excel files in excel파일/ directory
# 2. Checking generated DXF output in 작업완료/ directory
# 3. Manual validation of drawing accuracy and layer structure

# Debugging output files are saved to:
# - DXF files: 작업완료/ directory
# - Logs: data/log.txt (if logging is enabled)
```

## Architecture Overview

### Core Application Structure

**dawan_jamb.py** (Main Application - ~10,000+ lines)
- CAD drawing automation system for fire-resistant door jambs (방화쟘)
- Uses ezdxf for DXF/DWG file generation
- Excel-based data input via openpyxl
- Gooey GUI framework for user interface

### Key Components

1. **Drawing System**
   - Coordinate system using dictionary-based point management (`{x1, y1, x2, y2, ...}`)
   - Layer-based drawing: `'0'` (assembly), `'레이져'` (laser cutting), `'구성선'` (construction lines)
   - Support for two main jamb types: `막판유` and `막판무`

2. **Global Variables (lines 38-100)**
   - Material specifications: `thickness`, `br` (bending rate)
   - Dimensions: `OP`, `JE`, `JD`, `HH`, `MH`, `HPI_height`
   - Drawing parameters: `BasicXscale`, `BasicYscale`, `frame_scale`

3. **CPI Integration (lines 1660-1889)**
   - `draw_cpi_model()`: Unified CPI drawing function
   - Handles various CPI types and hole patterns
   - See README.md for detailed CPI documentation

4. **Assembly Drawing Functions**
   - `막판무` assembly: lines 2935-3074
   - `막판유` assembly: lines 2680-2731
   - Reinforcement sections: lines 4020-4044

### Data Flow

1. **Excel Input** → Load jamb specifications from `.xlsm` files
2. **Data Processing** → `load_excel()` function parses specifications
3. **Drawing Generation** → Coordinate calculation and DXF generation
4. **Output** → DXF files saved to `작업완료` directory

### Key Drawing Functions

- `set_point(dict, index, x, y)`: Set coordinate points
- `line(doc, x1, y1, x2, y2, layer)`: Draw lines
- `dim_linear()`, `dim_angular()`: Add dimension lines
- `insert_block()`: Insert standard blocks (holes, fixtures)
- `draw_hatshape()`: Draw reinforcement shapes
- `calculate_base()`: Calculate base dimensions from height and angle

### File Structure

```
C:\dawan\
├── dawan_jamb.py          # Main application
├── dawan_jamb.spec        # PyInstaller configuration
├── dim_stretch.py          # Dimension stretching utilities
├── data\                   # Configuration and data files
│   ├── settings.json       # Application settings
│   └── jamb.json          # Jamb specifications
├── excel파일\              # Excel templates
│   └── (각도입력폼) jamb신규쟘 작성양식.xlsm
├── dimstyle\              # DXF dimension styles
│   └── dawan_style.dxf
├── 작업완료\              # Output directory for generated DXF files
└── 연구자료\              # Reference DWG files

```

### Important Code Sections

- **Jamb Type Processing**: Lines 258-260 handle special `막판유` JD calculations
- **Wide Jamb Drawing**: Lines 2215-4065 (`draw_wide()` function)
- **Reinforcement Sections**: Lines 4020-4099 (cross-sections)
- **Frame Insertion**: Line 4061 (`insert_frame()`)
- **Hole Array Calculations**: `calcuteHoleArray()` function

### Layer Conventions

- `'0'`: Assembly drawings, main geometry
- `'레이져'`: Laser cutting paths
- `'구성선'`: Construction/reference lines
- `'치수선'`: Dimension lines

### Critical Variables

When modifying jamb specifications:
- `JD`: Jamb depth (기둥 깊이) - Note: `막판무` uses JD directly, `막판유` adds adjustments
- `OP`: Opening width
- `poleAngle`: Column angle for calculations
- `SW`, `SBW`: Side widths for reinforcement

### Error Handling

The application includes disk serial verification (`check_disk()`) and error message collection in the global `error_message` variable. GUI errors are handled through Gooey's validation system.

## Development Guidelines

### Working with Global Variables
- Most application state is managed through global variables (lines 38-100)
- When adding new functionality, follow the existing pattern of declaring globals at function start
- Key variables: `thickness`, `OP` (opening width), `JD` (jamb depth), `HH` (height), etc.

### Modifying Drawing Functions
- Core drawing logic is in coordinate dictionaries using `set_point()` pattern
- Layer conventions must be maintained: `'0'` (assembly), `'레이져'` (laser), `'구성선'` (construction)
- Always use the unified `draw_cpi_model()` function instead of legacy CPI functions

### Adding New Jamb Types
- Study existing type handling in lines 258-260 for `막판유` special cases
- New types should follow the coordinate-based drawing pattern
- Update the Excel parsing logic in `read_excel_rows()` if new columns are needed

### GUI and User Interface
- Application uses Gooey for GUI with Korean language support
- Main entry point is the `@Gooey` decorated `main()` function (line 4783)
- Error messages are displayed through `show_custom_error()` function
- Configuration is stored in `data/settings.json` and `data/jamb.json`

### Coordinate System and Scaling
- Uses `BasicXscale`, `BasicYscale` for coordinate transformation
- `frame_scale` controls overall drawing scale
- `set_point(dict, index, x, y)` is the standard way to manage coordinates
- Dimension positioning uses `saved_DimXpos`, `saved_DimYpos` for consistency

### Working with Excel Input
- Input templates are in `excel파일/` directory (.xlsm format)
- Main parsing happens in `read_excel_value()` and `read_excel_rows()`
- Cell mapping is defined in `variable_names` dictionary (lines 4841-4855)
- Always handle Unicode Korean text properly when reading Excel data