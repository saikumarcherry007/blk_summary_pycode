# Block Summary Processing Script - Refactored Structure

## Overview
This is a refactored version of the Block Summary Processing Script, organized into logical modules for better maintainability and code reusability.

## File Structure

### Core Modules

1. **config.py**
   - Contains all global constants and configuration parameters
   - Script metadata (version, author)
   - Excel header configurations
   - Block information
   - Easy to modify settings without touching logic

2. **utils.py**
   - Utility functions for printing and formatting
   - `custom_print()` - Controlled print wrapper
   - `toggle_print()` - Enable/disable print statements
   - `print_header()` - Script execution header
   - `check_clean_status()` - DataFrame status checker

3. **physical_verification.py**
   - Physical verification checks module
   - `process_drc_value()` - Design Rule Check processing
   - `process_lvs_value()` - Layout vs Schematic processing
   - `process_erc_value()` - Electrical Rule Check processing
   - `process_ant_value()` - Antenna check processing

4. **timing_analysis.py**
   - Timing-related metrics processing
   - `process_hold_data()` - HOLD timing analysis
   - `process_fmax_data()` - Maximum frequency analysis
   - `process_tcq_data()` - Clock-to-Q timing analysis
   - `process_min_pulse_width()` - Minimum pulse width analysis

5. **design_checks.py**
   - Design quality checks module
   - `process_drv_data()` - Design Rule Violations (transition/capacitance)
   - `process_ir_value_to_csv()` - IR drop analysis
   - `process_formality_value()` - Formal verification checks

6. **excel_processor.py**
   - Excel file processing and output generation
   - `process_excel_file()` - Process single Excel file
   - `create_output_excel()` - Generate formatted Excel report

7. **main.py**
   - Main entry point that orchestrates all modules
   - Handles file discovery and processing workflow
   - Generates final summary reports

## Usage

### Running the Script
```bash
python main.py
```

### Modifying Configuration
Edit `config.py` to change:
- Output directories
- Project paths
- Block information
- Excel formatting options

### Adding New Blocks
Update the `BLOCK_INFO` list in `config.py`:
```python
BLOCK_INFO = [
    {"block_name": "your_block", "compiler": "Compiler Name", "owner": "Owner Name"},
    # Add more blocks here
]
```

## Benefits of Refactored Structure

1. **Modularity**: Each module has a single responsibility
2. **Maintainability**: Easier to locate and fix bugs
3. **Reusability**: Functions can be imported and reused
4. **Testability**: Individual modules can be tested independently
5. **Scalability**: Easy to add new features or checks
6. **Readability**: Clear separation of concerns

## Module Dependencies

```
main.py
├── config.py
├── utils.py
└── excel_processor.py
    ├── config.py
    ├── utils.py
    ├── physical_verification.py
    ├── timing_analysis.py
    └── design_checks.py
```

## Requirements
- Python 3.x
- pandas
- numpy
- xlsxwriter
- openpyxl

## Original File
The original monolithic script is preserved as:
`script_with_grn_effect_copy_11_wo_blk_dimensions.py`

## Author
Sai Kumar Malluru

## Version
1.0.0 (Refactored)
