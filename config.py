

# Directory and file paths
ALL_BLOCK_CSV_FILES_DIR = "all_block_csv_files"  # For Debug purpose dir will be created.
proj_dir_path = "scdc/wefw/rwfrwg/dveqw/"  # Example project directory path
Output_xls_name = "output_summary_latest.xlsx"

# Script metadata
SCRIPT_VERSION = "1.0.0"
AUTHOR = "Sai Kumar Malluru"

# Print control flag
ENABLE_PRINT = False

# Excel headers configuration
MAIN_HEADERS = [
    "Compiler", "Block Name", "Block Owner", 
    "Dashboard", "Dashboard", "Dashboard",
    "HOLD", "FMAX", "DRV", "TCQ", "MPW",
    "PHYSICAL VERIFICATION", "PHYSICAL VERIFICATION", 
    "PHYSICAL VERIFICATION", "PHYSICAL VERIFICATION",
    "IR DROP", "IR DROP", "Formality"
]

SUB_HEADERS = [
    "", "", "", 
    "PARA ERRORS", "NOT ANNOTATED", "MPW VIOLATION",
    "Clk_groups", "", "", "", "",
    "DRC", "LVS", "ERC", "ANT",
    "VDD", "VSS", ""
]

# Column widths for Excel output
COLUMN_WIDTHS = {
    0: 20, 1: 15, 2: 15, 3: 15, 4: 15, 5: 15, 
    6: 40, 7: 60, 8: 40, 9: 55, 10: 15, 11: 15, 
    12: 15, 13: 15, 14: 15, 15: 15, 16: 15, 17: 15
}

# Block information (can be moved to external config file or database)
BLOCK_INFO = [
    {"block_name": "CDM_top", "compiler": "Generic", "owner": "Sai Kumar"},
    {"block_name": "i36_i50", "compiler": "HSPRAM", "owner": "John Doe"},
    {"block_name": "i36_i50_i12", "compiler": "HS1PRF", "owner": "Alice Johnson"}
]
