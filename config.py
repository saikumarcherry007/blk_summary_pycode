

# Directory and file paths
ALL_BLOCK_CSV_FILES_DIR = "all_block_csv_files"  # For Debug purpose dir will be created.
proj_dir_path = "scdc/wefw/rwfrwg/dveqw/"  # Example project directory path
Output_xls_name = f"{proj_dir_path.rstrip('/').split('/')[-1]}_block_summary.xlsx"

# Script metadata
SCRIPT_VERSION = "1.0.0"
AUTHOR = "Sai Kumar Malluru"

# Print control flag
ENABLE_PRINT = False

# Detailed info control flag for PASS/FAIL status or CLEAN / NOT CLEAN status
DETAILED_INFO = 0  # Set to 1 for detailed WNS/TNS/FEP or violations info, 0 for simple NOT CLEAN status

# Function control flags for highest_only parameter (Most off the tymes lead might need only worst case violations so keeping below varibale always to 1 )
# Set to 1 to return only highest/worst-case values, 0 to return all values (all list of all violations will be reported.)
FMAX_HIGHEST_ONLY = 1      # Controls process_fmax_data function
TCQ_HIGHEST_ONLY = 1       # Controls process_tcq_data function
MPW_HIGHEST_ONLY = 1       # Controls process_min_pulse_width function
DRV_HIGHEST_ONLY = 1       # Controls process_drv_data function (default 0 for detailed DRV info)

# Excel headers configuration
MAIN_HEADERS = [
    "Compiler", "Block Name", "Block Owner",
    "PNR FP CHECK",
    "PNR PG Checks",
    "PNR GENERAL CHECKS",
    "PNR SANITY CHECKS",
    "Dashboard", "Dashboard", "Dashboard",
    "HOLD", "FMAX", "DRV", "TCQ", "MPW",
    "PHYSICAL VERIFICATION", "PHYSICAL VERIFICATION", 
    "PHYSICAL VERIFICATION", "PHYSICAL VERIFICATION",
    "IR DROP", "IR DROP", "Formality"
]

SUB_HEADERS = [
    "", "", "",
    "PINS ON TRACK",
    "PG SHORTS", "PG OPENS", "PG MISSING VIAS",
    "MEM TAPPING", "SHIELD COVERAGE", "MAX ROUTE LENGTH", "NO ERRORS P&R LOG",
    "1'B/ASSIGN", "LVS SHORT", "LVS OPEN", "DRC <500", "MAX ROUTE", "LEF CHECK", "TERMINAL CREATION", "TERMINAL LAYER", "SYN VS PNR", "MEM PLEF VS NDM",
    "PARA ERRORS", "NOT ANNOTATED", "MPW VIOLATION",
    "Clk_groups", "", "", "", "",
    "DRC", "LVS", "ERC", "ANT",
    "VDD", "VSS", ""
]

# Column widths for Excel output
COLUMN_WIDTHS = {
    0: 20, 1: 15, 2: 15,  # Compiler, Block Name, Block Owner
    3: 15,  # PNR FP CHECK (PINS ON TRACK)
    4: 15, 5: 15, 6: 18,  # PNR PG Checks sub-columns
    7: 15, 8: 18, 9: 18, 10: 18,  # PNR GENERAL CHECKS sub-columns (MEM TAPPING, SHIELD COVERAGE, MAX ROUTE LENGTH, NO ERRORS P&R LOG)
    11: 14, 12: 14, 13: 14, 14: 14, 15: 14, 16: 14, 17: 14, 18: 14, 19: 14, 20: 14,  # PNR SANITY CHECKS sub-columns (10 checks)
    21: 15, 22: 18, 23: 18,  # Dashboard sub-columns (PARA ERRORS, NOT ANNOTATED, MPW VIOLATION)
    24: 40,  # HOLD
    25: 60,  # FMAX
    26: 40,  # DRV
    27: 55,  # TCQ
    28: 40,  # MPW
    29: 15, 30: 15, 31: 15, 32: 15,  # PHYSICAL VERIFICATION sub-columns (DRC, LVS, ERC, ANT)
    33: 15, 34: 15,  # IR DROP sub-columns (VDD, VSS)
    35: 15  # Formality
}

# Block information (can be moved to external config file or database)
BLOCK_INFO = [
    {"block_name": "CDM_top", "compiler": "Generic", "owner": "Sai Kumar"},
    {"block_name": "i36_i50", "compiler": "HSPRAM", "owner": "John Doe"},
    {"block_name": "i36_i50_i12", "compiler": "HS1PRF", "owner": "Alice Johnson"}
]
