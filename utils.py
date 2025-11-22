

import sys
import platform
import builtins
from datetime import datetime
from config import ENABLE_PRINT, proj_dir_path, Output_xls_name, ALL_BLOCK_CSV_FILES_DIR, SCRIPT_VERSION, AUTHOR


# Save the original print function
builtins._original_print = builtins.print


def custom_print(*args, **kwargs):
    
    if ENABLE_PRINT:
        builtins._original_print(*args, **kwargs)


def toggle_print(enable=True):
    
    import config
    config.ENABLE_PRINT = enable
    custom_print(f"[INFO] Print statements {'enabled' if enable else 'disabled'}")


def print_header():
    
    current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    env_info = f"Python {sys.version.split()[0]}, OS: {platform.system()} {platform.release()}"
    header = f"""
{'=' * 100}
                                \033[93mBlock Summary Processing Script\033[0m
{'=' * 100}
Project Directory: {proj_dir_path}
Block Summary OP : {Output_xls_name}
Execution Time   : {current_date}
Author           : {AUTHOR}
Script Version   : {SCRIPT_VERSION}
Environment      : {env_info}
Outputs Dir      : {ALL_BLOCK_CSV_FILES_DIR} (DEBUG PURPOSE)
Purpose          : Process Excel files for block metrics and PNR checks finally generates summary / detailed  report
{'=' * 100}
    
    [ Processing Blocks ] ▶ ▷ ▶
    """
    
    print(header)


def check_clean_status(df, column_name):
   
    if column_name in df.columns:
        if df[column_name].astype(str).str.contains("Not Clean", case=False).any():
            return "NOT CLEAN"
        else:
            import pandas as pd
            numeric_column = pd.to_numeric(df[column_name], errors='coerce')
            if numeric_column.sum() == 0:
                return "CLEAN"
            else:
                return "NOT CLEAN"
    return "N/A"
