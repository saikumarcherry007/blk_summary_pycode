"""
Physical Verification Module
Contains functions for processing DRC, LVS, ERC, and ANT verification results
"""

import os


def process_drc_value(excel_file, proj_dir_path):

    try:
        block_name_with_ext = os.path.basename(excel_file)
        block_name = os.path.splitext(block_name_with_ext)[0]
        if block_name.endswith("_metrics"):
            block_name = block_name[:-len("_metrics")]
            
        drc_file_path = os.path.join(proj_dir_path, "PV", "drc", block_name, "icv_mf_drc_run", f"{block_name}.RESULTS")

        if os.path.exists(drc_file_path):
            with open(drc_file_path, 'r') as f:
                first_line = f.readline().strip()
            if "RESULTS: CLEAN" in first_line:
                return "CLEAN"
            elif "RESULTS: NOT CLEAN" in first_line:
                return "NOT CLEAN"
            else:
                return first_line
        else:
            return "DRC File Not Found"
    except Exception as e:
        return f"Error reading DRC file: {e}"


def process_lvs_value(excel_file, proj_dir_path):

    try:
        block_name_with_ext = os.path.basename(excel_file)
        block_name = os.path.splitext(block_name_with_ext)[0]
        if block_name.endswith("_metrics"):
            block_name = block_name[:-len("_metrics")]

        lvs_file_path = os.path.join(proj_dir_path, "PV", "lvs", block_name, "icv_mf_lvs_run", f"{block_name}.RESULTS")

        if os.path.exists(lvs_file_path):
            with open(lvs_file_path, 'r') as f:
                first_line = f.readline().strip()
            if "LVS Compare Results: PASS" in first_line:
                return "CLEAN"
            elif "LVS Compare Results: NOT CLEAN" in first_line:
                return "NOT CLEAN"
            else:
                return first_line
        else:
            return "LVS File Not Found"
    except Exception as e:
        return f"Error reading LVS file: {e}"


def process_erc_value(excel_file, proj_dir_path):

    try:
        block_name_with_ext = os.path.basename(excel_file)
        block_name = os.path.splitext(block_name_with_ext)[0]
        if block_name.endswith("_metrics"):
            block_name = block_name[:-len("_metrics")]

        erc_file_path = os.path.join(proj_dir_path, "PV", "lvs", block_name, "icv_mf_lvs_run", f"{block_name}.RESULTS")

        if os.path.exists(erc_file_path):
            with open(erc_file_path, 'r') as f:
                line = None
                for i in range(11):  # Read up to 11 lines
                    line = f.readline()
                    if not line:  # Break if end of file is reached before 11 lines
                        break
                if line:
                    eleventh_line = line.strip()
                    if "DRC and Extraction Results: CLEAN" in eleventh_line:
                        return "CLEAN"
                    elif "DRC and Extraction Results: NOT CLEAN" in eleventh_line:
                        return "NOT CLEAN"
                    else:
                        return eleventh_line
                else:
                    return "ERC File Empty or Less than 11 lines"
        else:
            return "ERC File Not Found"
    except Exception as e:
        return f"Error reading ERC file: {e}"


def process_ant_value(excel_file, proj_dir_path):

    try:
        block_name_with_ext = os.path.basename(excel_file)
        block_name = os.path.splitext(block_name_with_ext)[0]
        if block_name.endswith("_metrics"):
            block_name = block_name[:-len("_metrics")]

        ant_file_path = os.path.join(proj_dir_path, "PV", "ant", block_name, "icv_mf_ant_run", f"{block_name}.RESULTS")

        if os.path.exists(ant_file_path):
            with open(ant_file_path, 'r') as f:
                first_line = f.readline().strip()
            if "RESULTS: CLEAN" in first_line:
                return "CLEAN"
            elif "RESULTS: NOT CLEAN" in first_line:
                return "NOT CLEAN"
            else:
                return first_line
        else:
            return "ANT File Not Found"
    except Exception as e:
        return f"Error reading ANT file: {e}"
