import os
import glob
import csv
import pandas as pd
import numpy as np

from datetime import datetime
import time
import sys
import platform
import builtins

ALL_BLOCK_CSV_FILES_DIR = "all_block_csv_files"  # For Debug purpose dir will be created.
proj_dir_path = "scdc/wefw/rwfrwg/dveqw/"  # Example project directory path
Output_xls_name = "output_summary_latest.xlsx"

# Save the original print function
builtins._original_print = builtins.print

def custom_print(*args, **kwargs):
    """Wrapper around print that respects the ENABLE_PRINT flag."""
    if ENABLE_PRINT:
        builtins._original_print(*args, **kwargs)

def toggle_print(enable=True):
    """Toggle print statements on or off."""
    global ENABLE_PRINT
    ENABLE_PRINT = enable
    custom_print(f"[INFO] Print statements {'enabled' if enable else 'disabled'}")

def print_header():
    """Prints a formatted header for the script execution."""
    current_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    script_version = "1.0.0"  # Update this as needed
    env_info = f"Python {sys.version.split()[0]}, OS: {platform.system()} {platform.release()}"
    header = f"""
{'=' * 100}
                                \033[93mBlock Summary Processing Script\033[0m
{'=' * 100}
Project Directory: {proj_dir_path}
Block Summary OP : {Output_xls_name}
Execution Time   : {current_date}
Author           : Sai Kumar Malluru
Script Version   : {script_version}
Environment      : {env_info}
Outputs Dir      : {ALL_BLOCK_CSV_FILES_DIR} (DEBUG PURPOSE)
Purpose          : Process Excel files for block metrics and generate summary report
{'=' * 100}
    
    [ Processing Blocks ] ▶ ▷ ▶
    """
    
    print(header)

def process_drc_value(excel_file):
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

def process_lvs_value(excel_file):
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

def process_erc_value(excel_file):
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

def process_ant_value(excel_file):
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

def process_ir_value_to_csv(excel_file, proj_dir_path):
    try:
        block_name_with_ext = os.path.basename(excel_file)
        block_name = os.path.splitext(block_name_with_ext)[0]
        if block_name.endswith("_metrics"):
            block_name = block_name[:-len("_metrics")]

        ir_file_pattern = os.path.join(proj_dir_path, "ir_drop_rh", block_name, "func*", f"voltage*.rpt")
        ir_files = glob.glob(ir_file_pattern)

        if not ir_files:
            custom_print(f"[INFO] No IR voltage report files found for block: {block_name} using pattern: {ir_file_pattern}")
            return None, None

        custom_print(f"[INFO] Found IR voltage report files for block {block_name}: {ir_files}")

        vdd_value_str = None
        vss_value_str = None

        for ir_file in ir_files:
            try:
                with open(ir_file, 'r') as f_vdd:
                    for line in f_vdd:
                        if "/VDD" in line:
                            columns = line.split()
                            if columns:
                                vdd_value_str = columns[-1]
                                custom_print(f"[INFO] First VDD Value Found in {ir_file}: {vdd_value_str}")
                                break
                if vdd_value_str is not None:
                    with open(ir_file, 'r') as f_vss:
                        for line in f_vss:
                            if "/VSS" in line:
                                columns = line.split()
                                if columns:
                                    vss_value_str = columns[-1]
                                    custom_print(f"[INFO] First VSS Value Found in {ir_file}: {vss_value_str}")
                                    break
            except Exception as e:
                custom_print(f"[WARNING] Error reading IR report file {ir_file}: {e}")

            if vdd_value_str is not None and vss_value_str is not None:
                break

        return vdd_value_str, vss_value_str
    except Exception as e:
        custom_print(f"[WARNING] Error processing IR value for {excel_file}: {e}")
        return None, None

def process_formality_value(excel_file, proj_dir_path):
    try:
        block_name_with_ext = os.path.basename(excel_file)
        block_name = os.path.splitext(block_name_with_ext)[0]
        if block_name.endswith("_metrics"):
            block_name = block_name[:-len("_metrics")]
        log_file_path = os.path.join(proj_dir_path, "formality", block_name, f"fm.log")

        if not os.path.exists(log_file_path):
            return "Log File Not Found"

        with open(log_file_path, 'r') as f:
            log_content = f.read()
            if "Verification SUCCEEDED" in log_content:
                return "PASSING"
            else:
                return "NOT PASSING"
    except Exception as e:
        custom_print(f"[WARNING] Error processing formality log for {excel_file}: {e}")
        return "Error Processing Log"

def process_excel_file(excel_file, output_excel_file="output.xlsx"):
    try:
        xls = pd.ExcelFile(excel_file)
        base_name = os.path.splitext(excel_file)[0]
        output_dir = os.path.join(ALL_BLOCK_CSV_FILES_DIR, base_name + "_csv")

        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            custom_print(f"[CREATED] Created directory: {output_dir}")
        else:
            custom_print(f"[WARNING] Directory already exists: {output_dir}. Files might be overwritten.")

        first_sheet_name = xls.sheet_names[0]
        df_main = xls.parse(first_sheet_name)

        para_status = check_clean_status(df_main, "PARA ERRORS")
        not_annotated_status = check_clean_status(df_main, "NOT ANNOTATED")

        mpw_details = process_min_pulse_width(excel_file, output_dir)
        mpw_violation_status = "CLEAN" if mpw_details == "CLEAN" else "NOT CLEAN"
        hold_clk_grp_output = process_hold_data(excel_file, output_dir)

        clock_groups = base_name.split("_")
        fmax_details = process_fmax_data(excel_file, clock_groups, xls, output_dir)

        drv_details = process_drv_data(xls, output_dir)

        tcq_percentage = process_tcq_data(excel_file, output_dir)

        mpw_details = process_min_pulse_width(excel_file, output_dir)

        drc_value = process_drc_value(excel_file)
        lvs_value = process_lvs_value(excel_file)
        erc_value = process_erc_value(excel_file)
        ant_value = process_ant_value(excel_file)

        vdd_value_str, vss_value_str = process_ir_value_to_csv(excel_file, proj_dir_path)
        calculated_vdd = "Vol*.rpt File Not Found"
        calculated_vss = "Vol*.rpt File Not Found"
        vdd_numeric = None
        vss_numeric = None

        if vdd_value_str is not None:
            try:
                vdd_numeric = float(vdd_value_str.replace("D", "e"))
                calculated_vdd = f"{vdd_numeric / 0.825 * 100:.2f}%"
            except ValueError:
                calculated_vdd = "Error"

        if vss_value_str is not None:
            try:
                vss_numeric = float(vss_value_str.replace("D", "e"))
                calculated_vss = f"{vss_numeric / 0.825 * 100:.2f}%"
            except ValueError:
                calculated_vss = "Error"

        formality_value = process_formality_value(excel_file, proj_dir_path)

        block_name_with_ext = os.path.basename(excel_file)
        block_name = os.path.splitext(block_name_with_ext)[0]
        if block_name.endswith("_metrics"):
            block_name = block_name[:-len("_metrics")]

        output_data = [
            [block_name, para_status, not_annotated_status, mpw_violation_status,
             hold_clk_grp_output, fmax_details, drv_details, tcq_percentage, mpw_details,
             drc_value, lvs_value, erc_value, ant_value, calculated_vdd, calculated_vss, formality_value]
        ]

        custom_print(f"\nExcel Output Table for {excel_file}:")

        main_header_row_parts = []
        main_header_row_parts.append("| ")
        main_header_row_parts.append(main_headers[0] + " |")
        dashboard_width = len(" | ".join(sub_headers[1:4]))
        main_header_row_parts.append(main_headers[1].center(dashboard_width) + " |")
        for header in main_headers[4:9]:
            main_header_row_parts.append(header + " |")
        drc_width = len(" | ".join(sub_headers[9:13]))
        main_header_row_parts.append(main_headers[9].center(drc_width) + " |")
        ir_drop_width = len(" | ".join(sub_headers[13:15]))
        main_header_row_parts.append(main_headers[13].center(ir_drop_width) + " |")
        main_header_row_parts.append(main_headers[15] + " |")
        main_header_row = "".join(main_header_row_parts)

        header_row = "| " + " | ".join(sub_headers) + " |"

        separator_row = "+" + "-+-".join(['-' * (len(h) if len(h) > 0 else 15) for h in sub_headers]) + "-+"

        custom_print(separator_row)
        custom_print(main_header_row)
        custom_print(header_row)
        custom_print(separator_row)

        for row in output_data:
            output_row = "| " + " | ".join(str(item) for item in row) + " |"
            custom_print(output_row)
        custom_print(separator_row)

        custom_print("\nMain Headers (List format - as before):")
        custom_print(main_headers[:10])
        return output_data

    except Exception as e:
        custom_print(f"[WARNING] Error processing file {excel_file}: {e}")
        return ["Error processing file"]

def check_clean_status(df, column_name):
    if column_name in df.columns:
        if df[column_name].astype(str).str.contains("Not Clean", case=False).any():
            return "NOT CLEAN"
        else:
            numeric_column = pd.to_numeric(df[column_name], errors='coerce')
            if numeric_column.sum() == 0:
                return "CLEAN"
            else:
                return "NOT CLEAN"
    return "N/A"

def process_drv_data(xls, output_dir, highest_only=1):
    drv_details = ""

    if "DRV" in xls.sheet_names:
        df_excel = pd.read_excel(xls, sheet_name='DRV')

        corners_col = df_excel.columns[0]
        maxtran_cols = df_excel.columns[1:4]
        maxcap_cols = df_excel.columns[4:7]

        df_excel[maxtran_cols] = df_excel[maxtran_cols].apply(pd.to_numeric, errors='coerce')
        df_excel[maxcap_cols] = df_excel[maxcap_cols].apply(pd.to_numeric, errors='coerce')

        df_excel.dropna(inplace=True)

        tran_wns_col = maxtran_cols[0]
        cap_wns_col = maxcap_cols[0]

        if not df_excel.empty and (df_excel[tran_wns_col] == 0).all():
            drv_details = "TRAN: CLEAN"
            if (df_excel[cap_wns_col] == 0).all():
                drv_details += " | CAP: CLEAN"
                block_tran_cap_file = os.path.join(output_dir, "block_tran_cap.csv")
                with open(block_tran_cap_file, 'w') as f:
                    f.write("TRAN : CLEAN | CAP: CLEAN")
                custom_print(f"[CREATED] Dumped 'block_tran_cap.csv' to: {block_tran_cap_file}")
            elif not df_excel.empty:
                if highest_only == 1:
                    maxcap_farthest_index = df_excel[cap_wns_col].abs().idxmax()
                    maxcap_result = df_excel.loc[maxcap_farthest_index]
                    drv_details += f" | CAP: {maxcap_result[corners_col]} - WNS:{maxcap_result[cap_wns_col]}; BEP:{maxcap_result[maxcap_cols[1]]}; FEP:{maxcap_result[maxcap_cols[2]]}."
                else:
                    cap_not_zero = df_excel[df_excel[cap_wns_col] != 0]
                    cap_details_list = []
                    for index, row in cap_not_zero.iterrows():
                        cap_details = f"CAP: {row[corners_col]} - WNS:{row[cap_wns_col]}; BEP:{row[maxcap_cols[1]]}; FEP: {row[maxcap_cols[2]]}"
                        cap_details_list.append(cap_details)
                    if cap_details_list:
                        drv_details += " | " + ", ".join(cap_details_list)
                    else:
                        drv_details += " | CAP: CLEAN"
        elif not df_excel.empty:
            maxtran_not_zero = df_excel[df_excel[tran_wns_col] != 0]
            block_tran_csv_file = os.path.join(output_dir, "block_tran_cap.csv")
            if highest_only == 1 and not maxtran_not_zero.empty:
                maxtran_farthest_index = maxtran_not_zero[tran_wns_col].abs().idxmax()
                maxtran_result = maxtran_not_zero.loc[[maxtran_farthest_index]]
                maxtran_result.to_csv(block_tran_csv_file, index=False)
                custom_print(f"[CREATED] Dumped 'DRV' sheet (Highest TRAN WNS) to: {block_tran_csv_file}")
                maxtran_to_report = maxtran_result
            else:
                maxtran_not_zero.to_csv(block_tran_csv_file, index=False)
                custom_print(f"[CREATED] Dumped 'DRV' sheet (TRAN WNS != 0) to: {block_tran_csv_file}")
                maxtran_to_report = maxtran_not_zero

            drv_details_list = []
            for index, row in maxtran_to_report.iterrows():
                tran_details = f"TRAN: {row[corners_col]} - WNS:{row[tran_wns_col]}; BEP:{row[maxtran_cols[1]]}; FEP: {row[maxtran_cols[2]]}"
                drv_details_list.append(tran_details)
            drv_details = ", ".join(drv_details_list)

            if (df_excel[cap_wns_col] == 0).all():
                drv_details += " | CAP: CLEAN"
            else:
                if highest_only == 1:
                    maxcap_farthest_index = df_excel[cap_wns_col].abs().idxmax()
                    maxcap_result = df_excel.loc[maxcap_farthest_index]
                    drv_details += f" | CAP: {maxcap_result[corners_col]} - WNS:{maxcap_result[cap_wns_col]}; BEP:{maxcap_result[maxcap_cols[1]]}; FEP:{maxcap_result[maxcap_cols[2]]}."
                else:
                    cap_not_zero = df_excel[df_excel[cap_wns_col] != 0]
                    cap_details_list = []
                    for index, row in cap_not_zero.iterrows():
                        cap_details = f"CAP: {row[corners_col]} - WNS:{row[cap_wns_col]}; BEP:{row[maxcap_cols[1]]}; FEP: {row[maxcap_cols[2]]}"
                        cap_details_list.append(cap_details)
                    if cap_details_list:
                        drv_details += " | " + ", ".join(cap_details_list)
                    else:
                        drv_details += " | CAP: CLEAN"
        else:
            drv_details = "TRAN: CLEAN | CAP: CLEAN"
            block_tran_cap_file = os.path.join(output_dir, "block_tran_cap.csv")
            with open(block_tran_cap_file, 'w') as f:
                f.write("TRAN : CLEAN | CAP: CLEAN")
            custom_print(f"[CREATED] Dumped 'block_tran_cap.csv' to: {block_tran_cap_file}")

    else:
        drv_details = ""

    return drv_details

def process_hold_data(excel_file, output_dir):
    hold_sheet = "HOLD_MASTER_CLK"
    summary_sheet = "HOLD_MASTER_CLK_SUM"

    try:
        df_summary = pd.read_excel(excel_file, sheet_name=summary_sheet)
        clk_grps = df_summary.iloc[0:, 0].tolist()

        clk_grps_csv = os.path.join(output_dir, "HOLD_MASTER_CLK_SUM_allclk_grps.csv")
        with open(clk_grps_csv, 'w') as f:
            for clk_grp in clk_grps:
                f.write(f"{clk_grp}\n")
        custom_print(f"[CREATED] Clock groups list saved to: {clk_grps_csv}")

        if not clk_grps:
            custom_print("[WARNING] No clock groups found in the HOLD_MASTER_CLK_SUM sheet.")
            return "No clock groups found"

        df_hold = pd.read_excel(excel_file, sheet_name=hold_sheet)
        clk_grp_indices = [i for i, col in enumerate(df_hold.columns) if col.startswith("clk_grp")]

        if not clk_grp_indices:
            custom_print("[WARNING] No 'clk_grp' columns found in the HOLD_MASTER_CLK sheet.")
            return "No clk_grp columns found"

        hold_corners = df_hold.iloc[:, [0]]
        clk_group_results = []
        all_clean = True

        for clk_grp_name in clk_grps:
            matching_indices = [i for i, col in enumerate(df_hold.columns) if col == clk_grp_name]
            if not matching_indices:
                custom_print(f"[WARNING] Clock group '{clk_grp_name}' not found in HOLD_MASTER_CLK sheet.")
                continue

            start_idx = matching_indices[0]
            next_clk_indices = [i for i in clk_grp_indices if i > start_idx]
            end_idx = next_clk_indices[0] if next_clk_indices else len(df_hold.columns)

            df_clk_grp = df_hold.iloc[:, start_idx:end_idx]
            df_final = pd.concat([hold_corners, df_clk_grp], axis=1)
            df_final = df_final.iloc[1:].reset_index(drop=True)

            csv_file = os.path.join(output_dir, f"{clk_grp_name}_grouped.csv")
            df_final.to_csv(csv_file, sep=' ', index=False, header=False)
            custom_print(f"[CREATED] Data for {clk_grp_name} saved to CSV: {csv_file}")

            df_csv = pd.read_csv(csv_file, sep=' ', header=None)
            if df_csv.shape[1] > 3:
                df_csv[1] = pd.to_numeric(df_csv[1], errors='coerce')

                if not (df_csv[1] < 0).any():
                    custom_print(f"{clk_grp_name} : CLEAN")
                    clk_group_results.append(f"{clk_grp_name} : CLEAN")
                else:
                    all_clean = False
                    max_abs_index = df_csv[1].abs().idxmax()
                    max_abs_row = df_csv.loc[max_abs_index]

                    func_name = max_abs_row[0]
                    wns = max_abs_row[1]
                    tns = max_abs_row[2]
                    fep = max_abs_row[3]

                    formatted_row = f"{clk_grp_name} {func_name} : WNS: {wns}; TNS: {tns}; FEP: {fep}"
                    clk_group_results.append(formatted_row)

                    with open(csv_file, "w") as f:
                        f.write(formatted_row + "\n")

                    custom_print(f"[CREATED] Farthest value from zero in {clk_grp_name}: {formatted_row}")
            else:
                custom_print(f"[WARNING] {clk_grp_name} does not have enough columns to process.")
                all_clean = False

        output_string = " | ".join(clk_group_results) + "."
        if all_clean and clk_group_results:
            return "HOLD CLEAN"
        elif not clk_group_results:
            return "No clock groups found or processed"
        else:
            return output_string

    except Exception as e:
        custom_print(f"[WARNING] Error processing HOLD data: {e}")
        return f"Error processing HOLD data: {str(e)}"

def process_fmax_data(excel_file, clock_groups, xls, output_dir, highest_only=1):
    try:
        excel_filename_without_ext = os.path.splitext(os.path.basename(excel_file))[0]
        base_name = excel_filename_without_ext.replace('_metrics', '')
        if base_name in ["CDM_top", "PLL", "setuphold", "clk_jtag_pll_cntrl"]:
            custom_print(f"[INFO] FMAX sheet marked as Not Applicable for file: {excel_file}")
            return "FMAX Not Applicable"

        if 'FMAX' in xls.sheet_names:
            custom_print("FMAX sheet is present.")
            df_fmax = pd.read_excel(excel_file, sheet_name='FMAX')

            excel_filename_without_ext = os.path.splitext(os.path.basename(excel_file))[0]
            csv_file = os.path.join(output_dir, f"FMAX_{excel_filename_without_ext}.csv")
            df_fmax.to_csv(csv_file, index=False, sep=' ', header=False)
            custom_print(f"[CREATED] Sheet 'FMAX' converted to CSV: {csv_file} (space delimited, no header)")

            df_fmax = pd.read_csv(csv_file, delimiter=' ', header=None)
            df_fmax.iloc[:, 0] = df_fmax.iloc[:, 0].fillna('')
            filtered_fmax = df_fmax[df_fmax.iloc[:, 0].str.startswith(('func', 'test', 'fbist'))]
            filtered_fmax.to_csv(csv_file, index=False, sep=' ', header=False)
            custom_print(f"[CREATED] Main CSV file '{csv_file}' has been filtered.")

            raw_blocks = excel_filename_without_ext.split('_')
            blocks = [block for block in raw_blocks if block != "metrics"]
            custom_print(f"Identified blocks from filename: {blocks}")

            part_size = 7
            fmax_results = []
            tcc_blocks = []

            for idx, part in enumerate(blocks):
                new_csv_file = os.path.join(output_dir, f"{part}_fmax.csv")
                custom_print(f"Creating CSV file for block {part}: {new_csv_file}")

                start_col = 1 + idx * part_size
                end_col = start_col + part_size

                if start_col >= filtered_fmax.shape[1]:
                    custom_print(f"[WARNING] Not enough columns for {part}.")
                    fmax_results.append(f"{part}: Not enough data")
                    continue

                columns_to_extract = [0] + list(range(start_col, min(end_col, filtered_fmax.shape[1])))
                part_data = filtered_fmax.iloc[:, columns_to_extract].copy()

                if part_data.shape[1] > 1:
                    mask = part_data.iloc[:, 1] != "-"
                    part_data = part_data[mask]

                part_data.to_csv(new_csv_file, index=False, sep=' ', header=False)
                custom_print(f"[CREATED] Created {new_csv_file} with filtered data")

                if part_data.shape[1] >= 4:
                    limit_col_idx = min(part_data.shape[1] - 3, part_data.shape[1] - 1)
                    tccmargin_col_idx = min(part_data.shape[1] - 2, part_data.shape[1] - 1)
                    holdmargin_col_idx = min(part_data.shape[1] - 1, part_data.shape[1] - 1)

                    if part_data.empty:
                        custom_print(f"[WARNING] No valid data for {part} after filtering.")
                        fmax_results.append(f"{part}: No valid data")
                        continue

                    limit_col = part_data.iloc[:, limit_col_idx]
                    custom_print(f"[DEBUG] Unique values in limit_col for {part}: {limit_col.unique()}")
                    filtered_part_data = part_data[limit_col.str.contains('Memory|SMS', na=False, regex=True)]

                    custom_print(f"[DEBUG] Original rows: {len(part_data)}, Filtered rows: {len(filtered_part_data)}")
                    if filtered_part_data.empty:
                        custom_print(f"[DEBUG] No rows matched 'Memory|SMS' filter for {part}. Using all data instead.")
                        filtered_part_data = part_data

                    highest_sms = {"value": -1, "corner": "", "limit_value": "", "hold_margin": ""}
                    highest_memory = {"value": -1, "corner": "", "limit_value": "", "hold_margin": ""}
                    highest_general = {"value": -1, "corner": "", "limit_value": "", "hold_margin": ""}

                    all_lines = []
                    has_tcc = False
                    for idx_row, row in filtered_part_data.iterrows():
                        limit_value = row.iloc[limit_col_idx] if limit_col_idx < len(row) else "Unknown"
                        if "TCC" in str(limit_value):
                            has_tcc = True
                            tcc_blocks.append(part)
                            break

                    custom_print(f"[DEBUG] Block {part} has TCC: {has_tcc}")
                    if has_tcc:
                        custom_print(f"[INFO] Block {part} has TCC data - using simplified format")
                        fmax_results.append(f"{part}: TCC")
                        continue

                    for idx_row, row in filtered_part_data.iterrows():
                        corner = row.iloc[0]
                        limit_value = row.iloc[limit_col_idx] if limit_col_idx < len(row) else "Unknown"
                        tcc_margin = row.iloc[tccmargin_col_idx] if tccmargin_col_idx < len(row) and pd.notna(row.iloc[tccmargin_col_idx]) else 'NA'
                        hold_margin = row.iloc[holdmargin_col_idx] if holdmargin_col_idx < len(row) and pd.notna(row.iloc[holdmargin_col_idx]) else 'NA'

                        tcc_margin_str = str(tcc_margin).rstrip('%')
                        hold_margin_str = str(hold_margin).rstrip('%')

                        formatted_line = f"{corner} ({limit_value}: {tcc_margin_str}%); Hold_margin: {hold_margin_str}"
                        all_lines.append(formatted_line)
                        custom_print(f"[DEBUG] Processing row - Corner: {corner}, Limit: {limit_value}, Margin: {tcc_margin_str}, Hold: {hold_margin_str}")

                        try:
                            margin_value = float(str(tcc_margin).rstrip('%'))
                            if "SMS" in str(limit_value) and margin_value > highest_sms["value"]:
                                highest_sms = {
                                    "value": margin_value,
                                    "corner": corner,
                                    "limit_value": limit_value,
                                    "hold_margin": hold_margin_str
                                }
                                custom_print(f"[DEBUG] New highest SMS: {margin_value}% for {corner}")
                            elif "Memory" in str(limit_value) and margin_value > highest_memory["value"]:
                                highest_memory = {
                                    "value": margin_value,
                                    "corner": corner,
                                    "limit_value": limit_value,
                                    "hold_margin": hold_margin_str
                                }
                                custom_print(f"[DEBUG] New highest Memory: {margin_value}% for {corner}")
                            elif margin_value > highest_general["value"]:
                                highest_general = {
                                    "value": margin_value,
                                    "corner": corner,
                                    "limit_value": limit_value,
                                    "hold_margin": hold_margin_str
                                }
                                custom_print(f"[DEBUG] New highest general: {margin_value}% for {corner} ({limit_value})")
                        except (ValueError, TypeError) as e:
                            custom_print(f"[DEBUG] Error converting margin value '{tcc_margin}' to float: {e}")
                            continue

                    with open(new_csv_file, 'w') as outfile:
                        outfile.write('\n'.join(all_lines) if all_lines else f"{part}: No data")

                    if highest_only == 0:
                        if all_lines:
                            fmax_results.append(f"{part}: {', '.join(all_lines)}")
                        else:
                            fmax_results.append(f"{part}: No data")
                    else:
                        part_summary = []
                        if highest_sms["value"] > -1:
                            sms_line = f"{highest_sms['corner']} ({highest_sms['limit_value']}: {highest_sms['value']}%); Hold_margin: {highest_sms['hold_margin']}"
                            part_summary.append(sms_line)
                        if highest_memory["value"] > -1:
                            memory_line = f"{highest_memory['corner']} ({highest_memory['limit_value']}: {highest_memory['value']}%); Hold_margin: {highest_memory['hold_margin']}"
                            part_summary.append(memory_line)
                        if len(part_summary) == 0 and highest_general["value"] > -1:
                            general_line = f"{highest_general['corner']} ({highest_general['limit_value']}: {highest_general['value']}%); Hold_margin: {highest_general['hold_margin']}"
                            part_summary.append(general_line)

                        custom_print(f"[DEBUG] Part summary for {part}: {part_summary}")
                        if part_summary:
                            fmax_results.append(f"{part}: {', '.join(part_summary)}")
                        else:
                            if all_lines:
                                custom_print(f"[DEBUG] Warning: Have {len(all_lines)} lines but empty part_summary for {part}")
                                fmax_results.append(f"{part}: {all_lines[0]}")
                            else:
                                fmax_results.append(f"{part}: No valid margin data")
                else:
                    custom_print(f"[WARNING] {part} does not have enough columns (needs at least 4).")
                    fmax_results.append(f"{part}: Not enough columns")

            return " | ".join(fmax_results) + "."
        else:
            custom_print("FMAX sheet is not present.")
            return "FMAX sheet not found."
    except Exception as e:
        custom_print(f"[WARNING] Error processing FMAX data: {e}")
        return f"Error processing FMAX data: {e}"

def process_tcq_data(excel_file, output_dir, highest_only=0):
    try:
        base_name = os.path.splitext(excel_file)[0]
        if base_name.endswith('_metrics'):
            base_name = base_name[:-len('_metrics')]
        expected_block_names = base_name.split('_')
        
        df = pd.read_excel(excel_file, sheet_name='TCQ')
        
        if df.empty:
            custom_print(f"[WARNING] TCQ Sheet in {excel_file} is empty.")
            return "Empty Sheet"
        
        data_columns = df.columns[1:]
        actual_blocks = []
        actual_column_groups = []
        
        if len(data_columns) > 0:
            current_group = []
            current_block = None
            
            for col in data_columns:
                col_str = str(col).lower()
                block_match = None
                for block in expected_block_names:
                    if block.lower() in col_str:
                        block_match = block
                        break
                
                if block_match and block_match != current_block:
                    if current_group:
                        actual_column_groups.append(current_group)
                        actual_blocks.append(current_block)
                    current_group = [col]
                    current_block = block_match
                else:
                    current_group.append(col)
            
            if current_group:
                actual_column_groups.append(current_group)
                actual_blocks.append(current_block)
        
        if not actual_blocks:
            custom_print(f"[WARNING] Couldn't identify block patterns in columns. Trying to detect based on column count.")
            corr_matrix = df.iloc[:, 1:].corr(method='pearson', min_periods=5)
            non_nan_patterns = df.iloc[:, 1:].notna().sum()
            pattern_changes = np.diff(non_nan_patterns.values)
            potential_breaks = [i+1 for i, change in enumerate(pattern_changes) if abs(change) > df.shape[0]/4]
            
            if potential_breaks:
                potential_breaks = [0] + potential_breaks + [len(data_columns)]
                for i in range(len(potential_breaks)-1):
                    start_idx = potential_breaks[i]
                    end_idx = potential_breaks[i+1]
                    if i < len(expected_block_names):
                        actual_blocks.append(expected_block_names[i])
                        actual_column_groups.append(data_columns[start_idx:end_idx])
            else:
                num_expected_blocks = len(expected_block_names)
                if len(data_columns) % num_expected_blocks == 0:
                    columns_per_block = len(data_columns) // num_expected_blocks
                    for i, block in enumerate(expected_block_names):
                        start_idx = i * columns_per_block
                        end_idx = start_idx + columns_per_block
                        if start_idx < len(data_columns):
                            actual_blocks.append(block)
                            actual_column_groups.append(data_columns[start_idx:end_idx])
                else:
                    custom_print(f"[WARNING] Column count doesn't match expected blocks. Treating as single block.")
                    actual_blocks = [expected_block_names[0]]
                    actual_column_groups = [data_columns]
        
        custom_print(f"Detected blocks: {actual_blocks}")
        tcq_percentage_entries = []
        all_not_applicable = True
        
        for idx, block in enumerate(actual_blocks):
            block_columns = actual_column_groups[idx]
            block_data = df.loc[:, block_columns]
            combined_data = pd.concat([df.iloc[:, 0], block_data], axis=1)
            
            block_csv_file = os.path.join(output_dir, f"{block}_tcq_data.csv")
            combined_data.to_csv(block_csv_file, index=False, sep=' ', header=False)
            
            block_df = pd.read_csv(block_csv_file, sep=' ', header=None)
            block_df = block_df.iloc[1:].reset_index(drop=True)
            
            if block_df.empty:
                tcq_result_string = "TCQ Not Applicable"
                tcq_percentage_entries.append(f"{block}: {tcq_result_string}")
                custom_print(f"{block}: TCQ Data not applicable (empty after header removal). Reporting: {tcq_result_string}")
                continue
            
            all_not_applicable = False
            if block_df.shape[1] < 2:
                custom_print(f"[WARNING] Block {block} has fewer than 2 columns. Skipping.")
                tcq_percentage_entries.append(f"{block}: Insufficient data")
                continue
            
            try:
                block_df[1] = pd.to_numeric(block_df[1], errors='coerce')
            except Exception as conv_err:
                custom_print(f"[WARNING] Error converting column 1 to numeric for block {block}: {conv_err}")
                raise
            
            try:
                block_df.iloc[:, -1] = pd.to_numeric(block_df.iloc[:, -1], errors='coerce')
            except Exception as conv_err:
                custom_print(f"[WARNING] Error converting last column to numeric for block {block}: {conv_err}")
                raise
            
            tcq_percentage = (block_df.iloc[:, -1] / block_df[1]) * 100
            block_df['raw_percentage'] = tcq_percentage
            tcq_percentage_formatted = tcq_percentage.apply(lambda x: f"{x:.2f}%" if not pd.isna(x) else "")
            block_df['tcq_percentage'] = tcq_percentage_formatted
            
            filtered_data = block_df[abs(block_df['raw_percentage']) >= 10][[0, 'tcq_percentage', 'raw_percentage']]
            filtered_data.to_csv(block_csv_file, index=False, sep=' ', header=False)
            
            if not filtered_data.empty:
                block_entries = []
                if highest_only == 1:
                    max_percentage_idx = filtered_data['raw_percentage'].abs().idxmax()
                    max_row = filtered_data.loc[max_percentage_idx]
                    corner = max_row[0]
                    percentage = max_row['tcq_percentage']
                    block_entries.append(f"{corner}: {percentage}")
                else:
                    for _, row in filtered_data.iterrows():
                        corner = row[0]
                        percentage = row['tcq_percentage']
                        block_entries.append(f"{corner}: {percentage}")
                
                tcq_percentage_entries.append(f"{block}: " + ", ".join(block_entries))
            else:
                tcq_percentage_entries.append(f"{block}: CLEAN")
        
        if all_not_applicable:
            return "TCQ Not Applicable"
        elif tcq_percentage_entries:
            return " | ".join(tcq_percentage_entries) + "."
        else:
            return "No TCQ data available."
        
    except Exception as e:
        custom_print(f"[WARNING] TCQ Not Applicable: {e}")
        return "TCQ Not Applicable"

def process_min_pulse_width(excel_file, output_dir, highest_only=1):
    try:
        sheet_name = "MIN_PULSE_WIDTH"
        csv_file = os.path.join(output_dir, "MIN_PULSE_WIDTH.csv")
        df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)

        if df.empty:
            return "Empty Sheet"

        df.to_csv(csv_file, sep=' ', header=None, index=False)
        custom_print(f"Sheet '{sheet_name}' successfully converted to '{csv_file}' with space as delimiter.")

        df = pd.read_csv(csv_file, sep=r'\s+', engine='python', header=None)
        corner_col = df.columns[0]
        wns_col = df.columns[1]
        fep_col = df.columns[2]

        df[wns_col] = pd.to_numeric(df[wns_col], errors='coerce')
        df = df.dropna(subset=[wns_col])

        if df.empty:
            return "No valid MIN_PULSE_WIDTH data"

        if (df[wns_col] == 0).all():
            return "CLEAN"

        if highest_only == 1:
            farthest_row = df.loc[df[wns_col].abs().idxmax()]
            if farthest_row[wns_col] == 0:
                return "CLEAN"

            formatted_output = f"{farthest_row[corner_col]} - WNS: {farthest_row[wns_col]}; FEP: {farthest_row[fep_col]}"
            with open(csv_file, "w") as f:
                f.write(formatted_output)

            custom_print(f"File '{csv_file}' has been updated with the farthest WNS row.")
            return formatted_output
        elif highest_only == 0:
            output_lines = []
            with open(csv_file, "w") as f:
                for index, row in df.iterrows():
                    if row[wns_col] != 0:
                        formatted_output = f"{row[corner_col]} - WNS: {row[wns_col]}; FEP: {row[fep_col]}"
                        output_lines.append(formatted_output)
                        f.write(formatted_output + "\n")

            if not output_lines:
                custom_print(f"File '{csv_file}' has been updated (no WNS != 0 rows found).")
                return "CLEAN"
            else:
                custom_print(f"File '{csv_file}' has been updated with rows where WNS is not zero.")
                return "\n".join(output_lines)
        else:
            return "Invalid value for 'highest_only' parameter. Use 0 or 1."

    except Exception as e:
        custom_print(f"[WARNING] Error processing MIN_PULSE_WIDTH data: {e}")
        return "Error processing MIN_PULSE_WIDTH data"

def adjust_column_widths(worksheet, data, sub_headers, main_headers, start_row=2):
    column_widths = {
        0: 15, 1: 15, 2: 15, 3: 15, 4: 15, 5: 60, 6: 40, 7: 40, 8: 55, 9: 15, 10: 15, 11: 15, 12: 15, 13: 15, 14: 15, 15: 15
    }
    for col_index, width in column_widths.items():
        worksheet.set_column(col_index, col_index, width)
    for row_num, file_output in enumerate(data):
        worksheet.set_row(row_num + 2, None, None, {'hidden': False, 'level': 0, 'collapsed': False})

def create_output_excel(all_output_data, sub_headers, main_headers, blocks_comp_names, blocks_owners, proj_dir_path, output_file=Output_xls_name):
    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        updated_main_headers = ["Compiler", "Block Name", "Block Owner", "Dashboard", "Dashboard", "Dashboard",
                               "HOLD", "FMAX", "DRV", "TCQ", "MPW",
                               "PHYSICAL VERIFICATION", "PHYSICAL VERIFICATION", "PHYSICAL VERIFICATION", "PHYSICAL VERIFICATION",
                               "IR DROP", "IR DROP", "Formality"]
        updated_sub_headers = ["", "", "", "PARA ERRORS", "NOT ANNOTATED", "MPW VIOLATION",
                              "Clk_groups", "", "", "", "",
                              "DRC", "LVS", "ERC", "ANT",
                              "VDD", "VSS", ""]
        
        updated_output_data = []
        for row_data in all_output_data:
            if row_data == ["Error processing file"]:
                continue
            block_name = row_data[0] if row_data[0] != "File Not Found" else row_data[1]
            if row_data[0] == "File Not Found":
                new_row = []
                comp_name = blocks_comp_names.get(block_name, "N/A")
                block_owner = blocks_owners.get(block_name, "N/A")
                new_row.append(comp_name)
                new_row.append(block_name)
                new_row.append(block_owner)
                new_row.extend(["File Not Found"] * 14)
                updated_output_data.append(new_row)
                continue
            
            comp_name = blocks_comp_names.get(block_name, "N/A")
            block_owner = blocks_owners.get(block_name, "N/A")
            new_row = [comp_name, block_name, block_owner]
            if len(row_data) >= 16:
                new_row.extend(row_data[1:16])
            else:
                new_row.extend(row_data[1:])
                missing_cols = 18 - len(new_row)
                if missing_cols > 0:
                    new_row.extend(["N/A"] * missing_cols)
            updated_output_data.append(new_row)

        if updated_output_data:
            output_df = pd.DataFrame(updated_output_data, columns=updated_main_headers)
            output_df.to_excel(writer, sheet_name="Summary", index=False, startrow=2, header=False)
            workbook = writer.book
            worksheet = writer.sheets["Summary"]

            header_format = workbook.add_format({"bold": True, "align": "center", "valign": "vcenter", "border": 2, "bg_color": "#6EACDA", "text_wrap": True})
            subheader_format = workbook.add_format({"bold": True, "align": "center", "valign": "vcenter", "border": 2, "bg_color": "#F8CBAD", "text_wrap": True})
            cell_format = workbook.add_format({"align": "center", "valign": "vcenter", "border": 1, "text_wrap": True})
            clean_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100', "align": "center", "valign": "vcenter", "border": 1, "text_wrap": True})
            not_applicable_format = workbook.add_format({'bg_color': '#FFACAC', 'font_color': '#BE3144', "align": "center", "valign": "vcenter", "border": 1, "text_wrap": True})
            light_orange_format = workbook.add_format({'bg_color': '#FFE0B2', 'font_color': '#000000', "align": "center", "valign": "vcenter", "border": 1, "text_wrap": True})
            light_red_format = workbook.add_format({'bg_color': '#FFCDD2', 'font_color': '#000000', "align": "center", "valign": "vcenter", "border": 1, "text_wrap": True})
            light_green_format = workbook.add_format({'bg_color': '#C8E6C9', 'font_color': '#006100', "align": "center", "valign": "vcenter", "border": 1, "text_wrap": True})
            file_not_found_format = workbook.add_format({'bg_color': '#FFACAC', 'font_color': '#BE3144', "align": "center", "valign": "vcenter", "border": 1, "text_wrap": True})
            block_name_format = workbook.add_format({'bg_color': '#E6E6FA', 'font_color': '#4B0082', 'bold': True, 'italic': False, 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'font_size': 10})

            worksheet.set_row(0, 30)
            worksheet.set_row(1, 30)

            worksheet.merge_range(0, 0, 1, 0, updated_main_headers[0], header_format)
            worksheet.merge_range(0, 1, 1, 1, updated_main_headers[1], header_format)
            worksheet.merge_range(0, 2, 1, 2, updated_main_headers[2], header_format)
            worksheet.merge_range(0, 3, 0, 5, updated_main_headers[3], header_format)
            worksheet.merge_range(0, 6, 1, 6, updated_main_headers[6], header_format)
            worksheet.merge_range(0, 7, 1, 7, updated_main_headers[7], header_format)
            worksheet.merge_range(0, 8, 1, 8, updated_main_headers[8], header_format)
            worksheet.merge_range(0, 9, 1, 9, updated_main_headers[9], header_format)
            worksheet.merge_range(0, 10, 0, 13, updated_main_headers[10], header_format)
            worksheet.merge_range(0, 14, 0, 15, updated_main_headers[14], header_format)
            worksheet.merge_range(0, 16, 1, 16, updated_main_headers[16], header_format)

            worksheet.write(1, 3, updated_sub_headers[3], subheader_format)
            worksheet.write(1, 4, updated_sub_headers[4], subheader_format)
            worksheet.write(1, 5, updated_sub_headers[5], subheader_format)
            worksheet.write(1, 6, updated_sub_headers[6], subheader_format)
            worksheet.write(1, 10, updated_sub_headers[10], subheader_format)
            worksheet.write(1, 11, updated_sub_headers[11], subheader_format)
            worksheet.write(1, 12, updated_sub_headers[12], subheader_format)
            worksheet.write(1, 13, updated_sub_headers[13], subheader_format)
            worksheet.write(1, 14, updated_sub_headers[14], subheader_format)
            worksheet.write(1, 15, updated_sub_headers[15], subheader_format)

            worksheet.autofilter(1, 0, 1, len(updated_main_headers) - 1)
            worksheet.freeze_panes(2, 3)

            formatting_rules = {
                3: {"CLEAN": clean_format, "NOT CLEAN": not_applicable_format, "FILE NOT FOUND": file_not_found_format},
                4: {"CLEAN": clean_format, "NOT CLEAN": not_applicable_format, "FILE NOT FOUND": file_not_found_format},
                5: {"CLEAN": clean_format, "NOT CLEAN": not_applicable_format, "FILE NOT FOUND": file_not_found_format},
                6: {"HOLD CLEAN": clean_format, "FILE NOT FOUND": file_not_found_format},
                7: {"ALL TCC": clean_format, "FMAX NOT APPLICABLE": not_applicable_format, "FMAX SHEET NOT FOUND.": not_applicable_format,
                    "ERROR PROCESSING FMAX DATA": not_applicable_format, "FILE NOT FOUND": file_not_found_format},
                8: {"TRAN: CLEAN | CAP: CLEAN": clean_format, "FILE NOT FOUND": file_not_found_format},
                9: {"TCQ NOT APPLICABLE": not_applicable_format, "EMPTY SHEET": not_applicable_format, "FILE NOT FOUND": file_not_found_format, "OR": True},
                10: {"CLEAN": clean_format, "EMPTY SHEET": not_applicable_format, "NO VALID MIN_PULSE_WIDTH DATA": not_applicable_format, "FILE NOT FOUND": file_not_found_format},
                11: {"CLEAN": clean_format, "NOT CLEAN": not_applicable_format, "DRC FILE NOT FOUND": not_applicable_format, "FILE NOT FOUND": file_not_found_format},
                12: {"CLEAN": clean_format, "NOT CLEAN": not_applicable_format, "LVS FILE NOT FOUND": not_applicable_format, "FILE NOT FOUND": file_not_found_format},
                13: {"CLEAN": clean_format, "NOT CLEAN": not_applicable_format, "ERC FILE NOT FOUND": not_applicable_format,
                     "ERC FILE EMPTY OR LESS THAN 11 LINES": not_applicable_format, "FILE NOT FOUND": file_not_found_format, "OR": True},
                14: {"CLEAN": clean_format, "NOT CLEAN": not_applicable_format, "ANT FILE NOT FOUND": not_applicable_format, "FILE NOT FOUND": file_not_found_format},
                15: {"CLEAN": clean_format, "Vol*.rpt File Not Found": not_applicable_format, "Error processing IR value": not_applicable_format, "FILE NOT FOUND": file_not_found_format, "OR": True},
                16: {"CLEAN": clean_format, "Vol*.rpt File Not Found": not_applicable_format, "Error processing IR value": not_applicable_format, "FILE NOT FOUND": file_not_found_format, "OR": True},
                17: {"PASSING": clean_format, "NOT PASSING": not_applicable_format, "LOG FILE NOT FOUND": not_applicable_format, "ERROR PROCESSING LOG": not_applicable_format, "FILE NOT FOUND": file_not_found_format}
            }

            for row_num, file_output in enumerate(updated_output_data):
                for col_num, value in enumerate(file_output):
                    fmt = cell_format
                    val_upper = str(value).upper()

                    if col_num == 1:
                        worksheet.write(row_num + 2, col_num, value, block_name_format)
                        continue

                    if val_upper == "FILE NOT FOUND":
                        worksheet.write(row_num + 2, col_num, value, file_not_found_format)
                        continue

                    if col_num == 7:
                        if isinstance(value, str):
                            parts = [part.strip() for part in value.split('|')]
                            all_tcc = len(parts) > 0 and all(part.endswith(": TCC") or part.endswith(": TCC.") for part in parts)
                            if all_tcc:
                                worksheet.write(row_num + 2, col_num, value, clean_format)
                                continue
                            elif val_upper in formatting_rules[7]:
                                fmt = formatting_rules[7][val_upper]
                    elif col_num == 9:
                        if isinstance(value, str):
                            if "TCQ NOT APPLICABLE" in val_upper or "EMPTY SHEET" in val_upper:
                                fmt = not_applicable_format
                            else:
                                parts = [part.strip() for part in value.split('|')]
                                all_clean = len(parts) > 0 and all(": CLEAN" in part.upper() for part in parts)
                                if all_clean:
                                    worksheet.write(row_num + 2, col_num, value, light_green_format)
                                    continue

                    for cols, conditions in formatting_rules.items():
                        if isinstance(cols, int) and col_num == cols:
                            if val_upper in conditions:
                                fmt = conditions[val_upper]
                            elif "OR" in conditions and any(val_upper == cond.upper() for cond in conditions if cond != "OR"):
                                fmt = conditions[next(cond for cond in conditions if cond != "OR" and val_upper == cond.upper())]
                            elif "AND" in conditions:
                                all_conditions_met = True
                                for cond_part in conditions:
                                    if cond_part != "AND" and cond_part.split(":")[0] in val_upper and cond_part.split(":")[1] in val_upper:
                                        continue
                                    elif cond_part != "AND":
                                        all_conditions_met = False
                                        break
                                if all_conditions_met:
                                    fmt = conditions[next(cond for cond in conditions if cond != "AND" and cond.split(":")[0] in val_upper and cond.split(":")[1] in val_upper)]
                        elif isinstance(cols, tuple) and col_num in cols:
                            if val_upper in conditions:
                                fmt = conditions[val_upper]

                    worksheet.write(row_num + 2, col_num, value, fmt)

                try:
                    vdd_value = updated_output_data[row_num][15]
                    if isinstance(vdd_value, str) and vdd_value.endswith("%"):
                        percentage = float(vdd_value[:-1])
                        if percentage <= 1.0:
                            worksheet.write(row_num + 2, 15, vdd_value, light_green_format)
                        elif percentage <= 1.5:
                            worksheet.write(row_num + 2, 15, vdd_value, light_orange_format)
                        elif percentage > 1.6:
                            worksheet.write(row_num + 2, 15, vdd_value, light_red_format)
                except (ValueError, IndexError):
                    pass

                try:
                    vss_value = updated_output_data[row_num][16]
                    if isinstance(vss_value, str) and vss_value.endswith("%"):
                        percentage = float(vss_value[:-1])
                        if percentage <= 1.0:
                            worksheet.write(row_num + 2, 16, vss_value, light_green_format)
                        elif percentage <= 1.5:
                            worksheet.write(row_num + 2, 16, vss_value, light_orange_format)
                        elif percentage > 1.6:
                            worksheet.write(row_num + 2, 16, vss_value, light_red_format)
                except (ValueError, IndexError):
                    pass

                for col_num in range(len(file_output)):
                    value = file_output[col_num]
                    if isinstance(value, str) and ": CLEAN" in value and "|" in value:
                        parts = [part.strip() for part in value.split('|')]
                        all_clean = len(parts) > 0 and all(": CLEAN" in part.upper() for part in parts)
                        if all_clean:
                            worksheet.write(row_num + 2, col_num, value, light_green_format)

                worksheet.set_row(row_num + 2, None, None, {'hidden': False, 'level': 0, 'collapsed': False})

            column_widths = {
                0: 20, 1: 15, 2: 15, 3: 15, 4: 15, 5: 60, 6: 40, 7: 40, 8: 55, 9: 15, 10: 15, 11: 15, 12: 15, 13: 15, 14: 15, 15: 15
            }
            for col_index, width in column_widths.items():
                worksheet.set_column(col_index, col_index, width)

            custom_print(f"[CREATED] Saved consolidated output to {output_file} with fixed column widths, dynamic row heights, and fixed headers.")
        else:
            custom_print("[WARNING] No data to write to Excel. Skipping.")

if __name__ == "__main__":
    try:
        toggle_print(False)
        builtins.custom_print = custom_print
        print_header()
        time.sleep(3)

        main_headers = ["Compiler", "Block Name", "Block Owner", "Dashboard", "Dashboard", "Dashboard", "HOLD", "FMAX", "DRV", "TCQ", "MPW", "PHYSICAL VERIFICATION", "PHYSICAL VERIFICATION", "PHYSICAL VERIFICATION", "PHYSICAL VERIFICATION", "IR DROP", "IR DROP", "Formality"]
        sub_headers = ["", "", "", "PARA ERRORS", "NOT ANNOTATED", "MPW VIOLATION", "Clk_groups", "", "", "", "", "DRC", "LVS", "ERC", "ANT", "VDD", "VSS", ""]

        if not os.path.exists(ALL_BLOCK_CSV_FILES_DIR):
            os.makedirs(ALL_BLOCK_CSV_FILES_DIR)
            custom_print(f"[CREATED] Created main directory: {ALL_BLOCK_CSV_FILES_DIR}")
        else:
            custom_print(f"[WARNING] Main directory already exists: {ALL_BLOCK_CSV_FILES_DIR}. Files might be overwritten.")

        block_info = [
            {"block_name": "CDM_top", "compiler": "Generic", "owner": "Sai Kumar"},
            {"block_name": "i36_i50", "compiler": "HSPRAM", "owner": "John Doe"},
            {"block_name": "i36_i50_i12", "compiler": "HS1PRF", "owner": "Alice Johnson"}
        ]

        i_blocks = [block for block in block_info if block["block_name"].lower().startswith('i')]
        i_blocks_sorted = sorted(i_blocks, key=lambda x: x["block_name"].lower())
        non_i_blocks = [block for block in block_info if not block["block_name"].lower().startswith('i')]
        non_i_blocks_sorted = sorted(non_i_blocks, key=lambda x: x["block_name"].lower())
        block_info_sorted = i_blocks_sorted + non_i_blocks_sorted
        custom_print("[DEBUG] Sorted block_info:", [block["block_name"] for block in block_info_sorted])

        blocks_comp_names = {block["block_name"]: block["compiler"] for block in block_info_sorted}
        blocks_owners = {block["block_name"]: block["owner"] for block in block_info_sorted}
        
        all_output_data = []
        for block in block_info_sorted:
            excel_file = f"{block['block_name']}_metrics.xlsx"
            if os.path.exists(excel_file):
                custom_print(f"Processing Excel file: {excel_file}")
                output_data = process_excel_file(excel_file)
                custom_print(f"Output for {block['block_name']}: {output_data}")
                if output_data and output_data[0] != "Error processing file":
                    all_output_data.append(output_data[0])
                else:
                    all_output_data.append(["File Not Found", block['block_name']])
            else:
                custom_print(f"File not found: {excel_file}")
                all_output_data.append(["File Not Found", block['block_name']])
        
        custom_print("Final all_output_data:", all_output_data)
        print("    [ Done Processing! ]", flush=True)

        print("\n" + "=" * 100, flush=True)
        print("                                     \033[93mFINAL PROCESSING SUMMARY\033[0m          ", flush=True)
        print("=" * 100, flush=True)

        valid_files = [f"{b['block_name']}_metrics.xlsx" for b in block_info_sorted 
                      if os.path.exists(f"{b['block_name']}_metrics.xlsx")]
        failed_files = [f"{b['block_name']}_metrics.xlsx" for b in block_info_sorted 
                       if not os.path.exists(f"{b['block_name']}_metrics.xlsx")]

        custom_print(f"Valid files: {valid_files}", flush=True)
        custom_print(f"Failed files: {failed_files}", flush=True)

        print("Processing Results:", flush=True)
        for file in valid_files:
            print(f" \033[32m+ VALID\033[0m: {file}", flush=True)
        for file in failed_files:
            print(f" \033[31m- FAILED\033[0m: {file}", flush=True)

        print("\nStatistics:", flush=True)
        print(f"Total files processed: {len(block_info_sorted)}", flush=True)
        if len(valid_files) != 0:
            print(f"Successfully processed: \033[32m{len(valid_files)}\033[0m", flush=True)
        else:
            print(f"Successfully processed: {len(valid_files)}", flush=True)
        if len(failed_files) != 0:
            print(f"Failed to process: \033[31m{len(failed_files)}\033[0m", flush=True)
        else:
            print(f"Failed to process: {len(failed_files)}", flush=True)

        if len(failed_files) == 0:
            print("\n" + "*" * 100, flush=True)
            print(f"\033[1;32mSUCCESS\033[0m: All files processed successfully!", flush=True)
            print("*" * 100, flush=True)
            print(f"\033[1;31mNOTE\033[0m: Please check the output file '{Output_xls_name}' for the processed results.", flush=True)
        else:
            print("\n" + "!" * 100, flush=True)
            print("\033[31mNOTE\033[0m: Some files could not be processed. Please verify:", flush=True)
            print("- Block names are correct", flush=True)
            print("- Excel files exist in the expected directory", flush=True)
            print("!" * 100, flush=True)

        print("=" * 100 + "\n", flush=True)
        create_output_excel(all_output_data, sub_headers, main_headers, blocks_comp_names, blocks_owners, proj_dir_path)

    except Exception as e:
        print(f"Error during execution: {e}", flush=True)