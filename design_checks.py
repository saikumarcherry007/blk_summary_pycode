"""
Design Checks Module
Contains functions for processing DRV, IR drop, and formality checks
"""

import os
import glob
import pandas as pd
from utils import custom_print


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
