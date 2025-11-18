"""
Excel Processor Module
Contains functions for processing Excel files and creating output reports
"""

import os
import pandas as pd
from config import ALL_BLOCK_CSV_FILES_DIR, proj_dir_path, COLUMN_WIDTHS, Output_xls_name
from utils import custom_print, check_clean_status
from physical_verification import process_drc_value, process_lvs_value, process_erc_value, process_ant_value
from timing_analysis import process_hold_data, process_fmax_data, process_tcq_data, process_min_pulse_width
from design_checks import process_drv_data, process_ir_value_to_csv, process_formality_value


def process_excel_file(excel_file, main_headers, sub_headers):

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

        drc_value = process_drc_value(excel_file, proj_dir_path)
        lvs_value = process_lvs_value(excel_file, proj_dir_path)
        erc_value = process_erc_value(excel_file, proj_dir_path)
        ant_value = process_ant_value(excel_file, proj_dir_path)

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


def create_output_excel(all_output_data, sub_headers, main_headers, blocks_comp_names, blocks_owners, output_file=Output_xls_name):

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

            # Define cell formats
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

            # Merge cells for main headers
            worksheet.merge_range(0, 0, 1, 0, updated_main_headers[0], header_format)
            worksheet.merge_range(0, 1, 1, 1, updated_main_headers[1], header_format)
            worksheet.merge_range(0, 2, 1, 2, updated_main_headers[2], header_format)
            worksheet.merge_range(0, 3, 0, 5, updated_main_headers[3], header_format)
            worksheet.merge_range(0, 6, 1, 6, updated_main_headers[6], header_format)
            worksheet.merge_range(0, 7, 1, 7, updated_main_headers[7], header_format)
            worksheet.merge_range(0, 8, 1, 8, updated_main_headers[8], header_format)
            worksheet.merge_range(0, 9, 1, 9, updated_main_headers[9], header_format)
            worksheet.merge_range(0, 10, 1, 10, updated_main_headers[10], header_format)
            worksheet.merge_range(0, 11, 0, 14, updated_main_headers[11], header_format)
            worksheet.merge_range(0, 15, 0, 16, updated_main_headers[15], header_format)
            worksheet.merge_range(0, 17, 1, 17, updated_main_headers[17], header_format)

            # Write sub-headers
            worksheet.write(1, 3, updated_sub_headers[3], subheader_format)
            worksheet.write(1, 4, updated_sub_headers[4], subheader_format)
            worksheet.write(1, 5, updated_sub_headers[5], subheader_format)
            worksheet.write(1, 6, updated_sub_headers[6], subheader_format)
            worksheet.write(1, 11, updated_sub_headers[11], subheader_format)
            worksheet.write(1, 12, updated_sub_headers[12], subheader_format)
            worksheet.write(1, 13, updated_sub_headers[13], subheader_format)
            worksheet.write(1, 14, updated_sub_headers[14], subheader_format)
            worksheet.write(1, 15, updated_sub_headers[15], subheader_format)
            worksheet.write(1, 16, updated_sub_headers[16], subheader_format)

            worksheet.autofilter(1, 0, 1, len(updated_main_headers) - 1)
            worksheet.freeze_panes(2, 3)

            # Formatting rules for different columns
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

            # Apply formatting to data cells
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

                    worksheet.write(row_num + 2, col_num, value, fmt)

                # Apply special formatting for VDD/VSS percentage values
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

                # Check for multiple CLEAN entries
                for col_num in range(len(file_output)):
                    value = file_output[col_num]
                    if isinstance(value, str) and ": CLEAN" in value and "|" in value:
                        parts = [part.strip() for part in value.split('|')]
                        all_clean = len(parts) > 0 and all(": CLEAN" in part.upper() for part in parts)
                        if all_clean:
                            worksheet.write(row_num + 2, col_num, value, light_green_format)

                worksheet.set_row(row_num + 2, None, None, {'hidden': False, 'level': 0, 'collapsed': False})

            # Set column widths
            for col_index, width in COLUMN_WIDTHS.items():
                worksheet.set_column(col_index, col_index, width)

            custom_print(f"[CREATED] Saved consolidated output to {output_file} with fixed column widths, dynamic row heights, and fixed headers.")
        else:
            custom_print("[WARNING] No data to write to Excel. Skipping.")
