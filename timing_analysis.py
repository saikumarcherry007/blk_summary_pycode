"""
Timing Analysis Module
Contains functions for processing timing-related metrics: HOLD, FMAX, TCQ, and MIN_PULSE_WIDTH
"""

import os
import pandas as pd
import numpy as np
from utils import custom_print


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
