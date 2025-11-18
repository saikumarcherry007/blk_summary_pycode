def process_fmax_data(excel_file, clock_groups, xls, output_dir, highest_only=1):
    try:
        excel_filename_without_ext = os.path.splitext(os.path.basename(excel_file))[0]
        base_name = excel_filename_without_ext.replace('_metrics', '')  # Handle potential '_metrics' suffix

        if base_name in ["CDM_top", "PLL", "setuphold", "clk_jtag_pll_cntrl"]:
            print(f"[INFO] FMAX sheet marked as Not Applicable for file: {excel_file}")
            return "FMAX Not Applicable"

        if 'FMAX' in xls.sheet_names:
            print("FMAX sheet is present.")
            df_fmax = pd.read_excel(excel_file, sheet_name='FMAX')

            excel_filename_without_ext = os.path.splitext(os.path.basename(excel_file))[0]
            csv_file = os.path.join(output_dir, f"FMAX_{excel_filename_without_ext}.csv")
            df_fmax.to_csv(csv_file, index=False, sep=' ', header=False)
            print(f"[CREATED] Sheet 'FMAX' converted to CSV: {csv_file} (space delimited, no header)")

            df_fmax = pd.read_csv(csv_file, delimiter=' ', header=None)
            df_fmax.iloc[:, 0] = df_fmax.iloc[:, 0].fillna('')
            filtered_fmax = df_fmax[df_fmax.iloc[:, 0].str.startswith(('func', 'test', 'fbist'))]
            filtered_fmax.to_csv(csv_file, index=False, sep=' ', header=False)
            print(f"[CREATED] Main CSV file '{csv_file}' has been filtered.")

            raw_blocks = excel_filename_without_ext.split('_')
            blocks = [block for block in raw_blocks if block != "metrics"]
            print(f"Identified blocks from filename: {blocks}")

            part_size = 7
            fmax_results = []
            
            # Keep track of which blocks have TCC data
            tcc_blocks = []

            for idx, part in enumerate(blocks):
                new_csv_file = os.path.join(output_dir, f"{part}_fmax.csv")
                print(f"Creating CSV file for block {part}: {new_csv_file}")

                start_col = 1 + idx * part_size
                end_col = start_col + part_size

                if start_col >= filtered_fmax.shape[1]:
                    print(f"[WARNING] Not enough columns for {part}.")
                    fmax_results.append(f"{part}: Not enough data")
                    continue

                columns_to_extract = [0] + list(range(start_col, min(end_col, filtered_fmax.shape[1])))
                part_data = filtered_fmax.iloc[:, columns_to_extract].copy()

                if part_data.shape[1] > 1:
                    mask = part_data.iloc[:, 1] != "-"
                    part_data = part_data[mask]

                part_data.to_csv(new_csv_file, index=False, sep=' ', header=False)
                print(f"[CREATED] Created {new_csv_file} with filtered data")

                if part_data.shape[1] >= 4:
                    limit_col_idx = min(part_data.shape[1] - 3, part_data.shape[1] - 1)
                    tccmargin_col_idx = min(part_data.shape[1] - 2, part_data.shape[1] - 1)
                    holdmargin_col_idx = min(part_data.shape[1] - 1, part_data.shape[1] - 1)

                    if part_data.empty:
                        print(f"[WARNING] No valid data for {part} after filtering.")
                        fmax_results.append(f"{part}: No valid data")
                        continue

                    limit_col = part_data.iloc[:, limit_col_idx]
                    
                    # Debug: Print the unique values in limit_col to check what's available
                    print(f"[DEBUG] Unique values in limit_col for {part}: {limit_col.unique()}")
                    
                    # This line might be filtering out valid data if regex pattern doesn't match
                    filtered_part_data = part_data[limit_col.str.contains('Memory|SMS', na=False, regex=True)]
                    
                    # Debug: Check how many rows are filtered
                    print(f"[DEBUG] Original rows: {len(part_data)}, Filtered rows: {len(filtered_part_data)}")
                    
                    # If filtered_part_data is empty, we may need to adjust the filter
                    if filtered_part_data.empty:
                        print(f"[DEBUG] No rows matched 'Memory|SMS' filter for {part}. Using all data instead.")
                        filtered_part_data = part_data  # Use all data if no Memory/SMS specific data is found

                    highest_sms = {"value": -1, "corner": "", "limit_value": "", "hold_margin": ""}
                    highest_memory = {"value": -1, "corner": "", "limit_value": "", "hold_margin": ""}
                    # Add a general highest for any other categories
                    highest_general = {"value": -1, "corner": "", "limit_value": "", "hold_margin": ""}

                    all_lines = []

                    # Check if all rows have 'TCC' in their limit_value
                    has_tcc = False
                    for idx_row, row in filtered_part_data.iterrows():
                        limit_value = row.iloc[limit_col_idx] if limit_col_idx < len(row) else "Unknown"
                        if "TCC" in str(limit_value):
                            has_tcc = True
                            break
                    
                    if has_tcc:
                        tcc_blocks.append(part)
                    
                    print(f"[DEBUG] Block {part} has TCC: {has_tcc}")

                    for idx_row, row in filtered_part_data.iterrows():
                        corner = row.iloc[0]
                        limit_value = row.iloc[limit_col_idx] if limit_col_idx < len(row) else "Unknown"
                        tcc_margin = row.iloc[tccmargin_col_idx] if tccmargin_col_idx < len(row) and pd.notna(row.iloc[tccmargin_col_idx]) else 'NA'
                        hold_margin = row.iloc[holdmargin_col_idx] if holdmargin_col_idx < len(row) and pd.notna(row.iloc[holdmargin_col_idx]) else 'NA'

                        # FIX: Remove any existing % symbol before adding one
                        tcc_margin_str = str(tcc_margin).rstrip('%')
                        hold_margin_str = str(hold_margin).rstrip('%')

                        # Consistently use TCC in the CSV files
                        if "TCC" in str(limit_value):
                            formatted_line = f"{corner} (TCC: {tcc_margin_str}%); Hold_margin: {hold_margin_str}"
                        else:
                            formatted_line = f"{corner} ({limit_value}: {tcc_margin_str}%); Hold_margin: {hold_margin_str}"
                        
                        all_lines.append(formatted_line)
                        
                        # Debug the values being processed
                        print(f"[DEBUG] Processing row - Corner: {corner}, Limit: {limit_value}, Margin: {tcc_margin_str}, Hold: {hold_margin_str}")

                        try:
                            # Strip % when converting to float
                            margin_value = float(str(tcc_margin).rstrip('%'))

                            if "SMS" in str(limit_value) and margin_value > highest_sms["value"]:
                                highest_sms = {
                                    "value": margin_value,
                                    "corner": corner,
                                    "limit_value": limit_value,
                                    "hold_margin": hold_margin_str
                                }
                                print(f"[DEBUG] New highest SMS: {margin_value}% for {corner}")
                            elif "Memory" in str(limit_value) and margin_value > highest_memory["value"]:
                                highest_memory = {
                                    "value": margin_value,
                                    "corner": corner,
                                    "limit_value": limit_value,
                                    "hold_margin": hold_margin_str
                                }
                                print(f"[DEBUG] New highest Memory: {margin_value}% for {corner}")
                            # Add a catch-all for any other types
                            elif margin_value > highest_general["value"]:
                                highest_general = {
                                    "value": margin_value,
                                    "corner": corner,
                                    "limit_value": limit_value,
                                    "hold_margin": hold_margin_str
                                }
                                print(f"[DEBUG] New highest general: {margin_value}% for {corner} ({limit_value})")
                        except (ValueError, TypeError) as e:
                            print(f"[DEBUG] Error converting margin value '{tcc_margin}' to float: {e}")
                            continue

                    # Write all data to CSV file regardless of highest_only setting
                    with open(new_csv_file, 'w') as outfile:
                        outfile.write('\n'.join(all_lines) if all_lines else f"{part}: No data")

                    # But decide what to add to results based on highest_only setting
                    if highest_only == 0:
                        # Add all lines to the results
                        if all_lines:
                            fmax_results.append(f"{part}: {', '.join(all_lines)}")
                        else:
                            fmax_results.append(f"{part}: No data")
                    else:
                        # Only add highest values to results
                        part_summary = []
                        if highest_sms["value"] > -1:
                            if "TCC" in str(highest_sms["limit_value"]):
                                sms_line = f"{highest_sms['corner']} (TCC: {highest_sms['value']}%); Hold_margin: {highest_sms['hold_margin']}"
                            else:
                                sms_line = f"{highest_sms['corner']} ({highest_sms['limit_value']}: {highest_sms['value']}%); Hold_margin: {highest_sms['hold_margin']}"
                            part_summary.append(sms_line)

                        if highest_memory["value"] > -1:
                            if "TCC" in str(highest_memory["limit_value"]):
                                memory_line = f"{highest_memory['corner']} (TCC: {highest_memory['value']}%); Hold_margin: {highest_memory['hold_margin']}"
                            else:
                                memory_line = f"{highest_memory['corner']} ({highest_memory['limit_value']}: {highest_memory['value']}%); Hold_margin: {highest_memory['hold_margin']}"
                            part_summary.append(memory_line)
                            
                        # Also include highest general if we didn't find specific SMS or Memory types
                        if len(part_summary) == 0 and highest_general["value"] > -1:
                            if "TCC" in str(highest_general["limit_value"]):
                                general_line = f"{highest_general['corner']} (TCC: {highest_general['value']}%); Hold_margin: {highest_general['hold_margin']}"
                            else:
                                general_line = f"{highest_general['corner']} ({highest_general['limit_value']}: {highest_general['value']}%); Hold_margin: {highest_general['hold_margin']}"
                            part_summary.append(general_line)

                        print(f"[DEBUG] Part summary for {part}: {part_summary}")
                        
                        if part_summary:
                            fmax_results.append(f"{part}: {', '.join(part_summary)}")
                        else:
                            # If we have lines but no summary, it means our categorization might be wrong
                            if all_lines:
                                print(f"[DEBUG] Warning: Have {len(all_lines)} lines but empty part_summary for {part}")
                                # As a fallback, use the first line
                                fmax_results.append(f"{part}: {all_lines[0]}")
                            else:
                                fmax_results.append(f"{part}: No valid margin data")
                else:
                    print(f"[WARNING] {part} does not have enough columns (needs at least 4).")
                    fmax_results.append(f"{part}: Not enough columns")

            # Modified section: Instead of using "ALL TCC", list specific blocks with "TCC"
            if len(tcc_blocks) > 1:
                print(f"[INFO] Multiple blocks with TCC data found: {tcc_blocks}")
                
                # Create TCC-specific results for each block that has TCC data
                tcc_results = []
                for block in tcc_blocks:
                    tcc_results.append(f"{block}: TCC")
                
                # Join them with the pipe separator
                tcc_final_result = " | ".join(tcc_results)
                print(f"[INFO] Using specific TCC blocks in result: {tcc_final_result}")
                return tcc_final_result
            
            return " | ".join(fmax_results) + "."
        else:
            print("FMAX sheet is not present.")
            return "FMAX sheet not found."
    except Exception as e:
        print(f"[WARNING] Error processing FMAX data: {e}")
        return f"Error processing FMAX data: {e}"
