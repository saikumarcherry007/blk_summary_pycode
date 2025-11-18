import os
import time
import builtins
from config import (
    ALL_BLOCK_CSV_FILES_DIR, 
    proj_dir_path, 
    MAIN_HEADERS, 
    SUB_HEADERS, 
    BLOCK_INFO
)
from utils import toggle_print, print_header, custom_print
from excel_processor import process_excel_file, create_output_excel


def main():

    try:
        # Initialize print control
        toggle_print(False)
        builtins.custom_print = custom_print
        print_header()
        time.sleep(3)

        # Create output directory if it doesn't exist
        if not os.path.exists(ALL_BLOCK_CSV_FILES_DIR):
            os.makedirs(ALL_BLOCK_CSV_FILES_DIR)
            custom_print(f"[CREATED] Created main directory: {ALL_BLOCK_CSV_FILES_DIR}")
        else:
            custom_print(f"[WARNING] Main directory already exists: {ALL_BLOCK_CSV_FILES_DIR}. Files might be overwritten.")

        # Sort block information: i-blocks first, then others
        i_blocks = [block for block in BLOCK_INFO if block["block_name"].lower().startswith('i')]
        i_blocks_sorted = sorted(i_blocks, key=lambda x: x["block_name"].lower())
        non_i_blocks = [block for block in BLOCK_INFO if not block["block_name"].lower().startswith('i')]
        non_i_blocks_sorted = sorted(non_i_blocks, key=lambda x: x["block_name"].lower())
        block_info_sorted = i_blocks_sorted + non_i_blocks_sorted
        custom_print("[DEBUG] Sorted block_info:", [block["block_name"] for block in block_info_sorted])

        # Create dictionaries for block metadata
        blocks_comp_names = {block["block_name"]: block["compiler"] for block in block_info_sorted}
        blocks_owners = {block["block_name"]: block["owner"] for block in block_info_sorted}
        
        # Process each Excel file
        all_output_data = []
        for block in block_info_sorted:
            excel_file = f"{block['block_name']}_metrics.xlsx"
            if os.path.exists(excel_file):
                custom_print(f"Processing Excel file: {excel_file}")
                output_data = process_excel_file(excel_file, MAIN_HEADERS, SUB_HEADERS)
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

        # Print processing summary
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

        # Print final status
        if len(failed_files) == 0:
            print("\n" + "*" * 100, flush=True)
            print(f"\033[1;32mSUCCESS\033[0m: All files processed successfully!", flush=True)
            print("*" * 100, flush=True)
            from config import Output_xls_name
            print(f"\033[1;31mNOTE\033[0m: Please check the output file '{Output_xls_name}' for the processed results.", flush=True)
        else:
            print("\n" + "!" * 100, flush=True)
            print("\033[31mNOTE\033[0m: Some files could not be processed. Please verify:", flush=True)
            print("- Block names are correct", flush=True)
            print("- Excel files exist in the expected directory", flush=True)
            print("!" * 100, flush=True)

        print("=" * 100 + "\n", flush=True)
        
        # Create consolidated Excel output
        create_output_excel(all_output_data, SUB_HEADERS, MAIN_HEADERS, blocks_comp_names, blocks_owners)

    except Exception as e:
        print(f"Error during execution: {e}", flush=True)
        raise


if __name__ == "__main__":
    main()
