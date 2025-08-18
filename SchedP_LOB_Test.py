import xml.etree.ElementTree as ET
import os
from typing import List, Dict, Any, Optional
import logging
import glob

# --- Setup basic logging ---
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

def get_cell_data(row_element: ET.Element, cell_index: int, ns: Dict[str, str]) -> Optional[str]:
    """
    Finds a cell in a row by its logical index and returns its text data.
    """
    if cell_index is None or row_element is None:
        return None

    current_index = 1
    for cell in row_element.findall('ss:Cell', ns):
        idx_attr = cell.get(f'{{{ns["ss"]}}}Index')
        if idx_attr:
            current_index = int(idx_attr)

        if current_index == cell_index:
            data_element = cell.find('ss:Data', ns)
            if data_element is not None and data_element.text:
                return data_element.text.strip()
            return None

        current_index += 1
    return None

def test_schedule_p_lob_identification(worksheet: ET.Element, ns: Dict[str, str]):
    """
    Tests a single worksheet to see if it's a valid Schedule P LOB report (AL or APD).
    """
    sheet_name = worksheet.get(f'{{{ns["ss"]}}}Name')
    all_rows = list(worksheet.findall('.//ss:Row', ns))
    
    if len(all_rows) < 3:
        print(f"  - Status: SKIPPED (Not enough rows to check header)")
        return

    # --- Logic Check 1: Check the header text in the third row ---
    header_row_text = get_cell_data(all_rows[2], 1, ns) # Text is usually in the first cell of the third row
    
    assigned_lob = None
    is_summary = False

    if header_row_text:
        header_text_upper = header_row_text.upper()
        if "COMMERCIAL AUTO LIABILITY" in header_text_upper:
            assigned_lob = "AL"
        elif "AUTO PHYSICAL DAMAGE" in header_text_upper:
            assigned_lob = "APD"
        elif "SUMMARY" in header_text_upper:
            is_summary = True

    # --- Logic Check 2: Check the tab name as a confirmation ---
    sheet_name_upper = sheet_name.upper()
    if "COMM'L AUTO L" in sheet_name_upper and assigned_lob is None:
        assigned_lob = "AL"
    elif "AUTO PHYS" in sheet_name_upper and assigned_lob is None:
        assigned_lob = "APD"

    # --- Print Results ---
    print(f"  - Header Text Found: '{header_row_text}'")
    if is_summary:
        print("  - Result: Identified as SUMMARY. Would be SKIPPED.")
    elif assigned_lob:
        print(f"  - Result: Identified as LOB = {assigned_lob}. Would be PROCESSED.")
    else:
        print("  - Result: Not identified as a required LOB. Would be SKIPPED.")


# --- Main Script Execution ---
if __name__ == "__main__":
    input_directory = "./Inputs/"
    xml_files = glob.glob(os.path.join(input_directory, '*.xml'))

    if not xml_files:
        logging.critical(f"No XML files found in the directory: {input_directory}")
    else:
        logging.info(f"Found {len(xml_files)} XML files to test.\n")
        
        for file_path in xml_files:
            try:
                tree = ET.parse(file_path)
                root = tree.getroot()
                ns = {'ss': 'urn:schemas-microsoft-com:office:spreadsheet'}
                worksheets = root.findall('ss:Worksheet', ns)

                # Heuristic to check if this is likely a Schedule P file
                is_sched_p_file = any("PG33" in ws.get(f'{{{ns["ss"]}}}Name') or "PG35" in ws.get(f'{{{ns["ss"]}}}Name') for ws in worksheets)

                if is_sched_p_file:
                    print(f"=====================================================")
                    print(f"Testing Schedule P File: {os.path.basename(file_path)}")
                    print(f"=====================================================")

                    for worksheet in worksheets:
                        sheet_name = worksheet.get(f'{{{ns["ss"]}}}Name')
                        # Process only sheets that are likely Schedule P pages
                        if "PG33" in sheet_name or "PG35" in sheet_name:
                            print(f"\n--- Testing Worksheet: {sheet_name} ---")
                            test_schedule_p_lob_identification(worksheet, ns)
                else:
                    logging.info(f"Skipping non-Schedule P file: {os.path.basename(file_path)}")

            except Exception as e:
                logging.critical(f"An unexpected error occurred while processing {file_path}: {e}", exc_info=True)
