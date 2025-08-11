import pandas as pd
import re
import os
import xml.etree.ElementTree as ET
from typing import List, Dict, Any, Optional
import logging
import glob
from datetime import datetime

# --- Setup basic logging ---
# This will print messages to your console.
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

# --- Utility Functions (Common to Both Parsers) ---

def clean_numeric(value: Any) -> Optional[float]:
    """
    Cleans a value by removing commas, handling parentheses for negatives,
    and converting it to a float. Returns None if conversion fails.
    """
    if value is None or not isinstance(value, str):
        return None
    
    value = value.strip()
    if 'XXX' in value.upper() or value == "" or value == 'NA':
        return None
        
    try:
        if value.startswith('(') and value.endswith(')'):
            value = '-' + value[1:-1]
        return float(value.replace(',', ''))
    except (ValueError, AttributeError):
        return None

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

def identify_report_type(rows: List[ET.Element], ns: Dict[str, str]) -> Optional[str]:
    """
    Peeks inside the first few rows of a worksheet to determine if it's
    a Page 19 or Schedule P report.
    """
    for row in rows[:5]: # Check the top 5 rows for keywords
        for i in range(1, 5):
            cell_text = get_cell_data(row, i, ns)
            if cell_text:
                text = cell_text.upper()
                if "EXHIBIT OF PREMIUMS AND LOSSES" in text:
                    return "Page19"
                if "SCHEDULE P - PART 1" in text:
                    return "ScheduleP"
    return None

# --- Page 19 Specific Functions ---

def find_page19_header_map(rows: List[ET.Element], ns: Dict[str, str]) -> Dict[str, int]:
    number_to_ss_index_map = {}
    for i, row in enumerate(rows[:10]):
        cell_1_text = get_cell_data(row, 8, ns)
        cell_2_text = get_cell_data(row, 9, ns)
        if cell_1_text == '1' and cell_2_text == '2':
            number_row = rows[i]
            current_index = 1
            for cell in number_row.findall('ss:Cell', ns):
                idx_attr = cell.get(f'{{{ns["ss"]}}}Index')
                if idx_attr: current_index = int(idx_attr)
                data_element = cell.find('ss:Data', ns)
                if data_element is not None and data_element.text and data_element.text.strip().isdigit():
                    num = int(data_element.text.strip())
                    number_to_ss_index_map[num] = current_index
                current_index += 1
            break
    if not number_to_ss_index_map: return {}
    return {
        "Direct Premiums Written": number_to_ss_index_map.get(1),
        "Direct Premiums Earned": number_to_ss_index_map.get(2),
        "Direct Losses Incurred": number_to_ss_index_map.get(6),
        "Direct Defense and Cost Containment Expense Incurred": number_to_ss_index_map.get(9)
    }

def process_page19_worksheet(worksheet: ET.Element, ns: Dict[str, str]) -> List[Dict[str, Any]]:
    sheet_name = worksheet.get(f'{{{ns["ss"]}}}Name')
    all_lobs_data = []
    try:
        all_rows = list(worksheet.findall('.//ss:Row', ns))
        if not all_rows: return []
        header_string = get_cell_data(all_rows[0], 2, ns)
        if not header_string: return []
        match = re.search(r"(\d{4}) OF THE (.*?) ?(?:\(NAIC #(\S+)\))?$", header_string)
        if not match: return []
        year, company_name, naic = int(match.group(1)), match.group(2).strip(), match.group(3).strip(')') if match.group(3) else "N/A"
        state = "GRAND_TOTAL"
        for row in all_rows[:5]:
            cell_text = get_cell_data(row, 2, ns)
            if cell_text and "DIRECT BUSINESS IN THE STATE OF" in cell_text.upper():
                for j in range(3, 10):
                    state_val = get_cell_data(row, j, ns)
                    if state_val and state_val.strip():
                        state = state_val.strip().upper()
                        if "GRAND TOTAL" in state: state = "GRAND_TOTAL"
                        break
                break
        header_map = find_page19_header_map(all_rows, ns)
        if not header_map or any(v is None for v in header_map.values()): return []
        
        for i, row in enumerate(all_rows):
            lob_identifier = get_cell_data(row, 2, ns)
            if lob_identifier:
                lob_identifier_clean = lob_identifier.strip()
                lob_code = None
                liability_type = None
                try:
                    lob_float = float(lob_identifier_clean)
                    rounded_lob = round(lob_float, 1)
                    if rounded_lob == 19.3:
                        lob_code = "19.3"
                        liability_type = 'AL'
                    elif rounded_lob == 19.4:
                        lob_code = "19.4"
                        liability_type = 'AL'
                    elif rounded_lob == 21.2:
                        lob_code = "21.2"
                        liability_type = 'APD'
                except (ValueError, TypeError):
                    continue
                if lob_code:
                    data_row = all_rows[i+1] if lob_code == "19.3" else row
                    gwp = clean_numeric(get_cell_data(data_row, header_map["Direct Premiums Written"], ns))
                    ep = clean_numeric(get_cell_data(data_row, header_map["Direct Premiums Earned"], ns))
                    losses = clean_numeric(get_cell_data(data_row, header_map["Direct Losses Incurred"], ns))
                    dcc = clean_numeric(get_cell_data(data_row, header_map["Direct Defense and Cost Containment Expense Incurred"], ns))
                    all_lobs_data.append({
                        "YEAR": year, "Compan_Name": company_name, "NAIC": naic, "State": state,
                        "Liability": liability_type, "LOB": lob_code, "GWP": gwp, "EP": ep,
                        "LOSSES_INCURRED": (losses or 0) + (dcc or 0),
                        "DIRECT_LOSSES_INC": losses, "DCC": dcc
                    })
        return all_lobs_data
    except Exception as e:
        logging.error(f"Error in Page 19 parser for '{sheet_name}': {e}", exc_info=True)
        return []

# --- Schedule P Specific Functions ---

def identify_sched_p_lob(worksheet: ET.Element, rows: List[ET.Element], ns: Dict[str, str]) -> Optional[str]:
    """
    Identifies the Line of Business (AL or APD) from the worksheet header or name.
    Returns None if it's an unknown type. Returns "SUMMARY" for summary sheets.
    """
    # Check 1: Header text in the third row
    if len(rows) >= 3:
        header_text = get_cell_data(rows[2], 1, ns)
        if header_text:
            header_text_upper = header_text.upper()
            if "COMMERCIAL AUTO LIABILITY" in header_text_upper: return "AL"
            if "AUTO PHYSICAL DAMAGE" in header_text_upper: return "APD"
            if "SUMMARY" in header_text_upper: return "SUMMARY"

    # Check 2: Fallback to worksheet name
    sheet_name = worksheet.get(f'{{{ns["ss"]}}}Name').upper()
    if "COMM'L AUTO L" in sheet_name:
        return "AL"
    if "AUTO PHYS" in sheet_name:
        return "APD"
    
    # Check 3: Final check for summary in sheet name (e.g., PG33)
    if "PG33" in sheet_name:
        return "SUMMARY"

    return None

def find_schedule_p_headers(rows: List[ET.Element], ns: Dict[str, str]) -> Dict[str, int]:
    number_to_ss_index_map = {}
    targets_to_find = {'1', '25', '26'}
    found_targets = set()
    for row in rows[:50]:
        current_index = 1
        for cell in row.findall('ss:Cell', ns):
            idx_attr = cell.get(f'{{{ns["ss"]}}}Index')
            if idx_attr: current_index = int(idx_attr)
            data_element = cell.find('ss:Data', ns)
            if data_element is not None and data_element.text:
                text = data_element.text.strip()
                if text in targets_to_find and text not in found_targets:
                    number_to_ss_index_map[int(text)] = current_index
                    found_targets.add(text)
            current_index += 1
        if len(found_targets) == len(targets_to_find): break
    if len(found_targets) != len(targets_to_find): return {}
    return {
        "EP": number_to_ss_index_map.get(1),
        "LOSSES_INC": number_to_ss_index_map.get(26),
        "CLAIMS": number_to_ss_index_map.get(25)
    }

def process_schedule_p_worksheet(worksheet: ET.Element, ns: Dict[str, str], lob: str) -> List[Dict[str, Any]]:
    sheet_name = worksheet.get(f'{{{ns["ss"]}}}Name')
    parsed_data = []
    try:
        all_rows = list(worksheet.findall('.//ss:Row', ns))
        if not all_rows: return []
        header_string = get_cell_data(all_rows[0], 2, ns)
        if not header_string: return []
        match = re.search(r"(\d{4}) OF THE (.*?) ?(?:\(NAIC #(\S+)\))?$", header_string)
        if not match: return []
        report_year, company_name, naic = int(match.group(1)), match.group(2).strip(), match.group(3).strip(')') if match.group(3) else "N/A"
        column_map = find_schedule_p_headers(all_rows, ns)
        if not all(column_map.values()): return []
        prior_row_indices = []
        for i, row in enumerate(all_rows):
            year_val = get_cell_data(row, 3, ns) 
            if year_val and "Prior" in year_val:
                prior_row_indices.append(i)
        if len(prior_row_indices) < 3: return []
        start_row_ep, start_row_claims, start_row_losses = prior_row_indices[0], prior_row_indices[1], prior_row_indices[2]
        for i in range(12):
            row_index_ep, row_index_claims, row_index_losses = start_row_ep + i, start_row_claims + i, start_row_losses + i
            if not all(idx < len(all_rows) for idx in [row_index_ep, row_index_claims, row_index_losses]): break
            row_ep, row_claims, row_losses = all_rows[row_index_ep], all_rows[row_index_claims], all_rows[row_index_losses]
            year = get_cell_data(row_ep, 3, ns)
            if not year: continue
            parsed_data.append({
                "REPORT_YEAR": report_year, "Company_Name": company_name, "NAIC": naic, "LOB": lob, "YEAR": year,
                "EP": clean_numeric(get_cell_data(row_ep, column_map["EP"], ns)),
                "LOSSES_INC": clean_numeric(get_cell_data(row_losses, column_map["LOSSES_INC"], ns)),
                "CLAIMS": clean_numeric(get_cell_data(row_claims, column_map["CLAIMS"], ns))
            })
        return parsed_data
    except Exception as e:
        logging.error(f"Error in Schedule P parser for '{sheet_name}': {e}", exc_info=True)
        return []

# --- Main Execution Logic ---
if __name__ == "__main__":
    input_directory = "./Inputs/"
    output_directory = "./Output/"
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = os.path.join(output_directory, f"Combined_Output_{timestamp}.xlsx")

    os.makedirs(output_directory, exist_ok=True)
    xml_files = glob.glob(os.path.join(input_directory, '*.xml'))

    if not xml_files:
        logging.critical(f"No XML files found in the directory: {input_directory}")
    else:
        logging.info(f"Found {len(xml_files)} XML files to process.")
        
        all_page19_data = []
        all_sched_p_data = []

        for file_path in xml_files:
            logging.info(f"--- Processing file: {file_path} ---")
            try:
                tree = ET.parse(file_path)
                root = tree.getroot()
                ns = {'ss': 'urn:schemas-microsoft-com:office:spreadsheet'}
                worksheets = root.findall('ss:Worksheet', ns)
                
                for worksheet in worksheets:
                    sheet_name = worksheet.get(f'{{{ns["ss"]}}}Name')
                    all_rows = list(worksheet.findall('.//ss:Row', ns))
                    report_type = identify_report_type(all_rows, ns)
                    
                    if report_type == "Page19":
                        all_page19_data.extend(process_page19_worksheet(worksheet, ns))
                    elif report_type == "ScheduleP":
                        lob = identify_sched_p_lob(worksheet, all_rows, ns)
                        if lob in ["AL", "APD"]:
                            logging.info(f"  -> Processing Schedule P - {lob} sheet: {sheet_name}...")
                            all_sched_p_data.extend(process_schedule_p_worksheet(worksheet, ns, lob))
                        else:
                            logging.info(f"  -> Skipping Schedule P - Summary/Other sheet: {sheet_name}")
            except Exception as e:
                logging.critical(f"Fatal error processing file {file_path}: {e}", exc_info=True)

        # --- Final Output Generation ---
        with pd.ExcelWriter(output_filename, engine='xlsxwriter') as writer:
            if all_page19_data:
                pg19_df = pd.DataFrame(all_page19_data)
                pg19_df.drop_duplicates(subset=['NAIC', 'YEAR', 'State', 'LOB'], keep='first', inplace=True)
                pg19_df_sorted = pg19_df.sort_values(by=["Compan_Name", "YEAR", "State", "Liability", "LOB"]).reset_index(drop=True)
                pg19_df_sorted.to_excel(writer, sheet_name='Page 14 Data', index=False)
                logging.info(f"\nProcessed {len(pg19_df_sorted)} rows of Page 14 data.")
            else:
                logging.warning("No Page 14 data was processed.")

            if all_sched_p_data:
                sched_p_column_order = ["REPORT_YEAR", "Company_Name", "NAIC", "LOB", "YEAR", "EP", "LOSSES_INC", "CLAIMS"]
                sched_p_df = pd.DataFrame(all_sched_p_data, columns=sched_p_column_order)
                sched_p_df_sorted = sched_p_df.sort_values(by=["Company_Name", "REPORT_YEAR", "LOB", "YEAR"]).reset_index(drop=True)
                sched_p_df_sorted.to_excel(writer, sheet_name='Schedule P Data', index=False)
                logging.info(f"Processed {len(sched_p_df_sorted)} rows of Schedule P data.")
            else:
                logging.warning("No Schedule P data was processed.")
        
        if all_page19_data or all_sched_p_data:
             logging.info(f"\nSuccessfully saved all data to '{output_filename}'")
        else:
            logging.warning("No data was successfully processed from any files.")
