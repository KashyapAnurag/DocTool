# pdf_extractor/core.py

import os
from pathlib import Path
from pypdf import PdfReader
import pandas as pd


def extract_raw_data(pdf_path: str) -> pd.DataFrame:
    """
    Extracts raw text line-by-line from a PDF and returns it as a pandas DataFrame.
    Each line is split into tokens, forming columns.
    """
    all_rows = []
    try:
        reader = PdfReader(pdf_path)
        for page in reader.pages:
            text = page.extract_text()
            if text:
                for line in text.splitlines():
                    tokens = line.split()
                    all_rows.append(tokens)
    except Exception as e:
        print(f"Error reading PDF {pdf_path}: {e}")
        return pd.DataFrame() # Return empty DataFrame on error

    if not all_rows:
        return pd.DataFrame() # Return empty if no text extracted

    max_cols = max(len(row) for row in all_rows)
    raw_columns = [f"Col{i+1}" for i in range(max_cols)]
    df_raw = pd.DataFrame(all_rows, columns=raw_columns).fillna("")
    return df_raw

def extract_incoming_pipes(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Extracts incoming pipe data from the raw DataFrame based on fixed row/column indices.
    """
    if df_raw.empty or df_raw.shape[0] < 70 or df_raw.shape[1] < 11:
        print("Warning: Raw DataFrame too small for incoming pipe extraction.")
        return pd.DataFrame(columns=["ID", "UPSTREAM REFERENCE", "PIPE SHAPE", "PIPE SIZE (mm)", 
                                     "BACKDROP DIAM (mm)", "PIPE MATERIAL", "LINING", 
                                     "DEPTH FROM COVER (m)", "INVERT LEVEL (m AD)"])

    pipe_data_in = df_raw.iloc[63:70, 0:11].reset_index(drop=True)
    pipe_data_in_final = pd.DataFrame()
    pipe_data_in_final["ID"] = pipe_data_in.iloc[:, 0]
    pipe_data_in_final["UPSTREAM REFERENCE"] = pipe_data_in.iloc[:, 1]
    pipe_data_in_final["PIPE SHAPE"] = pipe_data_in.iloc[:, 2]
    
    # Handle potential empty cells before concatenation
    pipe_size_in = (
        pipe_data_in.iloc[:, 3].fillna('').astype(str).str.strip() + " " +
        pipe_data_in.iloc[:, 4].fillna('').astype(str).str.strip() + " " +
        pipe_data_in.iloc[:, 5].fillna('').astype(str).str.strip()
    )
    pipe_data_in_final["PIPE SIZE (mm)"] = pipe_size_in.str.strip()
    pipe_data_in_final["BACKDROP DIAM (mm)"] = pipe_data_in.iloc[:, 6]
    pipe_data_in_final["PIPE MATERIAL"] = pipe_data_in.iloc[:, 7]
    pipe_data_in_final["LINING"] = pipe_data_in.iloc[:, 8]
    pipe_data_in_final["DEPTH FROM COVER (m)"] = pipe_data_in.iloc[:, 9]
    pipe_data_in_final["INVERT LEVEL (m AD)"] = pipe_data_in.iloc[:, 10]
    return pipe_data_in_final

def extract_outgoing_pipes(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Extracts outgoing pipe data from the raw DataFrame based on fixed row/column indices.
    """
    if df_raw.empty or df_raw.shape[0] < 85 or df_raw.shape[1] < 12:
        print("Warning: Raw DataFrame too small for outgoing pipe extraction.")
        return pd.DataFrame(columns=["ID", "UPSTREAM REFERENCE", "PIPE SHAPE", "PIPE SIZE (mm)", 
                                     "COND", "CRITY", "PIPE MATERIAL", "LINING", 
                                     "DEPTH FROM COVER (m)", "INVERT LEVEL (m AD)"])

    pipe_data_out = df_raw.iloc[83:85, 0:12].reset_index(drop=True)
    pipe_data_out_final = pd.DataFrame()
    pipe_data_out_final["ID"] = pipe_data_out.iloc[:, 0]
    pipe_data_out_final["UPSTREAM REFERENCE"] = pipe_data_out.iloc[:, 1]
    pipe_data_out_final["PIPE SHAPE"] = pipe_data_out.iloc[:, 2]
    
    # Handle potential empty cells before concatenation
    pipe_data_out_final["PIPE SIZE (mm)"] = (
        pipe_data_out.iloc[:, 3].fillna('').astype(str).str.strip() + " " +
        pipe_data_out.iloc[:, 4].fillna('').astype(str).str.strip() + " " +
        pipe_data_out.iloc[:, 5].fillna('').astype(str).str.strip()
    ).str.strip() # Strip again after concat

    pipe_data_out_final["COND"] = pipe_data_out.iloc[:, 6]
    pipe_data_out_final["CRITY"] = pipe_data_out.iloc[:, 7]
    pipe_data_out_final["PIPE MATERIAL"] = pipe_data_out.iloc[:, 8]
    pipe_data_out_final["LINING"] = pipe_data_out.iloc[:, 9]
    pipe_data_out_final["DEPTH FROM COVER (m)"] = pipe_data_out.iloc[:, 10]
    pipe_data_out_final["INVERT LEVEL (m AD)"] = pipe_data_out.iloc[:, 11]
    return pipe_data_out_final

def extract_location_data(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Extracts location and other general details from the raw DataFrame.
    Uses index-based access, with bounds checks for robustness.
    """
    # Helper to safely get value from DataFrame
    def get_val(df, row, col):
        return str(df.iloc[row, col]).strip() if df.shape[0] > row and df.shape[1] > col else ""

    # Helper to safely join a row range
    def join_row_range(df, row, start_col, end_col=None):
        if df.shape[0] <= row: return ""
        cols_to_use = df.iloc[row, start_col:end_col].tolist() if end_col is not None else df.iloc[row, start_col:].tolist()
        return " ".join(str(x).strip() for x in cols_to_use if pd.notna(x) and str(x).strip())


    node_reference_value = get_val(df_raw, 8, 2)

    coord_text = get_val(df_raw, 9, 2)
    x_ref = y_ref = ""
    if "," in coord_text:
        parts = coord_text.split(",")
        if len(parts) >= 2:
            x_ref, y_ref = parts[0].strip(), parts[1].strip()

    location_str = join_row_range(df_raw, 10, 1)
    drainage_code = get_val(df_raw, 12, 1)
    survey_date = join_row_range(df_raw, 13, 2)
    Year_Laid = get_val(df_raw, 15, 2)
    status_pr = get_val(df_raw, 15, 4)
    Function_pr = get_val(df_raw, 15, 6)
    Node_type = get_val(df_raw, 16, 1)
    Cover_shape = get_val(df_raw, 19, 2)
    Hinged_st = get_val(df_raw, 19, 4)
    Lock_st = get_val(df_raw, 19, 6)
    Duty_st = get_val(df_raw, 19, 8)
    Cover_size = join_row_range(df_raw, 19, 9)
    Side_entry = get_val(df_raw, 26, 1)
    Reg_course = get_val(df_raw, 27, 1)
    Depth = get_val(df_raw, 28, 1)
    Shaft_size = join_row_range(df_raw, 28, 2, 5) # Explicit end_col
    Soffit = get_val(df_raw, 30, 2)
    Steps = get_val(df_raw, 32, 0)
    Ladders = get_val(df_raw, 34, 1)
    Landings = get_val(df_raw, 34, 3)
    Chamber_size = join_row_range(df_raw, 34, 4, 7) # Explicit end_col
    Flow_depth = get_val(df_raw, 38, 2)
    Silting = get_val(df_raw, 39, 2)
    Surch = get_val(df_raw, 44, 0)
    MH_cover = get_val(df_raw, 47, 2)

    notes_95 = get_val(df_raw, 95, 0)
    notes_96 = join_row_range(df_raw, 96, 0)
    Notes = (notes_95 + " " + notes_96).strip()

    location_data = {
        "COMPONENTS": [
            "Node Reference", "Coordinates X", "Coordinates Y", "Location", "Drainage Area Code", "Survey Date",
            "Year Laid", "Status", "Function", "Node Type",
            "Cover Shape", "Hinged", "Lock", "Duty", "Cover Size",
            "Side Entry", "Reg Course", "Depth", "Shaft Size",
            "Soffit", "Steps", "Ladders", "Landings", "Chamber Size",
            "Depth of Flow (mm)", "Depth of Silt (mm)", "Surch Height (mm)", "Cover Level (m AD)", "Notes",
        ],
        "VALUE": [
            node_reference_value, x_ref, y_ref, location_str, drainage_code, survey_date,
            Year_Laid, status_pr, Function_pr, Node_type,
            Cover_shape, Hinged_st, Lock_st, Duty_st, Cover_size,
            Side_entry, Reg_course, Depth, Shaft_size,
            Soffit, Steps, Ladders, Landings, Chamber_size,
            Flow_depth, Silting, Surch, MH_cover, Notes,
        ]
    }
    return pd.DataFrame(location_data)

def process_pdf(pdf_path: str, output_excel: str) -> None:
    """
    Processes a single PDF file, extracts data, and saves it to a multi-sheet Excel file.
    """
    df_raw = extract_raw_data(pdf_path)
    if df_raw.empty:
        print(f"Skipping {pdf_path}: No raw data extracted.")
        return

    pipe_data_in_final = extract_incoming_pipes(df_raw)
    pipe_data_out_final = extract_outgoing_pipes(df_raw)
    df_location = extract_location_data(df_raw)

    # Extract node reference for the Excel filename (or sheet content)
    # Using a more robust check in case column 2 is missing or too short
    node_reference_value = df_raw.iloc[8, 2] if df_raw.shape[0] > 8 and df_raw.shape[1] > 2 else ''
    if not node_reference_value: # Fallback if direct access fails
        # Try to find it using a more general regex if direct indexing is unreliable
        text_content = PdfReader(pdf_path).pages[0].extract_text()
        match = re.search(r"NODE REFERENCE\s*([A-Z0-9]+)", text_content)
        if match:
            node_reference_value = match.group(1)
        else:
            node_reference_value = "UNKNOWN_NODE" # Default if not found

    with pd.ExcelWriter(output_excel, engine="xlsxwriter") as writer:
        df_raw.to_excel(writer, sheet_name="Raw Data", index=False)

        workbook = writer.book
        bold_format = workbook.add_format({'bold': True})
        
        sheets_with_node = {
            "Location Details": df_location,
            "Incoming Pipes": pipe_data_in_final,
            "Outgoing Pipes": pipe_data_out_final
        }

        for sheet_name, df in sheets_with_node.items():
            ws = workbook.add_worksheet(sheet_name)
            
            # Write Node Reference
            ws.write('A1', "Node Reference", bold_format)
            ws.write('B1', node_reference_value) # Write value in B1
            
            # Write DataFrame headers starting from row 2 (index 1)
            # and data starting from row 3 (index 2)
            for col_num, col_name in enumerate(df.columns):
                ws.write(2, col_num, col_name, bold_format)
                for row_num, value in enumerate(df[col_name]):
                    # Adjust row_num by 3 because A1/B1 for node ref, and row 2 for headers
                    ws.write(row_num + 3, col_num, value)
            
            # Auto-adjust column widths
            for i, col in enumerate(df.columns):
                series = df[col].astype(str)
                # Max length should consider header (row 2) and data
                max_len = max(series.map(len).max(), len(str(col))) + 2
                ws.set_column(i, i, max_len)
            
            # Adjust column width for Node Reference columns
            # Assuming "Node Reference" in A1 and its value in B1, adjust A and B
            ws.set_column('A:A', max(len("Node Reference"), len(node_reference_value)) + 2)
            ws.set_column('B:B', len(node_reference_value) + 2) # If B1 is value, adjust its column


# Main script execution logic (now part of the package's main usage or an example)
# This part would typically be in a separate script that imports your package
# and calls the process_pdf function.
def run_processing_batch(input_folder: str, output_folder: str) -> list:
    """
    Walks through the input folder, processes all PDFs, and saves results to the output folder.
    Returns a list of paths to PDFs that failed processing.
    """
    os.makedirs(output_folder, exist_ok=True)
    failed_pdfs = []

    for root, dirs, files in os.walk(input_folder):
        for file in files:
            if file.lower().endswith(".pdf"):
                pdf_path = os.path.join(root, file)
                excel_filename = Path(file).stem + ".xlsx"
                output_excel_path = os.path.join(output_folder, excel_filename)
                print(f"Processing {pdf_path} -> {output_excel_path}")
                try:
                    process_pdf(pdf_path, output_excel_path)
                except Exception as e:
                    print(f"Error processing {pdf_path}: {e}")
                    failed_pdfs.append(pdf_path)
    return failed_pdfs