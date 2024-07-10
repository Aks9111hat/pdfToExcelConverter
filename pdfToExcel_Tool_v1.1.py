import fitz  # PyMuPDF
import pandas as pd
import re

def extract_text_from_pdf(pdf_path):
    document = fitz.open(pdf_path)
    text = ""
    for page_num in range(len(document)):
        page = document[page_num]
        text += page.get_text()
    return text

def parse_table_from_text(text):
    lines = text.split("\n")
    tables = []
    current_table = []
    current_columns = None

    for line in lines:
        # Skip empty lines or lines that are likely headers/footers
        if not line.strip() or "Page" in line or "Generation Date" in line:
            continue

        # Split line into columns and strip whitespace
        columns = [col.strip() for col in re.split(r'\s{2,}', line) if col.strip()]
        
        # Ensure column consistency
        if current_columns is None:
            current_columns = len(columns)
        
        if len(columns) == current_columns:
            current_table.append(columns)
        else:
            if current_table:
                tables.append(current_table)
                current_table = []
            current_table.append(columns)
            current_columns = len(columns)

    if current_table:
        tables.append(current_table)

    return tables

def tables_to_excel(tables, excel_path):
    with pd.ExcelWriter(excel_path) as writer:
        for i, table in enumerate(tables):
            df = pd.DataFrame(table)
            df.to_excel(writer, sheet_name=f'Table_{i+1}', index=False, header=False)

def main(pdf_path, excel_path):
    text = extract_text_from_pdf(pdf_path)
    tables = parse_table_from_text(text)
    tables_to_excel(tables, excel_path)
    print(f'Tables extracted and saved to {excel_path}')

if __name__ == "__main__":
    tests = [3, 5, 6]  # Update with actual PDF test file numbers
    for i, test in enumerate(tests):
        pdf_path = f"test{test}.pdf"
        excel_path = f"output_file_for_tests{test}_v1.1.xlsx"
        main(pdf_path, excel_path)
