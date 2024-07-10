import fitz 
import pandas as pd

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

    for line in lines:
        if line.strip() == "":
            if current_table:
                tables.append(current_table)
                current_table = []
        else:
            current_table.append(line.split())

    if current_table:
        tables.append(current_table)

    return tables

def tables_to_excel(tables, excel_path):
    with pd.ExcelWriter(excel_path) as writer:
        for i, table in enumerate(tables):
            df = pd.DataFrame(table)
            df.to_excel(writer, sheet_name=f'Table_{i+1}', index=False, header=True)

def main(pdf_path, excel_path):
    text = extract_text_from_pdf(pdf_path)
    tables = parse_table_from_text(text)
    tables_to_excel(tables, excel_path)
    print(f'Tables extracted and saved to {excel_path}')

if __name__ == "__main__":

    tests = [3,5,6]
    for i, test in enumerate(tests):
        pdf_path = f"test{test}.pdf"
        excel_path = f"output_file_for_tests{test}.xlsx"
        main(pdf_path, excel_path)
