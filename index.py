import pandas as pd
import tabula
from PyPDF2 import PdfReader, PdfWriter
import os

def create_pdf_subset(pdf_path, startpage, endpage, subset_pdf_path):
    reader = PdfReader(pdf_path)
    writer = PdfWriter()

    for i in range(startpage - 1, endpage):
        writer.add_page(reader.pages[i])

    with open(subset_pdf_path, "wb") as f:
        writer.write(f)

def pdf_to_excel_with_custom_margins(original_pdf_path, excel_output_path, startpage, endpage):
    # Create a subset PDF with only the desired pages
    subset_pdf_path = "subset_pdf.pdf"
    create_pdf_subset(original_pdf_path, startpage, endpage, subset_pdf_path)
    
    # Create an Excel writer
    writer = pd.ExcelWriter(excel_output_path, engine='openpyxl')
    
    # Custom margins for each page in points (1 inch = 72 points)
    margins = {
    1: 2 * 72,    # 2 inches for the first page
    2: 1.3 * 72,  # 1.3 inches for the second page
    3: 1.5 * 72,  # 1.5 inches for the third page
    4: 1 * 72     # 1 inch for the fourth page
    }

    # Extract tables with custom margins
    for page in range(startpage, endpage + 1):
        top_margin = margins[page]
        bottom_margin = 792 - top_margin
        area = [top_margin, 0, bottom_margin, 612]
        tables = tabula.read_pdf(subset_pdf_path, pages=page-startpage+1, area=area, multiple_tables=True)

        for i, table in enumerate(tables):
            # Clean column names
            table.columns = table.columns.str.replace('^Unnamed: \d+', '', regex=True)
            table.columns = [' ' if col.isdigit() else col for col in table.columns]

            # Write each table to a separate sheet
            sheet_name = f'Page_{page}_Table_{i+1}'
            table.to_excel(writer, sheet_name=sheet_name, index=False)
    
    # Save the Excel file
    writer._save()
    # Remove the temporary subset PDF
    os.remove(subset_pdf_path)

# Example usage
original_pdf_path = "RLI.pdf"
excel_output_path = "output.xlsx"
startpage = 61
endpage = 62
pdf_to_excel_with_custom_margins(original_pdf_path, excel_output_path, startpage, endpage)
