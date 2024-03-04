import openpyxl
from docx import Document
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_COLOR_INDEX

def write_columns_to_word(file_path, sheet_name, end_column, output_doc):
    # Load the Excel workbook
    workbook = openpyxl.load_workbook(file_path)

    # Select the desired sheet
    sheet = workbook[sheet_name]

    # Initialize a Word document
    doc = Document()

    # Iterate through each row and add values from columns A to F to the document
    for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=6):
        doc.add_paragraph()
        p = doc.add_paragraph()
        for index, cell in enumerate(row):
            cell_value = str(cell.value)
            
            # Add a newline character between each element (except the first one)
            if index > 0:
                p.add_run('\n')

            # Add the cell value to the document
            run = p.add_run(cell_value)

            # Check the first element for highlighting
            if index == 0:
                if row[1].value == 'Green':
                    run.font.highlight_color = WD_COLOR_INDEX.BRIGHT_GREEN
                elif row[1].value == 'Red':
                    run.font.highlight_color = WD_COLOR_INDEX.RED
                elif row[1].value == 'Yellow':
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW

        doc.add_paragraph()  # Add a new line after each set of row data

    # Save the Word document
    doc.save(output_doc)

# Example usage
file_path = 'new.xlsx'
sheet_name = 'Export'
end_column = 'F'  # Set the end column to F
output_document = 'output_document6.docx'  # Set the desired output document name

write_columns_to_word(file_path, sheet_name, end_column, output_document)
