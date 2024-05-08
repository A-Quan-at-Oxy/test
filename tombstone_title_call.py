import pandas as pd
from reportlab.platypus import SimpleDocTemplate, PageBreak, Table
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib import colors

def read_book_titles_from_spreadsheet(spreadsheet_filename):
    df = pd.read_csv(spreadsheet_filename)
    titles = df["Title"].tolist()
    call_numbers = df["LC Call Number"].tolist()
    return titles, call_numbers

def create_pdf(output_filename, book_titles, call_numbers, num_columns):
    doc = SimpleDocTemplate(output_filename, pagesize=landscape(letter), leftMargin=0, rightMargin=0, topMargin=0,
                            bottomMargin=0)

    num_titles = len(book_titles)
    start_index = 0
    end_index = 0

    while start_index < num_titles:
        end_index = min(start_index + num_columns, num_titles)
        titles_page = book_titles[start_index:end_index], call_numbers[start_index:end_index]

        title_data = [[''] * num_columns for _ in range(3)]  # Initialize data for the table

        min_length = min(len(titles_page[0]), len(titles_page[1]))
        
                # Add 'BARCODE' to the first row
        for col in range(num_columns):
            title_data[0][col] = 'BARCODE'

        for i in range(min(5, min_length)):
            title_data[1][i] = titles_page[0][i]  # Populate titles
            title_data[2][i] = titles_page[1][i]  # Populate call numbers

        table = Table(title_data, colWidths=144, rowHeights=[54, 54, 54], repeatRows=1)
        table.setStyle([('VALIGN', (0, 0), (-1, -1), 'TOP'),
                        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                        ('BACKGROUND', (0, 0), (-1, -1), colors.white),
                        ('TOPPADDING', (0, 0), (-1, -1), 0),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
                        ('TOP', (0, 0), (-1, -1), 1, colors.lightgrey),
                        ('BOTTOM', (0, 0), (-1, -1), 1, colors.lightgrey),
                        ('LEFT', (0, 0), (-1, -1), 1, colors.lightgrey),
                        ('RIGHT', (0, 0), (-1, -1), 1, colors.lightgrey),
                        ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.lightgrey),
                        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                        ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),
                        ('ALIGN', (0, 1), (-1, 1), 'CENTER'),
                        ('VALIGN', (0, 1), (-1, 1), 'TOP'),
                        ('ALIGN', (0, 2), (-1, 2), 'CENTER'),
                        ('VALIGN', (0, 2), (-1, 2), 'TOP')])

        doc.build([table])

        if end_index < num_titles:
            doc.build([PageBreak()])

        start_index = end_index

# Example usage
spreadsheet_filename = "/Users/yanni/Desktop/Code_Test/test/Morrison-Huntington-Titles.csv"
book_titles, call_numbers = read_book_titles_from_spreadsheet(spreadsheet_filename)

# Filter out NaN values from book_titles and call_numbers
book_titles = [title for title in book_titles if isinstance(title, str)]
call_numbers = [call_number for call_number in call_numbers if isinstance(call_number, str)]

if book_titles:
    output_filename_prefix = "output"
    num_columns = 5
    page_number = 1

    while book_titles:
        output_filename = f"{output_filename_prefix}_{page_number}.pdf"

        if len(book_titles) >= num_columns:
            create_pdf(output_filename, book_titles[:num_columns], call_numbers[:num_columns], num_columns)
            book_titles = book_titles[num_columns:]
            call_numbers = call_numbers[num_columns:]
        else:
            create_pdf(output_filename, book_titles, call_numbers, len(book_titles))
            break

        page_number += 1

        if len(book_titles) >= num_columns:
            book_titles = book_titles[num_columns:]
            call_numbers = call_numbers[num_columns:]
        else:
            break
else:
    print("No book titles found. Exiting.")
