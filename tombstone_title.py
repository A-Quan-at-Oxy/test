import pandas as pd
from reportlab.platypus import SimpleDocTemplate, PageBreak, Table, Paragraph
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet, PropertySet
from reportlab.lib import units
from reportlab.platypus.tables import TableStyle
import math
import textwrap


# THIS SECTION GETS THE SOURCE CSV AND GETS A LIST OF TITLES FROM THE "TITLE" COLUMN IN THE CSV FILE.
# IT RETURNS AND PRINTS A VARIABLE NAMED "TITLE"
# NEED TO FIGURE OUT HOW TO TELL IT TO IGNORE ROWS WITHOUT DATA IN THE TITLE COLUMN

def read_book_titles_from_spreadsheet(spreadsheet_filename):
    # Read the spreadsheet
    df = pd.read_csv(spreadsheet_filename)

    # Extract titles from the "Title" column
    titles = df["Title"].tolist()

    return titles


def create_pdf(output_filename, book_titles, num_columns):
    # Use landscape orientation
    doc = SimpleDocTemplate(output_filename, pagesize=landscape(letter), leftMargin=0, rightMargin=0, topMargin=0,
                            bottomMargin=0)

    num_titles = len(book_titles)

    # FOLLOWING SECTION DETERMINES HOW MANY COLUMNS TO CREATE BASED ON HOW MANY INSTANCES OF THE "TITLE" VARIABLE WERE RETURNED
    # I DON'T REALLY UNDERSTAND IT

    start_index = 0
    end_index = 1  # Initialize end index to ensure each column contains one title

    while start_index < num_titles:
        # Ensure the end index does not exceed the number of titles
        end_index = min(start_index + num_columns, num_titles)

        titles_page = book_titles[start_index:end_index]

        title_data = [[''] * num_columns for _ in range(2)]  # Initialize data for the table

        # THIS SECTION IS DEFINING THE TABLE THAT THE LIST OF TITLES WILL BE PLOPPED INTO.
        # EACH VARIABLE IS INTENDED TO PRINT IN THE SECOND CELL OF THE COLUMN.
        # THE FIRST CELL IS WHERE THE BARCODE WILL BE STUCK AND IS FORMATTED TO BE THE CORRECT SIZE

        # Add 'BARCODE' to the first row
        for col in range(num_columns):
            title_data[0][col] = 'BARCODE'

        # Place titles in the second row, one title per column
        for col, title in enumerate(titles_page):
            title_data[1][col] = title

        # THIS SECTION SETS THE SIZE AND STYLE FOR THE TABLE

        # Create the table with fixed column widths
        table = Table(title_data, colWidths=144, rowHeights=[54, 54], repeatRows=1)

        # Set style for cells
        table.setStyle([('VALIGN', (0, 0), (-1, -1), 'TOP'),
                        ('TEXTCOLOR', (0, 0), (-1, -1), colors.black),
                        ('BACKGROUND', (0, 0), (-1, -1), colors.white),
                        ('TOPPADDING', (0, 0), (-1, -1), 0),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
                        ('TOP', (0, 0), (-1, -1), 1, colors.lightgrey),  # Set top border
                        ('BOTTOM', (0, 0), (-1, -1), 1, colors.lightgrey),  # Set bottom border
                        ('LEFT', (0, 0), (-1, -1), 1, colors.lightgrey),
                        ('RIGHT', (0, 0), (-1, -1), 1, colors.lightgrey),
                        ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.lightgrey),  # Set border width for vertical lines
                        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),  # Center alignment for the first row
                        ('VALIGN', (0, 0), (-1, 0), 'MIDDLE'),  # Middle vertical alignment for the first row
                        ('ALIGN', (0, 1), (-1, 1), 'CENTER'),  # Center alignment for the second row
                        ('VALIGN', (0, 1), (-1, 1), 'TOP')])  # Middle vertical alignment for the second row

        # Build the PDF document
        doc.build([table])

        # THIS SECTION TELLS IT TO CREATE NEW PAGES IF THERE ARE MORE THAN 5 COLUMNS

        # Insert page break if there are more titles
        if end_index < num_titles:
            doc.build([PageBreak()])

        start_index = end_index


# NO CLUE WHAT THIS IS

# Example usage
spreadsheet_filename = "/Users/aquan3/Documents/Code/Morrison-Huntington-Titles-copy.csv"
book_titles = read_book_titles_from_spreadsheet(spreadsheet_filename)

# Filter out NaN values from book_titles
book_titles = [title for title in book_titles if isinstance(title, str)]

if book_titles:  # Check if book_titles is not empty
    output_filename_prefix = "output"
    num_columns = 5
    page_number = 1

    # Create PDFs until all data is processed
    while book_titles:
        output_filename = f"{output_filename_prefix}_{page_number}.pdf"
        create_pdf(output_filename, book_titles[:num_columns], num_columns)
        page_number += 1

        # Remove processed titles from the list
        book_titles = book_titles[num_columns:]
else:
    print("No book titles found. Exiting.")
