import re

import pdfplumber

import PyPDF4

import pandas as pd

import os

import glob

#create a function
def find_all_pdf_in_directory():
    script_path = os.path.abspath(__file__)
    script_directory = os.path.dirname(script_path)
    print("Script Directory:", script_directory)
    pdf_files = glob.glob(os.path.join(script_directory, '*.pdf'))
    pdf_filenames = [os.path.basename(pdf_file) for pdf_file in pdf_files]
    print(pdf_filenames)
    return pdf_filenames


def getPageCount(filepath):
    file = open(filepath, 'rb')
    # creating a pdf reader object
    readpdf = PyPDF4.PdfFileReader(file)
    # get total number of pages in pdf file
    totalpages = readpdf.numPages
    # printing number of pages in pdf file
    return totalpages


def extract_book_names_from_pdf(list_of_files):
    # Create a list to store the book names
    book_names = []
    
    # Define the regex pattern to match book names
    pattern = r'([A-Z]+\d+)\s+([A-Z\s,()]+)\s+([\s\S]+?)(?=[A-Z]+\d+|$)'

    for pdf_file in list_of_files:
        num_of_pages = getPageCount(pdf_file)
        # Read the PDF file
        pdf = pdfplumber.open(pdf_file)

        for page_num in range(num_of_pages):
            page = pdf.pages[page_num]
            text = page.extract_text()
            print(text)
            matches = re.findall(pattern, text)
            for match in matches:
                book_code, book_name ,book_info = match
                print("Book Code:", book_code)
                print("Book Name:", book_name)
                print()
                book_names.append(match)

    final_book_names = sorted(book_names)
    return final_book_names


def convert_to_excel(list_of_files):
    ful_data = extract_book_names_from_pdf(list_of_files)
    # Convert the list of tuples to a DataFrame
    df = pd.DataFrame(ful_data, columns=['Index', 'Book Name', 'Book Info'])

    # Create an Excel writer object
    writer = pd.ExcelWriter('book-names-lists.xlsx', engine='xlsxwriter')

    # Convert the DataFrame to an Excel sheet
    df.to_excel(writer, sheet_name='Sheet1', index=False)

    # Save the Excel file
    writer.close()
    print("Excel file 'output.xlsx' created successfully!")

# Test the function
list_of_files = find_all_pdf_in_directory()
extract_book_names_from_pdf(list_of_files)