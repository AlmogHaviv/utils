import re

import pdfplumber

import PyPDF4

import pandas as pd

#create a function
def getPageCount(filepath):
    file = open(filepath, 'rb')
    # creating a pdf reader object
    readpdf = PyPDF4.PdfFileReader(file)
    # get total number of pages in pdf file
    totalpages = readpdf.numPages
    # printing number of pages in pdf file
    return totalpages


def extract_book_names_from_pdf(pdf_file):
    num_of_pages = getPageCount(pdf_file)
    # Read the PDF file
    pdf = pdfplumber.open(pdf_file)

    # Define the regex pattern to match book names
    pattern = r'([A-Z]+\d+)\s(.+)'  # Match a capital letter(s) followed by one or more digits

    # Create a list to store the book names
    book_names = []

    for page_num in range(num_of_pages):
        page = pdf.pages[page_num]
        text = page.extract_text()
        print(text)
        matches = re.findall(pattern, text)
        for match in matches:
            book_code, book_name = match
            final_book_name = ""
            word_list = book_name.split()
            counter = 0
            while (word_list[counter].isupper()) and (counter + 1 < len(word_list)):
                final_book_name += word_list[counter] + " "
                counter += 1
            print("Book Code:", book_code)
            print("Book Name:", book_name)
            print()
            book_names.append((book_code, final_book_name[:-1]))

    
    return book_names


def convert_to_excel(filename):
    ful_data = extract_book_names_from_pdf(filename)
    # Convert the list of tuples to a DataFrame
    df = pd.DataFrame(ful_data, columns=['Index', 'Book Name'])

    # Create an Excel writer object
    writer = pd.ExcelWriter('book-names-lists.xlsx', engine='xlsxwriter')

    # Convert the DataFrame to an Excel sheet
    df.to_excel(writer, sheet_name='Sheet1', index=False)

    # Save the Excel file
    writer.close()
    print("Excel file 'output.xlsx' created successfully!")

# Test the function
convert_to_excel('sample-pdf.pdf')

