'''
Luke Katsel
Cyber Jobs Webscraper created
for UArizona Cyber Security Clinic
12/14/24
'''
debug = True

import openpyxl
import re

#path to excel file for url input and program output
excel_file_path = 'Cyber-Jobs.xlsx'

# regular expression pattern for urls
url_pattern = r'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\\(\\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+'

# open the Excel file
workbook = openpyxl.load_workbook(excel_file_path)

# initialize a list to store extracted URLs
extracted_urls = []

# create a variable for the input sheet
# the second sheet (1) in the workbook will always be the input sheet
sheet_name = workbook.sheetnames[1]
worksheet = workbook[sheet_name]

#iterate through each row
for row in worksheet.iter_rows(values_only=True):
    #iterate through each cell in row
    for cell_value in row:
        if cell_value is not None and isinstance(cell_value, str):
            # find URLs in the cell's text using the regular expression
            urls = re.findall(url_pattern, cell_value)
            # add URLs to extracted_urls list
            extracted_urls.extend(urls)
            
# url duplication sanitization 
unique_urls = list(set(extracted_urls))

# for debugging 
if debug: 
    print(f"Debug: unique urls: {unique_urls}")


# urls are ready for use
for url in extracted_urls: