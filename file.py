### Program to parse and split the three parts of the excel file to have the fields and subfields accessible
## Field 1 - Sales Order form
## Field 2 - Supplier
## Field 3 - Files
## FIeld 4 - Quality

# Using pandas and OpenPyxl to access Excel with Python

'''Pandas - Pandas is a Python library that provides data structures and tools for working with data sets. 
It is built on top of numpy. Pandas stands for “Panel Data” and "Python Data Analysis" '''

''' OpenPyxl - Openpyxl is a Python library that is used for reading and writing Excel files with...
.... extensions like xlsx, xlsm, xltx, xltm'''

# Both are Open source
# Another option was to use xlrd but last release was in 2020


# importing libraries
import pandas as pd  
from openpyxl import load_workbook
import re
import datetime
from tabulate import tabulate

# loading the workbook
workbook = load_workbook(filename='Test_data.xlsx')
worksheet = workbook.active

# test to load and see if the file loads
# rb - read binary (python can read text files which are human understandable and also binary with aid of libraries
df = pd.read_excel(open('Test_data.xlsx', 'rb'),) 
print (df.head())

print('THE EXCEL SHEET LOADS OKAY')
print()
print()

# access file and sheet
file = "Test_data.xlsx"
sheet = "Test Data"


##################################################################################################################################
# Sales Order Form 

# As mentioned above, openpyxl supports the newer file format for MS Excel documents
reader = pd.ExcelFile(file, engine = 'openpyxl')
input_data = pd.read_excel(reader,sheet_name=sheet,nrows = 6, keep_default_na=False, na_values=['NULL'])
# Here we have the arguments 
# - nrows(mentions how many rows to skip at the start of the file)
# - skiprows(nrows=6 to load the next 6 rows into a DataFrame) The number can be varied based on SPECIFIC requirements
# - If keep_default_na is False, and na_values are specified, only the NaN values specified na_values are used for parsing.
headers_0 = input_data['Sales order Form']
input_data_1 = pd.read_excel(reader, sheet_name=sheet, nrows = 2, usecols = [3,4], keep_default_na=False, na_values=['NULL'])
headers_1 = input_data_1.iloc[:, 0].rename('Sales order Form_a')

#print (headers_1)
#print(headers_1.dtype)

headers = headers_0.append(headers_1, ignore_index=True)
# f string to print (new)
print(f'THE VALUES OF THE HEADERS IN THE SALES ORDER FORM ARE \n{headers}.')
print()
# the array headers will contain  all 5 Field Headers stored in it. To access the values, we use indexing
# to access different headers 
# print(headers[0])

# to access subfields from the excel file from section Sales Order Form
data_1 = pd.read_excel('Test_data.xlsx',header = [1,2,3,4,5], na_values=['NULL'])
data_2 = pd.read_excel('Test_data.xlsx', nrows = 2, na_values=['NULL'])
# arguments
# - The header parameter is used to specify which row(s) in the Excel sheet.....
# ...... should be used as the header of the resulting DataFrame.

# first part
cell_value_1 = data_1.columns.values[1]
# second part
cell_value_2 = data_2.iloc[:, 4]

print()
print ('EXTRACTING ALL THE FIELDS AND SUBFIELDS FROM SOURCE INTO CALLABLE LISTS/VARIABLES.......')
print()

# to print and check
for i in range(len(headers)):
    if i < 5:
        print((f'The name of the {headers[i]} is {cell_value_1[i]}'))
    else:
        print((f'The name of the {headers[i]} is {cell_value_2[i-5]}'))

# Contact name, phone and mail

data_3 = pd.read_excel('Test_data.xlsx',header = [8,9], keep_default_na=False,na_values=['NULL'])
cell_value_3 = data_3.columns.values[0]
cell_value_4 = data_3.columns.values[1]
cell_value_5 = data_3.columns.values[2]

for header, value in zip(['Contact names', 'Contact phone', 'Contact mail'], [cell_value_3, cell_value_4, cell_value_5]):
    print(f'The {header} of the suppliers are {value}')
print()
##################################################################################################################################
# Supplier 

input_data_2 = pd.read_excel(reader, sheet_name=sheet,  header =11 , keep_default_na=False, na_values=['NULL'])
headers_2 = input_data_2.iloc[0,:].rename('Supplier')

data_4 = pd.read_excel('Test_data.xlsx',header = [13], keep_default_na=False,na_values=['NULL'])
col_values = []
for i in range(len(data_4.columns)):
    # Get the column name from headers_2
    col_name = headers_2[i]  
    # Get the column value from data_4
    col_value = data_4.iloc[:, i].name  
    col_values.append(col_value)

    # Print the result
    print(f'The Suppliers {col_name} is {col_value}')
print()
##################################################################################################################################
# Files

input_data_3 = pd.read_excel(reader, sheet_name=sheet, usecols='A:B,D:E', header=15, keep_default_na=False, na_values=['NULL'])
headers_3 = input_data_3.iloc[0, :].rename('Files')

 # List of header values to iterate over
header_values = [18, 19, 20] 
for header in header_values:
    data_5 = pd.read_excel('Test_data.xlsx', usecols='A:B,D:E', header=[header], keep_default_na=False, na_values=['NULL'])
    for i in range(len(data_5.columns)):
        # Get the column name from headers_3
        col_name = headers_3[i]  
        # Get the column value from data_5
        col_value = data_5.iloc[:, i].name  
        # Print the result
        print(f'The title {col_name} is {col_value}')
print()
#################################################################################################################################
# Quality 

input_data_4 = pd.read_excel(reader, sheet_name=sheet, usecols='A:E', header=22, keep_default_na=False, na_values=['NULL'])
headers_4 = input_data_4.iloc[0, :].rename('Quality')

# List of header values to iterate over
header_values = [24,25] 
#previous_values = [None] * len(header_values)
for header in header_values:
    data_6 = pd.read_excel('Test_data.xlsx', usecols='A:E', header=[header], keep_default_na=False, na_values=['NULL'])
    for i in range(len(data_6.columns)):
        # Get the column name from headers_4
        col_name = headers_4[i] 
        # Get the column value from data_6 
        col_value = data_6.iloc[:, i].name  
        # Print the result
        #if col_value != previous_values[j]:
        #    print(f'The Quality {col_name} is {col_value}')
        #    previous_values[j] = col_value
        print(f'The Quality {col_name} is {col_value}')
        
print()
#################################################################################################################################

# Parsing using regular Expression for Sales Order Form

print()
print('PARSING THE TEST_DATA USING REGULAR EXPRESSIONS.........................')
print()

df = pd.read_excel('Test_data.xlsx')
# Get the values of cells D4 and D5 because it hasnt been parsed above
d4_value = df.loc[2, 'Unnamed: 3']
d5_value = df.loc[3, 'Unnamed: 3']
# Print the values
print(f'the contents in D4 cell are {d4_value}')
print(f'the contents in D5 cell are {d5_value}')
# parsing them using '=' operator, here split is used
values = [d4_value, d5_value]
for value in values:
    parsed = value.split('=')
    print(f'The parsed values using "=" split are {parsed}')

print()
# cell_value_1 contains contents/subfields from the second column in the first block
print (f'THE VALUES OF CELL_VALUE_1 ARE {cell_value_1}')

for item in cell_value_1:
    if isinstance(item, str):
        # Parse based on spaces, periods, capital letters, forward slash, and digits
        parsed = re.findall(r'[A-Z][a-z]*|[A-Z0-9]+|[./]+|[0-9]+', item)
        print(f'The parsed cell value contents {parsed}')
    else:
        # Skipped parsing for date because its not a string
        # can however be converted using str and then parsed if needed
        print(f"Skipping item {item} because it's not a string")


# to be parsed in the first block include emails which is implemented by 

print()
print('FINDING EMAILS IN THE SPREADSHEET........................................')
print()

emails = r'\S+@\S+'
for col in df.columns:
    for index, row in df.iterrows():
        if isinstance(row[col], str):
            match = re.match(emails, row[col])
            if match:
                print(f"Found match in column {col}, row {index}: {match.group(0)}")

print()
# for the date a simple regex can be used 
# however accompanying logic can be used to correctly wield out dates from mishaps like 99.99.99

# In the supplier section we have the subfields under col_values variable

print('SUPPLIER SECTION.........................................................')
print()

print (f'THE VALUES OF SUBFIELDS  ARE {col_values}')

for value in col_values:
    # Extract words separated by spaces, slashes, or hyphens
    subparts = re.findall(r'[a-zA-Z0-9/-]+', str(value))
    print(f'The subparts parsed from col_values are {subparts}')

    # Extract the value of the boolean variable (True or False)
    if isinstance(value, bool):
        boolean = value
        print(f'The boolean value is {boolean}')

# Similar logic can be applied to any required subfields where we extract data based on our requirements... 
# ...and pass it on to a new list










