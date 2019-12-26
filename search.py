# Import the modules
import re
from docx import *
import xlsxwriter
import datetime

# Note start time
time_start = datetime.datetime.now()
# Open the docx file
document = Document("test_report_wbsamb_permit_generation_Cycle_2.0.docx")

# Read the test report table
test_report = document.tables[5]

# Initialize Cycle 1 counters to zero
cycle1_complied = cycle1_notcomplied_inconclusive = cycle1_notapplicable = 0
# Initialize Cycle 2 counters to zero
cycle2_complied = cycle2_notcomplied_inconclusive = cycle2_notapplicable = 0

# Read through the rows of the test report table
for row in test_report.rows:

    # Stripping everything but alphanumeric chars from a string in Python
    status = re.sub(r'\W+', '', row.cells[2].text).lower()
    #print(row.cells[2].text)

    # Keep a count of Cycle 1 statuses
    if status.find('cycle1complied') != -1:
        cycle1_complied = cycle1_complied + 1
    elif status.find('cycle1notcomplied') != -1:
        cycle1_notcomplied_inconclusive = cycle1_notcomplied_inconclusive + 1
    elif status.find('cycle1inconclusive') != -1:
        cycle1_notcomplied_inconclusive = cycle1_notcomplied_inconclusive + 1
    elif status.find('cycle1notapplicable') != -1:
        cycle1_notapplicable = cycle1_notapplicable + 1

    # Keep a count of Cycle 2 statuses
    if status.find('cycle2complied') != -1:
        cycle2_complied = cycle2_complied + 1
    elif status.find('cycle2notcomplied') != -1:
        cycle2_notcomplied_inconclusive = cycle2_notcomplied_inconclusive + 1
    elif status.find('cycle2inconclusive') != -1:
        cycle2_notcomplied_inconclusive = cycle2_notcomplied_inconclusive + 1
    elif status.find('cycle2notapplicable') != -1:
        cycle2_notapplicable = cycle2_notapplicable + 1

# Create an excel workbook of compliance status
workbook = xlsxwriter.Workbook("Compliance Status.xlsx")
# Create an excel sheet of compliance status
worksheet = workbook.add_worksheet("Compliance Status")
# Header row of the excel sheet
excel_row = 0
worksheet.write(excel_row, 0, "Cycle")
worksheet.write(excel_row, 1, "Date")
worksheet.write(excel_row, 2, "Number of Requirements")
worksheet.write(excel_row, 3, "Complied")
worksheet.write(excel_row, 4, "Not complied / inconclusive")
worksheet.write(excel_row, 5, "Not Applicable")

# Cycle 1 row of the excel sheet
excel_row = excel_row + 1
worksheet.write(excel_row, 0, "Cycle-1")
worksheet.write(excel_row, 1, "DD-MMM-YYYY")
worksheet.write(excel_row, 2, "43")
worksheet.write(excel_row, 3, cycle1_complied)
worksheet.write(excel_row, 4, cycle1_notcomplied_inconclusive)
worksheet.write(excel_row, 5, cycle1_notapplicable)

# Cycle 2 row of the excel sheet
excel_row = excel_row + 1
worksheet.write(excel_row, 0, "Cycle-2")
worksheet.write(excel_row, 1, "DD-MMM-YYYY")
worksheet.write(excel_row, 2, "43")
worksheet.write(excel_row, 3, cycle2_complied)
worksheet.write(excel_row, 4, cycle2_notcomplied_inconclusive)
worksheet.write(excel_row, 5, cycle2_notapplicable)

# Close workbook
workbook.close()

print("Successfully generated \"Compliance Status.xlsx\" in : " + str(
    (datetime.datetime.now() - time_start).total_seconds()) + " seconds")
