# Import the modules
import datetime
import re
import sys
import xlsxwriter
from docx import *
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph


def count(test_report_table):
    # Initialize Cycle 1 counters to zero
    cycle1_complied = cycle1_notcomplied_inconclusive = cycle1_notapplicable = 0
    # Initialize Cycle 2 counters to zero
    cycle2_complied = cycle2_notcomplied_inconclusive = cycle2_notapplicable = 0

    # Read through the rows of the test report table
    for row in test_report_table.rows:
        # print(row.cells[0].text)
        # Stripping everything but keeping only alphanumeric chars from a string
        status = re.sub(r'\W+', '', row.cells[2].text).lower()
        # print(row.cells[2].text)

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
    workbook = xlsxwriter.Workbook("ComplianceStatus.xlsx")
    # Create an excel sheet of compliance status
    worksheet = workbook.add_worksheet("ComplianceStatus")
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
    worksheet.write(excel_row, 2, cycle1_complied + cycle1_notcomplied_inconclusive + cycle1_notapplicable)
    worksheet.write(excel_row, 3, cycle1_complied)
    worksheet.write(excel_row, 4, cycle1_notcomplied_inconclusive)
    worksheet.write(excel_row, 5, cycle1_notapplicable)

    if (cycle2_complied + cycle2_notcomplied_inconclusive + cycle2_notapplicable != 0):
        # Cycle 2 row of the excel sheet
        excel_row = excel_row + 1
        worksheet.write(excel_row, 0, "Cycle-2")
        worksheet.write(excel_row, 1, "DD-MMM-YYYY")
        worksheet.write(excel_row, 2, cycle2_complied + cycle2_notcomplied_inconclusive + cycle2_notapplicable)
        worksheet.write(excel_row, 3, cycle2_complied)
        worksheet.write(excel_row, 4, cycle2_notcomplied_inconclusive)
        worksheet.write(excel_row, 5, cycle2_notapplicable)

    # Close workbook
    workbook.close()

    print("Successfully generated \"ComplianceStatus.xlsx\" in : " + str(
        (datetime.datetime.now() - time_start).total_seconds()) + " seconds")


if __name__ == "__main__":
    # Note start time
    time_start = datetime.datetime.now()
    # Help: python search.py <absolutepath>/test_report_wbsamb_checkpost_verification_Cycle_1.0.docx
    docx_file = sys.argv[1]
    # report_cycle = 2  # sys.argv[2]
    # Open the docx file
    document = Document(docx_file)
    start_count = False
    # select only paragraphs or table nodes
    for child in document.element.body.xpath('w:p | w:tbl'):
        # Read paragraph
        if isinstance(child, CT_P):
            paragraph = Paragraph(child, document)
            # If paragraph contains text "Test Report Summary"
            if paragraph.text == "Test Report Summary":
                # Set start count flag
                start_count = True
        # If start count flag is set
        if start_count:
            # If the element is table
            if isinstance(child, CT_Tbl):
                # extract the table
                table = Table(child, document)
                # count the table statuses
                count(table)
                # Terminate loop
                break
