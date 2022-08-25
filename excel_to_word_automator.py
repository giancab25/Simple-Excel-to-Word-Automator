import os, sys
import xlwings as xw
from docxtpl import DocxTemplate
from openpyxl import load_workbook


# Change path to current working directory
os.chdir(sys.path[0])


def find_data(worksheet):
    # Create storage system
    storage = {"first_name": [], "last_name": [], "work": [], "studying": []}

    # Select location of data
    first_name = worksheet.range("B2").value
    last_name = worksheet.range("A2").value
    message = worksheet.range("C2").value
    row = 2

    # Store data from cells while moving down row by row
    while (first_name is not None) and (last_name is not None) and (message is not None):
        message = message.lower()
        taxes_pt = message.find("work")
        fafsa_pt = message.find("studying")

        storage["first_name"].append(first_name)
        storage["last_name"].append(last_name)
        
        numbers = [int(s) for s in message.split() if s.isdigit()]
        if (taxes_pt > fafsa_pt):
            storage["studying"].append(numbers[0])
            storage["work"].append(numbers[1])
        else:
            storage["work"].append(numbers[0])
            storage["studying"].append(numbers[1])

        row += 1
        first_name = worksheet.range(f"B{row}").value
        last_name = worksheet.range(f"A{row}").value
        message = worksheet.range(f"C{row}").value
    
    # Return all saved data
    return storage

def main():

    # Open the desired Excel spreadsheet
    wb = xw.Book.caller()

    # Select which sheet has the data
    data_sheet = wb.sheets["Data"]

    # Collect data from Excel
    data_storage = find_data(data_sheet)

    # Load in the Word Document template
    doc = DocxTemplate("Letter template copy.docx")

    # Insert all values into unique templates, respectively, within one Word Document
    for index, item in enumerate(data_storage["first_name"]):
        # Use main template as a reference for the sub-document
        sd = doc.new_subdoc("Letter template.docx")

        context = {
            "first_name" : item,
            "last_name" : data_storage["last_name"][index],
            "work" : data_storage["work"][index],
            "studying" : data_storage["studying"][index],
            "copy" : sd,
        }

        doc.render(context)

        doc.save("Letter template copy.docx")


if __name__ == "__main__":
    xw.Book("excel_to_word_automator.xlsm").set_mock_caller()
    main()