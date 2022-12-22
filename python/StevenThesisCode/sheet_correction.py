import openpyxl


def correct_sheet(file_path, conversion_dict):
    print("Converting the A1 values to corrected A1 values")
    # Open the Excel file
    wb = openpyxl.load_workbook(file_path)

    # Iterate over the sheet names in the workbook
    for sheet_name in wb.sheetnames:
        # Select the sheet you want to work with
        sheet = wb[sheet_name]

        # Iterate over the cells in the column you want to read
        for cell in sheet['C']:
            # Get the value of the cell
            old_value = cell.value
            # Convert the value using the dictionary
            new_value = conversion_dict.get(old_value, old_value)
            # Write the new value to the cell in the other column
            cell.offset(column=1).value = new_value

    # Save the changes to the file
    wb.save(file_path)
