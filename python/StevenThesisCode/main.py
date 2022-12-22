import ocr
from excel import Excel

# Create an Ocr object
#ocr = Ocr("image.jpg")

# Extract text from an image
#text = ocr.extract_text()
#print(text)

# Create an Excel object
excel = Excel("C:/Users/Steven McEwen/OneDrive - University of Cape Town/Desktop/Masters thesis/thesis_data_testing", "C:/Users/Steven McEwen/OneDrive - University of Cape Town/Desktop/Masters thesis/thesis_data_testing/output.xlsx")

# Combine sheets from Excel files in the specified folder
excel.combine_sheets()

# Update values in the Excel file
#excel.update_values("key", "new_value")