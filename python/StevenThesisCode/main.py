import ocr
from excel import Excel
import sheet_correction
import matplotlib.pyplot as plt
from schmidt_diagram import create_schmidt_diagrams

OUTPUT_FILE_START_LOCATION = "C:/Users/Steven McEwen/OneDrive - University of Cape Town/Desktop/Masters thesis/thesis_data_testing"
OUTPUT_FILE_PATH = 'C:/Users/Steven McEwen/OneDrive - University of Cape Town/Desktop/Masters thesis/thesis_data_testing/combined/output.xlsx'

# # Extract text from an image
# # ocr = Ocr("image.jpg")
# # text = ocr.extract_text()
# # print(text)

# Combine sheets from Excel files in the specified folder
excel = Excel(OUTPUT_FILE_START_LOCATION, OUTPUT_FILE_PATH)
excel.combine_sheets()

# Update values in the Excel file


# correct sheets and do conversions
sheet_correction.correct_sheet(OUTPUT_FILE_PATH)

# make the schmidt_diagrams
create_schmidt_diagrams(OUTPUT_FILE_PATH)


