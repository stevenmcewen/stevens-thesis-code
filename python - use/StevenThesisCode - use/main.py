from excel import Excel
import sheet_correction
from schmidt_diagram import create_schmidt_diagrams
from fabric_diagrams import make_fabric_diagrams
import Statistics

# # real
# OUTPUT_FILE_START_LOCATION = "C:/Users/Steven McEwen/OneDrive - University of Cape Town/Desktop/Masters thesis/thesis_data"
# OUTPUT_FILE_PATH = 'C:/Users/Steven McEwen/OneDrive - University of Cape Town/Desktop/Masters thesis/thesis_data/combined/output.xlsx'
# IMAGES_FOLDER_FABRIC_DIAGRAM = 'C:/Users/Steven McEwen/OneDrive - University of Cape Town/Desktop/Masters thesis/thesis_data/combined/fabric_diagrams/'
# IMAGES_FOLDER_SCHMIDT_DIAGRAM = 'C:/Users/Steven McEwen/OneDrive - University of Cape Town/Desktop/Masters thesis/thesis_data/combined/schmidt_diagrams/'

# testing
OUTPUT_FILE_START_LOCATION = "C:/Users/Steven McEwen/OneDrive - University of Cape Town/Desktop/Masters thesis/thesis_data_testing"
OUTPUT_FILE_PATH = 'C:/Users/Steven McEwen/OneDrive - University of Cape Town/Desktop/Masters thesis/thesis_data_testing/combined/output.xlsx'
IMAGES_FOLDER_FABRIC_DIAGRAM = 'C:/Users/Steven McEwen/OneDrive - University of Cape Town/Desktop/Masters thesis/thesis_data_testing/combined/fabric_diagrams/'
IMAGES_FOLDER_SCHMIDT_DIAGRAM = 'C:/Users/Steven McEwen/OneDrive - University of Cape Town/Desktop/Masters thesis/thesis_data_testing/combined/schmidt_diagrams/'

# Combine sheets from Excel files in the specified folder
excel = Excel(OUTPUT_FILE_START_LOCATION, OUTPUT_FILE_PATH)
excel.combine_sheets()
#
# correct sheets and do conversions
sheet_correction.correct_sheet(OUTPUT_FILE_PATH)
#
# # make the schmidt_diagrams
# create_schmidt_diagrams(OUTPUT_FILE_PATH, IMAGES_FOLDER_SCHMIDT_DIAGRAM)
#
# # fabric diagrams
# make_fabric_diagrams(OUTPUT_FILE_PATH, IMAGES_FOLDER_FABRIC_DIAGRAM)

# statistics
Statistics.calculate_orientation(OUTPUT_FILE_PATH)
Statistics.spherical_aperture(OUTPUT_FILE_PATH)
