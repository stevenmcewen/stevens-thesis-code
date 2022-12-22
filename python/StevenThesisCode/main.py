import ocr
from excel import Excel
import sheet_correction

# Extract text from an image
# ocr = Ocr("image.jpg")
# text = ocr.extract_text()
# print(text)

# Combine sheets from Excel files in the specified folder
excel = Excel("C:/Users/Steven McEwen/OneDrive - University of Cape Town/Desktop/Masters thesis/thesis_data_testing",
              "C:/Users/Steven McEwen/OneDrive - University of Cape Town/Desktop/Masters thesis/thesis_data_testing/combined/output.xlsx")
excel.combine_sheets()

# Update values in the Excel file
file_path = 'C:/Users/Steven McEwen/OneDrive - University of Cape Town/Desktop/Masters thesis/thesis_data_testing/combined/output.xlsx'

conversion_dict = {
    '1EP': "1E", '2EP': "1.5E", '3EP': "2E", '4EP': "3E", '5EP': "4E", '6EP': "4.5E", '7EP': "5.5E", '8EP': "6E",
    '9EP': "7E", '10EP': "7.5E", '11EP': "8.5E", '12EP': "9E", '13EP': "10E", '14EP': "10.5E", '15EP': "11.5E",
    '16EP': "12E", '17EP': "13E", '18EP': "14E", '19EP': "14.5E", '20EP': "15E", '21EP': "16E", '22EP': "17E",
    '23EP': "17.5E", '24EP': "18E", '25EP': "19E", '26EP': "19.5E", '27EP': "20.5E", '28EP': "21E", '29EP': "22E",
    '30EP': "22.5E", '31EP': "23E", '32EP': "24E", '33EP': "24.5E", '34EP': "25E", '35EP': "26E", '36EP': "26.5E",
    '37EP': "27.5E", '38EP': "28E", '39EP': "28.5E", '40EP': "29.5E", '41EP': "30E", '42EP': "30.5E",
    '43EP': "31E", '44EP': "32E", '45EP': "32.5E", '46EP': "33E", '47EP': "34E", '48EP': "34.5E", '49EP': "35E",
    '50EP': "36E", '51EP': "36.5E", '52EP': "37E", '53EP': "37.5E", '54EP': "38E", '55EP': "38.5E", '56EP': "39E",
    '57EP': "40E", '58EP': "40.5E", '59EP': "41E", '60EP': "41.5E", '61EP': "42E", '62EP': "42.5E", '63EP': "43E",
    '64EP': "43.5E", '65EP': "44E", '66EP': "44E", '67EP': "44.5E", '68EP': "45E", '69EP': "45.5E", '70EP': "46E",
    '71EP': "46E", '72EP': "46.5E", '73EP': "46.5E", '74EP': "47E", '75EP': "47.5E", '1EE': "1E", '2EE': "2E",
    '3EE': "3E", '4EE': "4E", '5EE': "6E", '6EE': "7E", '7EE': "8E", '8EE': "9E", '9EE': "10E", '10EE': "11E",
    '11EE': "13E", '12EE': "14E", '13EE': "15E", '14EE': "16E", '15EE': "17E", '16EE': "18E", '17EE': "19E",
    '18EE': "20E", '19EE': "22E", '20EE': "23E", '21EE': "24E", '22EE': "25E", '23EE': "26E", '24EE': "27E",
    '25EE': "28E",
    '26EE': "29E", '27EE': "30E", '28EE': "32E", '29EE': "33E", '30EE': "34E", '31EE': "35E", '32EE': "36E",
    '33EE': "37E",
    '34EE': "38E", '35EE': "39E", '36EE': "41E", '37EE': "42E", '38EE': "43E", '39EE': "44E", '40EE': "45E",
    '41EE': "46E", '42EE': "47E",
    '43EE': "48E", '44EE': "50E", '45EE': "51E", '46EE': "52E", '47EE': "53E", '48EE': "54E", '49EE': "55E",
    '50EE': "56E",
    '51EE': "58E", '52EE': "59E", '53EE': "60E", '54EE': "61E", '55EE': "62E", '56EE': "63E", '57EE': "64E",
    '58EE': "65E",
    '1WP': "1W", '2WP': "1.5W", '3WP': "2W", '4WP': "3W", '5WP': "4W", '6WP': "4.5W", '7WP': "5.5W",
    '8WP': "6W", '9WP': "7W", '10WP': "7.5W", '11WP': "8.5W", '12WP': "9W", '13WP': "10W", '14WP': "10.5W",
    '15WP': "11.5W", '16WP': "12W", '17WP': "13W", '18WP': "14W", '19WP': "14.5W", '20WP': "15W", '21WP': "16W",
    '22WP': "17W", '23WP': "17.5W", '24WP': "18W", '25WP': "19W", '26WP': "19.5W", '27WP': "20.5W", '28WP': "21W",
    '29WP': "22W", '30WP': "22.5W", '31WP': "23W", '32WP': "24W", '33WP': "24.5W", '34WP': "25W", '35WP': "26W",
    '36WP': "26.5W", '37WP': "27.5W", '38WP': "28W", '39WP': "28.5W", '40WP': "29.5W", '41WP': "30W",
    '42WP': "30.5W", '43WP': "31W", '44WP': "32W", '45WP': "32.5W", '46WP': "33W", '47WP': "34W", '48WP': "34.5W",
    '49WP': "35W", '50WP': "36W", '51WP': "36.5W", '52WP': "37W", '53WP': "37.5W", '54WP': "38W", '55WP': "38.5W",
    '56WP': "39W", '57WP': "40W", '58WP': "40.5W", '59WP': "41W", '60WP': "41.5W", '61WP': "42W", '62WP': "42.5W",
    '63WP': "43W", '64WP': "43.5W", '65WP': "44W", '66WP': "44W", '67WP': "44.5W", '68WP': "45W", '69WP': "45.5W",
    '70WP': "46W", '71WP': "46W", '72WP': "46.5W", '73WP': "46.5W", '74WP': "47W", '75WP': "47.5W", '1WE': "1W",
    '2WE': "2W", '3WE': "3W", '4WE': "4W", '5WE': "6W", '6WE': "7W", '7WE': "8W", '8WE': "9W", '9WE': "10W",
    '10WE': "11W",
    '11WE': "13W", '12WE': "14W", '13WE': "15W", '14WE': "16W", '15WE': "17W", '16WE': "18W", '17WE': "19W",
    '18WE': "20W",
    '19WE': "22W", '20WE': "23W", '21WE': "24W", '22WE': "25W", '23WE': "26W", '24WE': "27W", '25WE': "28W",
    '26WE': "29W", '27WE': "30W", '28WE': "32W", '29WE': "33W", '30WE': "34W", '31WE': "35W", '32WE': "36W",
    '33WE': "37W",
    '34WE': "38W", '35WE': "39W", '36WE': "41W", '37WE': "42W", '38WE': "43W", '39WE': "44W", '40WE': "45W",
    '41WE': "46W",
    '42WE': "47W", '43WE': "48W", '44WE': "50W", '45WE': "51W", '46WE': "52W", '47WE': "53W", '48WE': "54W",
    '49WE': "55W",
    '50WE': "56W", '51WE': "58W", '52WE': "59W", '53WE': "60W", '54WE': "61W", '55WE': "62W", '56WE': "63W",
    '57WE': "64W",
    '58WE': "65W"
}

sheet_correction.correct_sheet(file_path, conversion_dict)
