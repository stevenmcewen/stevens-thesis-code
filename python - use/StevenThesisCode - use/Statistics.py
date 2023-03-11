# this file is going to define 2 functions that will go through the excel sheet and create a degree of orientation
# for each of the thin section samples, along with the spherical aperture of the schmidt diagrams
import pandas as pd
import math
import openpyxl


def calculate_orientation(filepath):
    # read the excel sheet
    wb = openpyxl.load_workbook(filepath)
    # calculate vector components for each c-axis and add them up
    for sheet_name in wb.sheetnames:
        # Select the sheet you want to work with
        sheet = wb[sheet_name]
        n = 0
        x_sum = 0
        y_sum = 0
        z_sum = 0

        # Iterate over the cells in the column you want to read

        for i, cell in enumerate(sheet['D']):
            if i == 0:
                continue
            # Extract the angle value and label from the cell value
            a4_value = cell.value[0:-2]  # Extract the angle value string
            a4_value = float(a4_value)
            label = (cell.value[-2:]).upper()

            if label == "WE" or label == "WP":
                west = True
            elif label == "EP" or label == "EE":
                west = False
            else:
                raise ValueError("Invalid label in row " + str(i) + "")

            if west:
                a4_value = a4_value * -1
            phi = math.radians(a4_value)
            theta_cell = sheet.cell(row=cell.row, column=2)
            theta_value = theta_cell.value
            theta = math.radians(theta_value)
            x_sum += math.sin(theta) * math.cos(phi)
            y_sum += math.sin(theta) * math.sin(phi)
            z_sum += math.cos(theta)
            n += 1

        # calculate vector sum and length
        vector_sum = (x_sum, y_sum, z_sum)
        vector_length = math.sqrt(x_sum ** 2 + y_sum ** 2 + z_sum ** 2)
        print(f"{vector_length} length of vector for {sheet_name}")
        # calculate degree of orientation
        R_percent = ((2 * abs(vector_length) - n) / n) * 100
        R_percent = round(R_percent, 2)
        print(f"{R_percent} R percent for {sheet_name}")

        # write degree of orientation to excel sheet
        sheet.cell(row=1, column=30).value = "Degree of Orientation"
        sheet.cell(row=2, column=30).value = R_percent

    wb.save(filepath)


def spherical_aperture(filepath):
    wb = openpyxl.load_workbook(filepath)
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        total_x = 0
        total_y = 0
        total_n = 0
        for i, cell in enumerate(sheet['AA']):
            if i == 0:
                continue
            x_cartesian_value = cell.value
            total_x += x_cartesian_value
            y_cartesian_value = sheet.cell(row=cell.row, column=28).value
            total_y += y_cartesian_value
            total_n += 1

        average_x = total_x / total_n
        average_y = total_y / total_n
        sheet.cell(row=1, column=29).value = "Distances to average point"
        distance_list = []

        for i, cell in enumerate(sheet['AA']):
            if i == 0:
                continue
            x_cartesian_value = cell.value
            y_cartesian_value = sheet.cell(row=cell.row, column=28).value
            x_difference = x_cartesian_value - average_x
            y_difference = y_cartesian_value - average_y
            distance = math.sqrt(x_difference ** 2 + y_difference ** 2)
            distance = round(distance, 2)
            sheet.cell(row=cell.row, column=29).value = distance
            distance_list.append(distance)

        nintyth_percentile = round(total_n * 0.9)
        distance_list.sort()
        ninety_percentile_distance = distance_list[nintyth_percentile]
        sa = (ninety_percentile_distance/10)*100
        sa = round(sa, 2)
        sheet.cell(row=1, column=31).value = "Spherical Aperture"
        sheet.cell(row=2, column=31).value = sa

    wb.save(filepath)
