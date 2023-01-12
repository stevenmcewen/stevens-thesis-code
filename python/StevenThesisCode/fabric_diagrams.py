import openpyxl
import math
from collections import defaultdict
import matplotlib.pyplot as plt
import numpy as np


def make_fabric_diagrams(filepath):
    wb = openpyxl.load_workbook(filepath)

    # Iterate over the sheet names in the workbook
    for sheet_name in wb.sheetnames:
        # Select the sheet you want to work with
        sheet = wb[sheet_name]

        grid = defaultdict(int)
        for x in range(-10, 11):
            for y in range(-10, 11):
                for i, (x_value, y_value) in enumerate(zip(sheet['AA'], sheet['AB'])):
                    # Skip the first row
                    if i == 0:
                        continue
                    x_value = x_value.value
                    y_value = y_value.value
                    distance = math.sqrt((x - x_value) ** 2 + (y - y_value) ** 2)

                    if distance <= 1:
                        grid[(x, y)] += 1
                    else:
                        grid[(x, y)] += 0

                    tuples_with_z = [(x, y, z) for (x, y), z in grid.items()]

        x, y, z = zip(*tuples_with_z)

        hist, xedges, yedges = np.histogram2d(x, y, weights=z, bins=(20, 20))
        x_midp = xedges[:-1] + (xedges[1] - xedges[0]) / 2
        y_midp = yedges[:-1] + (yedges[1] - yedges[0]) / 2

        plt.contour(x_midp, y_midp, hist.T, cmap='jet')
        plt.grid(color='black', linewidth=1)
        plt.xlabel('X')
        plt.ylabel('Y')
        plt.show()
