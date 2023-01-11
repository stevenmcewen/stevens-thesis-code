import matplotlib.pyplot as plt
import pandas as pd
import openpyxl
import math
from collections import defaultdict
import numpy as np


def create_density_plot(filepath):
    # Read in all the sheets from the xlsx file
    sheets = pd.read_excel(filepath, sheet_name=None)

    # Iterate through the sheets dictionary
    for sheet_name, df in sheets.items():
        # Extract the x and y values from the sheet
        x_values, y_values = df[['x cartesian values', 'y cartesian values']].values.T

        # Do distance calculations
        grid = defaultdict(int)
        for x in range(-10, 11):
            for y in range(-10, 11):
                for x_value, y_value in zip(x_values, y_values):
                    distance = math.sqrt((x - x_value) ** 2 + (y - y_value) ** 2)

                    if distance <= 1:
                        grid[(x, y)] += 1
                    else:
                        grid[(x, y)] += 0
                        continue

        r = df['r values'] / 6.5
        radians = df['radian values']
        # Set up a polar axis
        plt.figure(figsize=(6, 6), dpi=80)
        ax = plt.subplot(111, projection='polar')

        # Create a scatter plot using the r and theta values
        ax.scatter(radians, r, marker='o', c='g', alpha=1, zorder=1)
        ax.grid(False)
        ax.set_rlabel_position(0)
        ax.set_theta_zero_location("E")
        ax.get_xaxis().set_ticklabels([])
        ax.get_yaxis().set_ticklabels([])

        # Add a second set of axes with a 10x10 grid
        ax2 = plt.axes(projection='rectilinear')
        ax2.grid(True)
        ax2.set_xlim(-10, 10)
        ax2.set_ylim(-10, 10)
        ax2.set_facecolor((1, 1, 1, 0.1))
        ax2.xaxis.set_major_locator(plt.MultipleLocator(1))
        ax2.yaxis.set_major_locator(plt.MultipleLocator(1))

        grid_data = np.array([[grid[(x, y)] for y in range(-10, 11)] for x in range(-10, 11)])
        plt.contour(grid_data, cmap='jet')
        plt.colorbar()
        plt.grid()
        plt.xlabel('x-axis')
        plt.ylabel('y-axis')
        plt.title('Contour Plot')

        # Add a title and label to the plot
        ax.set_title(sheet_name)
        ax.text(0, 1.1, sheet_name, transform=ax.transAxes, ha='center')

        plt.show()
