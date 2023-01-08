import math
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

def create_density_plots(filepath):
    # Read in all the sheets from the xlsx file
    sheets = pd.read_excel(filepath, sheet_name=None)

    # Iterate through the sheets dictionary
    for sheet_name, df in sheets.items():
        # Extract the r and theta values from the sheet
        r = df['r values']
        theta = df['theta values']

        # Convert theta values from degrees to radians
        theta = np.radians(theta)

        r = np.array(r)
        theta = np.array(theta)

        # Set up a polar axis
        plt.figure(figsize=(6, 6), dpi=80)
        ax = plt.subplot(111, projection='polar')

        # Create a scatter plot using the r and theta values
        ax.scatter(theta, r, marker='o', c='g', alpha=1, zorder=1)
        ax.grid(False)
        ax.set_rlabel_position(0)
        ax.set_theta_zero_location("N")
        ax.set_theta_direction(-1)
        ax.get_xaxis().set_ticklabels([])
        ax.get_yaxis().set_ticklabels([])

        # Add a second set of axes with a 10x10 grid
        ax2 = plt.axes(projection='rectilinear')
        ax2.grid(True)
        ax2.set_xlim(-5, 5)
        ax2.set_ylim(-5, 5)
        ax2.set_facecolor((1, 1, 1, 0.1))
        ax2.xaxis.set_major_locator(plt.MultipleLocator(1))
        ax2.yaxis.set_major_locator(plt.MultipleLocator(1))
        ax2.xaxis.set_minor_locator(plt.MultipleLocator(0.1))
        ax2.yaxis.set_minor_locator(plt.MultipleLocator(0.1))

        # Initialize a 2D array to store the counts for each grid intersection
        grid_intersection_counter = [[0 for ii in range(-5, 6)] for jj in range(-5, 6)]
        # Initialize a dictionary to store the labels for each grid intersection
        grid_intersection_labels = {}

        # Set the desired radial distance
        radial_distance = 10

        # Convert theta values from radians to degrees
        theta_degrees = np.degrees(theta)

        # Initialize the variables ii and jj
        ii = 0
        jj = 0

        # Iterate over all the points on the first set of axes
        for x, y in zip(theta_degrees, r):

            # Find the indices of the data points that correspond to the desired x and y values
            i, = np.where(theta_degrees == x)
            j, = np.where(r == y)
            i = i[0]
            j = j[0]
            # Convert the polar coordinates to cartesian coordinates
            x_cartesian = y * np.cos(x)
            y_cartesian = y * np.sin(x)
            print(f'x_cartesian: {x_cartesian}, y_cartesian: {y_cartesian}')
            for i in range(-5, 6):
                for j in range(-5, 6):
                    # Calculate the Euclidean distance between the data point and the grid intersection
                    distance = math.sqrt(np.abs(x_cartesian - ii) ** 2 + np.abs(y_cartesian - jj) ** 2)
                    # Check if the distance is less than or equal to the radial distance
                    if distance <= radial_distance:
                        grid_intersection_counter[ii][jj] += 1
                        grid_intersection_labels[(ii, jj)] = f'{ii},{jj}'

        # Display the grid intersection counts on the second set of axes
        for ii in range(-5, 6):
            for jj in range(-5, 6):
                ax2.text(ii, jj, grid_intersection_counter[ii][jj], ha='center', va='center', fontsize=8)

        # Add a title and label to the plot
        ax.set_title(sheet_name)
        ax.text(0, 1.1, sheet_name, transform=ax.transAxes, ha='center')

        plt.show()