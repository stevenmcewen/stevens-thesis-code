import numpy as np
import matplotlib.pyplot as plt
import pandas as pd

def create_polar_plots(filepath):

    # Read in all the sheets from the xlsx file
    sheets = pd.read_excel(filepath, sheet_name=None)

    # Iterate through the sheets dictionary
    for sheet_name, df in sheets.items():
        # Extract the r and theta values from the sheet
        r = df['r values']
        theta = df['theta values']

        # Calculate the histogram of the data
        densities, theta_edges, r_edges = np.histogram2d(theta, r, bins=10)

        levels = np.linspace(densities.min(), densities.max(), num=10)

        # Create the angle and radius bins
        angle_bins = 0.5 * (theta_edges[1:] + theta_edges[:-1])
        radius_bins = 0.5 * (r_edges[1:] + r_edges[:-1])

        # Set up a polar axis
        plt.figure(figsize=(6, 6), dpi=80)
        ax = plt.subplot(111, projection='polar')

        # Create a contour plot using the angle, radius, and density arrays
        cf = ax.contourf(angle_bins, radius_bins, densities, levels, cmap='RdYlBu_r')

        # Add a title and label to the plot
        ax.set_title(sheet_name)
        ax.text(0, 1.1, sheet_name, transform=ax.transAxes, ha='center')

        plt.show()