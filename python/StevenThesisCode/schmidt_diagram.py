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

        # Set up a polar axis
        plt.figure(figsize=(6, 6), dpi=80)
        ax = plt.subplot(111, projection='polar')

        # Create a scatter plot using the r and theta values
        ax.scatter(theta, r, marker='o', c='r')

        # Add a title and label to the plot
        ax.set_title(sheet_name)
        ax.text(0, 1.1, sheet_name, transform=ax.transAxes, ha='center')

        plt.show()
