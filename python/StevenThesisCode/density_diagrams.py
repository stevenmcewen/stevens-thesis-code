import matplotlib.pyplot as plt
import pandas as pd

def create_density_plots(filepath):
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

        # Add a title and label to the plot
        ax.set_title(sheet_name)
        ax.text(0, 1.1, sheet_name, transform=ax.transAxes, ha='center')

        plt.show()
