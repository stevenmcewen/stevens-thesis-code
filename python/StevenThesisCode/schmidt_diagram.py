import matplotlib.pyplot as plt
import pandas as pd


def create_schmidt_diagrams(filepath):
    # Read in all the sheets from the xlsx file
    sheets = pd.read_excel(filepath, sheet_name=None)

    # Iterate through the sheets dictionary
    for sheet_name, df in sheets.items():
        # Extract the r and radian values from the sheet
        r = df['r values']/6.5
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

        # Add a third set of polar axes to the figure
        axes_dict = {'r_values': [10, ], 'theta_values': [0, ]}
        ax3 = plt.subplot(111, projection='polar')
        ax3.scatter(axes_dict['theta_values'], axes_dict['r_values'], marker='o', c='b', alpha=1, zorder=1)
        ax3.set_ylim(0, 10)

        # Add a second set of axes with a 10x10 grid
        ax2 = plt.axes(projection='rectilinear')
        ax2.grid(True)
        ax2.set_xlim(-10, 10)
        ax2.set_ylim(-10, 10)
        ax2.set_facecolor((1, 1, 1, 0.1))
        ax2.xaxis.set_major_locator(plt.MultipleLocator(1))
        ax2.yaxis.set_major_locator(plt.MultipleLocator(1))


        # Add a title and label to the plot
        ax.set_title(sheet_name)
        ax.text(0, 1.1, sheet_name, transform=ax.transAxes, ha='center')

        plt.show()
