import numpy as np
import pandas as pd

def count_points_in_nodes(filepath, num_nodes, node_radius):
    # Read in all the sheets from the xlsx file
    sheets = pd.read_excel(filepath, sheet_name=None)

    # Initialize a dictionary to store the counts for each sheet
    sheet_counts = {}

    # Iterate through the sheets dictionary
    for sheet_name, df in sheets.items():
        # Extract the r and theta values from the sheet
        r = df['r values']
        theta = df['theta values']

        # Generate the x and y coordinates for the nodes
        node_x = []
        node_y = []
        for i in range(num_nodes):
            for j in range(num_nodes):
                x = node_radius * i
                y = node_radius * j
                node_x.append(x)
                node_y.append(y)

        # Initialize a list to store the counts for each node
        node_counts = [0] * (num_nodes * num_nodes)

        # Iterate over the nodes
        for i in range(num_nodes * num_nodes):
            # Get the current node's coordinates
            x1, y1 = node_x[i], node_y[i]

            # Iterate over the data points
            for x2, y2 in zip(r, theta):
                # Calculate the distance between the node and the data point
                distance = np.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)

                # If the distance is less than or equal to the node radius, increment the count for the node
                if distance <= node_radius:
                    node_counts[i] += 1

        # Add the counts for the current sheet to the dictionary
        sheet_counts[sheet_name] = node_counts

    return sheet_counts
