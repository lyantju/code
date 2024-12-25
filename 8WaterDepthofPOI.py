import pandas as pd
import numpy as np
import glob
import os

def read_asc_matrix(asc_file):
    """
    Read an ASC file and return its matrix representation.

    Args:
        asc_file (str): Path to the ASC file.

    Returns:
        np.ndarray: 2D array representing the data in the ASC file.
    """
    with open(asc_file, 'r') as f:
        # Skip the first 6 header lines
        for _ in range(6):
            next(f)
        matrix_data = []
        for line in f:
            # Convert each line to a list of floats
            row = [float(x) for x in line.strip().split()]
            matrix_data.append(row)
    return np.array(matrix_data)


def process_all_asc_files(excel_file, asc_folder):
    """
    Process all ASC files in a folder and append their data to an Excel file.

    Args:
        excel_file (str): Path to the Excel file.
        asc_folder (str): Path to the folder containing ASC files.
    """
    # Read the Excel file
    df = pd.read_excel(excel_file)

    # Get all ASC files in the folder
    asc_files = glob.glob(os.path.join(asc_folder, "*.asc"))

    # Process each ASC file
    for asc_file in asc_files:
        # Extract the file name (without extension) for the column name
        file_name = os.path.splitext(os.path.basename(asc_file))[0]

        # Read the ASC matrix
        matrix = read_asc_matrix(asc_file)

        # Extract values based on row and column indices in the Excel file
        values = []
        for idx, row in df.iterrows():
            col_num = int(row.iloc[1])  # Column index (from the second column)
            row_num = int(row.iloc[2])  # Row index (from the third column)
            value = matrix[row_num - 1][col_num - 1]  # Adjust for 0-based indexing
            values.append(value)

        # Add the extracted values as a new column
        df[file_name] = values

    # Save the updated DataFrame back to the Excel file
    df.to_excel(excel_file, index=False)

excel_file = r'D:/WRR/Residential_Depth_m.xlsx'
asc_folder = r'D:/WRR'

# Process the data
process_all_asc_files(excel_file, asc_folder)
