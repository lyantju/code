import pandas as pd
import numpy as np
import os

def process_excel_data(input_excel, output_excel):
    """
    Process Excel data by applying transformations to specific columns.

    Args:
        input_excel (str): Path to the input Excel file.
        output_excel (str): Path to save the processed Excel file.
    """
    # Read the input Excel file
    df = pd.read_excel(input_excel)

    # Create a new DataFrame, keeping the first 3 columns unchanged
    new_df = df.iloc[:, :3].copy()

    # Iterate through the remaining columns to process their values
    for col in df.columns[3:]:
        new_values = []  # List to store the calculated values for the column
        for value in df[col]:
            if pd.isna(value):  # If the value is NaN
                new_values.append(np.nan)
            elif value == -9999:  # If the value equals -9999
                new_values.append(0)
            else:
                # Perform the transformation based on the provided formula
                calc_value = np.exp(1.178 * np.sqrt(value * 0.01) + 1.239)  # Public class
                # Uncomment the appropriate formula based on the category
                # calc_value = np.exp(1.015 * np.sqrt(value * 0.01) + 1.721)  # Commercial class
                # calc_value = np.exp(1.077 * np.sqrt(value * 0.01) + 1.454)  # Residential class
                new_values.append(calc_value)

        # Add the transformed column to the new DataFrame
        new_df[col] = new_values

    # Save the processed DataFrame to a new Excel file
    new_df.to_excel(output_excel, index=False)

input_file = r'D:/WRR/Residential_Depth_m.xlsx'
output_file = r'D:/WRR/Public_Loss.xlsx'

# Ensure the directories exist before running
os.makedirs(os.path.dirname(output_file), exist_ok=True)

# Process the data
process_excel_data(input_file, output_file)
