import pandas as pd
import os

def process_folders(matching_file, main_folder_path, output_folder_path):
    """
    Process all folders starting with 'out' in the main directory,
    merge CSV data with an Excel mapping file, and save results to an output directory.

    Args:
        matching_file (str): Path to the Excel file containing the mapping (Number and serial number).
        main_folder_path (str): Path to the main folder containing 'out*' subfolders.
        output_folder_path (str): Path to the output directory for saving results.
    """
    # Get all subfolders starting with 'out'
    out_folders = [
        folder for folder in os.listdir(main_folder_path)
        if os.path.isdir(os.path.join(main_folder_path, folder)) and folder.lower().startswith('out')
    ]

    # Read the mapping Excel file
    matching_df = pd.read_excel(matching_file)
    number_column = matching_df['Number']
    index_column = matching_df['Serialnumber']

    # Process each 'out' folder
    for out_folder in out_folders:
        out_folder_path = os.path.join(main_folder_path, out_folder)
        folder_path = os.path.join(out_folder_path, 'out15')
        # Explanation: out15 contains files for overflow data from all nodes in the region.
        if not os.path.exists(folder_path):
            print(f"Skipping folder '{out_folder}' as 'out15' does not exist.")
            continue

        # Initialize the result DataFrame
        result_df = pd.DataFrame()

        # Get the time column from the first CSV file
        first_file_path = os.path.join(folder_path, 'floodgw_060201-24420304-000028.csv')
        # Explanation: floodgw_060201-24420304-000028.csv is a node number.
        if not os.path.exists(first_file_path):
            print(f"Missing expected file '{first_file_path}'. Skipping folder '{out_folder}'.")
            continue

        first_file_df = pd.read_csv(first_file_path)
        result_df['Time'] = first_file_df.iloc[:, 0]

        # Process all files in the 'out15' folder
        for file in os.listdir(folder_path):
            if file.startswith('flood'):
                file_path = os.path.join(folder_path, file)
                file_name_without_extension = os.path.splitext(file)[0]

                df = pd.read_csv(file_path)
                data_column = df.iloc[:, 1]

                # Match file name to the mapping Excel and find the corresponding sequence number
                try:
                    matching_index = number_column[number_column == file_name_without_extension].index[0]
                    matching_number = index_column[matching_index]

                    # Aggregate data if the column already exists
                    if matching_number in result_df.columns:
                        result_df[matching_number] += data_column
                    else:
                        result_df[matching_number] = data_column
                except IndexError:
                    print(f"No matching entry for '{file_name_without_extension}' in the Excel file. Skipping.")

        result_df = result_df[['Time'] + sorted(result_df.columns[1:])]

        # Save the result DataFrame to an Excel file
        result_file_name = f"1D2D{out_folder.replace('out', '')}.xlsx"
        result_file_path = os.path.join(output_folder_path, result_file_name)
        result_df.to_excel(result_file_path, index=False)
        print(f"Results saved to: {result_file_path}")

# Explanation: Paths to files and folders
matching_file = r'D:/WRR/matching_file.xlsx'
main_folder_path = r'D:/WRR/swmm'
output_folder_path = r'D:/WRR'

# Ensure the output directory exists
os.makedirs(output_folder_path, exist_ok=True)

# Run the processing
process_folders(matching_file, main_folder_path, output_folder_path)
