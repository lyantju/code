import os
import pandas as pd

# Specify the path to the main folder containing all the outpz folders
main_folder_path = r'D:\WRR\swmm'

# Get all folders that start with 'outpz'
out_folders = [folder for folder in os.listdir(main_folder_path) if folder.lower().startswith('outpz')]

# Initialize an empty DataFrame for storing results
result_df = pd.DataFrame()

# Iterate through each outpz folder
for out_folder in out_folders:
    # Construct the path to the 'out15' folder
    out1_folder_path = os.path.join(main_folder_path, out_folder, 'out15')

    # Check if the 'out15' folder exists
    if os.path.exists(out1_folder_path):
        file_results = {}

        # Get the list of files in the 'out15' folder
        files = os.listdir(out1_folder_path)

        # Iterate through each file in the folder
        for file in files:
            # Check if the file starts with 'flood' and ends with '.csv'
            if file.lower().startswith("flood") and file.lower().endswith(".csv"):

                file_path = os.path.join(out1_folder_path, file)

                df = pd.read_csv(file_path, header=0, names=['Time', 'data'])
                times = df['Time']
                data = df['data']

                # Calculate the total overflow for each node in the file
                file_results[file] = data.sum()
        result_df[out_folder] = file_results

result_file_path = os.path.join(main_folder_path, 'overflow.xlsx')
result_df.to_excel(result_file_path)