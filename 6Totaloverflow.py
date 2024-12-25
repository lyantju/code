import os
import pandas as pd

# Folder path
out_path = 'D:/WRR'
base_folders = [base_folder for base_folder in os.listdir(out_path) if
              base_folder.startswith('swmm')]
# Iterate through each base folder
for base_folder in base_folders:
    # Folder path
    folder_path = os.path.join(out_path, base_folder)
    # Get the list of files that start with 'overflow'
    file_list = [file for file in os.listdir(folder_path) if file.startswith("overflow")]

    # Initialize result DataFrame for summary statistics
    result_df = pd.DataFrame(columns=["pipenode", "river"])

    # Iterate through each file in the list
    for file in file_list:
        sum_file_path = os.path.join(folder_path, file)
        if os.path.isfile(sum_file_path):
            sum_df = pd.read_excel(sum_file_path)

            gw = sum_df.loc[sum_df["name"].str.startswith("floodgw"), sum_df.columns[1:]].sum().sum()
            hd = sum_df.loc[sum_df["name"].str.startswith("floodhd"), sum_df.columns[1:]].sum().sum()

        # Append the results to the DataFrame
        result_df = pd.concat([result_df, pd.DataFrame({"name": [file], "pipenode": [gw],
                                                            "river": [hd]})],ignore_index=True, axis=0)

    # Save the summary statistics to an Excel file
    result_path = os.path.join(folder_path, "overflowDU.xlsx")
    result_df.to_excel(result_path, index=False)
