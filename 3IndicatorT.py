import os
import pandas as pd

# Replace with your folder path
folder_path = "D:/WRR/swmm/outpz/out15"
correspondence_table_file = "elevation.xlsx"

result_file = "D:/WRR/resultT.xlsx"

correspondence_table = pd.read_excel(os.path.join(folder_path, correspondence_table_file))

result_df = pd.DataFrame(columns=["name", "FirstT", "FirstZ", "FinalT", "FinalZ"])

for index, row in correspondence_table.iterrows():
    name = row["name"]
    value = row["value"]

    csv_file = os.path.join(folder_path, f"{name}.csv")

    df = pd.read_csv(csv_file, skiprows=1, names=["Time", "Z"])

    filtered_df = df[df["Z"] > value]

    if not filtered_df.empty:
        first_record = filtered_df.iloc[0]
        last_record = filtered_df.iloc[-1]

        result_df = pd.concat([result_df, pd.DataFrame([[name, first_record["Time"],first_record["Z"],last_record["Time"]
                                                         ,last_record["Z"]]], columns=["name", 'FirstT','FirstZ'
                                                                                       ,'FinalT','FinalZ'])])

result_df.to_excel(os.path.join(folder_path, result_file), index=False)

