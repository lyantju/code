import pandas as pd
import swmmtoolbox.swmmtoolbox as st
import os
import re

folder_path = 'D:/WRR/Drainagesysteminformation/'  # Excel file containing the drainage pipe network and nodes information
out_path = 'D:/WRR/swmm' # Output path for SWMM result files
file_list = [file for file in os.listdir(folder_path) if file.startswith('information')]# List of files in the folder that start with 'information'

# Get SWMM output file list
swmm_files = [swmm_file for swmm_file in os.listdir(out_path) if
              swmm_file.startswith('swmm') and swmm_file.endswith('.out')]

# Loop through each SWMM output file
for swmm_file in swmm_files:
    swmm_num = re.match(r'swmm(\w+)\.out', swmm_file).group(1)
    SWMM = os.path.join(out_path, swmm_file)

    for file in file_list:
        file_path = os.path.join(folder_path, file)

        #  Excel file drainage pipe nodes information
        df_gw = pd.read_excel(file_path, sheet_name='pipenode')

        gw = df_gw['name'].tolist()
        # Create output folder for saving results
        output_folder = os.path.join(out_path, f'out{swmm_num}',
                                     file.replace('information', 'out').replace('.xlsx', ''))
        os.makedirs(output_folder, exist_ok=True)

        # Read the overflow process of drainage pipe network nodes
        for num, node_id in enumerate(gw):
            flood_gw = st.extract(SWMM, ['node', node_id, 'Flow_lost_flooding'])
            flood_gw.to_csv(os.path.join(output_folder, f'floodgw_{node_id}.csv'))

        excel_file = pd.ExcelFile(file_path)
        sheet_names = excel_file.sheet_names

        # Read the overflow process of river nodes
        if 'rivernode' in sheet_names:
            df_hd = pd.read_excel(file_path, sheet_name='rivernode')
            hd = df_hd['name'].tolist()

            output_folder = os.path.join(out_path, f'out{swmm_num}',
                                         file.replace('information', 'out').replace('.xlsx', ''))
            os.makedirs(output_folder, exist_ok=True)

            for num, node_id in enumerate(hd):
                flood_hd = st.extract(SWMM, ['node', node_id, 'Flow_lost_flooding'])
                flood_hd.to_csv(os.path.join(output_folder, f'floodhd_{node_id}.csv'))

        # Read the water level process of river nodes
        if 'riverlevel' in sheet_names:
            df_hd1 = pd.read_excel(file_path, sheet_name='riverlevel')
            hd1 = df_hd1['name'].tolist()

            output_folder = os.path.join(out_path, f'out{swmm_num}',
                                         file.replace('information', 'out').replace('.xlsx', ''))
            os.makedirs(output_folder, exist_ok=True)

            for num, node_id in enumerate(hd1):
                water_level = st.extract(SWMM, ['node', node_id, 'Hydraulic_head'])
                water_level.to_csv(os.path.join(output_folder, f'Head_{node_id}.csv'))

        # Read the flow process of drainage pipe network
        if 'pipenetwork' in sheet_names:
            df_gd = pd.read_excel(file_path, sheet_name='pipenetwork')
            gd = df_gd['name'].tolist()

            output_folder = os.path.join(out_path, f'out{swmm_num}',
                                         file.replace('information', 'out').replace('.xlsx', ''))
            os.makedirs(output_folder, exist_ok=True)

            for num, link_id in enumerate(gd):
                Flow = st.extract(SWMM, ['link', link_id, 'Flow_rate'])
                Flow.to_csv(os.path.join(output_folder, f'Flow_{link_id}.csv'))