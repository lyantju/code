import os
from pyswmm import Simulation

folder_path = r'D:\WRR\swmm'  # Replace with your folder path
file_names = os.listdir(folder_path)

for file_name in file_names:
    if file_name.endswith('.inp'):
        file_path = os.path.join(folder_path, file_name)
        with Simulation(file_path) as sim:
            sim.execute()

        output_file_name = os.path.splitext(file_name)[0] + '_modified.inp'
        output_file_path = os.path.join(folder_path, output_file_name)