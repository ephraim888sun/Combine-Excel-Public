from openpyxl import load_workbook
import os
import pandas as pd


rootdir = r'C:\Users\Ephraim.Sun\Desktop\proj3\File'
for subdir, dirs, files in os.walk(rootdir):
    for file in files:
        loc = os.path.join(subdir, file)
        if file.endswith(".csv"):
            file_name = loc.replace('.csv', '.xlsx')
            print(file_name)
            read_file = pd.read_csv(loc)
            read_file.to_excel(file_name, index=None, header=True)
            os.remove(loc)
