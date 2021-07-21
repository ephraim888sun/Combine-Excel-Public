import os
import pandas as pd

filename = r'C:\Users\Ephraim.Sun\Desktop\proj3\Super Plant\Super_Plant_template.xlsx'

df_excel = pd.read_excel(filename, index_col=0)
print(len(df_excel.columns))

'''Clean Up col names'''
for col in df_excel.columns:
    df_excel = df_excel.rename(columns={col: col.replace('\n', '')})
count = 0
max_row = 0
rootdir = r'C:\Users\Ephraim.Sun\Desktop\proj3\Super Plant\File'
for subdir, dirs, files in os.walk(rootdir):
    for file in files:
        excel_loc = os.path.join(subdir, file)
        print(excel_loc)
        if file.endswith(".xlsx"):
            df = pd.read_excel(excel_loc, index_col=0)
        else:
            df = pd.read_csv(excel_loc, index_col=0)
        df.drop_duplicates()
        print(len(df.columns))
        max_row = max_row + len(df.index)
        print(max_row)
        count = count+1
        # print(count)
        '''Clean Up col names'''
        for col in df.columns:
            df = df.rename(columns={col: col.replace('\n', '')})

        #     if (col == 'FEE_AMOUNT'):
        #         df = df.rename(columns={col: col.replace('_',' ')})
        # print(len(df.columns))

        df_excel = pd.concat([df_excel, df], axis=0, ignore_index=False)
        # print(len(df_excel.columns))

print(max_row)
df_excel.to_excel('test1.xlsx', index=True)

