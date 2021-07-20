from openpyxl import load_workbook
import os
import pandas as pd

filename = r'C:\Users\Ephraim.Sun\Desktop\proj3\final.xlsx'
# wb = load_workbook(filename=filename)
# ws = wb.active
df_excel = pd.read_excel(filename, index_col=0)
# print(len(df_excel.columns))

'''Clean Up col names'''
for col in df_excel.columns:
    df_excel = df_excel.rename(columns={col: col.replace('\n', '')})


rootdir = r'C:\Users\Ephraim.Sun\Desktop\proj3\File'
for subdir, dirs, files in os.walk(rootdir):
    for file in files:
        excel_loc = os.path.join(subdir, file)
        print(excel_loc)
        if file.endswith(".xlsx"):
            df = pd.read_excel(excel_loc, index_col=0)
        else:
            df = pd.read_csv(excel_loc, index_col=0)
        '''Clean Up col names'''
        for col in df.columns:
            df = df.rename(columns={col: col.replace('\n', '')})
            # if col == 'FEE_AMOUNT':
            #     df = df.rename(columns={col: col.replace('_',' ')})
        print(len(df.columns))
        # pd.merge(df_excel, df, left_on=list(df_excel.columns),
        #          right_on=list(df.columns),
        #          how='left')

        df_excel = pd.concat([df_excel, df], axis=0)
        print(len(df_excel.columns))


df_excel.to_excel('concat.xlsx', index=True)
# wb.save(filename=filename)
