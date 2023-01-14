# %%
import xlwings as xw
import pandas as pd
import numpy as np
import os
import shutil


# %%
path = os.getcwd()
root_folder = 'file-Prepare'
end_string_0 = '0.xlsx'
end_string_1 = '1.xlsx'
folder_0 = 'folder_0'
folder_1 = 'folder_1'

folder_merge = 'folder_Merge'


for file_0 in os.listdir(os.path.join(path,folder_0)):
    os.remove(os.path.join(path,folder_0,file_0))
for file_1 in os.listdir(os.path.join(path,folder_1)):
    os.remove(os.path.join(path,folder_1,file_1))
for file_merge in os.listdir(os.path.join(path,folder_merge)):
    os.remove(os.path.join(path,folder_merge,file_merge))



for file in os.listdir(os.path.join(path,root_folder)): 
        if file.endswith(end_string_0):
                root = os.path.join(path , root_folder,  file)

                target = os.path.join(path, folder_0 ,file)

                shutil.move(root, target)
        if file.endswith(end_string_1):
                root = os.path.join(path , root_folder,  file)

                target = os.path.join(path, folder_1 ,file)

                shutil.move(root, target)

# %% [markdown]
# ### Merge File

# %%
print('--------------')
print('Bắt đầu Merge File...')

cwd = os.getcwd()
#Remove folder
# for file_0 in os.listdir(os.path.join(cwd,folder_0)):
#     os.remove(file_0)
# for file_1 in os.listdir(os.path.join(cwd,folder_1)):
#     os.remove(file_1)
# folders = ['#bs', '#main']
folder_0 = 'folder_0'
folder_1 = 'folder_1'
folder_merge = 'folder_Merge'
lst_0 = []
lst_1 = []
for file_0 in os.listdir(os.path.join(cwd,folder_0)):
    lst_0.append(file_0[:len(file_0)-7])

for file_1 in os.listdir(os.path.join(cwd,folder_1)):
    lst_1.append(file_1[:len(file_1)-7])

for i in lst_0:
    for j in lst_1:
        listFile =[]
        if i == j:
            listFile.append(pd.read_excel(os.path.join(cwd , folder_0, f'{i}_0.xlsx'), sheet_name=0, dtype=str, engine='openpyxl'))
            listFile.append(pd.read_excel(os.path.join(cwd , folder_1, f'{j}_1.xlsx'), sheet_name=0, dtype=str, engine='openpyxl'))
            listFile_master = pd.concat(listFile, axis=0).drop_duplicates(subset=['TRANS_NO'])
            listFile_master.to_excel(os.path.join(cwd , folder_merge, f'merge_{i}.xlsx'), index=False, engine='openpyxl')


# print('Kết Thúc Merge File...')

# %% [markdown]
# # Merge file có cùng tên

# %%
path = os.getcwd()
folder_statements = 'file-Statements'
folder_Merge = 'folder_Merge'
for root, dirs, files in os.walk(os.path.join(path,folder_statements)): #dirs is list folder
    for dir in dirs:
        subfolder = os.path.join(root, dir)
        for file_statement in os.listdir(subfolder):
            if file_statement.endswith(('.xlsx')):
                count_file = []
                for file_merge in os.listdir(os.path.join(path,folder_Merge)):
                    if file_merge[6:len(file_merge)-7] == file_statement[:len(file_statement)-5]:    
                        print(file_merge[6:len(file_merge)-7])
                        count_file.append(file_merge)
                if len(count_file) > 1:
                    lst_file = []
                    for i in count_file:
                        lst_file.append(pd.read_excel(os.path.join(path , folder_Merge,i ), sheet_name=0, dtype=str, engine='openpyxl'))
                    listFile_master = pd.concat(lst_file, axis=0).drop_duplicates(subset=['TRANS_NO'])
                    listFile_master.to_excel(os.path.join(path , folder_Merge, f'{i[:len(i)-7]}.xlsx'), index=False, engine='openpyxl')  # print(os.path.join(path , folder_Merge,i ))
            if file_statement.endswith(('.xls')):
                count_file = []
                for file_merge in os.listdir(os.path.join(path,folder_Merge)):
                    if file_merge[6:len(file_merge)-7] == file_statement[:len(file_statement)-4]:   
                        print(file_merge[6:len(file_merge)-7])
                        print(file_statement[:len(file_statement)-4])
                        count_file.append(file_merge)
                if len(count_file) > 1:
                    lst_file = []
                    for i in count_file:
                        lst_file.append(pd.read_excel(os.path.join(path , folder_Merge,i ), sheet_name=0, dtype=str, engine='openpyxl'))
                    listFile_master = pd.concat(lst_file, axis=0).drop_duplicates(subset=['TRANS_NO'])
                    listFile_master.to_excel(os.path.join(path , folder_Merge, f'{i[:len(i)-7]}.xlsx'), index=False, engine='openpyxl')  # print(os.path.join(path , folder_Merge,i ))


# %% [markdown]
# ### Check

# %%
path = os.getcwd()
folder_statements = 'file-Statements'
folder_Merge = 'folder_Merge'
for root, dirs, files in os.walk(os.path.join(path,folder_statements)): #dirs is list folder
    for dir in dirs:
        subfolder = os.path.join(root, dir)
        for file_statement in os.listdir(subfolder):
            for file_merge in os.listdir(os.path.join(path, folder_Merge)):
                if file_statement.endswith('.xls') and file_statement[:len(file_statement)-4] == file_merge[6:len(file_merge)-5]:
                # if file_statement.endswith('.xls'):
                    print(f'Loading file {file_merge}......')
                    print(f'Loading file {file_statement}......')

                    df_master = pd.read_excel(os.path.join(path,folder_Merge,file_merge), dtype=str)
                    df_sk = pd.read_excel(os.path.join(path,folder_statements, subfolder, file_statement), skipfooter=3, header=8)

                    full_row_sk = len(df_sk.index)
                    df_sk.index +=10
                    df_sk = df_sk.loc[df_sk['Số GD'].str.startswith('FT')]
                    df_sk['Số GD'] = df_sk['Số GD'].apply(lambda x: x[:12])

                    app = xw.App()
                    wb = xw.Book(os.path.join(path,folder_statements, subfolder, file_statement))
                    file_blank = open(os.path.join(path,folder_statements, subfolder, file_statement[:len(file_statement)-4] + '_blank.txt'), 'w')
                    wb.sheets[0].range(f'I9:K9').value = df_master.columns.values[1:]
                    wb.sheets[0].range(f'I9:K9').font.bold = True
                    wb.sheets[0].range(f'I9:K9').font.italic = True

                    #ALl Border
                    for i in range(7,13):
                        wb.sheets[0].range(f'I9:K{9 + full_row_sk}').api.Borders(i).LineStyle = 1

                    for index_master, row_master in df_master.iterrows():
                        for index_sk, row_sk in df_sk.iterrows():

                            if row_master['TRANS_NO'] == row_sk['Số GD']:
                                wb.sheets[0].range(f'I{index_sk}').number_format = '@'
                                wb.sheets[0].range('I9').options(index=False,format = str).value

                                wb.sheets[0].range(f'I{index_sk}:K{index_sk}').raw_value = row_master[1:]

                                df_sk = df_sk.drop([index_sk])
                                break
                    for i in df_sk['Số GD']:
                        file_blank.write(f"{i}\n")
                    wb.save()
                    wb.close()
                    file_blank.close()
                    app.quit()
                    break


