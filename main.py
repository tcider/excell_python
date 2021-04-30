import pandas as pd

def modify_name(x):
    if x and isinstance(x, str):
        x = x.replace(',', '')
        x = x.replace('!', '')
        x = x.replace('.', '')
        if x and x[0] == '-':
            x = x[1:]
        size = len(x)
        if size and x[size - 1] == '-':
            x = x[:size - 1]
        x = x.title()
        x = x.replace('Ё', 'Е')
    return x


def name_cmp(name1, fname_1, name2, fname_2):
    if not isinstance(name2, str):
        return False
    if name1 == name2 and ((not len(fname_1) and not len(fname_2)) or (fname_1 == fname_2)):
        return True
    if len(name2) == 2 and name1[0] == name2[0]:
        name2 = name2.upper()
        if not isinstance(fname_1, str) or fname_1[0] == name2[1]:
            return True
    if len(name2) == 1 and name1[0] == name2[0]:
        if not isinstance(fname_1, str) or not isinstance(fname_2, str):
            return True
        if fname_1[0] == fname_2[0]:
            return True
    return False

xl = pd.ExcelFile("book.xlsx")
df1 = xl.parse(xl.sheet_names[0])
df2 = xl.parse(xl.sheet_names[1])

new_df = df1['ФИО'].str.split(' ', n=3, expand=True)
new_df.columns = ['Фамилия ИП', 'Имя ИП', 'Отчество ИП', 'Прочее']
df1 = pd.concat([df1, new_df], axis=1).drop('ФИО', axis=1)
df1.rename(columns = {'Город проживания':'Адрес ИП'}, inplace=True)
df1 = df1[['Фамилия ИП', 'Имя ИП', 'Отчество ИП', 'Прочее', 'Адрес ИП']]
df1['Фамилия ИП'] = df1['Фамилия ИП'].map(lambda x: modify_name(x))
df1['Имя ИП'] = df1['Имя ИП'].map(lambda x: modify_name(x))
df1['Отчество ИП'] = df1['Отчество ИП'].map(lambda x: modify_name(x))
df1.drop_duplicates(inplace=True)

df2['ИП'] = 'ИП'
df2['Прочее'] = ''
df2 = df2[['Фамилия ИП', 'Имя ИП', 'Отчество ИП', 'Прочее', 'Адрес ИП', 'Регион регистрации', 'ИП']]
df2['Фамилия ИП'] = df2['Фамилия ИП'].map(lambda x: modify_name(x))
df2['Имя ИП'] = df2['Имя ИП'].map(lambda x: modify_name(x))
df2['Отчество ИП'] = df2['Отчество ИП'].map(lambda x: modify_name(x))
df2.drop_duplicates(inplace=True)
df2.sort_values(by =['Фамилия ИП', 'Имя ИП', 'Отчество ИП'], inplace=True)


res_df = pd.DataFrame()
for line in df2.itertuples():
    if line[2] and isinstance(line[2], str):
        tmp_df = df1.loc[df1['Фамилия ИП'] == line[1]]
        res_tmp_df = pd.DataFrame()
        for tmp_line in tmp_df.itertuples():
            if name_cmp(line[2], line[3], tmp_line[2], tmp_line[3]):
                res_tmp_df = pd.concat([res_tmp_df, tmp_df.loc[[tmp_line[0]]]])

        if not res_tmp_df.empty:
            res_tmp_df = pd.concat([df2.loc[[line[0]]], res_tmp_df])
            res_df = pd.concat([res_df, res_tmp_df])
            empty_row = {'Фамилия ИП': '-----', 'Имя ИП': '', 'Отчество ИП': '', 'Прочее': '',
                         'Адрес ИП': '', 'Регион регистрации': '', 'ИП': ''}
            res_df = res_df.append(empty_row, ignore_index=True)


#with pd.ExcelWriter('result.xlsx', mode='a') as res_file:
with pd.ExcelWriter('result.xlsx') as res_file:
    res_df.to_excel(res_file, index=False)
    res_file.sheets['Sheet1'].column_dimensions['A'].width = 16
    res_file.sheets['Sheet1'].column_dimensions['B'].width = 16
    res_file.sheets['Sheet1'].column_dimensions['C'].width = 16
    res_file.sheets['Sheet1'].column_dimensions['D'].width = 8
    res_file.sheets['Sheet1'].column_dimensions['E'].width = 20
    res_file.sheets['Sheet1'].column_dimensions['F'].width = 20
    res_file.sheets['Sheet1'].column_dimensions['G'].width = 5



