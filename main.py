import pandas as pd


def process_cell(cell):
    # Удаление неразрывающихся пробелов и замена запятой на точку
    cell = cell.replace('\xa0', '').replace(',', '.')
    return cell


def preprocess_frame(data_path):
    df = pd.read_csv(data_path, header=None)
    df = df.iloc[2:]
    df = df.drop(df.columns[6], axis=1)
    df = df.reset_index(drop=True)
    df[0] = df[0].str.replace(' ', '').str.replace('₽', '').str.replace(',', '.')
    # Применение функции к столбцу 'Сумма'
    df[0] = df[0].apply(process_cell)

    # Преобразование столбца в числовой формат (float)
    df[0] = df[0].astype(float)
    df[2] = df[2].str.replace(',', '').str.replace('%', '')
    df[2] = df[2].astype(float) / 100
    df[3] = df[3].str.replace(' ', '').str.replace('₽', '').str.replace(',', '.')
    # Применение функции к столбцу 'Сумма'
    df[3] = df[3].apply(process_cell)

    # Преобразование столбца в числовой формат (float)
    df[3] = df[3].astype(float)
    df[4] = df[4].str.replace(' ', '').str.replace('₽', '').str.replace(',', '.')
    # Применение функции к столбцу 'Сумма'
    df[4] = df[4].apply(process_cell)

    # Преобразование столбца в числовой формат (float)
    df[4] = df[4].astype(float)
    df[5] = df[5].str.replace(',', '').str.replace('%', '')
    df[5] = df[5].astype(float) / 100
    df_res = df.iloc[:, :6]
    df_res.columns = ['Loan issued', 'Рейтинг', 'Comission, %', 'Earned interest', 'Unpaid,  full amount', 'EL']
    return df_res


data_path = 'data.csv'
df_res = preprocess_frame(data_path)
pivot_table = df_res.pivot_table(values='Loan issued',  # Значение для агрегации
                                 index='Рейтинг',  # Ось X: Рейтинг займа
                                 columns='Comission, %',  # Ось Y: Комиссия
                                 aggfunc='sum',  # Функция агрегации (сумма)
                                 fill_value=0)  # Значение по умолчанию

file_path = 'output.xlsx'
with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
    pivot_table.to_excel(writer, sheet_name='Sheet1', index=True)

print(f"DataFrame успешно записан в файл '{file_path}'.")
