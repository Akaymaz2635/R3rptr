# import pandas as pd

# data = {
#         'Element': ['KN088_Distance3', 'KN088_Distance2', 'KN088_Distance1', 'KN087_Distance3', 'KN087_Distance2', 'KN087_Distance1', 'KN013', 'KN014'],
#         'Datum': ['', '', '', '', '', '', '', ''],
#         'Property': ['L', 'L', 'L', 'L', 'L', 'L', 'L', 'L'],
#         'Nominal': [3.7, 3.7, 3.7, 29.89, 29.89, 29.89, 30, 31],
#         'Actual': [3.669, 3.682, 3.69, 29.427, 29.414, 0, 31, 32],
#         'Actual1': [3.669, 3.682, 3.69, 29.427, 29.414, 0, 31, 32],
#         'Actual2': [3.669, 3.682, 3.69, 29.427, 29.414, 0, 31, 32],
#         'Actual3': [3.669, 3.682, 3.69, 29.427, 29.414, 0, 31, 32],
#         'Actual4': [3.669, 3.682, 3.69, 29.427, 29.414, 0, 31, 32],
#         'Tol-': [-0.1, -0.1, -0.1, -0.1, -0.1, -0.1, 0.5, 0.5],
#         'Tol+': [0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.5, 0.5],
#         'Dev': ['', '', '', '', '', '', '', ''],
#         'Check': ['', '', '', '', '', '', '', ''],
#         'Out': ['', '', '', '', '', '', '', '']
#     }

# df = pd.DataFrame(data)
# df['Element'] = df['Element'].str.split('_').str[0]

# # 'Dev','Actual', 'Out', 'Check' sütunlarını sil

# df = df.drop(columns=['Nominal','Datum', 'Property' ,'Tol-', 'Tol+', 'Dev' ,'Out' , 'Check'])

# #df_group = df.groupby('Element').agg(lambda x : x.head(1) if x.dtype=='object' else x.mean())

# # Sayısal sütunların minimum ve maksimum değerlerini bir arada göster, diğer sütunlar için ilk değeri al
# df_group = df.groupby('Element').agg(lambda x: f"{x.min()} / {x.max()}" if pd.api.types.is_numeric_dtype(x) else x.iloc[0])

# # Sayısal sütunların / işaretinden önceki ve sonraki değerlerini birbirine eşitse, sadece birini al, diğerini sil
# # def process_column(x):
# #     if pd.api.types.is_string_dtype(x):
# #         splits = x.str.split('/')
# #         if len(splits) > 1 and (splits.str[0] == splits.str[1]).all():
# #             return splits.str[0]
# #     return x.iloc[0]

# # df_group = df.apply(process_column)

# print(df_group)
################################################### ÇALIŞIYORRRR
# import pandas as pd

# data = {
#         'Element': ['KN088_Distance3', 'KN088_Distance2', 'KN088_Distance1', 'KN087_Distance3', 'KN087_Distance2', 'KN087_Distance1', 'KN013', 'KN014'],
#         'Datum': ['', '', '', '', '', '', '', ''],
#         'Property': ['L', 'L', 'L', 'L', 'L', 'L', 'L', 'L'],
#         'Nominal': [3.7, 3.7, 3.7, 29.89, 29.89, 29.89, 30, 31],
#         'Actual': [3.669, 3.682, 3.69, 29.427, 29.414, 0, 31, 32],
#         'Actual1': [3.669, 3.682, 3.69, 29.427, 29.414, 0, 31, 32],
#         'Actual2': [3.669, 3.682, 3.69, 29.427, 29.414, 0, 31, 32],
#         'Actual3': [3.669, 3.682, 3.69, 29.427, 29.414, 0, 31, 32],
#         'Actual4': [3.669, 3.682, 3.69, 29.427, 29.414, 0, 31, 32],
#         'Tol-': [-0.1, -0.1, -0.1, -0.1, -0.1, -0.1, 0.5, 0.5],
#         'Tol+': [0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.5, 0.5],
#         'Dev': ['', '', '', '', '', '', '', ''],
#         'Check': ['', '', '', '', '', '', '', ''],
#         'Out': ['', '', '', '', '', '', '', '']
#     }

# df = pd.DataFrame(data)
# df['Element'] = df['Element'].str.split('_').str[0]

# # 'Dev','Out', 'Check' sütunlarını sil
# df = df.drop(columns=['Nominal','Datum', 'Property' ,'Tol-', 'Tol+', 'Dev' ,'Out' , 'Check'])

# # Sayısal sütunların minimum ve maksimum değerlerini bir arada göster, diğer sütunlar için ilk değeri al
# df_group = df.groupby('Element').agg(lambda x: f"{x.min()} / {x.max()}" if pd.api.types.is_numeric_dtype(x) else x.iloc[0])

# # Veriyi istediğiniz formata dönüştürme
# #df_group = df_group.unstack(level=-1)
# print(df_group)


# # Her bir sütun için döngü
# for col in df_group.columns[0:]:
#     # Her bir hücre için döngü
#     for index, value in df_group[col].items():
#         # Eğer / işareti varsa
#         if '/' in value:
#             # / işaretinden önceki ve sonraki değerlerin aynı olup olmadığını kontrol et
#             parts = value.split('/')
#             if parts[0].strip() == parts[1].strip():
#                 df_group.at[index, col] = float(parts[0].strip())
#                 # Eğer eşitse, / işaretinden sonraki kısmı sil
#                 #df.at[index, col] = parts[0].strip()
#             else:
#                 # Eğer değerler farklı ise, işlemi atla
#                 continue

# print(df_group)

# df_group.to_excel("output.xlsx")

import pandas as pd

data = {
        'Element': ['KN088_Distance3', 'KN088_Distance2', 'KN088_Distance1', 'KN087_Distance3', 'KN087_Distance2', 'KN087_Distance1', 'KN013', 'KN014'],
        'Datum': ['', '', '', '', '', '', '', ''],
        'Property': ['L', 'L', 'L', 'L', 'L', 'L', 'L', 'L'],
        'Nominal': [3.7, 3.7, 3.7, 29.89, 29.89, 29.89, 30, 31],
        'Actual': [3.669, 3.682, 3.69, 29.427, 29.414, 0, 31, 32],
        'Actual1': [3.669, 3.682, 3.69, 29.427, 29.414, 0, 31, 32],
        'Actual2': [3.669, 3.682, 3.69, 29.427, 29.414, 0, 31, 32],
        'Actual3': [3.669, 3.682, 3.69, 29.427, 29.414, 0, 31, 32],
        'Actual4': [3.669, 3.682, 3.69, 29.427, 29.414, 0, 31, 32],
        'Tol-': [-0.1, -0.1, -0.1, -0.1, -0.1, -0.1, 0.5, 0.5],
        'Tol+': [0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.5, 0.5],
        'Dev': ['', '', '', '', '', '', '', ''],
        'Check': ['', '', '', '', '', '', '', ''],
        'Out': ['', '', '', '', '', '', '', '']
    }

df = pd.DataFrame(data)
print('################################ DATA FRAME ##########################################################')
print(df)
print('######################################################################################################')
df['Element'] = df['Element'].str.split('_').str[0]

# 'Dev','Out', 'Check' sütunlarını sil
df = df.drop(columns=['Nominal','Datum', 'Property' ,'Tol-', 'Tol+', 'Dev' ,'Out' , 'Check'])

# Sayısal sütunların minimum ve maksimum değerlerini bir arada göster, diğer sütunlar için ilk değeri al
df_group = df.groupby('Element').agg(lambda x: f"{x.min()} / {x.max()}" if pd.api.types.is_numeric_dtype(x) else x.iloc[0])

# Veriyi istediğiniz formata dönüştürme
#df_group = df_group.unstack(level=-1)
print('################################ DATA FRAME UPDATE ##########################################################')
print(df_group)
print('#############################################################################################################')
# Her bir sütun için döngü
for col in df_group.columns[0:]:
    # Her bir hücre için döngü
    for index, value in df_group[col].items():
        # Eğer / işareti varsa
        if '/' in value:
            # / işaretinden önceki ve sonraki değerlerin aynı olup olmadığını kontrol et
            parts = value.split('/')
            if parts[0].strip() == parts[1].strip():
                df_group.at[index, col] = float(parts[0].strip())
                # Eğer eşitse, / işaretinden sonraki kısmı sil
                #df.at[index, col] = parts[0].strip()
            else:
                # Eğer değerler farklı ise, işlemi atla
                continue

# Her bir sütun için döngü
for col in df_group.columns[0:]:
    # Her bir hücre için döngü
    for index, value in df_group[col].items():
        # Eğer değerde 0 varsa
        if isinstance(value, str):
            if '0.0' in value:
                parts = value.split('/')
                new_value = '/'.join(part for part in parts if part.strip() != '0.0')
                df_group.at[index, col] = new_value
        


print(df_group)

df_group.to_excel("output.xlsx")
