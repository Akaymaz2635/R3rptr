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
df['Element'] = df['Element'].str.split('_').str[0]

# 'Dev', 'Out', 'Check', 'Nominal', 'Property', 'Tol-', 'Tol+' sütunlarını sil
df = df.drop(columns=['Nominal', 'Property', 'Tol-', 'Tol+', 'Dev', 'Out', 'Check'])

# Sayısal sütunların / işaretinden önceki ve sonraki değerlerini birbirine eşitse, sadece birini al, diğerini sil
def process_column(x):
    if pd.api.types.is_string_dtype(x):
        if len(x.str.split('/')) > 1 and (x.str.split('/')[0] == x.str.split('/')[1]).all():
            return x.str.split('/')[0]
    return x.iloc[0]

df_group = df.groupby('Element').agg(lambda x: process_column(x))

print(df_group)
