df['Element'] = df['Element'].str.split('_').str[0]

# Boş bir DataFrame oluştur
results = pd.DataFrame(columns=['Element'])

for column in df.columns:
    if column != 'Element':
        result = df.groupby('Element').agg({
            column: lambda x: f"{x.min()} / {x.max()}"
        }).reset_index()
        # Sütunu sonuç DataFrame'ine ekle
        results = pd.concat([results, result[column]], axis=1)

print(results)
