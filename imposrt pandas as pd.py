import pandas as pd

output_path_excel = None

def data_merger_from_excel_file(input_path):

    global output_path_excel

    df = pd.read_excel(input_path, sheet_name='Sheet1')

    # Actual sütunundaki değeri 0'dan farklı olan satırların Check sütununa kopyalanması
    df.loc[df['Actual'] != 0, 'Check'] = df['Actual']

    # Check sütununu sayısal tipe dönüştür
    df['Check'] = pd.to_numeric(df['Check'], errors='coerce')

    df = df.drop(columns=['Datum', 'Property', 'Actual', 'Dev', 'Out'])

    # Sütun adlarından grup ismi oluştur
    df['Element'] = df['Element'].str.split('_').str[0]

    # Tüm kolonların sonuçlarını depolamak için bir sözlük oluştur
    results = {}

    # Gruplar oluştur ve her bir kolon için en küçük ve en büyük değerleri hesapla
    for column in df.columns:
        if column not in ['Element', 'Tol-', 'Tol+']:  # 'Element', 'Tol-' ve 'Tol+' sütunlarını işleme
            result = df.groupby('Element').agg({
                'Tol-': 'first',
                'Tol+': 'first',
                column: lambda x: f"{x.min()} / {x.max()}"
            }).reset_index()

            # Her bir kolonun sonucunu sözlüğe ekle
            results[column] = result

    # Sonuçları aynı Excel dosyasının 'Reduced' sayfasına yazdır
    with pd.ExcelWriter(output_path_excel, engine='openpyxl', mode='a') as writer:
        for column, result in results.items():
            result.to_excel(writer, sheet_name=column, index=False)

if __name__ == '__main__':
    input_path = input("Lütfen Excel dosyasının yolunu girin: ")
    if output_path_excel is None:
        output_path_excel = input("Lütfen sonucun yazılacağı Excel dosyasının yolunu girin: ")
    data_merger_from_excel_file(input_path)
