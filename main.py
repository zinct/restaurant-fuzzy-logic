# Fuzzy Logic System untuk Memilih 5 Restoran Terbaik
# Nama: Indra Mahesa (1302220067)
# Nama: Yazid Al Ghozali (1302223047)
# Kelas: SE4601
# Mata Kuliah: Kecerdasan Buatan

import openpyxl

def fungsi_segitiga(x, a, b, c):
    if a < x < b:
        return (x - a) / (b - a)
    elif b <= x < c:
        return (c - x) / (c - b)
    elif x == b:
        return 1.0
    else:
        return 0.0
    
def fungsi_trapesium(x, a, b, c, d):
    if x <= a or x >= d:
        return 0.0
    elif a < x < b:
        return (x - a) / (b - a)
    elif b <= x <= c:
        return 1.0
    elif c < x < d:
        return (d - x) / (d - c)
        
def membership_price(value):
    """
    Menghitung derajat keanggotaan bagian price (harga) menggunakan fungsi trapesium.
    Untuk mengetahui membership mana yang cocok dari (murah, sedang, mahal)
    """

    return {
        'murah': fungsi_trapesium(value, 0, 0, 25000, 50000),
        'sedang': fungsi_trapesium(value, 30000, 50000, 75000, 100000),
        'mahal': fungsi_trapesium(value, 80000, 110000, 150000, 150000)
    }


def membership_service(value):
    """
    Menghitung derajat keanggotaan bagian service (pelayanan) menggunakan fungsi segitiga.
    Untuk mengetahui membership mana yang cocok dari (Burug, Sedang, Bagus)
    """

    return {
        'buruk': fungsi_segitiga(value, 0, 0, 50),
        'sedang': fungsi_segitiga(value, 30, 60, 90),
        'bagus': fungsi_segitiga(value, 70, 100, 100)
    }

def inferensi_mamdani(service, price):
    rules = []

    rules += [('bagus', min(service['bagus'], price['murah']))]
    rules += [('biasa', min(service['bagus'], price['sedang']))]
    rules += [('biasa', min(service['bagus'], price['mahal']))]

    rules += [('biasa', min(service['sedang'], price['murah']))]
    rules += [('biasa', min(service['sedang'], price['sedang']))]
    rules += [('buruk', min(service['sedang'], price['mahal']))]

    rules += [('buruk', min(service['buruk'], price['murah']))]
    rules += [('buruk', min(service['buruk'], price['sedang']))]
    rules += [('buruk', min(service['buruk'], price['mahal']))]

    hasil = {'buruk': 0, 'biasa': 0, 'bagus': 0}
    for kategori, nilai in rules:
        hasil[kategori] = max(hasil[kategori], nilai)

    return hasil

# Defuzzification with Center of Gravity
def defuzzification(result):
    bobot = { 'buruk': 20, 'biasa': 60, 'bagus': 90 }

    total_bobot = (result['buruk'] * bobot['buruk'] +
                   result['biasa'] * bobot['biasa'] +
                   result['bagus'] * bobot['bagus'])

    total_degree = (result['buruk'] +
                         result['biasa'] +
                         result['bagus'])

    if total_degree == 0:
        return 0
    
    return total_bobot / total_degree

def result_output(filename, result):
   file = openpyxl.Workbook()
   fileActive = file.active
   fileActive.append(['Restoran Id', 'Pelayanan', 'Harga', 'Nilai Kelayakan', 'Kualitas'])
   for item in result:
      fileActive.append([item['id'], item['pelayanan'], item['harga'], item['skor'], item['keterangan']])
    
   file.save(filename)

def read_excel():
    file = openpyxl.open("restoran.xlsx")
    fileActive = file.active
    data = []
    for row in fileActive.iter_rows(min_row=2, values_only=True):
      id_customer, service, price = row 
      data.append({'id': id_customer, 'pelayanan': service, 'harga': price})
    return data

def quality(skor):
    if skor < 35:
        return "Buruk"
    elif skor < 65:
        return "Biasa Saja"
    else:
        return "Sangat Baik"
    
def main():
    data = read_excel()
    hasil = []

    for restoran in data:
        mb_service = membership_service(restoran['pelayanan'])
        mb_price = membership_price(restoran['harga'])
        inference_result = inferensi_mamdani(mb_service, mb_price)
        nilai_kelayakan = defuzzification(inference_result)
        score = quality(nilai_kelayakan)
        hasil.append({'id': restoran['id'], 'pelayanan': restoran['pelayanan'],
                      'harga': restoran['harga'], 'skor': nilai_kelayakan, 'keterangan': score})

    hasil = sorted(hasil, key=lambda x: x['skor'], reverse=True)
    top5 = hasil[:5]

    result_output('peringkat.xlsx', top5)

    print("5 Restoran Terbaik:")
    for item in top5:
        print(f"ID: {item['id']}, Pelayanan: {item['pelayanan']}, Harga: {item['harga']}, Skor: {item['skor']:.2f}, Kualitas: {item['keterangan']}")

if __name__ == "__main__":
    main()