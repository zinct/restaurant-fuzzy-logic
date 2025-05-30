# Fuzzy Logic System untuk Memilih 5 Restoran Terbaik
# Nama: Indra Mahesa (1302220067)
# Nama: Yazid Al Ghozali (1302223047)
# Kelas: SE4601
# Mata Kuliah: Kecerdasan Buatan

import openpyxl

# Fungsi Membership (Parafrase)

def membership_servis(servis):
    """
    Menghitung nilai keanggotaan fuzzy untuk kategori servis: buruk, sedang, bagus.
    Menggunakan bentuk segitiga untuk setiap kategori.
    """
    # Kategori Buruk: 0-50
    if servis <= 25:
        buruk = 1.0
    elif 25 < servis <= 50:
        buruk = (50 - servis) / 25.0
    else:
        buruk = 0.0

    # Kategori Sedang: 40-70
    if 40 < servis <= 55:
        sedang = (servis - 40) / 15.0
    elif 55 < servis <= 70:
        sedang = (70 - servis) / 15.0
    else:
        sedang = 0.0

    # Kategori Bagus: 60-100
    if 60 < servis <= 80:
        bagus = (servis - 60) / 20.0
    elif 80 < servis <= 100:
        bagus = 1.0
    else:
        bagus = 0.0

    return {
        'buruk': max(0.0, min(buruk, 1.0)),
        'sedang': max(0.0, min(sedang, 1.0)),
        'bagus': max(0.0, min(bagus, 1.0))
    }

def membership_harga(harga):
    """
    Menghitung nilai keanggotaan fuzzy untuk kategori harga: murah, sedang, mahal.
    Setiap kategori menggunakan fungsi segitiga.
    """
    # Murah: 0 - 30.000
    if harga <= 20000:
        murah = 1.0
    elif 20000 < harga <= 30000:
        murah = (30000 - harga) / 10000.0
    else:
        murah = 0.0

    # Sedang: 25.000 - 45.000
    if 25000 < harga <= 35000:
        sedang = (harga - 25000) / 10000.0
    elif 35000 < harga <= 45000:
        sedang = (45000 - harga) / 10000.0
    else:
        sedang = 0.0

    # Mahal: 40.000 - 70.000
    if 40000 < harga <= 55000:
        mahal = (harga - 40000) / 15000.0
    elif 55000 < harga <= 70000:
        mahal = 1.0
    else:
        mahal = 0.0

    return {
        'murah': max(0.0, min(murah, 1.0)),
        'sedang': max(0.0, min(sedang, 1.0)),
        'mahal': max(0.0, min(mahal, 1.0))
    }

# Fungsi Inferensi (Parafrase)
def fuzzy_inference(keanggotaan_servis, keanggotaan_harga):
    """
    Melakukan proses inferensi fuzzy berdasarkan aturan yang telah ditentukan.
    Output berupa derajat keanggotaan untuk setiap kategori kualitas.
    """
    aturan = [
        ('sangat_baik', min(keanggotaan_servis['bagus'], keanggotaan_harga['murah'])),
        ('biasa_saja', min(keanggotaan_servis['bagus'], keanggotaan_harga['sedang'])),
        ('biasa_saja', min(keanggotaan_servis['bagus'], keanggotaan_harga['mahal'])),
        ('biasa_saja', min(keanggotaan_servis['sedang'], keanggotaan_harga['murah'])),
        ('biasa_saja', min(keanggotaan_servis['sedang'], keanggotaan_harga['sedang'])),
        ('buruk', min(keanggotaan_servis['sedang'], keanggotaan_harga['mahal'])),
        ('buruk', min(keanggotaan_servis['buruk'], keanggotaan_harga['murah'])),
        ('buruk', min(keanggotaan_servis['buruk'], keanggotaan_harga['sedang'])),
        ('buruk', min(keanggotaan_servis['buruk'], keanggotaan_harga['mahal'])),
    ]
    hasil = {'buruk': 0.0, 'biasa_saja': 0.0, 'sangat_baik': 0.0}
    for kategori, nilai in aturan:
        if nilai > hasil[kategori]:
            hasil[kategori] = nilai
    return hasil

# Fungsi Defuzzifikasi (Parafrase)
def fuzzy_defuzzification(hasil_fuzzy):
    """
    Melakukan defuzzifikasi menggunakan metode rata-rata berbobot.
    Bobot:
    - buruk: 25
    - biasa_saja: 50
    - sangat_baik: 85
    """
    bobot_kategori = {
        'buruk': 25,
        'biasa_saja': 50,
        'sangat_baik': 85
    }
    total_nilai = sum(hasil_fuzzy[k] * bobot_kategori[k] for k in hasil_fuzzy)
    total_keanggotaan = sum(hasil_fuzzy.values())
    if total_keanggotaan == 0:
        return 0.0
    return total_nilai / total_keanggotaan

# Fungsi Keterangan Kualitas (Parafrase)
def kualitas_keterangan(nilai_skor):
    """
    Mengembalikan label kualitas berdasarkan skor defuzzifikasi.
    """
    if nilai_skor < 35:
        return "Buruk"
    elif nilai_skor < 65:
        return "Biasa Saja"
    else:
        return "Sangat Baik"

# Fungsi Membaca Data Excel (Parafrase)
def ambil_data_excel(namafile):
    """
    Membaca data restoran dari file Excel dan mengembalikan list dictionary.
    """
    workbook = openpyxl.load_workbook(namafile)
    sheet = workbook.active
    data_restoran = []
    for baris in sheet.iter_rows(min_row=2, values_only=True):
        id_restoran, nilai_servis, nilai_harga = baris
        data_restoran.append({
            'id': id_restoran,
            'pelayanan': nilai_servis,
            'harga': nilai_harga
        })
    return data_restoran

# Fungsi Simpan Output Excel (Parafrase)
def simpan_hasil_excel(namafile, data_hasil):
    """
    Menyimpan hasil peringkat restoran ke file Excel.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['ID Restoran', 'Pelayanan', 'Harga', 'Skor Kelayakan', 'Keterangan Kualitas'])
    for entri in data_hasil:
        ws.append([
            entri['id'],
            entri['pelayanan'],
            entri['harga'],
            entri['skor'],
            entri['keterangan']
        ])
    wb.save(namafile)

# Main Program (Parafrase)
def main():
    data = ambil_data_excel('restoran.xlsx')
    hasil_akhir = []
    for resto in data:
        keanggotaan_servis = membership_servis(resto['pelayanan'])
        keanggotaan_harga = membership_harga(resto['harga'])
        hasil_fuzzy = fuzzy_inference(keanggotaan_servis, keanggotaan_harga)
        skor = fuzzy_defuzzification(hasil_fuzzy)
        label = kualitas_keterangan(skor)
        hasil_akhir.append({
            'id': resto['id'],
            'pelayanan': resto['pelayanan'],
            'harga': resto['harga'],
            'skor': skor,
            'keterangan': label
        })
    hasil_akhir = sorted(hasil_akhir, key=lambda x: x['skor'], reverse=True)
    lima_terbaik = hasil_akhir[:5]
    simpan_hasil_excel('output/peringkat.xlsx', lima_terbaik)
    print("5 Restoran Terbaik:")
    for entri in lima_terbaik:
        print(f"ID: {entri['id']}, Pelayanan: {entri['pelayanan']}, Harga: {entri['harga']}, Skor: {entri['skor']:.2f}, Kualitas: {entri['keterangan']}")

if __name__ == "__main__":
    main()