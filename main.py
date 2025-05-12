from docxtpl import DocxTemplate
import datetime
import locale
import re

# Daftar nama bulan yang valid dalam bahasa Indonesia
BULAN_INDONESIA = {
    'januari': 1, 'februari': 2, 'maret': 3, 'april': 4, 'mei': 5, 'juni': 6,
    'juli': 7, 'agustus': 8, 'september': 9, 'oktober': 10, 'november': 11, 'desember': 12
}

def validate_date(date_str):
    try:
        # Memisahkan input menjadi hari, bulan, dan tahun
        pattern = r'^(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})$'
        match = re.match(pattern, date_str)
        
        if not match:
            return False, "Format tanggal harus: DD Bulan YYYY (contoh: 09 Mei 2025)"
        
        hari, bulan, tahun = match.groups()
        hari = int(hari)
        bulan = bulan.lower()
        tahun = int(tahun)
        
        # Validasi bulan
        if bulan not in BULAN_INDONESIA:
            return False, f"Bulan tidak valid. Gunakan salah satu dari: {', '.join(BULAN_INDONESIA.keys())}"
        
        # Validasi hari (1-31)
        if not (1 <= hari <= 31):
            return False, "Hari harus antara 1 dan 31"
        
        # Validasi tahun (tahun sekarang sampai 10 tahun ke depan)
        current_year = datetime.datetime.now().year
        if not (current_year <= tahun <= current_year + 10):
            return False, f"Tahun harus antara {current_year} dan {current_year + 10}"
        
        # Format ulang tanggal dengan kapitalisasi yang benar
        formatted_date = f"{hari:02d} {bulan.capitalize()} {tahun}"
        return True, formatted_date
        
    except Exception as e:
        return False, "Format tanggal tidak valid"

def get_user_input():
    print("\n=== Generator Laporan Atensi Pimpinan ===\n")
    
    # Get basic information
    tujuan = []
    print("Masukkan tujuan laporan (tekan Enter dua kali jika selesai):")
    while True:
        item = input("> ")
        if item == "":
            break
        tujuan.append(item)

    print("\nKEGIATAN")
    kegiatan = input("Masukkan nama kegiatan: ")
    
    print("\nWAKTU DAN TEMPAT")
    hari = input("Hari: ")
    
    # Input dan validasi tanggal
    while True:
        tanggal = input("Tanggal (contoh: 09 Mei 2025): ")
        is_valid, result = validate_date(tanggal)
        if is_valid:
            tanggal = result
            break
        print(f"Error: {result}")
    
    tempat = input("Tempat: ")
    
    print("\nPeserta dari (instansi):")
    peserta_instansi = []
    while True:
        item = input("> ")
        if item == "":
            break
        peserta_instansi.append(item)
    
    print("\nPeserta dari Instansi Lain:")
    peserta_lain = []
    while True:
        item = input("> ")
        if item == "":
            break
        peserta_lain.append(item)
    
    print("\nURAIAN KEGIATAN")
    print("Masukkan poin-poin uraian kegiatan (tekan Enter dua kali jika selesai):")
    uraian = []
    counter = 1
    while True:
        item = input(f"{counter}. ")
        if item == "":
            break
        uraian.append(item)
        counter += 1

    kota = input("\nKota: ")
    
    # Create context dictionary
    context = {
        "tujuan": tujuan,
        "kegiatan": kegiatan,
        "hari": hari,
        "tanggal": tanggal,
        "tempat": tempat,
        "peserta_instansi": peserta_instansi,
        "peserta_lain": peserta_lain,
        "uraian": uraian,
        "kota": kota,
        "tanggal_laporan": tanggal
    }
    
    return context

def generate_report(template_path, output_path, data):
    # Load the template
    doc = DocxTemplate(template_path)
    
    # Render the template with your data
    doc.render(data)
    
    # Save the generated document
    doc.save(output_path)

if __name__ == "__main__":
    # Get data from user
    context = get_user_input()
    
    # Generate the report
    generate_report(
        "template.docx",  # Your template file
        "laporan_atensi.docx",  # Output file
        context
    )
    print("\nLaporan berhasil dibuat!")
