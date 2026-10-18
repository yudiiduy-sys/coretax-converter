import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
import re

def clean_identity(val):
    """Membersihkan karakter tersembunyi (zero-width space) & spasi"""
    if pd.isna(val):
        return ""
    return re.sub(r'[^0-9]', '', str(val))

def convert_simkeu_to_coretax(input_file, masa_pajak=8, tahun_pajak=2026, npwp_pemotong="0001734664626000"):
    print(f"🔄 Memproses file sumber: {input_file}...")
    
    # 1. BACA FILE MENTAH SIMKEU (.xls / .html)
    try:
        dfs = pd.read_html(input_file)
        df_raw = dfs[0]
        df_raw.columns = df_raw.iloc[0]
        df_raw = df_raw[1:].reset_index(drop=True)
    except Exception as e:
        print(f"❌ Gagal membaca file: {e}")
        return

    # 2. PROSES CHECKER & TRANSLATOR DATA
    errors = []
    processed_rows = []

    for idx, row in df_raw.iterrows():
        row_num = idx + 2  # Posisi baris di Excel
        
        nama = str(row.get('Nama', '')).strip()
        nik = clean_identity(row.get('NIK', ''))
        npwp = clean_identity(row.get('NPWP', ''))
        gol = str(row.get('Gol', '')).strip()
        
        try:
            bruto = float(row.get('Jumlah Bruto', 0))
            tarif = float(row.get('Tarif', 0))
            pph = float(row.get('Jumlah PPh', 0))
        except ValueError:
            errors.append(f"Baris {row_num} ({nama}): Format nominal angka tidak valid.")
            continue

        # --- VALIDASI CHECKER ---
        if len(nik) != 16:
            errors.append(f"Baris {row_num} ({nama}): NIK '{nik}' tidak 16 digit!")

        # --- MAPPING IDENTITAS & ID TKU ---
        counterpart_tin = nik if len(nik) == 16 else npwp.zfill(16)
        id_tku = counterpart_tin + "000000"  # Format ID TKU Coretax (22 Digit)

        # --- MAPPING KODE OBJEK PAJAK ---
        # Golongan IV -> 21-402-03, Lainnya -> 21-402-02
        tax_object_code = "21-402-03" if gol == 'IV' else "21-402-02"

        # Nomor Dokumen Referensi SPJ
        doc_ref_no = f"SPJ-{tahun_pajak}{str(masa_pajak).zfill(2)}-{str(idx+1).zfill(3)}"

        processed_rows.append({
            'NPWP Pemotong': npwp_pemotong,
            'Masa Pajak': masa_pajak,
            'Tahun Pajak': tahun_pajak,
            'NPWP': counterpart_tin,
            'ID TKU Penerima Penghasilan': id_tku,
            'Status PTKP': 'K/0',
            'Fasilitas': 'N/A',
            'Kode Objek Pajak': tax_object_code,
            'Penghasilan': int(bruto),
            'Deemed': 100,
            'Tarif': int(tarif),
            'Jenis Dok. Referensi': 'Contract',
            'Nomor Dok. Referensi': doc_ref_no,
            'Tanggal Dok. Referensi': f"{tahun_pajak}-{str(masa_pajak).zfill(2)}-31",
            'ID TKU Pemotong': npwp_pemotong,
            'Tanggal Pemotongan': f"{tahun_pajak}-{str(masa_pajak).zfill(2)}-31"
        })

    # --- LAPORAN HASIL CHECKER ---
    print("\n🔍 --- LAPORAN HASIL VALIDASI CHECKER ---")
    if errors:
        print(f"⚠️ Ditemukan {len(errors)} peringatan pada data:")
        for err in errors:
            print("  -", err)
    else:
        print("✅ Seluruh data Lolos Validasi! Tidak ada error NIK/Format.")

    df_clean = pd.DataFrame(processed_rows)

    # 3. EXPORT KE EXCEL SIAP IMPOR
    excel_output = "Hasil_Impor_Coretax.xlsx"
    df_clean.to_excel(excel_output, index=False)
    print(f"📄 File Excel Impor berhasil dibuat: {excel_output}")

    # 4. GENERATE STRUKTUR XML CORETAX
    root = ET.Element("TaxData")
    header = ET.SubElement(root, "Header")
    ET.SubElement(header, "WithholdingAgentTIN").text = npwp_pemotong
    ET.SubElement(header, "TaxPeriodMonth").text = str(masa_pajak)
    ET.SubElement(header, "TaxPeriodYear").text = str(tahun_pajak)

    details = ET.SubElement(root, "WithholdingDetails")
    for item in processed_rows:
        row_xml = ET.SubElement(details, "BP21")
        ET.SubElement(row_xml, "TIN").text = item['NPWP Pemotong']
        ET.SubElement(row_xml, "TaxPeriodMonth").text = str(item['Masa Pajak'])
        ET.SubElement(row_xml, "TaxPeriodYear").text = str(item['Tahun Pajak'])
        ET.SubElement(row_xml, "CounterpartTin").text = item['NPWP']
        ET.SubElement(row_xml, "IDPlaceOfBusinessActivityOfIncomeRecipient").text = item['ID TKU Penerima Penghasilan']
        ET.SubElement(row_xml, "StatusTaxExemption").text = item['Status PTKP']
        ET.SubElement(row_xml, "TaxCertificate").text = item['Fasilitas']
        ET.SubElement(row_xml, "TaxObjectCode").text = item['Kode Objek Pajak']
        ET.SubElement(row_xml, "Gross").text = str(item['Penghasilan'])
        ET.SubElement(row_xml, "Deemed").text = str(item['Deemed'])
        ET.SubElement(row_xml, "TaxRate").text = str(item['Tarif'])
        ET.SubElement(row_xml, "RefDocType").text = item['Jenis Dok. Referensi']
        ET.SubElement(row_xml, "RefDocNumber").text = item['Nomor Dok. Referensi']
        ET.SubElement(row_xml, "RefDocDate").text = item['Tanggal Dok. Referensi']

    # Format XML Rapi (Pretty Print)
    xml_str = minidom.parseString(ET.tostring(root)).toprettyxml(indent="  ")
    xml_output = "Hasil_Impor_Coretax.xml"
    with open(xml_output, "w", encoding="utf-8") as f:
        f.write(xml_str)
    
    print(f"🚀 File XML Impor Coretax berhasil dibuat: {xml_output}")

# CARA MENGGUNAKAN SKRIP:
# Cukup jalankan fungsi ini dengan menentukan nama file SIMKEU Anda:
convert_simkeu_to_coretax('17077420.xls', masa_pajak=8, tahun_pajak=2026)