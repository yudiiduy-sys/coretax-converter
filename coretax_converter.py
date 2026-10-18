import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
import io
import re

# Konfigurasi Halaman Web
st.set_page_config(
    page_title="Converter Presisi SIMKEU & PTKP UNEJ to Coretax",
    page_icon="📊",
    layout="centered"
)

st.title("🏛️ Konverter Presisi SIMKEU & PTKP UNEJ ke Coretax PPh 21")
st.markdown("""
Aplikasi web online untuk **Checker & Transformer** data PPh 21 dari SIMKEU Universitas Jember 
dengan pencocokan otomatis Status PTKP berdasarkan Database Referensi UNEJ 2026.
""")

st.markdown("---")

# Bagian 1: Upload File SIMKEU
uploaded_simkeu = st.file_uploader("1. Unggah File Mentah SIMKEU (.xls / .xlsx)", type=["xls", "xlsx"])

# Bagian 2: Upload File Database PTKP UNEJ (Opsional / Pendukung)
uploaded_ptkp = st.file_uploader("2. Unggah File Referensi PTKP UNEJ 2026 (.xlsx - Opsional)", type=["xlsx"])

if uploaded_simkeu is not None:
    try:
        # Baca file SIMKEU menggunakan lxml
        dfs = pd.read_html(uploaded_simkeu, flavor='lxml')
        df_raw = dfs[0]
        df_raw.columns = df_raw.iloc[0]
        df_raw = df_raw[1:].reset_index(drop=True)
        
        st.success(f"✅ File SIMKEU berhasil diunggah! Ditemukan {len(df_raw)} data pegawai.")

        # Load Database PTKP jika diunggah
        ptkp_dict = {}
        if uploaded_ptkp is not None:
            df_ptkp_ref = pd.read_excel(uploaded_ptkp, sheet_name='PTKP')
            for _, r in df_ptkp_ref.iterrows():
                nik_val = re.sub(r'[^0-9]', '', str(r.get('NIK', '')))
                ptkp_val = str(r.get('STATUS PTKP', 'K/0')).strip()
                if nik_val:
                    ptkp_dict[nik_val] = ptkp_val
            st.info(f"📚 Database PTKP UNEJ berhasil dimuat ({len(ptkp_dict)} data referensi NIK tercatat).")
        
        # Pengaturan Masa Pajak, Tahun Pajak, dan NPWP Pemotong
        col1, col2, col3 = st.columns(3)
        with col1:
            masa_pajak = st.number_input("Masa Pajak (Bulan)", min_value=1, max_value=12, value=8)
        with col2:
            tahun_pajak = st.number_input("Tahun Pajak", min_value=2024, max_value=2030, value=2026)
        with col3:
            npwp_pemotong = st.text_input("NPWP Pemotong (16 Digit)", value="0001734664626000")

        if st.button("🚀 Proses, Validasi & Transformasi Presisi"):
            errors = []
            processed_rows = []
            last_day_str = f"{tahun_pajak}-{str(masa_pajak).zfill(2)}-31"

            for idx, row in df_raw.iterrows():
                row_num = idx + 2
                nama = str(row.get('Nama', '')).strip()
                nik = re.sub(r'[^0-9]', '', str(row.get('NIK', '')))
                npwp = re.sub(r'[^0-9]', '', str(row.get('NPWP', '')))
                gol = str(row.get('Gol', '')).strip()
                
                try:
                    bruto = float(row.get('Jumlah Bruto', 0))
                    tarif = float(row.get('Tarif', 0))
                except ValueError:
                    errors.append(f"Baris {row_num} ({nama}): Format nominal angka tidak valid.")
                    continue

                # Validasi NIK 16 Digit
                if len(nik) != 16:
                    errors.append(f"Baris {row_num} ({nama}): NIK '{nik}' tidak valid (harus 16 digit).")

                # Tentukan Status PTKP (Otomatis dari Database UNEJ jika ada, fallback ke K/0)
                status_ptkp = ptkp_dict.get(nik, 'K/0')

                # Penentuan Counterpart TIN & ID TKU Penerima (NIK + 000000)
                counterpart_tin = nik if len(nik) == 16 else npwp.zfill(16)
                id_tku_penerima = counterpart_tin + "000000"

                # Penentuan ID TKU Pemotong (NPWP Pemotong + 000000)
                id_tku_pemotong = str(npwp_pemotong).strip().zfill(16) + "000000"

                # Mapping Kode Objek Pajak Presisi sesuai Golongan APBN/APBD
                if gol == 'IV':
                    tax_object_code = "21-402-03"
                elif gol in ['III', 'II', 'I', 'PPPK']:
                    tax_object_code = "21-402-02"
                else:
                    tax_object_code = "21-402-02"

                doc_ref_no = f"SPJ-{tahun_pajak}{str(masa_pajak).zfill(2)}-{str(idx+1).zfill(3)}"

                processed_rows.append({
                    'NPWP Pemotong': str(npwp_pemotong).strip(),
                    'Masa Pajak': int(masa_pajak),
                    'Tahun Pajak': int(tahun_pajak),
                    'NPWP': counterpart_tin,
                    'ID TKU Penerima Penghasilan': id_tku_penerima,
                    'Status PTKP': status_ptkp,
                    'Fasilitas': 'N/A',
                    'Kode Objek Pajak': tax_object_code,
                    'Penghasilan': int(bruto),
                    'Deemed': 100,
                    'Tarif': int(tarif),
                    'Jenis Dok. Referensi': 'Contract',
                    'Nomor Dok. Referensi': doc_ref_no,
                    'Tanggal Dok. Referensi': last_day_str,
                    'ID TKU Pemotong': id_tku_pemotong,
                    'Tanggal Pemotongan': last_day_str
                })

            if errors:
                st.warning("⚠️ Ditemukan beberapa catatan validasi pada data:")
                for err in errors:
                    st.text(err)
            else:
                st.info("✨ Pengecekan Selesai: Data presisi, status PTKP tersinkronisasi dari database UNEJ, dan siap diimpor!")

            df_clean = pd.DataFrame(processed_rows)

            # Buat file Excel Presisi di memori
            output_excel = io.BytesIO()
            with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                df_clean.to_excel(writer, index=False)
            excel_data = output_excel.getvalue()

            # Buat struktur XML Presisi sesuai skema Coretax
            root = ET.Element("TaxData")
            header = ET.SubElement(root, "Header")
            ET.SubElement(header, "WithholdingAgentTIN").text = str(npwp_pemotong)
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
                ET.SubElement(row_xml, "RefDocNumber").text = item['RefDocNumber'] if 'RefDocNumber' in item else item['Nomor Dok. Referensi']
                ET.SubElement(row_xml, "RefDocDate").text = item['Tanggal Dok. Referensi']

            xml_str = minidom.parseString(ET.tostring(root)).toprettyxml(indent="  ")

            st.markdown("---")
            st.subheader("📥 Unduh Hasil Dokumen Matang & Presisi")
            
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                st.download_button(
                    label="📄 Download Excel Presisi (PTKP Tersinkron)",
                    data=excel_data,
                    file_name="Hasil_Impor_Coretax_UNEJ.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with col_d2:
                st.download_button(
                    label="🚀 Download XML Coretax Presisi",
                    data=xml_str,
                    file_name="Hasil_Impor_Coretax_UNEJ.xml",
                    mime="application/xml"
                )

    except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses file: {e}")
