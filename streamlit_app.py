import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from xml.dom import minidom
import io
import re

# Konfigurasi Halaman Web
st.set_page_config(
    page_title="Converter SIMKEU to Coretax PPh 21",
    page_icon="📊",
    layout="centered"
)

st.title("🏛️ Konverter Otomatis SIMKEU ke Coretax PPh 21")
st.markdown("""
Aplikasi web untuk **Checker & Transformer** data PPh 21 dari SIMKEU 
agar langsung siap diimpor ke portal Coretax DJP.
""")

st.markdown("---")

# Bagian Upload File
uploaded_file = st.file_uploader("Unggah File Mentah SIMKEU (.xls / .xlsx)", type=["xls", "xlsx"])

if uploaded_file is not None:
    try:
        # Baca file HTML/Excel
        dfs = pd.read_html(uploaded_file)
        df_raw = dfs[0]
        df_raw.columns = df_raw.iloc[0]
        df_raw = df_raw[1:].reset_index(drop=True)
        
        st.success(f"✅ File berhasil diunggah! Ditemukan {len(df_raw)} data pegawai.")
        
        # Pengaturan Masa dan NPWP Pemotong
        col1, col2, col3 = st.columns(3)
        with col1:
            masa_pajak = st.number_input("Masa Pajak (Bulan)", min_value=1, max_value=12, value=8)
        with col2:
            tahun_pajak = st.number_input("Tahun Pajak", min_value=2024, max_value=2030, value=2026)
        with col3:
            npwp_pemotong = st.text_input("NPWP Pemotong", value="0001734664626000")

        if st.button("🚀 Proses & Validasi Data"):
            errors = []
            processed_rows = []

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
                    errors.append(f"Baris {row_num} ({nama}): Nominal angka tidak valid.")
                    continue

                if len(nik) != 16:
                    errors.append(f"Baris {row_num} ({nama}): NIK '{nik}' tidak valid (harus 16 digit).")

                counterpart_tin = nik if len(nik) == 16 else npwp.zfill(16)
                id_tku = counterpart_tin + "000000"
                tax_object_code = "21-402-03" if gol == 'IV' else "21-402-02"
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

            if errors:
                st.warning("⚠️ Ditemukan beberapa peringatan validasi pada data:")
                for err in errors:
                    st.text(err)
            else:
                st.info("✨ Checker selesai: Seluruh data valid dan bersih dari anomali karakter!")

            df_clean = pd.DataFrame(processed_rows)

            # Buat file Excel di memori
            output_excel = io.BytesIO()
            with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
                df_clean.to_excel(writer, index=False)
            excel_data = output_excel.getvalue()

            # Buat file XML di memori
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

            xml_str = minidom.parseString(ET.tostring(root)).toprettyxml(indent="  ")

            st.markdown("---")
            st.subheader("📥 Unduh Hasil Dokumen Matang")
            
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                st.download_button(
                    label="📄 Download Excel Siap Impor",
                    data=excel_data,
                    file_name="Hasil_Impor_Coretax.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with col_d2:
                st.download_button(
                    label="🚀 Download XML Coretax",
                    data=xml_str,
                    file_name="Hasil_Impor_Coretax.xml",
                    mime="application/xml"
                )

    except Exception as e:
        st.error(f"Terjadi kesalahan saat membaca file: {e}")