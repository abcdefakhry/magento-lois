import streamlit as st
import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
import time
import io

# Konfigurasi Tampilan
st.set_page_config(page_title="Magento Lois to CSV/Excel", page_icon="logolois.png", layout="centered")

st.title("Magento Lois to CSV/Excel")
st.write("Website online untuk mengunduh data Magento Lois secara otomatis ke Excel/CSV.")

st.divider()

url_target = st.text_input(
    "URL Target Magento Grid:",
    placeholder="Contoh: https://loisjeans.id/jalanbelakang/lac/transaction/index/key/..."
)
admin_cookie = st.text_input(
    "Value Admin Cookie ('admin'):",
    placeholder="Contoh: r9ed7mjavu8pij1f6s5umjk193"
)

col1, col2 = st.columns(2)
with col1:
    from_date = st.text_input("Dari Tanggal (MM/DD/YYYY):", placeholder="Contoh: 08/05/2026")
with col2:
    to_date = st.text_input("Sampai Tanggal (MM/DD/YYYY):", placeholder="Contoh: 08/06/2026")

custom_file_name = st.text_input("Nama File Simpan (tanpa ekstensi):", placeholder="Contoh: magento_lois_2026")
output_format = st.radio("Format File Output:", ["Excel (.xlsx)", "CSV (.csv)"], horizontal=True)

start_btn = st.button("MULAI UNDUH DATA", type="primary", use_container_width=True)

if start_btn:
    if not url_target.strip() or not admin_cookie.strip():
        st.error("Mohon lengkapi URL Target dan Cookie Admin!")
    elif not custom_file_name.strip():
        st.error("Mohon isi nama file yang diinginkan!")
    else:
        status_box = st.status("Sedang memproses scraping data...", expanded=True)

        headers = {
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Cookie": f"admin={admin_cookie.strip()}"
        }

        session = requests.Session()
        session.headers.update(headers)

        all_records = []
        headers_list = [
            "Created At", "Transaction No.", "Transaction Type", "LAC No.",
            "Status", "Point", "Point Amount", "Actual Point",
            "Actual Point Amount", "Remark", "Expire At"
        ]

        import base64

        try:
            status_box.write("Menghubungkan ke server Magento...")

            # Clean URL (Hilangkan part filter jika ada)
            clean_url = url_target.strip().split('/filter/')[0].rstrip('/') + '/'

            f_from = from_date.strip()
            f_to = to_date.strip()

            if f_from and f_to:
                status_box.write("Menerapkan filter tanggal ke URL Magento...")
                filter_str = f"created_at[from]={f_from}&created_at[to]={f_to}"
                # Gunakan base64 encoding standar Magento
                filter_b64 = base64.b64encode(filter_str.encode('utf-8')).decode('utf-8')
                clean_url = clean_url + f"filter/{filter_b64}/"

            # 2. BUKA HALAMAN DENGAN LIMIT 200 PER PAGE
            res = session.get(clean_url, params={'limit': 200, 'p': 1}, timeout=25)

            if res.status_code != 200:
                st.error(f"Gagal mengakses web. HTTP Status: {res.status_code}. Pastikan Cookie 'admin' masih aktif/belum expired.")
                status_box.update(label="Proses Gagal", state="error")
            else:
                soup = BeautifulSoup(res.text, 'html.parser')
                page_text = soup.get_text()
                match = re.search(r'of\s+(\d+)', page_text)
                total_pages = int(match.group(1)) if match else 1

                status_box.write(f"Total Halaman Terdeteksi: {total_pages}")
                progress_bar = st.progress(0)

                for p_num in range(1, total_pages + 1):
                    status_box.write(f"[{p_num}/{total_pages}] Memproses Halaman {p_num}...")

                    p_res = session.get(clean_url, params={'limit': 200, 'p': p_num}, timeout=25)
                    p_soup = BeautifulSoup(p_res.text, 'html.parser')
                    rows = p_soup.select('table tbody tr')

                    count = 0
                    for row in rows:
                        if 'headings' in row.get('class', []) or row.find('input') or row.find('select'):
                            continue

                        cells = row.find_all('td')
                        if len(cells) >= 5:
                            row_data = {}
                            for idx, cell in enumerate(cells):
                                col_name = headers_list[idx] if idx < len(headers_list) else f"Col_{idx+1}"
                                row_data[col_name] = cell.get_text(strip=True)

                            all_records.append(row_data)
                            count += 1

                    progress_bar.progress(p_num / total_pages)
                    time.sleep(0.3)

                if all_records:
                    status_box.update(label="Scraping Selesai Secara Sempurna!", state="complete")
                    st.success(f"Berhasil menarik total {len(all_records)} baris data!")

                    df = pd.DataFrame(all_records)

                    clean_filename = re.sub(r'[\\/*?:"<>|]', "", custom_file_name.strip())
                    clean_filename = re.sub(r'\.(xlsx|csv)$', '', clean_filename, flags=re.IGNORECASE)

                    if output_format == "Excel (.xlsx)":
                        buffer = io.BytesIO()
                        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                            df.to_excel(writer, index=False)

                        final_filename = f"{clean_filename}.xlsx"
                        st.download_button(
                            label=f"DOWNLOAD FILE EXCEL ({final_filename})",
                            data=buffer.getvalue(),
                            file_name=final_filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary",
                            use_container_width=True
                        )
                    else:
                        csv_data = df.to_csv(index=False, encoding='utf-8-sig')
                        final_filename = f"{clean_filename}.csv"
                        st.download_button(
                            label=f"DOWNLOAD FILE CSV ({final_filename})",
                            data=csv_data,
                            file_name=final_filename,
                            mime="text/csv",
                            type="primary",
                            use_container_width=True
                        )

                    st.subheader("Preview 50 Data Teratas:")
                    st.dataframe(df.head(50))

                else:
                    st.warning("Tidak ada data ditemukan. Cek kembali format tanggal atau Cookie Admin.")

        except Exception as e:
            st.error(f"Terjadi kesalahan koneksi: {str(e)}")
            status_box.update(label="Proses Terhenti", state="error")
