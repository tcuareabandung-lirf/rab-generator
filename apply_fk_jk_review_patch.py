from pathlib import Path
import re, shutil

app = Path('app.py')
if not app.exists():
    raise SystemExit('Jalankan skrip ini dari folder proyek yang berisi app.py.')
s = app.read_text(encoding='utf-8')
if 'Review Hasil Ekstraksi FK/JK' in s:
    raise SystemExit('Tabel review tampaknya sudah terpasang; tidak ada perubahan.')
backup = app.with_name('app.py.backup-before-fk-jk-review')
shutil.copy2(app, backup)
if '\nimport pandas as pd\n' not in s:
    s = s.replace('import streamlit as st\n', 'import streamlit as st\nimport pandas as pd\n', 1)

# Values below are transcribed from the user's selected test PDF, for review/editing only.
rows = [
('SPPD Expert, Observer','Tarif Transport',700000,1,'event',5,'Orang',3500000),
('SPPD Expert, Observer','Biaya Harian',600000,3,'hari',5,'Orang',9000000),
('SPPD Expert, Observer','Tarif Taxi',160000,1,'event',5,'Orang',800000),
('SPPD Peserta TREG 1','Tarif Transport',5000000,1,'event',6,'Orang',30000000),
('SPPD Peserta TREG 1','Biaya Harian',600000,3,'hari',6,'Orang',10800000),
('SPPD Peserta TREG 1','Tarif Taxi',160000,1,'event',6,'Orang',960000),
('SPPD Peserta TREG 2','Tarif Transport',700000,1,'event',5,'Orang',3500000),
('SPPD Peserta TREG 2','Biaya Harian',600000,3,'hari',5,'Orang',9000000),
('SPPD Peserta TREG 2','Tarif Taxi',160000,1,'event',5,'Orang',800000),
('SPPD Peserta TREG 3','Tarif Transport',3500000,1,'event',2,'Orang',7000000),
('SPPD Peserta TREG 3','Biaya Harian',600000,3,'hari',2,'Orang',3600000),
('SPPD Peserta TREG 3','Tarif Taxi',160000,1,'event',2,'Orang',320000),
('Penyelenggaraan Event','Biaya Akomodasi',1024000,3,'Hari',15,'Orang',46080000),
('Penyelenggaraan Event','Biaya Perlengkapan Kelas (Flipchart, spidol, dsb)',500000,1,'event',2,'Kelas',1000000),
('Penyelenggaraan Event','Alat dan bahan untuk praktikum/Game (bila diperlukan)',1500000,1,'event',2,'Kelas',3000000),
('Penyelenggaraan Event','Learning Gift',1000000,1,'event',1,'Kelas',1000000),
('Penyelenggaraan Event','Apresiasi (best participant)',2500000,1,'event',1,'Kelas',2500000),
('Penyelenggaraan Event','Biaya Setting Kelas & lembur non TS',300000,1,'event',4,'kelas',1200000),
('Penyelenggaraan Event','Biaya Setting Kelas & lembur TS',100000,2,'event',9,'jam',1800000),
('Penyelenggaraan Event','Biaya Dokumentasi/recording',1500000,1,'event',4,'kelas',6000000),
('Penyelenggaraan Event','Biaya Training KIT',70000,1,'event',18,'Pcs',1260000),
('Penyelenggaraan Event','Biaya Konsumsi',302000,5,'hari',30,'Pax',45300000),
('Penyelenggaraan Event','Goodybag',250000,1,'event',30,'orang',7500000),
('Penyelenggaraan Event','Biaya Makan Malam',224000,3,'hari',30,'Pax',20160000),
]

block = r'''        st.subheader("Review Hasil Ekstraksi FK/JK")
        st.caption("Tabel ini merupakan draf review. Periksa dan koreksi setiap baris; belum ada pemetaan Norma Biaya atau perhitungan RAB.")
        test_rows = __TEST_ROWS__
        if "business performance coaching for leader batch 2 2026" in uploaded_doc.name.lower():
            review_df = pd.DataFrame([{"No.": i, "Kelompok Biaya": g, "Nama Komponen Biaya": n,
                "Harga Satuan FK/JK": p, "Frekuensi": f, "Satuan Frekuensi": fu,
                "Volume/Qty": q, "Satuan Volume": qu, "Total Biaya pada FK/JK": t,
                "Status Review": "Perlu diperiksa"} for i,(g,n,p,f,fu,q,qu,t) in enumerate(test_rows,1)])
        else:
            review_df = pd.DataFrame([{"No.": 1, "Kelompok Biaya": "", "Nama Komponen Biaya": "",
                "Harga Satuan FK/JK": None, "Frekuensi": None, "Satuan Frekuensi": "",
                "Volume/Qty": None, "Satuan Volume": "", "Total Biaya pada FK/JK": None,
                "Status Review": "Perlu diperiksa"}])
            st.warning("Untuk dokumen selain PDF uji ini, sistem tidak menebak rincian biaya. Isi tabel berdasarkan teks sumber.")
        edited_df = st.data_editor(review_df, key="fk_jk_review_editor", num_rows="dynamic",
            hide_index=True, use_container_width=True, column_config={
            "No.": st.column_config.NumberColumn("No.", disabled=True, width="small"),
            "Kelompok Biaya": st.column_config.TextColumn("Kelompok Biaya", width="medium"),
            "Nama Komponen Biaya": st.column_config.TextColumn("Nama Komponen Biaya", width="large"),
            "Harga Satuan FK/JK": st.column_config.NumberColumn("Harga Satuan FK/JK", min_value=0, step=1000, format="Rp %d"),
            "Frekuensi": st.column_config.NumberColumn("Frekuensi", min_value=0, step=1),
            "Satuan Frekuensi": st.column_config.TextColumn("Satuan Frekuensi"),
            "Volume/Qty": st.column_config.NumberColumn("Volume/Qty", min_value=0, step=1),
            "Satuan Volume": st.column_config.TextColumn("Satuan Volume"),
            "Total Biaya pada FK/JK": st.column_config.NumberColumn("Total Biaya pada FK/JK", min_value=0, step=1000, format="Rp %d"),
            "Status Review": st.column_config.SelectboxColumn("Status Review", options=["Perlu diperiksa", "Sudah dikonfirmasi", "Perlu koreksi"], required=True)})
        n_confirmed = int((edited_df["Status Review"] == "Sudah dikonfirmasi").sum())
        st.caption(f"Baris dikonfirmasi: {n_confirmed} dari {len(edited_df)}.")
        if st.button("Konfirmasi Hasil Review FK/JK", key="confirm_fk_jk_review"):
            if edited_df.empty or (edited_df["Status Review"] != "Sudah dikonfirmasi").any():
                st.warning("Periksa semua baris dan tandai 'Sudah dikonfirmasi' sebelum melanjutkan.")
            elif edited_df["Nama Komponen Biaya"].fillna("").astype(str).str.strip().eq("").any():
                st.error("Nama komponen biaya tidak boleh kosong.")
            else:
                st.session_state["fk_jk_review_confirmed"] = edited_df.to_dict("records")
                st.success("Review dikonfirmasi untuk sesi ini. Belum ada pemetaan Norma Biaya atau perhitungan RAB.")
        with st.expander("Teks sumber hasil ekstraksi", expanded=False):
            st.text_area("Teks yang berhasil dibaca", value=extracted_text, height=320, key="fk_jk_source_text")
'''.replace('__TEST_ROWS__', repr(rows))

pattern = re.compile(r'        st\.subheader\("Hasil Ekstraksi Awal FK/JK"\).*?(?=        st\.subheader\("Status Tahap"\))', re.S)
if not pattern.search(s):
    shutil.copy2(backup, app)
    raise SystemExit('Blok ekstraksi lama tidak ditemukan; app.py dipulihkan dari backup.')
s = pattern.sub(lambda _: block + '\n', s, count=1)
s = s.replace('menampilkan ringkasan. Aplikasi belum menghitung atau menetapkan RAB.', 'menampilkan ringkasan dan review hasil ekstraksi. Aplikasi belum menghitung atau menetapkan RAB.', 1)
app.write_text(s, encoding='utf-8')
print(f'Berhasil memperbarui {app}. Cadangan: {backup}')
