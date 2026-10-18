
from pathlib import Path
from datetime import date

import streamlit as st
from openpyxl import load_workbook

# =====================================================
# KONFIGURASI APLIKASI
# =====================================================

st.set_page_config(
    page_title="RAB Generator",
    page_icon="📊",
    layout="wide",
)

MASTER_FILE = Path("Master/Master_RAB_Generator.xlsx")

st.title("RAB Generator")
st.caption(
    "V1 — Form data event, upload FK/JK, dan pemeriksaan input"
)


# =====================================================
# MEMBACA MASTER EXCEL
# =====================================================

@st.cache_data
def load_master_data(file_path, modified_time):
    if not Path(file_path).exists():
        raise FileNotFoundError(
            f"File master tidak ditemukan: {file_path}"
        )

    wb = load_workbook(file_path, data_only=True)

    required_sheets = [
        "TEMPLATE RAB",
        "INDEKS RAB",
        "COST CENTER & BEBAN ANGGARAN",
        "NORMA BIAYA",
    ]

    missing = [s for s in required_sheets if s not in wb.sheetnames]

    if missing:
        raise ValueError(
            "Sheet berikut tidak ditemukan: " + ", ".join(missing)
        )

    ws_norma = wb["NORMA BIAYA"]
    ws_indeks = wb["INDEKS RAB"]
    ws_cc = wb["COST CENTER & BEBAN ANGGARAN"]

    # Kategori yang sudah teridentifikasi pada kolom B
    categories = sorted({
        str(row[0]).strip()
        for row in ws_norma.iter_rows(
            min_row=2, min_col=2, max_col=2, values_only=True
        )
        if row[0] is not None
        and str(row[0]).strip() in {"R1", "R2", "L1", "L2"}
    })

    # Indeks lokasi: kolom G = nama/lokasi, H = indeks
    locations = []
    for row in ws_indeks.iter_rows(
        min_row=2, min_col=7, max_col=8, values_only=True
    ):
        name, index = row
        if name is not None and str(name).strip():
            locations.append({
                "name": str(name).strip(),
                "index": index,
            })

    # Cost Center: kolom B = kode, C = nama/keterangan
    cost_centers = []
    for row in ws_cc.iter_rows(
        min_row=2, min_col=2, max_col=3, values_only=True
    ):
        code, description = row
        if code is not None and str(code).strip():
            cost_centers.append({
                "code": str(code).strip(),
                "description": (
                    str(description).strip()
                    if description is not None else ""
                ),
            })

    # Akun beban: kolom E = kode, F = nama/keterangan
    accounts = []
    for row in ws_cc.iter_rows(
        min_row=2, min_col=5, max_col=6, values_only=True
    ):
        code, description = row
        if code is not None and str(code).strip():
            accounts.append({
                "code": str(code).strip(),
                "description": (
                    str(description).strip()
                    if description is not None else ""
                ),
            })

    return categories, locations, cost_centers, accounts


try:
    if not MASTER_FILE.exists():
        st.error(
            "File master belum tersedia. Pastikan file "
            "`Master_RAB_Generator.xlsx` berada di folder `master/`."
        )
        st.stop()

    categories, locations, cost_centers, accounts = load_master_data(
        str(MASTER_FILE),
        MASTER_FILE.stat().st_mtime,
    )

except Exception as exc:
    st.error(f"Gagal membaca file master: {exc}")
    st.stop()


if not categories:
    st.warning(
        "Kategori R1/R2/L1/L2 belum terbaca. "
        "Periksa posisi data pada sheet NORMA BIAYA."
    )

if not locations:
    st.warning(
        "Daftar lokasi belum terbaca. "
        "Periksa kolom G-H pada sheet INDEKS RAB."
    )


# =====================================================
# FORM DATA EVENT
# =====================================================

st.header("1. Data Event")

with st.form("rab_form"):

    col1, col2 = st.columns(2)

    with col1:
        event_id = st.text_input(
            "Event ID *",
            placeholder="Masukkan ID event",
        )

        kategori = st.selectbox(
            "Kategori Event *",
            options=categories,
            index=None,
            placeholder="Pilih kategori",
        )

        
akomodasi = st.selectbox(
    "Akomodasi *",
    options=[
        "Ditanggung Penyelenggara",
        "Tidak Ditanggung Penyelenggara",
    ],
    index=None,
    placeholder="Pilih status akomodasi",
)


        location_labels = [
            (
                f"{item['name']} — Indeks: {item['index']}"
                if item["index"] is not None
                else item["name"]
            )
            for item in locations
        ]

        selected_location = st.selectbox(
            "Indeks Lokasi *",
            options=location_labels,
            index=None,
            placeholder="Pilih lokasi",
        )

        tanggal_mulai = st.date_input(
            "Tanggal Pelaksanaan *",
            value=None,
        )

    with col2:
        tanggal_selesai = st.date_input(
            "Tanggal Selesai *",
            value=None,
        )

        lokasi_event = st.text_input(
            "Lokasi Event *",
            placeholder="Kota atau lokasi kegiatan",
        )

        cc_labels = [
            f"{item['code']} — {item['description']}"
            for item in cost_centers
        ]

        selected_cc = st.selectbox(
            "Cost Center *",
            options=cc_labels,
            index=None,
            placeholder="Pilih Cost Center",
        )

        account_labels = [
            f"{item['code']} — {item['description']}"
            for item in accounts
        ]

        selected_account = st.selectbox(
            "Akun Beban Anggaran *",
            options=account_labels,
            index=None,
            placeholder="Pilih akun anggaran",
        )

    st.divider()

    st.header("2. Upload Dokumen FK/JK")

    uploaded_file = st.file_uploader(
        "Pilih dokumen FK/JK",
        type=["pdf", "png", "jpg", "jpeg"],
        help="Format yang didukung: PDF, PNG, JPG, dan JPEG.",
    )

    submitted = st.form_submit_button(
        "Create RAB",
        type="primary",
        use_container_width=True,
    )


# =====================================================
# VALIDASI V1
# =====================================================

if submitted:
    errors = []

    if not event_id.strip():
        errors.append("Event ID wajib diisi.")

    if not kategori:
        errors.append("Kategori Event wajib dipilih.")

    if not selected_location:
        errors.append("Indeks Lokasi wajib dipilih.")

    if not selected_cc:
        errors.append("Cost Center wajib dipilih.")

    if not selected_account:
        errors.append("Akun Beban Anggaran wajib dipilih.")

    if not lokasi_event.strip():
        errors.append("Lokasi Event wajib diisi.")

    if tanggal_mulai is None:
        errors.append("Tanggal Pelaksanaan wajib diisi.")

    if tanggal_selesai is None:
        errors.append("Tanggal Selesai wajib diisi.")

    if (
        tanggal_mulai is not None
        and tanggal_selesai is not None
        and tanggal_selesai < tanggal_mulai
    ):
        errors.append(
            "Tanggal Selesai tidak boleh lebih awal "
            "daripada Tanggal Pelaksanaan."
        )

    if not uploaded_file:
        errors.append("Dokumen FK/JK wajib diunggah.")

    if not akomodasi:
    errors.append("Akomodasi wajib dipilih.")

    if errors:
        st.error("Data belum dapat diproses:")
        for error in errors:
            st.write(f"- {error}")
    else:
        st.success("Input awal lolos validasi.")

        st.info(
            "Form V1 sudah menerima input. "
            "Tahap berikutnya adalah membaca isi FK/JK, "
            "menerapkan aturan master, menghitung RAB, "
            "dan menghasilkan Excel sesuai TEMPLATE RAB."
        )
