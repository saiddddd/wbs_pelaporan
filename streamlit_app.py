import hashlib
import json
from pathlib import Path

import pandas as pd
import streamlit as st

from classifier import ROOT, PROMPT_VERSION, GroqClassifier, categorize, public_output, read_input
from exporter import excel_bytes

st.set_page_config(page_title="Analisis laporan WBS | Agrinas", page_icon=":material/analytics:", layout="wide")
logo = ROOT / "assets/logo_agrinas.png"
if logo.exists():
    st.image(str(logo), width=260, alt="Logo Agrinas Pangan Nusantara")
st.title("Analisis laporan WBS")
st.caption("Kategorisasi laporan untuk membantu pemetaan masalah dan tindak lanjut.")
st.divider()

@st.cache_data(ttl=900, max_entries=8, show_spinner=False)
def load_source(data):
    return read_input(data)

st.subheader("File laporan")
upload = st.file_uploader("Unggah file Excel", type=["xlsx"],
                          help="Gunakan file yang memiliki kolom Detail Laporan.")
st.caption("Hasil unduhan mempertahankan kolom asli dan menambahkan kategori_pelaporan. Setiap kategori dari laporan yang sama ditampilkan pada baris terpisah.")
default = ROOT / "laporan-wbs-20261026-084953.xlsx"
if upload is not None:
    data, filename = upload.getvalue(), upload.name
elif default.exists():
    data, filename = default.read_bytes(), default.name
    st.caption(f"File tersedia: {filename}")
else:
    st.info("Unggah file yang memiliki kolom Detail Laporan.")
    st.stop()

try:
    frame, sheet = load_source(data)
except Exception as exc:
    st.error(str(exc))
    st.stop()
if frame.empty:
    st.info("File tidak berisi laporan.")
    st.stop()

fingerprint = hashlib.sha256(data).hexdigest()
key = fingerprint + PROMPT_VERSION
st.session_state.setdefault("hasil", {})
st.session_state.setdefault("excel", {})
st.session_state.setdefault("ringkasan", {})

# Reuse completed local processing for the same source file; never mix uploads.
saved = ROOT / "hasil/laporan_wbs_kategori.json"
meta_path = ROOT / "hasil/laporan_wbs_kategori_ringkasan.json"
if key not in st.session_state.hasil and saved.exists() and meta_path.exists():
    meta = json.loads(meta_path.read_text(encoding="utf-8"))
    if meta.get("source_sha256") == fingerprint and meta.get("prompt_version") == PROMPT_VERSION:
        raw = pd.DataFrame(json.loads(saved.read_text(encoding="utf-8")))
        if raw['status_kategorisasi'].eq('Berhasil').all():
            st.session_state.hasil[key] = public_output(raw, list(frame.columns))
            st.session_state.ringkasan[key] = int((raw.drop_duplicates('baris_sumber')['jumlah_kategori'] > 1).sum())

if st.button("Buat kategori", type="primary", icon=":material/category:"):
    progress = st.progress(0, text="Memproses laporan…")
    notice = st.empty()
    try:
        raw, errors = categorize(
            frame, GroqClassifier(),
            on_progress=lambda done,total: progress.progress(min(done/max(total,1),1),text=f"{done}/{total} laporan selesai"),
            on_wait=lambda seconds: notice.info(f"Menunggu kuota Groq, sekitar {seconds:.0f} detik. Proses akan dilanjutkan otomatis."),
        )
        notice.empty()
        if raw['status_kategorisasi'].isin(['Gagal','Belum diproses']).any():
            st.error("Sebagian laporan belum selesai. Klik Buat kategori untuk melanjutkan tanpa mengulang hasil yang sudah tersimpan.")
            if errors:
                st.caption(errors[0]['error'])
        else:
            st.session_state.hasil[key] = public_output(raw, list(frame.columns))
            st.session_state.ringkasan[key] = int((raw.drop_duplicates('baris_sumber')['jumlah_kategori'] > 1).sum())
            st.session_state.excel.pop(key, None)
    except Exception as exc:
        st.error(str(exc))

if key in st.session_state.hasil:
    result = st.session_state.hasil[key]
    st.success(f"Selesai: {len(frame)} laporan menjadi {len(result)} baris.")
    categories = result['kategori_pelaporan'].fillna('').astype(str).str.strip()
    counts = categories[categories.ne('')].value_counts().rename_axis('Kategori').reset_index(name='Jumlah laporan')
    st.subheader("Ringkasan laporan")
    with st.container(horizontal=True):
        st.metric("Laporan asal", len(frame), border=True)
        st.metric("Baris hasil", len(result), border=True)
        st.metric("Jenis kategori", len(counts), border=True)
        if key in st.session_state.ringkasan:
            st.metric("Laporan multikategori", st.session_state.ringkasan[key], border=True)
    if not counts.empty:
        st.bar_chart(counts.head(10), x='Kategori', y='Jumlah laporan', horizontal=True,
                     sort='-Jumlah laporan', alt="Sepuluh kategori dengan jumlah laporan terbanyak")
        st.dataframe(counts, hide_index=True, alt="Jumlah laporan untuk setiap kategori")
        st.caption("Satu laporan dapat masuk beberapa kategori. Jumlah kategori tidak sama dengan jumlah laporan asal.")
    st.subheader("Hasil kategorisasi")
    st.dataframe(result, hide_index=True, height=500, alt="Kolom laporan asli ditambah kategori_pelaporan")
    if key not in st.session_state.excel:
        try:
            with st.spinner("Menyiapkan Excel…"):
                st.session_state.excel[key] = excel_bytes(result)
        except RuntimeError as exc:
            st.error(str(exc))
    if key in st.session_state.excel:
        st.download_button("Unduh Excel", st.session_state.excel[key],
            file_name=Path(filename).stem+"-kategori.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:")
else:
    st.dataframe(frame.drop(columns="_baris_excel"), hide_index=True, height=420, alt="Isi file laporan yang diunggah")
