import hashlib
import json
from pathlib import Path

import pandas as pd
import streamlit as st

from classifier import ROOT, PROMPT_VERSION, GroqClassifier, categorize, public_output, read_input
from exporter import excel_bytes

st.set_page_config(page_title="Kategori laporan", page_icon=":material/category:", layout="wide")
st.title("Kategori laporan")
st.caption("Unggah Excel. Hasilnya berisi seluruh kolom asli ditambah kategori_pelaporan. Laporan dengan beberapa kategori menjadi beberapa baris.")

@st.cache_data(ttl=900, max_entries=8, show_spinner=False)
def load_source(data):
    return read_input(data)

upload = st.file_uploader("File laporan Excel", type=["xlsx"])
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

# Reuse completed local processing for the same source file; never mix uploads.
saved = ROOT / "hasil/laporan_wbs_kategori.json"
meta_path = ROOT / "hasil/laporan_wbs_kategori_ringkasan.json"
if key not in st.session_state.hasil and saved.exists() and meta_path.exists():
    meta = json.loads(meta_path.read_text(encoding="utf-8"))
    if meta.get("source_sha256") == fingerprint and meta.get("prompt_version") == PROMPT_VERSION:
        raw = pd.DataFrame(json.loads(saved.read_text(encoding="utf-8")))
        if raw['status_kategorisasi'].eq('Berhasil').all():
            st.session_state.hasil[key] = public_output(raw, list(frame.columns))

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
            st.session_state.excel.pop(key, None)
    except Exception as exc:
        st.error(str(exc))

if key in st.session_state.hasil:
    result = st.session_state.hasil[key]
    st.success(f"Selesai: {len(frame)} laporan menjadi {len(result)} baris.")
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