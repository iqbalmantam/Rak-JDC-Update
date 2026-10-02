import json
import os
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

from baca_excel import baca_rak

st.set_page_config(page_title="Peta Rak JDC - Tampak Samping", layout="wide")
TEMPLATE = (Path(__file__).parent / "tampak_samping.html").read_text(encoding="utf-8")

st.sidebar.header("Data")
up = st.sidebar.file_uploader("Unggah Excel layout rak (mis. Rak_JDC.xlsx)", type=["xlsx"])
contoh = os.environ.get("RAK_XLSX")  # opsional: path file default saat uji lokal

if up is None and contoh and os.path.exists(contoh):
    sumber = contoh
elif up is not None:
    sumber = up
else:
    st.title("Peta Rak JDC · Tampak Samping")
    st.info("Unggah file Excel layout rak di panel kiri untuk menampilkan peta.")
    st.stop()


@st.cache_data(show_spinner="Membaca Excel...")
def muat(data: bytes | str):
    import io
    f = io.BytesIO(data) if isinstance(data, bytes) else data
    return baca_rak(f)


try:
    data = muat(sumber.getvalue() if hasattr(sumber, "getvalue") else sumber)
except Exception as e:
    st.error(f"Gagal membaca file: {e}")
    st.stop()

if not data["racks"]:
    st.warning("Kode rak (A01, B01, C01, D01, ...) tidak ditemukan. Pastikan layout sama dengan Rak_JDC.xlsx.")
    st.stop()

html = TEMPLATE.replace("__DATA__", json.dumps(data, separators=(",", ":")))
components.html(html, height=1500, scrolling=True)
