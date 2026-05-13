# 13 Mei 2026 13.44 Updated

import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH # Tambahan untuk rata tengah
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import requests
import copy

st.set_page_config(page_title="GESIT - BPS1372", layout="centered")

st.title("📝 GESIT (Generate SK Instan)")
st.write("Aplikasi ini akan mengotomasi pengisian dokumen SK Kuasa Pengguna Anggaran BPS Kota Solok.")
st.write("Website beralih kesini: https://gesit-bps1372.streamlit.app/ (sementara)")
st.write("Website dalam Proses Pengembangan lebih lanjut dan maintenance")
