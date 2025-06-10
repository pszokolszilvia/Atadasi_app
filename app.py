import streamlit as st
import os
import shutil
from tempfile import NamedTemporaryFile
from boritok import generate_boritok
from elvalasztok import generate_elvalasztok
from utils import zip_directory

st.set_page_config(page_title="Borító és oldalelválasztó generátor", layout="centered")
st.title("📄 Word generátor – borítók és oldalelválasztók")

excel_file = st.file_uploader("📋 Tartalomjegyzék (.xlsx)", type="xlsx")
borito_template = st.file_uploader("📄 Borító sablon (.docx)", type="docx")
elvalaszto_template = st.file_uploader("📄 Oldalelválasztó sablon (.docx)", type="docx")

if st.button("📎 Generálás"):
    if not (excel_file and borito_template and elvalaszto_template):
        st.warning("📌 Kérlek, töltsd fel az összes fájlt.")
    else:
        with NamedTemporaryFile(delete=False, suffix=".xlsx") as tf_excel:
            tf_excel.write(excel_file.read())
            excel_path = tf_excel.name
        with NamedTemporaryFile(delete=False, suffix=".docx") as tf_borito:
            tf_borito.write(borito_template.read())
            borito_path = tf_borito.name
        with NamedTemporaryFile(delete=False, suffix=".docx") as tf_elv:
            tf_elv.write(elvalaszto_template.read())
            elvalaszto_path = tf_elv.name

        borito_dir = os.path.join("output", "boritok")
        elv_dir = os.path.join("output", "elvalasztok")
        os.makedirs(borito_dir, exist_ok=True)
        os.makedirs(elv_dir, exist_ok=True)

        generate_boritok(excel_path, borito_path, borito_dir)
        generate_elvalasztok(borito_dir, elvalaszto_path, elv_dir)

        zip_path = "output/boritok_elvalasztok.zip"
        zip_directory("output", zip_path)

        with open(zip_path, "rb") as f:
            st.success("✅ Generálás kész!")
            st.download_button("📥 Letöltés ZIP fájlként", f, file_name="boritok_elvalasztok.zip")
