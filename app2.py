import streamlit as st
from pathlib import Path
import tempfile
from openpyxl import load_workbook

# Import fungsi dari script asli
from converter_faktur_coretax_v2_2 import (
    read_sheet,
    build_xml,
    SHEET_FAKTUR,
    SHEET_DETAIL,
    FAKTUR_HEADER_ROW,
    DETAIL_HEADER_ROW
)

# ============================
# STREAMLIT UI
# ============================

st.title("🚀 Converter Faktur CoreTax Online")
st.write("Upload file Excel template CoreTax, dan aplikasi akan mengonversinya menjadi XML.")

uploaded_file = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    # Simpan file upload ke file sementara
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(uploaded_file.read())
        temp_path = Path(tmp.name)

    st.success("File berhasil di-upload. Sedang diproses...")

    try:
        # Baca workbook
        wb = load_workbook(temp_path, data_only=True)

        # Baca sheet Faktur & Detail
        faktur = read_sheet(wb[SHEET_FAKTUR], FAKTUR_HEADER_ROW)
        detail = read_sheet(wb[SHEET_DETAIL], DETAIL_HEADER_ROW)

        # Generate XML
        xml_tree = build_xml(faktur, detail)

        # Save to temp XML
        xml_output_path = temp_path.with_suffix(".xml")
        xml_tree.write(xml_output_path, encoding="utf-8", xml_declaration=True)

        # Read kembali file XML untuk download
        with open(xml_output_path, "rb") as f:
            xml_bytes = f.read()

        st.success("XML berhasil dibuat!")

        st.download_button(
            label="⬇ Download XML",
            data=xml_bytes,
            file_name="output.xml",
            mime="application/xml"
        )

    except Exception as e:
        st.error(f"Terjadi error saat memproses file: {e}")
