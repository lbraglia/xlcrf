from src.xlcrf.CRF import CRF
import tempfile
import streamlit as st

st.markdown("""# Test streamlit


Questo è un [link](https://www.google.it)


""")

# Uploader file
xlsx = st.file_uploader(
    label = "Upload CRF structure",
    type = ["xlsx"],
    accept_multiple_files = False
)

if xlsx is not None:
    struc = tempfile.NamedTemporaryFile(suffix = '.xlsx')
    out = tempfile.NamedTemporaryFile(suffix = '.xlsx')
    strucfile = struc.name
    outfile = out.name
    # salvo per comodità il file in un file
    with open(strucfile, "wb") as f:
        f.write(xlsx.getbuffer())
    ex1 = CRF()
    ex1.read_structure(strucfile)
    ex1.create(outfile)
    with open(outfile, "rb") as f:
        btn = st.download_button(
            label = "Download CRF",
            data = f,
            file_name = "crf.xlsx")
    
