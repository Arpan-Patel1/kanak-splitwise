# streamlit_app.py
import os, json, tempfile, sqlite3
from typing import List

import openpyxl
import streamlit as st
import boto3
from oletools.olevba import VBA_Parser

# ---------- Bedrock client ----------
bedrock = boto3.client("bedrock-runtime", region_name="us-east-1")

EMBED_MODEL_ID = "amazon.titan-embed-text-v2:0"  # <-- correct suffix “:0”

# ---------- helper: Excel upload ----------
def save_file(uploaded):
    path = os.path.join(os.getcwd(), uploaded.name)
    with open(path, "wb") as f:
        f.write(uploaded.getbuffer())
    # strip VBA if .xlsm, otherwise keep as-is
    if path.lower().endswith(".xlsm"):
        wb = openpyxl.load_workbook(path, keep_vba=False)
        new_path = os.path.splitext(path)[0] + ".xlsx"
        wb.save(new_path)
        os.remove(path)
        return new_path
    return path

# ---------- helper: extract VBA ----------
def extract_vba(file_path: str) -> str:
    parser = VBA_Parser(file_path)
    modules = [code for *_ , code in parser.extract_macros() if code.strip()]
    return "\n\n".join(modules)

# ---------- helper: embed with Titan ----------
def titan_embed(text: str) -> List[float]:
    payload = {"inputText": text}
    resp = bedrock.invoke_model(
        modelId     = EMBED_MODEL_ID,
        body        = json.dumps(payload),
        contentType = "application/json",
        accept      = "application/json",
    )
    return json.loads(resp["body"].read())["embedding"]

# ---------- Streamlit UI ----------
st.set_page_config("Macro → Titan Embeddings", layout="wide")
st.title("📐 Titan v2 Embedding Demo for VBA Macros")

uploaded = st.file_uploader("Upload Excel file (.xlsm, .xlsb, .xls)")

if uploaded:
    # Save & extract VBA
    saved_path = save_file(uploaded)
    with st.spinner("Extracting VBA macros…"):
        vba_text = extract_vba(saved_path)

    if not vba_text:
        st.error("No VBA macros found.")
        st.stop()

    st.subheader("Extracted VBA")
    st.code(vba_text[:1500] + (" …" if len(vba_text) > 1500 else ""), language="vb")

    # Embed
    with st.spinner("Calling Titan Text Embeddings v2…"):
        vector = titan_embed(vba_text)

    st.subheader("Titan v2 Embedding (first 50 dims)")
    st.write(vector[:50])
    st.caption(f"Vector length: {len(vector)}")

else:
    st.info("Upload a macro-enabled workbook to begin.")
