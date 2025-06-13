import os
import json
import streamlit as st
import boto3
from oletools.olevba import VBA_Parser

# === Titan Embedding Model Config ===
bedrock = boto3.client("bedrock-runtime", region_name="us-east-1")
EMBED_MODEL_ID = "amazon.titan-embed-text-v2:0"

# === Progress Placeholder ===
progress = st.empty()

# === Your Extract VBA Logic (wrapped for standalone use) ===
def extract_vba(file_path: str) -> str:
    with st.spinner("Extracting VBA..."):
        parser = VBA_Parser(file_path)
        modules = [code.strip() for _, _, _, code in parser.extract_macros() if code.strip()]
        if not modules:
            st.error("No VBA macros found.")
            st.stop()
        vba_code = "\n\n".join(modules)
    with st.expander("Step 1: Extracted VBA code"):
        st.code(vba_code, language="vb")
    progress.progress(20)
    return vba_code

# === Titan Embedding Call ===
def titan_embed(text: str):
    if len(text) > 25000:
        st.warning("Truncating VBA code to 25,000 characters for embedding.")
        text = text[:25000]
    payload = {"inputText": text}
    response = bedrock.invoke_model(
        modelId=EMBED_MODEL_ID,
        contentType="application/json",
        accept="application/json",
        body=json.dumps(payload),
    )
    return json.loads(response['body'].read())["embedding"]

# === Streamlit UI ===
st.set_page_config(page_title="VBA Macro → Titan Embedding", layout="wide")
st.title("📌 Extract VBA Macro and View Titan Embedding")

uploaded_file = st.file_uploader("Upload a macro-enabled Excel file (.xlsm)", type=["xlsm", "xls", "xlsb"])

if uploaded_file:
    # Save uploaded file to disk
    file_path = os.path.join(os.getcwd(), uploaded_file.name)
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    # Extract full VBA using your function
    vba_code = extract_vba(file_path)

    # Titan Embedding
    with st.spinner("Generating Titan Embedding..."):
        embedding = titan_embed(vba_code)

    st.subheader("Step 2: Titan Text Embedding")
    st.text(f"Vector Length: {len(embedding)}")
    st.write(embedding[:50])  # Show first 50 values
    progress.progress(100)
else:
    st.info("Upload an Excel file with VBA macros to begin.")
