import os
import json
import streamlit as st
import boto3
from oletools.olevba import VBA_Parser

# === AWS Bedrock Setup ===
bedrock = boto3.client("bedrock-runtime", region_name="us-east-1")
EMBED_MODEL_ID = "amazon.titan-embed-text-v2:0"

# === VBA Extractor from your original code ===
def extract_vba(file_path: str) -> str:
    parser = VBA_Parser(file_path)
    modules = [code.strip() for _, _, _, code in parser.extract_macros() if code.strip()]
    return "\n\n".join(modules)

# === Titan Embedding Function ===
def titan_embed(text: str):
    payload = {"inputText": text}
    response = bedrock.invoke_model(
        modelId=EMBED_MODEL_ID,
        contentType="application/json",
        accept="application/json",
        body=json.dumps(payload),
    )
    return json.loads(response['body'].read())['embedding']

# === Streamlit UI ===
st.set_page_config(page_title="VBA → Titan Embedding", layout="wide")
st.title("📌 VBA Macro Embedding via Titan v2")

uploaded_file = st.file_uploader("Upload your .xlsm Excel file", type=["xlsm", "xls", "xlsb"])

if uploaded_file:
    # Save the uploaded file to disk
    file_path = os.path.join(os.getcwd(), uploaded_file.name)
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    # Extract VBA
    with st.spinner("Extracting VBA macros..."):
        vba_code = extract_vba(file_path)

    if not vba_code:
        st.error("❌ No VBA macros found in this file.")
        st.stop()

    st.subheader("Step 1: Extracted VBA")
    st.code(vba_code[:2000] + ("..." if len(vba_code) > 2000 else ""), language="vb")

    # Get Titan Embedding
    with st.spinner("Generating Titan Embedding..."):
        embedding = titan_embed(vba_code)

    st.subheader("Step 2: Titan Text Embedding (v2)")
    st.text(f"Length: {len(embedding)}")
    st.write(embedding[:50])
else:
    st.info("Please upload a macro-enabled Excel file (.xlsm) to get started.")
