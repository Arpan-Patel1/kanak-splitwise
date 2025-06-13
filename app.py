import os
import json
import sqlite3
import numpy as np
import streamlit as st
import boto3
from oletools.olevba import VBA_Parser

# === Configs ===
bedrock = boto3.client("bedrock-runtime", region_name="us-east-1")
EMBED_MODEL_ID = "amazon.titan-embed-text-v2:0"
DB_PATH = "macro_embeddings.db"

# === Init DB ===
def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS macros (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT,
                macro TEXT,
                embedding TEXT
            )
        """)
init_db()

# === Extract VBA macro ===
def extract_vba(file_path: str) -> str:
    parser = VBA_Parser(file_path)
    modules = [code.strip() for _, _, _, code in parser.extract_macros() if code.strip()]
    if not modules:
        return ""
    return "\n\n".join(modules)

# === Generate embedding using Titan ===
def get_embedding(text: str):
    if len(text) > 25000:
        text = text[:25000]
    payload = {"inputText": text}
    response = bedrock.invoke_model(
        modelId=EMBED_MODEL_ID,
        contentType="application/json",
        accept="application/json",
        body=json.dumps(payload),
    )
    return json.loads(response["body"].read())["embedding"]

# === Cosine Similarity ===
def cosine_similarity(vec1, vec2):
    v1 = np.array(vec1)
    v2 = np.array(vec2)
    return float(np.dot(v1, v2) / (np.linalg.norm(v1) * np.linalg.norm(v2)))

# === Search Best Match ===
def find_best_match(new_embedding):
    best_score = -1
    best_row = None
    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.execute("SELECT name, macro, embedding FROM macros")
        for name, macro, emb_str in cur.fetchall():
            try:
                old_embedding = json.loads(emb_str)
                sim = cosine_similarity(new_embedding, old_embedding)
                if sim > best_score:
                    best_score = sim
                    best_row = (name, macro, sim)
            except: continue
    return best_row

# === Save Macro + Embedding ===
def store_macro(name, macro, embedding):
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("INSERT INTO macros (name, macro, embedding) VALUES (?, ?, ?)",
                     (name, macro, json.dumps(embedding)))

# === Streamlit App ===
st.set_page_config(page_title="Macro Matcher", layout="wide")
st.title("🧠 Macro Matching Using Titan Embeddings")

uploaded_file = st.file_uploader("Upload Excel file with macro", type=["xlsm", "xls", "xlsb"])

if uploaded_file:
    # Save locally
    file_path = os.path.join(os.getcwd(), uploaded_file.name)
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    # Extract VBA
    macro_code = extract_vba(file_path)
    if not macro_code:
        st.error("No VBA macros found.")
        st.stop()

    st.subheader("🔍 Extracted VBA Code")
    st.code(macro_code, language="vb")

    # Generate embedding
    with st.spinner("Generating Titan Embedding..."):
        embedding = get_embedding(macro_code)

    # Match against DB
    match = find_best_match(embedding)

    st.subheader("📊 Match Results")
    if match:
        name, matched_macro, score = match
        st.markdown(f"**Closest Match:** `{name}`")
        st.markdown(f"**Matching Confidence:** `{round(score * 100, 2)}%`")
        with st.expander("🧾 Matched Macro Code"):
            st.code(matched_macro, language="vb")
    else:
        st.warning("No existing embeddings to compare against yet.")

    if st.button("✅ Save This Macro to DB"):
        store_macro(uploaded_file.name, macro_code, embedding)
        st.success("Macro and its embedding saved!")

