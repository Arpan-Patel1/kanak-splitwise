import os
import json
import sqlite3
import tempfile
from typing import TypedDict, Optional

import numpy as np
import openpyxl
import pandas as pd
import streamlit as st
import boto3
from oletools.olevba import VBA_Parser

# === Configs ===
bedrock = boto3.client("bedrock-runtime", region_name="us-east-1")
EMBED_MODEL_ID = "amazon.titan-embed-text-v2:0"
DB_PATH = "macro_embeddings.db"

PROMPTS = {
    "pivot_table": """I have the following VBA code that creates a Pivot Table in Excel:\n{vba_code}
Please write equivalent Python code that:
Produces the same summarized data the pivot table would show (e.g., group by fields, aggregation like SUM, COUNT, AVERAGE).
Uses pandas to perform the summary using pivot_table() or groupby().
Saves the resulting table into a sheet where it is suppose to be in the same Excel file using pandas.ExcelWriter or openpyxl.
Does not create a real Excel PivotTable, and does not use any fake or unsupported APIs like openpyxl.worksheet.table.tables.Table.
Make sure all Python libraries used are valid and the code runs end-to-end.""",
    "pivot_chart": "...",
    "user_form": "...",
    "formula": "...",
    "normal_operations": "..."
}

# === DB Init ===
def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS macro_matches (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT,
                vba_macro TEXT,
                category TEXT,
                embedding TEXT,
                generated_code TEXT,
                feedback INTEGER,
                timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)
init_db()

# === Embedding & Matching ===
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

def cosine_similarity(v1, v2):
    v1, v2 = np.array(v1), np.array(v2)
    return float(np.dot(v1, v2) / (np.linalg.norm(v1) * np.linalg.norm(v2)))

def find_best_match(new_embedding, threshold=0.5):
    best_score = -1
    best_row = None
    with sqlite3.connect(DB_PATH) as conn:
        for row in conn.execute("SELECT id, name, vba_macro, embedding, generated_code FROM macro_matches"):
            try:
                old = json.loads(row[3])
                sim = cosine_similarity(new_embedding, old)
                if sim > best_score:
                    best_score = sim
                    best_row = {
                        "id": row[0],
                        "name": row[1],
                        "vba_macro": row[2],
                        "embedding": old,
                        "generated_code": row[4],
                        "score": sim
                    }
            except: continue
    return best_row if best_row and best_row["score"] >= threshold else None

def store_result(name, macro, category, embedding, pycode, feedback):
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("""
            INSERT INTO macro_matches (name, vba_macro, category, embedding, generated_code, feedback)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (name, macro, category, json.dumps(embedding), pycode, feedback))

# === Claude Stream ===
def stream_claude(prompt: str):
    payload = {
        "anthropic_version": "bedrock-2023-05-31",
        "messages": [{"role": "user", "content": prompt}],
        "max_tokens": 4000,
        "temperature": 0,
        "top_p": 1.0,
        "top_k": 1,
    }
    resp = bedrock.invoke_model_with_response_stream(
        modelId="arn:aws:bedrock:us-east-1:137360334857:inference-profile/us.anthropic.claude-3-7-sonnet-20250219-v1:0",
        body=json.dumps(payload),
    )
    for event in resp["body"]:
        chunk = json.loads(event["chunk"]["bytes"])
        if delta := chunk.get("delta"):
            if text := delta.get("text"):
                yield text

class VBAState(TypedDict):
    file_path: Optional[str]
    vba_code: Optional[str]
    category: Optional[str]
    generated_code: Optional[str]
    embedding: Optional[list]
    match: Optional[dict]

# === Streamlit App ===
st.set_page_config(page_title="VBA2PyGen+", layout="wide")
st.title("🧠 VBA2PyGen + Titan Matching")
progress = st.progress(0)

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsm", "xls", "xlsb"])
if not uploaded_file:
    st.stop()

suffix = os.path.splitext(uploaded_file.name)[1]
with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
    tmp.write(uploaded_file.getbuffer())
    tmp_path = tmp.name

state: VBAState = {"file_path": tmp_path}

with st.spinner("Step 1: Extracting VBA..."):
    parser = VBA_Parser(tmp_path)
    modules = [code.strip() for _, _, _, code in parser.extract_macros() if code.strip()]
    if not modules:
        st.error("No VBA macros found.")
        st.stop()
    state["vba_code"] = "\n\n".join(modules)
    with st.expander("Step 1: Extracted VBA code"):
        st.code(state["vba_code"], language="vb")
progress.progress(20)

with st.spinner("Step 2: Embedding & Matching"):
    state["embedding"] = get_embedding(state["vba_code"])
    match = find_best_match(state["embedding"])
    state["match"] = match
    if match:
        st.markdown(f"**Reference Found:** `{match['name']}` — `{round(match['score']*100, 2)}%`)" )
        with st.expander("Matched VBA Macro"):
            st.code(match['vba_macro'], language="vb")
        with st.expander("Matched Python Code"):
            st.code(match['generated_code'], language="python")
progress.progress(50)

with st.spinner("Step 3: Categorizing"):
    cat_prompt = (
        "Classify this VBA code as one of: formulas, pivot_table, pivot_chart, user_form, normal_operations. Return only the category.\n\n"
        + state["vba_code"]
    )
    detected = "".join(stream_claude(cat_prompt)).strip().lower()
    state["category"] = detected if detected in PROMPTS else "normal_operations"
    st.markdown(f"**Detected Category:** `{state['category']}`")
progress.progress(70)

final_prompt = PROMPTS[state["category"]].format(vba_code=state["vba_code"])
with st.expander("Prompt Used"):
    st.code(final_prompt, language="text")

with st.spinner("Step 4: Generating Python Code"):
    full = "".join(stream_claude(final_prompt))
    if "```python" in full:
        s = full.find("```python") + len("```python")
        e = full.find("```", s)
        code = full[s:e].strip()
    else:
        code = full.strip()
    state["generated_code"] = code
with st.expander("Step 4: Generated Python Code"):
    st.code(code, language="python")
progress.progress(100)

if match:
    col1, col2 = st.columns(2)
    if col1.button("👍 Upvote: This match helped"):
        store_result(
            name=uploaded_file.name,
            macro=state["vba_code"],
            category=state["category"],
            embedding=state["embedding"],
            pycode=state["generated_code"],
            feedback=1
        )
        st.success("Saved with +1 feedback")
    if col2.button("👎 Downvote: This match didn’t help"):
        store_result(
            name=uploaded_file.name,
            macro=state["vba_code"],
            category=state["category"],
            embedding=state["embedding"],
            pycode=state["generated_code"],
            feedback=-1
        )
        st.success("Saved with -1 feedback")
