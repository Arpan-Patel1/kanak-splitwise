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
    "pivot_table": """I have the following VBA code that creates a Pivot Table in Excel:\n{vba_code}\nPlease write equivalent Python code that:\nProduces the same summarized data the pivot table would show (e.g., group by fields, aggregation like SUM, COUNT, AVERAGE).\nUses pandas to perform the summary using pivot_table() or groupby().\nSaves the resulting table into a sheet where it is suppose to be in the same Excel file using pandas.ExcelWriter or openpyxl.\nDoes not create a real Excel PivotTable, and does not use any fake or unsupported APIs like openpyxl.worksheet.table.tables.Table.\nMake sure all Python libraries used are valid and the code runs end-to-end.""",
    "pivot_chart": "I have the following VBA code that creates a Pivot Chart in Excel:\n{vba_code}\nPlease generate equivalent Python code that uses pandas for summarization and plotly/matplotlib for the chart. Ensure it matches the VBA chart type and runs end-to-end.",
    "user_form": "I have the following VBA code that creates a UserForm:\n{vba_code}\nGenerate equivalent Python code using Tkinter or PyQt5, capturing user input and writing to Excel via pandas/openpyxl.",
    "formula": "I have the following VBA or Excel formula logic:\n{vba_code}\nWrite Python code using pandas/numpy to compute the same results, and save back to Excel.",
    "normal_operations": "I have the following VBA code performing normal Excel operations:\n{vba_code}\nWrite equivalent Python using openpyxl or pandas for sheet structure and value operations, including formatting via openpyxl.styles."
}

# === DB Init ===
def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS macro_matches (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            vba_macro TEXT,
            category TEXT,
            embedding TEXT,
            generated_code TEXT,
            feedback INTEGER DEFAULT 0,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """
    )
    conn.close()
init_db()

# === Helpers ===
def get_embedding(text: str):
    if len(text) > 25000:
        text = text[:25000]
    payload = {"inputText": text}
    resp = bedrock.invoke_model(
        modelId=EMBED_MODEL_ID,
        contentType="application/json",
        accept="application/json",
        body=json.dumps(payload),
    )
    return json.loads(resp["body"].read())["embedding"]

def cosine_similarity(v1, v2):
    a, b = np.array(v1), np.array(v2)
    return float(np.dot(a, b) / (np.linalg.norm(a) * np.linalg.norm(b)))

def find_best_match(new_emb, threshold=0.5):
    conn = sqlite3.connect(DB_PATH)
    best = None
    best_score = -1.0
    for row in conn.execute("SELECT id, name, vba_macro, embedding, generated_code FROM macro_matches"):
        try:
            old_emb = json.loads(row[3])
            sim = cosine_similarity(new_emb, old_emb)
            if sim > best_score:
                best_score = sim
                best = {"id": row[0], "name": row[1], "vba_macro": row[2], "generated_code": row[4], "score": sim}
        except:
            pass
    conn.close()
    return best if best and best["score"] >= threshold else None

def update_feedback_by_delta(record_id: int, delta: int):
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        "UPDATE macro_matches SET feedback = feedback + ? WHERE id = ?",
        (delta, record_id)
    )
    conn.close()

def insert_record(name, macro, category, emb, code, initial_feedback: int):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute(
        "INSERT INTO macro_matches (name, vba_macro, category, embedding, generated_code, feedback) VALUES (?, ?, ?, ?, ?, ?)",
        (name, macro, category, json.dumps(emb), code, initial_feedback)
    )
    record_id = cur.lastrowid
    conn.close()
    return record_id

# === Claude Stream ===
def stream_claude(prompt: str):
    payload = {"anthropic_version": "bedrock-2023-05-31", "messages": [{"role":"user","content":prompt}], "max_tokens":4000, "temperature":0}
    resp = bedrock.invoke_model_with_response_stream(
        modelId="arn:aws:bedrock:us-east-1:137360334857:inference-profile/us.anthropic.claude-3-7-sonnet-20250219-v1:0",
        body=json.dumps(payload),
    )
    for event in resp.get("body", []):
        chunk = json.loads(event.get("chunk", {}).get("bytes", b"{}"))
        delta = chunk.get("delta", {})
        text = delta.get("text")
        if text:
            yield text

class VBAState(TypedDict):
    vba_code: str
    category: str
    embedding: list
    match: Optional[dict]
    generated_code: str

# === Streamlit App ===
st.set_page_config(page_title="VBA2PyGen+", layout="wide")
st.title("🧠 VBA2PyGen + Titan Matching")

progress = st.progress(0)
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsm","xls","xlsb"])
if not uploaded_file:
    st.stop()

file_id = uploaded_file.name
if "processed_file_id" not in st.session_state:
    st.session_state["processed_file_id"] = None
if "voted" not in st.session_state:
    st.session_state["voted"] = False

should_process = (st.session_state["processed_file_id"] != file_id) and not st.session_state["voted"]

if should_process:
    with st.spinner("Processing VBA to Python..."):
        # Steps (progress increments omitted for brevity)
        suffix = os.path.splitext(file_id)[1]
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
        tmp.write(uploaded_file.getbuffer()); tmp.flush(); tmp_path = tmp.name; tmp.close()
        modules = [code.strip() for *_, code in VBA_Parser(tmp_path).extract_macros() if code.strip()]
        if not modules: st.error("No VBA macros found."); st.stop()
        vba_code = "\n\n".join(modules)
        with st.expander("Extracted VBA Code"): st.code(vba_code, language="vb")

        emb = get_embedding(vba_code)
        match = find_best_match(emb)
        if match:
            with st.expander("Matched Python Code"): st.code(match['generated_code'], language="python")

        category = ("".join(stream_claude(f"Classify into: formulas,...\n\n{vba_code}")).strip().lower())
        category = category if category in PROMPTS else "normal_operations"
        st.markdown(f"**Detected Category:** `{category}`")

        final_prompt = PROMPTS[category].format(vba_code=vba_code)
        full = "".join(stream_claude(final_prompt))
        gen_code = (full.split("```python")[-1].split("```",1)[0].strip() if "```python" in full else full.strip())
        with st.expander("Generated Python Code"): st.code(gen_code, language="python")

        state = VBAState(vba_code=vba_code, category=category, embedding=emb, match=match, generated_code=gen_code)
        st.session_state["vba_state"] = state
        st.session_state["processed_file_id"] = file_id

state = st.session_state.get("vba_state")
if state and not should_process:
    with st.expander("Generated Python Code"): st.code(state["generated_code"], language="python")

# Voting UI
col1, col2 = st.columns(2)
if not st.session_state["voted"]:
    if col1.button("👍 Helpful"):
        # always update existing record if match, store new record on upvote
        if state.match and state.match.get("id"):
            update_feedback_by_delta(state.match["id"], 1)
            insert_record(file_id, state.vba_code, state.category, state.embedding, state.generated_code, 1)
        else:
            insert_record(file_id, state.vba_code, state.category, state.embedding, state.generated_code, 1)
        st.session_state["voted"] = True
        st.experimental_rerun()
    if col2.button("👎 Not Helpful"):
        # only update existing record feedback if match
        if state.match and state.match.get("id"):
            update_feedback_by_delta(state.match["id"], -1)
        else:
            insert_record(file_id, state.vba_code, state.category, state.embedding, state.generated_code, -1)
        st.session_state["voted"] = True
        st.experimental_rerun()
