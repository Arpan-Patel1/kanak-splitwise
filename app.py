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
from langgraph.graph import StateGraph, END

# === Configs ===
bedrock = boto3.client("bedrock-runtime", region_name="us-east-1")
EMBED_MODEL_ID = "amazon.titan-embed-text-v2:0"
DB_PATH = "macro_embeddings.db"

# === Prompt Templates ===
PROMPTS = {
    "pivot_table": "...",  # same content as earlier
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
                matched_score REAL,
                final_prompt TEXT,
                generated_code TEXT,
                feedback TEXT,
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

def find_best_match(new_embedding):
    best_score, best_row = -1, None
    with sqlite3.connect(DB_PATH) as conn:
        for row in conn.execute("SELECT id, name, vba_macro, embedding FROM macro_matches"):
            try:
                old = json.loads(row[3])
                sim = cosine_similarity(new_embedding, old)
                if sim > best_score:
                    best_score = sim
                    best_row = (row[0], row[1], row[2], sim)
            except: continue
    return best_row

def store_result(name, macro, category, embedding, score, prompt, pycode):
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("""
            INSERT INTO macro_matches (name, vba_macro, category, embedding, matched_score, final_prompt, generated_code, feedback)
            VALUES (?, ?, ?, ?, ?, ?, ?, 'pending')
        """, (name, macro, category, json.dumps(embedding), score, prompt, pycode))

def update_feedback(id_, vote):
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("UPDATE macro_matches SET feedback = ? WHERE id = ?", (vote, id_))

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
    final_prompt: Optional[str]
    generated_code: Optional[str]
    embedding: Optional[list]
    match_score: Optional[float]
    match_id: Optional[int]

# === Streamlit UI ===
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
    st.code(state["vba_code"], language="vb")
progress.progress(20)

with st.spinner("Step 2: Embedding and Matching..."):
    state["embedding"] = get_embedding(state["vba_code"])
    match = find_best_match(state["embedding"])
    state["match_score"] = match[3] if match else 0.0
    state["match_id"] = match[0] if match else None
    if match:
        st.markdown(f"**Closest Match:** `{match[1]}` — `{round(match[3]*100, 2)}%`")
        with st.expander("Matched Macro"):
            st.code(match[2], language="vb")
progress.progress(40)

with st.spinner("Step 3: Categorizing Macro..."):
    cat_prompt = (
        "Classify this VBA code as one of: formulas, pivot_table, pivot_chart, user_form, normal_operations. Return only the category.\n\n"
        + state["vba_code"]
    )
    detected = "".join(stream_claude(cat_prompt)).strip().lower()
    state["category"] = detected if detected in PROMPTS else "normal_operations"
    st.markdown(f"**Detected Category:** `{state['category']}`")
progress.progress(60)

state["final_prompt"] = PROMPTS[state["category"]].format(vba_code=state["vba_code"])
with st.expander("AI Prompt"):
    st.code(state["final_prompt"], language="text")

with st.spinner("Step 4: Generating Python Code..."):
    full = "".join(stream_claude(state["final_prompt"]))
    state["generated_code"] = full.strip()
    st.code(state["generated_code"], language="python")
progress.progress(80)

store_result(
    name=uploaded_file.name,
    macro=state["vba_code"],
    category=state["category"],
    embedding=state["embedding"],
    score=state["match_score"],
    prompt=state["final_prompt"],
    pycode=state["generated_code"]
)

st.success("✅ Conversion complete and saved!")
progress.progress(100)

if state["match_id"]:
    col1, col2 = st.columns(2)
    if col1.button("👍 Upvote Match"):
        update_feedback(state["match_id"], "upvote")
        st.success("Thanks for your feedback!")
    if col2.button("👎 Downvote Match"):
        update_feedback(state["match_id"], "downvote")
        st.success("Feedback noted!")
