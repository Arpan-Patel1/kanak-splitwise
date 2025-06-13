import os
import json
import tempfile
import sqlite3
import numpy as np
from typing import TypedDict, Optional

import openpyxl
import pandas as pd
import streamlit as st
import boto3
from oletools.olevba import VBA_Parser
from langgraph.graph import StateGraph, END

# ==================== DB INIT ====================
DB_PATH = "references.db"

def init_db():
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS references (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                vba TEXT,
                embedding TEXT,
                python_code TEXT,
                score INTEGER DEFAULT 0
            )
        """)

init_db()

# ==================== AWS Titan & Claude Config ====================
bedrock = boto3.client("bedrock-runtime")

TITAN_MODEL_ID = "amazon.titan-embed-text-v2:0"
CLAUDE_MODEL_ID = "arn:aws:bedrock:us-east-1:137360334857:inference-profile/us.anthropic.claude-3-7-sonnet-20250219-v1:0"

# ==================== PROMPTS ====================
PROMPTS = {
    "pivot_table": """...""",
    "pivot_chart": """...""",
    "user_form": """...""",
    "formula": """...""",
    "normal_operations": """..."""
}

# ==================== Claude Wrapper ====================
def stream_claude(prompt: str):
    try:
        payload = {
            "anthropic_version": "bedrock-2023-05-31",
            "messages": [{"role": "user", "content": prompt}],
            "max_tokens": 4000,
            "temperature": 0,
            "top_p": 1.0,
            "top_k": 1,
        }
        resp = bedrock.invoke_model_with_response_stream(
            modelId=CLAUDE_MODEL_ID,
            body=json.dumps(payload),
        )
        for event in resp["body"]:
            chunk = json.loads(event["chunk"]["bytes"])
            if delta := chunk.get("delta"):
                if text := delta.get("text"):
                    yield text
    except Exception as e:
        st.error(f"Claude error: {e}")
        st.stop()

# ==================== Embedding ====================
def get_embedding(text: str) -> list[float]:
    response = bedrock.invoke_model(
        modelId=TITAN_MODEL_ID,
        body=json.dumps({"inputText": text}),
        contentType="application/json",
        accept="application/json",
    )
    result = json.loads(response["body"].read())
    return result["embedding"]

# ==================== State ====================
class VBAState(TypedDict):
    file_path: Optional[str]
    vba_code: Optional[str]
    category: Optional[str]
    final_prompt: Optional[str]
    generated_code: Optional[str]
    embedding: Optional[list[float]]

# ==================== Match Logic ====================
def cosine_similarity(vec1, vec2):
    a = np.array(vec1)
    b = np.array(vec2)
    return float(np.dot(a, b) / (np.linalg.norm(a) * np.linalg.norm(b)))

def find_best_reference(new_emb: list[float]):
    with sqlite3.connect(DB_PATH) as conn:
        cur = conn.execute("SELECT id, embedding, python_code, score FROM references")
        best = (None, -1.0)
        for row in cur.fetchall():
            db_id, emb_str, py, score = row
            db_emb = json.loads(emb_str)
            sim = cosine_similarity(new_emb, db_emb)
            if sim > best[1]:
                best = (row, sim)
        return best

# ==================== Core Steps ====================
def save_uploaded_file(uploaded_file) -> tuple[str, str]:
    path = os.path.join(os.getcwd(), uploaded_file.name)
    with open(path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    if path.lower().endswith(".xlsm"):
        wb = openpyxl.load_workbook(path, keep_vba=False)
        no_macro = os.path.splitext(path)[0] + ".xlsx"
        wb.save(no_macro)
        os.remove(path)
        return no_macro, os.path.splitext(uploaded_file.name)[0]
    return path, os.path.splitext(uploaded_file.name)[0]

def extract_vba(state: VBAState) -> VBAState:
    with st.spinner("Extracting VBA..."):
        parser = VBA_Parser(state["file_path"])
        modules = [code.strip() for _, _, _, code in parser.extract_macros() if code.strip()]
        if not modules:
            st.error("No VBA macros found.")
            st.stop()
        state["vba_code"] = "\n\n".join(modules)
    st.code(state["vba_code"], language="text")
    return state

def categorize_vba(state: VBAState) -> VBAState:
    prompt = "Classify the following VBA code into: formulas, pivot_table, pivot_chart, user_form, normal_operations. Return only the category.\n\n" + state["vba_code"]
    cat = "".join(stream_claude(prompt)).strip().lower()
    state["category"] = cat if cat in PROMPTS else "normal_operations"
    st.markdown(f"**Category:** `{state['category']}`")
    return state

def embed_vba(state: VBAState) -> VBAState:
    state["embedding"] = get_embedding(state["vba_code"])
    return state

def build_prompt(state: VBAState) -> VBAState:
    state["final_prompt"] = PROMPTS[state["category"]].format(vba_code=state["vba_code"])
    st.code(state["final_prompt"], language="text")
    return state

def generate_python_code(state: VBAState) -> VBAState:
    full = "".join(stream_claude(state["final_prompt"]))
    state["generated_code"] = full
    st.code(full, language="python")
    return state

def save_reference(state: VBAState):
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            "INSERT INTO references (vba, embedding, python_code, score) VALUES (?, ?, ?, 0)",
            (state["vba_code"], json.dumps(state["embedding"]), state["generated_code"])
        )

def build_graph():
    steps = [extract_vba, categorize_vba, embed_vba, build_prompt, generate_python_code]
    g = StateGraph(VBAState)
    for fn in steps:
        g.add_node(fn.__name__, fn)
    g.set_entry_point(steps[0].__name__)
    for a, b in zip(steps, steps[1:]):
        g.add_edge(a.__name__, b.__name__)
    g.add_edge(steps[-1].__name__, END)
    return g.compile()

# ==================== Streamlit UI ====================
st.set_page_config(page_title="VBA2PyGen", layout="wide")
st.title("VBA2PyGen")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsm", "xls"])
if uploaded_file:
    xlsx_path, base_name = save_uploaded_file(uploaded_file)
    suffix = os.path.splitext(uploaded_file.name)[1]
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded_file.getbuffer())
        tmp_path = tmp.name
    graph = build_graph()
    for state in graph.stream({"file_path": tmp_path}):
        final = state

    st.success("✅ Conversion completed!")
    save_reference(final)

    # Matching Section
    match_row, match_score = find_best_reference(final["embedding"])
    if match_row:
        db_id, _, _, ref_code, score = match_row
        st.markdown(f"### 🔍 Closest Matching Code (Score: `{match_score:.2%}`)")
        st.code(ref_code, language="python")
        col1, col2 = st.columns(2)
        if col1.button("✅ Works", key="upvote"):
            with sqlite3.connect(DB_PATH) as conn:
                conn.execute("UPDATE references SET score = score + 1 WHERE id = ?", (db_id,))
        if col2.button("❌ Failed", key="downvote"):
            with sqlite3.connect(DB_PATH) as conn:
                conn.execute("UPDATE references SET score = score - 1 WHERE id = ?", (db_id,))

else:
    st.info("Upload an Excel file to begin.")
