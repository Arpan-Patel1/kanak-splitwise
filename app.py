import os
import json
import tempfile
import sqlite3
from typing import TypedDict, Optional

import openpyxl
import pandas as pd
import streamlit as st
import boto3
from oletools.olevba import VBA_Parser
from langgraph.graph import StateGraph, END

# === EMBEDDING + REFERENCE STORAGE ===
db_path = "reference_store.db"

def init_db():
    with sqlite3.connect(db_path) as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS references (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                macro TEXT NOT NULL,
                embedding BLOB NOT NULL,
                python_code TEXT NOT NULL,
                score INTEGER DEFAULT 0
            )
        """)
init_db()

bedrock = boto3.client("bedrock-runtime")

def get_embedding(text: str) -> list[float]:
    payload = {
        "inputText": text,
        "embeddingConfig": {"outputEmbeddingLength": 1536}
    }
    response = bedrock.invoke_model(
        modelId="amazon.titan-embed-text-v2",
        contentType="application/json",
        accept="application/json",
        body=json.dumps(payload),
    )
    embedding = json.loads(response['body'].read())['embedding']
    return embedding

def cosine_similarity(a, b):
    import numpy as np
    a, b = np.array(a), np.array(b)
    return float(np.dot(a, b) / (np.linalg.norm(a) * np.linalg.norm(b)))

def find_best_reference(vba_code):
    query_vec = get_embedding(vba_code)
    with sqlite3.connect(db_path) as conn:
        cur = conn.execute("SELECT id, macro, embedding, python_code, score FROM references")
        best_id, best_code, best_score = None, None, -1
        for rid, macro, emb_blob, py_code, score in cur.fetchall():
            try:
                ref_vec = json.loads(emb_blob)
                sim = cosine_similarity(query_vec, ref_vec)
                if sim > 0.90 and score >= best_score:
                    best_id, best_code, best_score = rid, py_code, score
            except: pass
    return best_code

def save_reference(macro, python_code):
    embedding = get_embedding(macro)
    with sqlite3.connect(db_path) as conn:
        conn.execute("INSERT INTO references (macro, embedding, python_code, score) VALUES (?, ?, ?, 1)",
                     (macro, json.dumps(embedding), python_code))

def update_reference_feedback(macro, is_good: bool):
    with sqlite3.connect(db_path) as conn:
        cur = conn.execute("SELECT id FROM references WHERE macro=? ORDER BY id DESC LIMIT 1", (macro,))
        row = cur.fetchone()
        if row:
            rid = row[0]
            conn.execute("UPDATE references SET score = score + ? WHERE id=?", (1 if is_good else -1, rid))

# === PROMPTS ===
PROMPTS = {
    "pivot_table": """I have the following VBA code that creates a Pivot Table in Excel:\n{vba_code} ...""",
    "pivot_chart": """I have the following VBA code that creates a Pivot Chart in Excel:\n{vba_code} ...""",
    "user_form": """I have the following VBA code that creates and handles a UserForm in Excel: \n{vba_code} ...""",
    "formula": """I have the following VBA or Excel formula-based code:\n{vba_code} ...""",
    "normal_operations": """I have the following VBA code that performs normal Excel operations (like inserting rows, copying values, deleting columns, formatting cells, renaming sheets, etc.):\n{vba_code} ...""",
}

class VBAState(TypedDict):
    file_path: Optional[str]
    vba_code: Optional[str]
    category: Optional[str]
    final_prompt: Optional[str]
    generated_code: Optional[str]

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
            modelId=("arn:aws:bedrock:us-east-1:137360334857:"
                     "inference-profile/us.anthropic.claude-3-7-sonnet-20250219-v1:0"),
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
    with st.expander("Step 1: Extracted VBA code"):
        st.code(state["vba_code"], language="text")
    progress.progress(20)
    return state

def categorize_vba(state: VBAState) -> VBAState:
    with st.spinner("Categorizing code..."):
        prompt = "Classify this VBA code into: formulas, pivot_table, pivot_chart, user_form, normal_operations. Return only the category.\n\n" + state["vba_code"]
        cat = "".join(stream_claude(prompt)).strip().lower()
        state["category"] = cat if cat in PROMPTS else "normal_operations"
    with st.expander("Step 2: Detected Category"):
        st.markdown(f"**Category detected:** `{state['category']}`")
    progress.progress(40)
    return state

def build_prompt(state: VBAState) -> VBAState:
    match = find_best_reference(state["vba_code"])
    if match:
        state["final_prompt"] = f"Use the same structure as this good Python code for similar macro:\n{match}\n\nNow convert this:\n{state['vba_code']}"
    else:
        state["final_prompt"] = PROMPTS[state["category"]].format(vba_code=state["vba_code"])
    with st.expander("Step 3: AI Prompt"):
        st.code(state["final_prompt"], language="text")
    progress.progress(60)
    return state

def generate_python_code(state: VBAState) -> VBAState:
    with st.spinner("Generating Python code..."):
        full = "".join(stream_claude(state["final_prompt"]))
        code = full
        if "```python" in full:
            s = full.find("```python") + len("```python")
            e = full.find("```", s)
            code = full[s:e].strip()
        state["generated_code"] = code
    with st.expander("Step 4: Generated Code"):
        st.code(code, language="python")
    if st.button("✅ Looks Good"):
        save_reference(state["vba_code"], code)
        update_reference_feedback(state["vba_code"], True)
        st.success("Saved as reference.")
    elif st.button("❌ Needs Fixing"):
        update_reference_feedback(state["vba_code"], False)
        st.warning("Feedback noted.")
    py_path = os.path.splitext(st.session_state['xlsx_path'])[0] + ".py"
    try:
        with open(py_path, "w") as f:
            f.write(code)
        st.markdown(f"**Saved Python code at:** `{py_path}`")
    except Exception as e:
        st.error(f"Error saving Python file: {e}")
    progress.progress(80)
    return state

def build_graph():
    steps = [extract_vba, categorize_vba, build_prompt, generate_python_code]
    g = StateGraph(VBAState)
    for fn in steps:
        g.add_node(fn.__name__, fn)
    g.set_entry_point(steps[0].__name__)
    for a, b in zip(steps, steps[1:]):
        g.add_edge(a.__name__, b.__name__)
    g.add_edge(steps[-1].__name__, END)
    return g.compile()

st.set_page_config(page_title="VBA2PyGen", layout="wide")
st.markdown("""
<style>
  body {background:#0e1117; color:#c7d5e0}
  .stTextArea textarea, .stTextInput input {background:#1e222d; color:#c7d5e0}
</style>
""", unsafe_allow_html=True)
st.title("VBA2PyGen")
st.markdown("Upload your Excel file")
progress = st.progress(0)

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsm","xlsb","xls"])
if not uploaded_file:
    st.session_state.pop("generated_code", None)
    st.session_state.pop("xlsx_path", None)
    st.info("Please upload a file to continue.")
    st.stop()

xlsx_path, base_name = save_uploaded_file(uploaded_file)
st.session_state['xlsx_path'] = xlsx_path
st.markdown(f"**Macro-stripped copy:** `{xlsx_path}`")

if "generated_code" not in st.session_state:
    st.session_state["generated_code"] = None

if st.session_state["generated_code"] is None:
    suffix = os.path.splitext(uploaded_file.name)[1]
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded_file.getbuffer())
        tmp_path = tmp.name
    graph = build_graph()
    for state in graph.stream({"file_path": tmp_path}):
        final = state
    st.success("✅ Conversion completed!")
else:
    st.success("✅ Already processed.")







sqlite3.OperationalError: near "references": syntax error
Traceback:
File "C:\Users\arpapate\Desktop\Generate_macro_prompt\.venv\Lib\site-packages\streamlit\runtime\scriptrunner\exec_code.py", line 121, in exec_func_with_error_handling
    result = func()
             ^^^^^^
File "C:\Users\arpapate\Desktop\Generate_macro_prompt\.venv\Lib\site-packages\streamlit\runtime\scriptrunner\script_runner.py", line 640, in code_to_exec
    exec(code, module.__dict__)
File "C:\Users\arpapate\Desktop\Generate_macro_prompt\test.py", line 28, in <module>
    init_db()
File "C:\Users\arpapate\Desktop\Generate_macro_prompt\test.py", line 19, in init_db
    conn.execute("""
