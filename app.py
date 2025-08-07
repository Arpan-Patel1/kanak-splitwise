import os
import json
import sqlite3
import tempfile
from typing import TypedDict, Optional
import numpy as np
import streamlit as st
import boto3
from oletools.olevba import VBA_Parser

# === Configs ===
bedrock = boto3.client("bedrock-runtime", region_name="us-east-1")
EMBED_MODEL_ID = "amazon.titan-embed-text-v2:0"
DB_PATH = "macro_embeddings.db"

PROMPTS = {
    ...  # Keep all PROMPTS the same as your original code
}

# === DB Init ===
def init_db():
    try:
        conn = sqlite3.connect(DB_PATH)
        conn.execute(
            """
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
    except Exception as e:
        print("Failed to init DB:", e)

init_db()

# === Cachable Helpers ===
@st.cache_data(show_spinner=False)
def extract_vba(path: str) -> str:
    parser = VBA_Parser(path)
    modules = [code.strip() for _, _, _, code in parser.extract_macros() if code.strip()]
    return "\n\n".join(modules)

@st.cache_data(show_spinner=False)
def get_embedding(text: str):
    payload = {"inputText": text[:25000]}
    resp = bedrock.invoke_model(
        modelId=EMBED_MODEL_ID,
        contentType="application/json",
        accept="application/json",
        body=json.dumps(payload),
    )
    return json.loads(resp["body"].read())["embedding"]

@st.cache_data(show_spinner=False)
def classify_vba(vba_code: str) -> str:
    prompt = (
        "Classify into: formulas, pivot_table, pivot_chart, user_form, normal_operations. Return only the category.\n\n"
        + vba_code
    )
    cat = "".join(stream_claude(prompt)).strip().lower()
    return cat if cat in PROMPTS else "normal_operations"

# === Matching & DB ===
def cosine_similarity(v1, v2):
    a, b = np.array(v1), np.array(v2)
    return float(np.dot(a, b) / (np.linalg.norm(a) * np.linalg.norm(b)))

def find_best_match(emb, threshold=0.5):
    if not os.path.exists(DB_PATH):
        return None

    try:
        conn = sqlite3.connect(DB_PATH)
        cur = conn.execute("SELECT id,name,vba_macro,embedding,generated_code FROM macro_matches")
        best, score = None, -1.0

        for row in cur:
            try:
                old = json.loads(row[3])
                sim = cosine_similarity(emb, old)
                if sim > score:
                    best, score = row, sim
            except Exception as e:
                print("Bad embedding record skipped:", e)

        conn.close()

        if best and score >= threshold:
            id_, name, vba, embstr, code = best[0], best[1], best[2], best[3], best[4]
            return {"id": id_, "name": name, "vba_macro": vba, "generated_code": code, "score": score}

    except Exception as e:
        print("Error accessing DB:", e)

    return None

def insert_record(name, vba_macro, category, emb, code, feedback):
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute(
        "INSERT INTO macro_matches (name,vba_macro,category,embedding,generated_code,feedback) VALUES (?,?,?,?,?,?)",
        (name, vba_macro, category, json.dumps(emb), code, feedback)
    )
    conn.commit()
    conn.close()
    return cur.lastrowid

def update_feedback(record_id, delta):
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        "UPDATE macro_matches SET feedback = feedback + ? WHERE id = ?",
        (delta, record_id)
    )
    conn.commit()
    conn.close()

# === Claude Stream ===
def stream_claude(prompt: str):
    payload = {"anthropic_version": "bedrock-2023-05-31", "messages": [{"role": "user", "content": prompt}], "max_tokens": 4000, "temperature": 0}
    resp = bedrock.invoke_model_with_response_stream(
        modelId=("arn:aws:bedrock:us-east-1:137360334857:inference-profile/us.anthropic.claude-3-7-sonnet-20250219-v1:0"),
        body=json.dumps(payload),
    )
    for event in resp.get("body", []):
        chunk = json.loads(event.get("chunk", {}).get("bytes", b"{}"))
        text = chunk.get("delta", {}).get("text")
        if text:
            yield text

class VBAState(TypedDict):
    vba_code: str
    category: str
    embedding: list
    match: Optional[dict]
    py_code: str

# === Streamlit App ===
st.set_page_config(page_title="VBA2PyGen+", layout="wide")
st.title("\U0001F9E0 VBA2PyGen + Titan Matching")

uploaded_file = st.file_uploader("Upload Excel file (.xlsm/.xls/.xlsb)")
if not uploaded_file:
    st.stop()
file_id = uploaded_file.name

if "state" not in st.session_state:
    st.session_state["state"] = None
if "voted" not in st.session_state:
    st.session_state["voted"] = False
if "processed_file_id" not in st.session_state:
    st.session_state["processed_file_id"] = None

do_process = (st.session_state["processed_file_id"] != file_id) and not st.session_state["voted"]

progress = st.progress(0)

if do_process:
    with st.spinner("Step 1: Extracting VBA..."):
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(file_id)[1])
        tmp.write(uploaded_file.getbuffer()); tmp.flush(); tmp_path = tmp.name; tmp.close()
        vba_code = extract_vba(tmp_path)
    progress.progress(20)

    with st.spinner("Step 2: Embedding & matching..."):
        emb = get_embedding(vba_code)
        match = find_best_match(emb)
    progress.progress(40)

    with st.expander("Extracted VBA Code"):
        st.code(vba_code, language="vb")

    if match:
        st.markdown(f"**Reference Found:** `{match['name']}` — `{match['score']*100:.1f}%`")
        with st.expander("Matched VBA Macro"):
            st.code(match['vba_macro'], language="vb")
        with st.expander("Matched Python Code"):
            st.code(match['generated_code'], language="python")
    else:
        st.info("No similar macro found in database. Generating code without reference.")

    with st.spinner("Step 3: Categorizing VBA..."):
        category = classify_vba(vba_code)
        st.markdown(f"**Detected Category:** `{category}`")
    progress.progress(60)

    with st.spinner("Step 4: Building prompt..."):
        if match:
            prompt_text = (
                PROMPTS[category].format(vba_code=vba_code) +
                "\n\nUse This Python Code as reference:\n" +
                match['generated_code'] +
                f"\n\nThis code is `{match['score']*100:.1f}%` similar to what we want."
            )
        else:
            prompt_text = PROMPTS[category].format(vba_code=vba_code)

    with st.expander("Prompt Used"):
        st.code(prompt_text, language="text")
    progress.progress(80)

    with st.spinner("Step 5: Generating Python code..."):
        full = "".join(stream_claude(prompt_text))
        py_code = (full.split("```python", 1)[1].split("```", 1)[0].strip() if "```python" in full else full.strip())
    with st.expander("Generated Python Code"):
        st.code(py_code, language="python")
    progress.progress(100)

    st.session_state["state"] = VBAState(vba_code=vba_code, category=category, embedding=emb, match=match, py_code=py_code)
    st.session_state["processed_file_id"] = file_id

state = st.session_state.get("state")
if state and not do_process:
    with st.expander("Extracted VBA Code"):
        st.code(state['vba_code'], language="vb")
    if state.get("match"):
        st.markdown(f"**Reference Found:** `{state['match']['name']}` — `{state['match']['score']*100:.1f}%`")
        with st.expander("Matched VBA Macro"):
            st.code(state['match']['vba_macro'], language="vb")
        with st.expander("Matched Python Code"):
            st.code(state['match']['generated_code'], language="python")
    st.markdown(f"**Detected Category:** `{state['category']}`")
    with st.expander("Prompt Used"):
        st.code(PROMPTS[state['category']].format(vba_code=state['vba_code']), language="text")
    with st.expander("Generated Python Code"):
        st.code(state['py_code'], language="python")

def upvote():
    rec_id = insert_record(
        file_id,
        state['vba_code'],
        state['category'],
        state['embedding'],
        state['py_code'],
        1
    )
    if state.get('match'):
        update_feedback(state['match']['id'], 1)
    st.session_state['voted'] = True

def downvote():
    if state.get('match'):
        update_feedback(state['match']['id'], -1)
    st.session_state['voted'] = True

col1, col2 = st.columns(2)
col1.button("\U0001F44D Helpful", on_click=upvote, disabled=st.session_state['voted'])
col2.button("\U0001F44E Not Helpful", on_click=downvote, disabled=st.session_state['voted'])
