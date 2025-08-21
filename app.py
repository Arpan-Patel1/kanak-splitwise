# streamlit_app.py
# Run: streamlit run streamlit_app.py

import os
import re
import json
import sqlite3
import tempfile
import hashlib
from datetime import datetime
from typing import TypedDict, Optional, List, Tuple

import numpy as np
import streamlit as st
import boto3
from oletools.olevba import VBA_Parser
from openpyxl import load_workbook

# ==============================
# === Config & Constants =======
# ==============================
APP_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(APP_DIR, "macro_embeddings.db")
OUTPUT_ROOT = os.path.join(APP_DIR, "outputs")
os.makedirs(OUTPUT_ROOT, exist_ok=True)

REGION = "us-east-1"
bedrock = boto3.client("bedrock-runtime", region_name=REGION)

EMBED_MODEL_ID = "amazon.titan-embed-text-v2:0"
# Replace with your Bedrock Claude model/inference profile if different:
CLAUDE_MODEL_ID = "arn:aws:bedrock:us-east-1:137360334857:inference-profile/us.anthropic.claude-3-7-sonnet-20250219-v1:0"

PROMPTS = {
    "pivot_table": (
        "I have the following VBA code that creates a Pivot Table in Excel:\n{vba_code}\n"
        "Please write equivalent Python code that:\n"
        "- Produces the same summarized data (group by fields with SUM/COUNT/AVERAGE).\n"
        "- Uses pandas (pivot_table or groupby).\n"
        "- Saves the result to the same Excel file (pandas.ExcelWriter or openpyxl).\n"
        "- Does NOT create a real Excel PivotTable or call unsupported APIs.\n"
        "- Use only real, installable Python libs, end-to-end runnable."
    ),
    "pivot_chart": (
        "I have the following VBA code that creates a Pivot Chart in Excel:\n{vba_code}\n"
        "Please generate Python that:\n"
        "- Summarizes data with pandas (as the PivotTable would).\n"
        "- Builds an equivalent chart using matplotlib or plotly matching the VBA chart type.\n"
        "- Saves the chart to an image or embeds via openpyxl/xlsxwriter if possible.\n"
        "- Avoids non-existent Excel chart APIs. Use only real, runnable code."
    ),
    "user_form": (
        "I have the following VBA that builds a UserForm (buttons, init, validation, maybe charts).\n"
        "Generate Python that:\n"
        "- Uses tkinter (with ttk) or PyQt to recreate the UI widgets.\n"
        "- Uses openpyxl for Excel I/O; use openpyxl.chart for a standard supported chart if needed.\n"
        "- DO NOT create pivot tables/charts. No placeholder sheets/text.\n"
        "- If database ops exist, use pyodbc with parameterized queries.\n"
        "- Convert each Private Sub to an event handler; preserve logic.\n"
        "- Only real imports; single-file code.\n\n"
        "Here is the VBA code:\n{vba_code}"
    ),
    # NOTE: use "formulas" (plural) key to match classifier aliasing below.
    "formulas": (
        "I have the following VBA or Excel formulas:\n{vba_code}\n"
        "Generate Python that:\n"
        "- Reproduces the same calculations with pandas/numpy/openpyxl.\n"
        "- If row-wise, use vectorized ops or apply; if ranges, load with read_excel/openpyxl.\n"
        "- Compute results in Python (do not embed Excel formulas) and write values back.\n"
        "- Use only valid libraries and real APIs."
    ),
    "normal_operations": (
        "I have the following VBA that does general Excel operations:\n{vba_code}\n"
        "Generate Python that:\n"
        "- Mirrors sheet-level edits with openpyxl (insert/delete rows/cols, rename, copy data).\n"
        "- Uses openpyxl/pandas for value-level changes.\n"
        "- Replicates formatting with openpyxl.styles if present.\n"
        "- Uses only supported APIs; runnable end-to-end."
    )
}

CATEGORY_ALIASES = {
    "formula": "formulas",
    "formulas": "formulas",
    "pivot": "pivot_table",
    "pivot table": "pivot_table",
    "pivot tables": "pivot_table",
    "pivot_chart": "pivot_chart",
    "pivot chart": "pivot_chart",
    "userform": "user_form",
}

# ==============================
# === DB Setup =================
# ==============================
def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
    CREATE TABLE IF NOT EXISTS macro_matches (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT,
        vba_full TEXT,
        category TEXT,
        embedding TEXT,
        generated_code TEXT,
        feedback INTEGER DEFAULT 0,
        name_hash TEXT,
        timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
    )
    """)
    # add missing columns if upgrading
    try:
        conn.execute("ALTER TABLE macro_matches ADD COLUMN name_hash TEXT")
    except sqlite3.OperationalError:
        pass
    conn.close()

init_db()

# ==============================
# === Utils ====================
# ==============================
def cosine_similarity(v1: List[float], v2: List[float]) -> float:
    a = np.array(v1, dtype=np.float32)
    b = np.array(v2, dtype=np.float32)
    denom = (np.linalg.norm(a) * np.linalg.norm(b)) or 1.0
    return float(np.dot(a, b) / denom)

def strip_comments(vba: str) -> str:
    lines = []
    for line in vba.splitlines():
        pos = line.find("'")
        lines.append(line if pos < 0 else line[:pos])
    return "\n".join(lines)

def normalize_identifiers(code: str) -> str:
    # light canonicalization to stabilize fingerprints & embeddings
    KEYWORDS = set(map(str.lower, """
        sub function end if then else elseif for next while wend do loop select case
        dim as set let get call with endwith public private const option explicit
    """.split()))
    tokens = re.findall(r'[A-Za-z_][A-Za-z0-9_]*|\S', code)
    out, v_i, seen = [], 1, {}
    for t in tokens:
        tl = t.lower()
        if re.match(r'^[A-Za-z_][A-Za-z0-9_]*$', t) and tl not in KEYWORDS:
            if t not in seen:
                seen[t] = f'v{v_i}'; v_i += 1
            out.append(seen[t])
        else:
            out.append(t)
    return "".join(out)

def vba_fingerprint(vba_text: str) -> str:
    norm = normalize_identifiers(strip_comments(vba_text))
    return hashlib.sha256(norm.encode("utf-8", errors="ignore")).hexdigest()

def extract_code_block(full_text: str) -> str:
    m = re.search(r"```python\s+(.*?)\s+```", full_text, flags=re.S | re.I)
    return m.group(1).strip() if m else full_text.strip()

# ==============================
# === Bedrock ==================
# ==============================
def get_embedding(text: str) -> List[float]:
    payload = {"inputText": (text or " ")[:25000]}
    resp = bedrock.invoke_model(
        modelId=EMBED_MODEL_ID,
        contentType="application/json",
        accept="application/json",
        body=json.dumps(payload),
    )
    body = json.loads(resp["body"].read())
    return body["embedding"]

def claude_complete(prompt: str, max_tokens: int = 4000, temperature: float = 0.0) -> str:
    payload = {
        "anthropic_version": "bedrock-2023-05-31",
        "messages": [{"role": "user", "content": prompt}],
        "max_tokens": max_tokens,
        "temperature": temperature,
    }
    resp = bedrock.invoke_model(
        modelId=CLAUDE_MODEL_ID,
        body=json.dumps(payload)
    )
    body = json.loads(resp["body"].read())
    parts = body.get("content", [])
    return "".join(p.get("text") or "" for p in parts)

# ==============================
# === Extraction & Classify ====
# ==============================
@st.cache_data(show_spinner=False)
def extract_vba(path: str) -> str:
    parser = VBA_Parser(path)
    modules = [code.strip() for _, _, _, code in parser.extract_macros() if code and code.strip()]
    return "\n\n".join(modules)

def classify_vba(vba_code: str) -> str:
    prompt = (
        "Return ONLY one of: formulas, pivot_table, pivot_chart, user_form, normal_operations.\n\n"
        f"{vba_code[:12000]}"
    )
    raw = claude_complete(prompt, max_tokens=16, temperature=0.0).strip().lower()
    cat = (raw.split() or ["normal_operations"])[0]
    cat = CATEGORY_ALIASES.get(cat, cat)
    return cat if cat in PROMPTS else "normal_operations"

# ==============================
# === DB I/O ===================
# ==============================
def insert_record(name: str, vba_full: str, category: str, emb: List[float], code: str, feedback: int) -> int:
    name_hash = vba_fingerprint(vba_full)
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute(
        "INSERT INTO macro_matches (name, vba_full, category, embedding, generated_code, feedback, name_hash) "
        "VALUES (?,?,?,?,?,?,?)",
        (name, vba_full, category, json.dumps(emb), code, feedback, name_hash)
    )
    conn.commit()
    rid = cur.lastrowid
    conn.close()
    return rid

def update_feedback(record_id: int, delta: int):
    conn = sqlite3.connect(DB_PATH)
    conn.execute("UPDATE macro_matches SET feedback = feedback + ? WHERE id = ?", (delta, record_id))
    conn.commit()
    conn.close()

def lookup_exact_by_hash(fp: str) -> Optional[Tuple[int, str]]:
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute("SELECT id, generated_code FROM macro_matches WHERE name_hash=?", (fp,))
    row = cur.fetchone()
    conn.close()
    return (row[0], row[1]) if row else None

def find_best_match(emb: List[float], category: str, threshold: float = 0.5) -> Optional[dict]:
    if not os.path.exists(DB_PATH):
        return None
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute("SELECT id,name,vba_full,category,embedding,generated_code,feedback FROM macro_matches")
    best, best_score, best_row = None, -1.0, None
    for row in cur:
        rid, name, vba_full, cat, emb_json, gen_code, fb = row
        try:
            old = json.loads(emb_json)
            score = cosine_similarity(emb, old)
            # prefer same category; give small boost if same
            if cat == category:
                score += 0.02
            # tiny reinforcement by feedback
            score *= (1.0 + 0.03 * (fb or 0))
            if score > best_score:
                best_score, best_row = score, row
        except Exception:
            continue
    conn.close()
    if best_row and best_score >= threshold:
        rid, name, vba_full, cat, emb_json, gen_code, fb = best_row
        return {"id": rid, "name": name, "vba_macro": vba_full, "generated_code": gen_code, "score": float(best_score)}
    return None

# ==============================
# === Streamlit State ==========
# ==============================
class VBAState(TypedDict):
    vba_code: str
    category: str
    embedding: List[float]
    match: Optional[dict]
    py_code: str
    # artifacts
    out_dir: Optional[str]
    py_path: Optional[str]
    xlsx_path: Optional[str]

# ==============================
# === App ======================
# ==============================
st.set_page_config(page_title="VBA2PyGen (Cosine Matching)", layout="wide")
st.title("🧠 VBA2PyGen — Cosine Matching (Clean)")

uploaded_file = st.file_uploader("Upload Excel file (.xlsx / .xlsm recommended)")

if not uploaded_file:
    st.stop()

file_id = uploaded_file.name
base_name, ext = os.path.splitext(file_id)

# Reset state per new file
if st.session_state.get("processed_file_id") != file_id:
    st.session_state["voted"] = False
    st.session_state["state"] = None

if "voted" not in st.session_state:
    st.session_state["voted"] = False
if "state" not in st.session_state:
    st.session_state["state"] = None
if "processed_file_id" not in st.session_state:
    st.session_state["processed_file_id"] = None

do_process = (st.session_state["processed_file_id"] != file_id)
progress = st.progress(0)

if do_process:
    # Step 1: Extract VBA
    with st.spinner("Step 1: Extracting VBA..."):
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
        tmp.write(uploaded_file.getbuffer()); tmp.flush(); tmp_path = tmp.name; tmp.close()
        try:
            vba_code = extract_vba(tmp_path)
        except Exception as e:
            st.error(f"Failed to parse macros: {e}")
            st.stop()
        if not vba_code.strip():
            st.error("No macros found in this workbook.")
            st.stop()
    progress.progress(20)

    # Optional: short-circuit exact match by fingerprint
    fp = vba_fingerprint(vba_code)
    exact = lookup_exact_by_hash(fp)

    with st.expander("Extracted VBA Code"):
        st.code(vba_code, language="vb")

    # Step 2: Embedding (full text)
    with st.spinner("Step 2: Embedding..."):
        norm = normalize_identifiers(strip_comments(vba_code))
        emb_full = get_embedding(norm if norm.strip() else " ")
    progress.progress(40)

    # Step 3: Category detection
    with st.spinner("Step 3: Categorizing..."):
        category = classify_vba(vba_code)
        st.markdown(f"**Detected Category:** `{category}`")
    progress.progress(60)

    # Step 4: Find best reference (cosine only)
    match = None
    with st.spinner("Step 4: Searching prior references (cosine)..."):
        # if exact match exists, treat it as a strong reference
        if exact:
            st.success("Exact match found from a previous run.")
            match = {"id": exact[0], "name": file_id, "vba_macro": vba_code, "generated_code": exact[1], "score": 1.0}
        else:
            match = find_best_match(emb_full, category, threshold=0.5)
        if match:
            st.markdown(f"**Reference Found:** {match['name']} — score {match['score']:.2f}")
            with st.expander("Matched VBA Macro"):
                st.code(match["vba_macro"], language="vb")
            if match.get("generated_code"):
                with st.expander("Matched Python Code (reference)"):
                    st.code(match["generated_code"], language="python")
        else:
            st.info("No similar macro found in database. Generating without reference.")
    progress.progress(75)

    # Step 5: Build prompt
    with st.spinner("Step 5: Building prompt..."):
        prompt_text = PROMPTS[category].format(vba_code=vba_code)
        if match and match.get("generated_code"):
            prompt_text += "\n\nUse this Python code as reference:\n" + match["generated_code"]
        with st.expander("Prompt Used"):
            st.code(prompt_text, language="text")
    progress.progress(85)

    # Step 6: Generate Python code
    with st.spinner("Step 6: Generating Python..."):
        full = claude_complete(prompt_text, max_tokens=14000, temperature=0.0)
        py_code = extract_code_block(full)
        with st.expander("Generated Python Code"):
            st.code(py_code, language="python")
    progress.progress(92)

    # Step 7: Save artifacts (.py + .xlsx replica w/o macros)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = os.path.join(OUTPUT_ROOT, f"{base_name}_{ts}")
    os.makedirs(out_dir, exist_ok=True)

    py_path = os.path.join(out_dir, f"{base_name}.py")
    with open(py_path, "w", encoding="utf-8") as f:
        f.write(py_code)

    xlsx_path = os.path.join(out_dir, f"{base_name}_replica.xlsx")
    replica_ok = False
    if ext.lower() in (".xlsx", ".xlsm"):
        try:
            wb = load_workbook(tmp_path, data_only=False, keep_vba=False)  # saving to .xlsx drops macros
            wb.save(xlsx_path)
            replica_ok = True
        except Exception as e:
            st.warning(f"Could not create macro-free .xlsx replica: {e}")
    else:
        st.warning("Replica only supported for .xlsx / .xlsm source files.")

    st.subheader("Downloads")
    with open(py_path, "rb") as f:
        st.download_button("⬇️ Download Python (.py)", data=f.read(), file_name=os.path.basename(py_path), mime="text/x-python")
    if replica_ok and os.path.exists(xlsx_path):
        with open(xlsx_path, "rb") as f:
            st.download_button("⬇️ Download Excel Replica (.xlsx, no macros)",
                               data=f.read(),
                               file_name=os.path.basename(xlsx_path),
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    progress.progress(100)

    st.session_state["state"] = VBAState(
        vba_code=vba_code,
        category=category,
        embedding=emb_full,
        match=match,
        py_code=py_code,
        out_dir=out_dir,
        py_path=py_path,
        xlsx_path=xlsx_path if replica_ok else None
    )
    st.session_state["processed_file_id"] = file_id

# Show cached state (if not reprocessing)
state: Optional[VBAState] = st.session_state.get("state")
if state and not do_process:
    with st.expander("Extracted VBA Code"):
        st.code(state["vba_code"], language="vb")
    if state.get("match"):
        st.markdown(f"**Reference (cached):** score {state['match'].get('score', 0):.2f}")
        with st.expander("Matched VBA Macro"):
            st.code(state["match"]["vba_macro"], language="vb")
        if state["match"].get("generated_code"):
            with st.expander("Matched Python Code (reference)"):
                st.code(state["match"]["generated_code"], language="python")
    st.markdown(f"**Detected Category:** `{state['category']}`")
    with st.expander("Generated Python Code"):
        st.code(state["py_code"], language="python")

    st.subheader("Downloads")
    if state.get("py_path") and os.path.exists(state["py_path"]):
        with open(state["py_path"], "rb") as f:
            st.download_button("⬇️ Download Python (.py)", data=f.read(),
                               file_name=os.path.basename(state["py_path"]), mime="text/x-python")
    if state.get("xlsx_path") and state["xlsx_path"] and os.path.exists(state["xlsx_path"]):
        with open(state["xlsx_path"], "rb") as f:
            st.download_button("⬇️ Download Excel Replica (.xlsx, no macros)", data=f.read(),
                               file_name=os.path.basename(state["xlsx_path"]),
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ==============================
# === Voting ===================
# ==============================
def upvote():
    parent_id = insert_record(
        file_id,
        state["vba_code"],
        state["category"],
        state["embedding"],
        state["py_code"],
        feedback=1
    )
    if state.get("match") and state["match"]:
        update_feedback(state["match"]["id"], +1)
    st.session_state["voted"] = True

def downvote():
    if state.get("match") and state["match"]:
        update_feedback(state["match"]["id"], -1)
    st.session_state["voted"] = True

col1, col2 = st.columns(2)
col1.button("👍 Helpful", on_click=upvote, disabled=st.session_state["voted"] or not state)
col2.button("👎 Not Helpful", on_click=downvote, disabled=st.session_state["voted"] or not state)

st.caption("Cosine-only matching with Titan embeddings. Upvotes save this conversion and boost the matched reference; downvotes penalize it.")
