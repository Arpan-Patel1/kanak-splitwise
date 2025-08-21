# streamlit_app.py
# Run: streamlit run streamlit_app.py

import os
import re
import json
import math
import sqlite3
import tempfile
from typing import TypedDict, Optional, List, Tuple
from collections import Counter
import numpy as np
import streamlit as st
import boto3
from oletools.olevba import VBA_Parser

# ==============================
# === Configs & Constants ===
# ==============================
REGION = "us-east-1"
bedrock = boto3.client("bedrock-runtime", region_name=REGION)

EMBED_MODEL_ID = "amazon.titan-embed-text-v2:0"
# Replace with your Bedrock Claude model or inference profile ID:
CLAUDE_MODEL_ID = "arn:aws:bedrock:us-east-1:137360334857:inference-profile/us.anthropic.claude-3-7-sonnet-20250219-v1:0"

DB_PATH = "macro_embeddings.db"

PROMPTS = {
    "pivot_table": (
        "I have the following VBA code that creates a Pivot Table in Excel:\n{vba_code}\n"
        "Please write equivalent Python code that:\n"
        "- Produces the same summarized data the pivot table would show (e.g., group by fields, aggregation like SUM, COUNT, AVERAGE).\n"
        "- Uses pandas to perform the summary using pivot_table() or groupby().\n"
        "- Saves the resulting table into a sheet where it is supposed to be in the same Excel file using pandas.ExcelWriter or openpyxl.\n"
        "- Does not create a real Excel PivotTable, and does not use any fake or unsupported APIs like openpyxl.worksheet.table.tables.Table.\n"
        "- Make sure all Python libraries used are valid and the code runs end-to-end."
    ),
    "pivot_chart": (
        "I have the following VBA code that creates a Pivot Chart in Excel:\n{vba_code}\n"
        "Use pandas to perform the same data summarization (as done by the PivotTable feeding the chart).\n"
        "Generate a chart that visually represents the same data, using a real Python charting library like matplotlib or plotly.\n"
        "The chart type should match what’s used in the VBA (e.g., column chart, line chart, pie chart, etc.).\n"
        "Save the resulting chart to an image file (PNG/JPG) or embed it into a new sheet of the same Excel workbook using openpyxl or xlsxwriter (if possible).\n"
        "Avoid using any non-existent Excel chart APIs. Make sure all code is real, valid, and executable with standard Python libraries.\n"
        "Do not use functions like ws.clear_rows()."
    ),
    "user_form": (
        "I have the following VBA code that creates and manages a UserForm in Excel, including Private Subs for buttons, form initialization, validation, and possibly creating charts.\n"
        "Please generate equivalent Python code that meets all of these requirements:\n"
        "1) UI Framework: Use tkinter (with ttk) or PyQt. Recreate forms/labels/textboxes/dropdowns/buttons.\n"
        "2) Excel Operations: Use openpyxl for reading/writing. Use openpyxl.chart for a supported chart. Do NOT attempt pivot tables/charts.\n"
        "   Do not add explanatory text or placeholder sheets; only create real charts supported by openpyxl.chart.\n"
        "3) DB Access: If VBA uses DB, use pyodbc with parameterized queries.\n"
        "4) No external custom modules; keep single-file.\n"
        "5) Convert each Private Sub to Python event handlers; preserve logic.\n"
        "6) Convert MsgBox to tkinter.messagebox / PyQt dialogs.\n"
        "7) Only keep comments from the original VBA; do not add TODOs.\n\n"
        "Here is the VBA code to convert:\n{vba_code}"
    ),
    # NOTE: key is "formulas" (plural) to align with classifier result handling.
    "formulas": (
        "I have the following VBA or Excel formula-based code:\n{vba_code}\n"
        "Please generate equivalent Python code that:\n"
        "- Replicates the same logic and calculations performed by the formulas.\n"
        "- Uses pandas/numpy/openpyxl to evaluate the logic.\n"
        "- If formulas are row-wise, apply via pandas.apply() or vectorized ops.\n"
        "- If they reference ranges, load with pandas.read_excel() or openpyxl and apply accordingly.\n"
        "- Compute results in Python (do NOT embed Excel formulas) and write final values back to Excel.\n"
        "- Use only valid Python libraries and real APIs."
    ),
    "normal_operations": (
        "I have the following VBA code that performs normal Excel operations:\n{vba_code}\n"
        "Please write equivalent Python code that:\n"
        "- Performs the same operations using openpyxl or pandas.\n"
        "- If VBA modifies sheet structure (insert/delete rows/cols, rename sheets, copy data), mirror with openpyxl.\n"
        "- For value-level ops, use openpyxl or pandas appropriately.\n"
        "- If formatting is applied, replicate with openpyxl.styles.\n"
        "- Use only supported APIs.\n"
        "- The final code should be fully executable and equivalent in logic."
    )
}

CATEGORY_ALIASES = {
    "formula": "formulas",
    "formulas": "formulas",
    "pivot": "pivot_table",
    "pivot tables": "pivot_table",
    "pivot table": "pivot_table",
    "pivot chart": "pivot_chart",
    "userform": "user_form",
}

CATEGORY_MUST = {
    "pivot_table": re.compile(r'Pivot(Cache|Caches|Table|Tables|Field|Fields)', re.I),
    "pivot_chart": re.compile(r'Chart|ChartObjects|SeriesCollection|PlotArea', re.I),
    "user_form":   re.compile(r'UserForm|CommandButton|TextBox|ComboBox|Label|Initialize', re.I),
    "formulas":    re.compile(r'Formula', re.I),  # loose
    "normal_operations": re.compile(r'.', re.I)
}

THRESH = {
    "pivot_table": 0.85,
    "pivot_chart": 0.82,
    "user_form":   0.80,
    "formulas":    0.80,
    "normal_operations": 0.78
}

# ==============================
# === DB Init ==================
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
        timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
    )
    """)
    conn.execute("""
    CREATE TABLE IF NOT EXISTS macro_units (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        parent_id INTEGER,      -- references macro_matches.id
        name TEXT,              -- Sub/Function name if detected
        category TEXT,
        vba_raw TEXT,
        vba_norm TEXT,
        tokens TEXT,            -- JSON list
        embedding TEXT,         -- JSON list
        feedback INTEGER DEFAULT 0,
        created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_units_cat ON macro_units(category)")
    conn.close()

init_db()

# ==============================
# === Helpers: VBA parsing ====
# ==============================
SUB_SPLIT = re.compile(
    r'(?im)^\s*(Public|Private)?\s*(Sub|Function)\s+([A-Za-z0-9_]+)\b.*?^\s*End\s+\2\s*$',
    re.S
)

def split_procedures(vba_text: str) -> List[str]:
    matches = [m.group(0) for m in SUB_SPLIT.finditer(vba_text)]
    return matches or [vba_text]

def strip_comments(vba: str) -> str:
    out = []
    for line in vba.splitlines():
        pos = line.find("'")
        out.append(line if pos < 0 else line[:pos])
    return "\n".join(out)

KEYWORDS = set(map(str.lower, """
sub function end if then else elseif for next while wend do loop select case
dim as set let get call with endwith public private const option explicit
""".split()))

def normalize_identifiers(code: str) -> str:
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

def tokenize_api_focus(code: str) -> List[str]:
    # API-ish tokens: PivotCaches.Create, ChartObjects, UserForm, Range, Cells, Worksheets, etc.
    toks = re.findall(r'[A-Za-z][A-Za-z0-9_.]{2,}', code)
    return [t.lower() for t in toks]

# ==============================
# === Tiny BM25 ===============
# ==============================
class BM25:
    def __init__(self, docs: List[List[str]], k1=1.5, b=0.75):
        self.k1, self.b = k1, b
        self.docs = docs
        self.N = len(docs)
        self.df = Counter()
        self.len = [len(d) for d in docs]
        self.avg_len = (sum(self.len)/self.N) if self.N else 0.0
        for d in docs:
            for t in set(d):
                self.df[t] += 1
        # log-saturated IDF
        self.idf = {t: math.log(1 + (self.N - n + 0.5) / (n + 0.5)) for t, n in self.df.items()}

    def score(self, q: List[str], doc_idx: int) -> float:
        if not self.docs:
            return 0.0
        tf = Counter(self.docs[doc_idx])
        L = self.len[doc_idx] or 1
        s = 0.0
        for t in q:
            if t not in tf:
                continue
            idf = self.idf.get(t, 0.0)
            numer = tf[t] * (self.k1 + 1)
            denom = tf[t] + self.k1 * (1 - self.b + self.b * L / (self.avg_len or 1))
            s += idf * (numer / (denom or 1))
        return s

def cosine(a, b) -> float:
    a, b = np.array(a, dtype=np.float32), np.array(b, dtype=np.float32)
    denom = (np.linalg.norm(a) * np.linalg.norm(b)) or 1.0
    return float(np.dot(a, b) / denom)

def hybrid_score(cos, bm25, alpha=0.7) -> float:
    return alpha * cos + (1 - alpha) * bm25

# ==============================
# === Bedrock calls ===========
# ==============================
def get_embedding(text: str) -> List[float]:
    # Titan expects {"inputText": "..."}
    payload = {"inputText": text[:25000]}
    resp = bedrock.invoke_model(
        modelId=EMBED_MODEL_ID,
        contentType="application/json",
        accept="application/json",
        body=json.dumps(payload)
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
    # Anthropic response shape: {'content': [{'type':'text','text':'...'}], ...}
    parts = body.get("content", [])
    texts = []
    for p in parts:
        t = p.get("text") or ""
        texts.append(t)
    return "".join(texts)

def extract_code_block(full_text: str) -> str:
    # Prefer fenced python block; else return full text
    m = re.search(r"```python\s+(.*?)\s+```", full_text, flags=re.S | re.I)
    if m:
        return m.group(1).strip()
    return full_text.strip()

# ==============================
# === Classification ==========
# ==============================
def classify_vba(vba_code: str) -> str:
    prompt = (
        "Classify the following VBA into one of these categories and return ONLY the single word:\n"
        "formulas, pivot_table, pivot_chart, user_form, normal_operations\n\n"
        f"{vba_code[:12000]}"
    )
    raw = claude_complete(prompt, max_tokens=20, temperature=0.0).strip().lower()
    cat = raw.split()[0] if raw else "normal_operations"
    cat = CATEGORY_ALIASES.get(cat, cat)
    return cat if cat in PROMPTS else "normal_operations"

# ==============================
# === VBA Extraction ==========
# ==============================
@st.cache_data(show_spinner=False)
def extract_vba(path: str) -> str:
    parser = VBA_Parser(path)
    modules = [code.strip() for _, _, _, code in parser.extract_macros() if code and code.strip()]
    return "\n\n".join(modules)

# ==============================
# === Matching Core ===========
# ==============================
def load_candidates_by_category(category: str) -> List[Tuple]:
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute(
        "SELECT id, parent_id, name, vba_raw, vba_norm, tokens, embedding, feedback FROM macro_units WHERE category=?",
        (category,)
    )
    rows = cur.fetchall()
    conn.close()
    cands = []
    for row in rows:
        uid, parent_id, name, raw, norm, toks_json, emb_json, fb = row
        try:
            toks = json.loads(toks_json) if toks_json else []
            emb = json.loads(emb_json) if emb_json else []
        except Exception:
            toks, emb = [], []
        cands.append((uid, parent_id, name, raw or "", norm or "", toks, emb, fb))
    return cands

def judge_pair_llm(query_text: str, cand_text: str) -> float:
    prompt = (
        "You are a strict code matcher. Output only a float between 0 and 1.\n\n"
        "VBA_QUERY:\n" + query_text[:8000] + "\n\n"
        "VBA_CANDIDATE:\n" + cand_text[:8000] + "\n\n"
        "Question: Do they perform the same task (ignoring variable names and formatting)?"
    )
    try:
        txt = claude_complete(prompt, max_tokens=20, temperature=0.0)
        m = re.search(r'([01](?:\.\d+)?)', txt)
        return float(m.group(1)) if m else 0.0
    except Exception:
        return 0.0

def find_best_match_units(query_vba: str, category: str, alpha=0.7, topk=20):
    # 1) Prepare query units
    q_units = split_procedures(query_vba)
    q_norms = [normalize_identifiers(strip_comments(u)) for u in q_units]
    q_tokens = [tokenize_api_focus(n) for n in q_norms]
    q_embs = [get_embedding(n if n.strip() else " ") for n in q_norms]

    # 2) Load candidate pool (prefilter by must-have tokens)
    candidates = load_candidates_by_category(category)
    if not candidates:
        return None

    must = CATEGORY_MUST.get(category, CATEGORY_MUST["normal_operations"])
    # Keep candidates that match must-have regex in raw or norm
    filtered = []
    for c in candidates:
        _, _, _, raw, norm, _, _, _ = c
        if must.search(raw) or must.search(norm):
            filtered.append(c)
    if not filtered:
        filtered = candidates  # fallback to all

    # 3) Build BM25 over candidate tokens
    bm25 = BM25([c[5] for c in filtered])  # tokens at index 5

    # 4) Score each candidate by best query unit
    scored = []
    for i, c in enumerate(filtered):
        cos_best, bm_best = 0.0, 0.0
        c_emb = c[6]
        for q_emb, q_tok in zip(q_embs, q_tokens):
            if c_emb:
                cos_best = max(cos_best, cosine(q_emb, c_emb))
            bm_best = max(bm_best, bm25.score(q_tok, i))
        scored.append((i, cos_best, bm_best))

    # Normalize BM25 to [0,1]
    if not scored:
        return None
    bm_vals = [b for _, _, b in scored]
    mn, mx = min(bm_vals), max(bm_vals) or 1.0

    ranked = []
    for i, ccos, bbm in scored:
        nb = 0.0 if mx == mn else (bbm - mn) / (mx - mn)
        s = hybrid_score(ccos, nb, alpha=alpha)
        ranked.append((s, i, ccos, nb))
    ranked.sort(reverse=True)
    top = ranked[:topk]

    # 5) LLM judge re-rank (use longest query unit text as representative)
    qtext = max(q_norms, key=len) if q_norms else query_vba
    best = None
    for s, i, ccos, nb in top:
        cand_raw = filtered[i][3]
        rr = judge_pair_llm(qtext, cand_raw)
        combined = 0.5 * s + 0.5 * rr
        if (best is None) or (combined > best[0]):
            best = (combined, i, rr, ccos, nb)

    if not best:
        return None

    combined, i, rr, ccos, nb = best
    uid, parent_id, name, raw, norm, toks, emb, fb = filtered[i]
    thresh = THRESH.get(category, 0.80)

    if combined < thresh:
        return None

    return {
        "unit_id": uid,
        "parent_id": parent_id,
        "unit_name": name or "",
        "vba_macro": raw,
        "score": combined,
        "judge": rr,
        "cosine": ccos,
        "nbm25": nb
    }

def fetch_parent_generated_code(parent_id: Optional[int]) -> Optional[str]:
    if not parent_id:
        return None
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute("SELECT generated_code FROM macro_matches WHERE id=?", (parent_id,))
    row = cur.fetchone()
    conn.close()
    return row[0] if row and row[0] else None

# ==============================
# === DB Write helpers ========
# ==============================
def insert_macro_match(filename: str, vba_full: str, category: str, emb: List[float], gen_code: str, feedback: int) -> int:
    conn = sqlite3.connect(DB_PATH)
    cur = conn.execute(
        "INSERT INTO macro_matches (name, vba_full, category, embedding, generated_code, feedback) VALUES (?,?,?,?,?,?)",
        (filename, vba_full, category, json.dumps(emb), gen_code, feedback)
    )
    conn.commit()
    last_id = cur.lastrowid
    conn.close()
    return last_id

def insert_units(parent_id: int, vba_full: str, category: str):
    units = split_procedures(vba_full)
    conn = sqlite3.connect(DB_PATH)
    for u in units:
        raw = u
        norm = normalize_identifiers(strip_comments(u))
        toks = tokenize_api_focus(norm)
        emb = get_embedding(norm if norm.strip() else " ")
        # Try to extract procedure name
        m = re.search(r'(?im)^\s*(Public|Private)?\s*(Sub|Function)\s+([A-Za-z0-9_]+)\b', raw)
        pname = m.group(3) if m else None
        conn.execute(
            "INSERT INTO macro_units (parent_id, name, category, vba_raw, vba_norm, tokens, embedding, feedback) "
            "VALUES (?,?,?,?,?,?,?,0)",
            (parent_id, pname, category, raw, norm, json.dumps(toks), json.dumps(emb))
        )
    conn.commit()
    conn.close()

def update_unit_feedback(unit_id: int, delta: int):
    conn = sqlite3.connect(DB_PATH)
    conn.execute("UPDATE macro_units SET feedback = feedback + ? WHERE id=?", (delta, unit_id))
    conn.commit()
    conn.close()

# ==============================
# === Streamlit App ===========
# ==============================
class VBAState(TypedDict):
    vba_code: str
    category: str
    embedding: List[float]
    match: Optional[dict]
    py_code: str
    parent_id: Optional[int]  # id of macro_matches row just inserted on upvote, else None

st.set_page_config(page_title="VBA2PyGen+ (Accurate Matching)", layout="wide")
st.title("🧠 VBA2PyGen + Accurate Matching")

uploaded_file = st.file_uploader("Upload Excel file (.xlsm/.xls; note: .xlsb may not extract macros reliably)")

if not uploaded_file:
    st.stop()

file_id = uploaded_file.name

# Reset vote when a new file arrives
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
    with st.spinner("Step 1: Extracting VBA..."):
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(file_id)[1])
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

    with st.spinner("Step 2: Embedding (full doc)..."):
        # Full-doc embedding (for storage/reference)
        full_norm = normalize_identifiers(strip_comments(vba_code))
        emb_full = get_embedding(full_norm if full_norm.strip() else " ")
    progress.progress(40)

    with st.expander("Extracted VBA Code"):
        st.code(vba_code, language="vb")

    with st.spinner("Step 3: Categorizing VBA..."):
        category = classify_vba(vba_code)
        st.markdown(f"**Detected Category:** `{category}`")
    progress.progress(60)

    with st.spinner("Step 4: Finding high-confidence reference..."):
        match = find_best_match_units(vba_code, category)
        if match:
            st.success(f"High-confidence reference found (unit #{match['unit_id']}) — score {match['score']:.2f}")
            with st.expander("Matched VBA Macro (reference)"):
                st.code(match["vba_macro"], language="vb")
            ref_python = fetch_parent_generated_code(match["parent_id"])
            if ref_python:
                with st.expander("Previously Generated Python (reference)"):
                    st.code(ref_python, language="python")
        else:
            st.info("No high-confidence reference found.")
    progress.progress(75)

    with st.spinner("Step 5: Building prompt..."):
        base_prompt = PROMPTS[category].format(vba_code=vba_code)
        if match:
            ref_code = fetch_parent_generated_code(match["parent_id"]) or ""
            if ref_code.strip():
                prompt_text = base_prompt + "\n\nUse this Python code as reference if relevant:\n" + ref_code
            else:
                prompt_text = base_prompt
        else:
            prompt_text = base_prompt
        with st.expander("Prompt Used"):
            st.code(prompt_text, language="text")
    progress.progress(85)

    with st.spinner("Step 6: Generating Python code..."):
        # We keep generation simple + robust (non-streaming); you can switch to streaming if you want
        full = claude_complete(prompt_text, max_tokens=14000, temperature=0.0)
        py_code = extract_code_block(full)
        with st.expander("Generated Python Code"):
            st.code(py_code, language="python")
    progress.progress(100)

    st.session_state["state"] = VBAState(
        vba_code=vba_code,
        category=category,
        embedding=emb_full,
        match=match,
        py_code=py_code,
        parent_id=None
    )
    st.session_state["processed_file_id"] = file_id

# Show prior state if not reprocessing
state: Optional[VBAState] = st.session_state.get("state")
if state and not do_process:
    with st.expander("Extracted VBA Code"):
        st.code(state['vba_code'], language="vb")
    if state.get("match"):
        m = state['match']
        st.markdown(f"**High-confidence reference (cached):** unit #{m['unit_id']} — score {m['score']:.2f}")
        with st.expander("Matched VBA Macro (reference)"):
            st.code(m["vba_macro"], language="vb")
        ref_python = fetch_parent_generated_code(m["parent_id"])
        if ref_python:
            with st.expander("Previously Generated Python (reference)"):
                st.code(ref_python, language="python")
    st.markdown(f"**Detected Category:** `{state['category']}`")
    with st.expander("Generated Python Code"):
        st.code(state['py_code'], language="python")

# ==============================
# === Voting Actions ==========
# ==============================
def upvote():
    # Insert whole-file record and per-unit rows
    parent_id = insert_macro_match(
        file_id,
        state['vba_code'],
        state['category'],
        state['embedding'],
        state['py_code'],
        feedback=1
    )
    insert_units(parent_id, state['vba_code'], state['category'])
    # If a match was used, boost its unit feedback
    if state.get("match") and state["match"]:
        update_unit_feedback(state["match"]["unit_id"], +1)
    st.session_state["voted"] = True
    # Store parent id (for any future use)
    state["parent_id"] = parent_id

def downvote():
    # Penalize matched unit if any
    if state.get("match") and state["match"]:
        update_unit_feedback(state["match"]["unit_id"], -1)
    st.session_state["voted"] = True

col1, col2 = st.columns(2)
col1.button("👍 Helpful", on_click=upvote, disabled=st.session_state['voted'] or not state)
col2.button("👎 Not Helpful", on_click=downvote, disabled=st.session_state['voted'] or not state)

st.caption("Tip: Votes directly influence future matches. Upvotes add your current conversion to the database and boost the matched unit; downvotes penalize the matched unit.")
