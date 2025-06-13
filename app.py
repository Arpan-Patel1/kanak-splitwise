# ----------------- standard libs ------------------
import os, json, tempfile, sqlite3, math
from typing import TypedDict, Optional, List

# ----------------- 3rd-party libs -----------------
import openpyxl                 # Excel I/O
import streamlit as st          # UI
import boto3                    # Bedrock
import numpy as np              # cosine sim
from oletools.olevba import VBA_Parser
from langgraph.graph import StateGraph, END

# ----------------- CONSTANTS ----------------------
DB_PATH   = "macro_references.db"
EMB_MODEL = "amazon.titan-embed-text-v2"
CLAUDE_ID = (
    "arn:aws:bedrock:us-east-1:137360334857:"
    "inference-profile/us.anthropic.claude-3-7-sonnet-20250219-v1:0"
)

# ---------- PROMPTS (unchanged – cropped for brevity) ----------
PROMPTS = {
    "pivot_table": """I have the following VBA code that creates a Pivot Table in Excel:
{vba_code}

Please write equivalent Python code that:
• Produces the same summarized data (group-by, SUM, COUNT…)
• Uses pandas (pivot_table or groupby)
• Writes the table back to the workbook (pandas.ExcelWriter or openpyxl)
Do NOT create a real Excel PivotTable and do NOT call unsupported APIs.
""",

    "pivot_chart": """I have the following VBA code that creates a Pivot Chart in Excel:
{vba_code}

Generate Python that:
• Summarises the data with pandas (matching the PivotTable)
• Plots the same chart type using matplotlib / seaborn / plotly
• Does not call openpyxl chart APIs that don’t exist
""",

    "user_form": """I have the following VBA code that creates / handles a UserForm:
{vba_code}

Create Python that:
• Replicates the form flow with Tkinter or PyQt5
• Captures input and writes to Excel via pandas / openpyxl
• Uses only real libraries; self-contained & runnable
""",

    "formula": """I have the following VBA or formula code:
{vba_code}

Generate Python that:
• Replicates the calculations with pandas / numpy
• Reads the sheet via pandas.read_excel or openpyxl
• Computes in Python (DO NOT embed formulas) then writes results back
""",

    "normal_operations": """I have the following VBA code that performs normal Excel ops:
{vba_code}

Create Python that:
• Performs the same inserts / deletes / formatting with openpyxl / pandas
• Uses openpyxl.styles for formatting
• No fake / missing APIs; fully runnable
"""
}

# ----------------- Bedrock client -----------------
bedrock = boto3.client("bedrock-runtime")

# ========= SQLite helpers (table name NOT reserved) =========
def init_db() -> None:
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS macro_refs (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                category      TEXT    NOT NULL,
                vba_code      TEXT    NOT NULL,
                python_code   TEXT    NOT NULL,
                embedding     TEXT    NOT NULL,
                upvotes       INTEGER DEFAULT 1,
                downvotes     INTEGER DEFAULT 0
            )
            """
        )

def get_embedding(text: str) -> List[float]:
    body = json.dumps({"inputText": text})
    resp = bedrock.invoke_model(
        modelId   = EMB_MODEL,
        body      = body,
        contentType = "application/json",
        accept      = "application/json",
    )
    return json.loads(resp["body"].read())["embedding"]

def cosine(a: List[float], b: List[float]) -> float:
    a, b = np.array(a), np.array(b)
    return float(np.dot(a, b) / (np.linalg.norm(a) * np.linalg.norm(b)))

def best_reference(vba_code: str, category: str):
    query_vec = get_embedding(vba_code)
    best_row, best_sim = None, 0.87  # threshold
    with sqlite3.connect(DB_PATH) as conn:
        for row in conn.execute(
            "SELECT id, vba_code, python_code, embedding, upvotes, downvotes "
            "FROM macro_refs WHERE category=?", (category,)
        ):
            emb = json.loads(row[3])
            sim = cosine(query_vec, emb)
            if sim > best_sim:
                best_sim, best_row = sim, row
    return best_row  # or None

def save_reference(cat: str, vba: str, py: str) -> None:
    emb = get_embedding(vba)
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(
            "INSERT INTO macro_refs (category, vba_code, python_code, embedding)"
            " VALUES (?, ?, ?, ?)",
            (cat, vba, py, json.dumps(emb))
        )

def vote(ref_id: int, good: bool) -> None:
    col = "upvotes" if good else "downvotes"
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute(f"UPDATE macro_refs SET {col}={col}+1 WHERE id=?", (ref_id,))

# ----------------- Claude streaming -----------------
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
        modelId = CLAUDE_ID,
        body    = json.dumps(payload),
    )
    for event in resp["body"]:
        chunk = json.loads(event["chunk"]["bytes"])
        if delta := chunk.get("delta"):
            if text := delta.get("text"):
                yield text

# ----------------- LangGraph state -----------------
class VBAState(TypedDict):
    file_path     : Optional[str]
    vba_code      : Optional[str]
    category      : Optional[str]
    final_prompt  : Optional[str]
    generated_code: Optional[str]
    ref_id        : Optional[int]

# ----------------- Helper: save uploaded file ---------------
def save_file(uploaded):
    path = os.path.join(os.getcwd(), uploaded.name)
    with open(path, "wb") as f:
        f.write(uploaded.getbuffer())
    if path.lower().endswith(".xlsm"):
        wb = openpyxl.load_workbook(path, keep_vba=False)
        new_path = os.path.splitext(path)[0] + ".xlsx"
        wb.save(new_path)
        os.remove(path)
        return new_path
    return path

# ------------- LangGraph step 1: extract VBA ----------------
def extract_vba(state: VBAState) -> VBAState:
    with st.spinner("Extracting VBA..."):
        parser  = VBA_Parser(state["file_path"])
        modules = [c.strip() for *_, c in parser.extract_macros() if c.strip()]
        if not modules:
            st.error("No VBA macros found."); st.stop()
        state["vba_code"] = "\n\n".join(modules)
    st.expander("Step 1 – VBA").code(state["vba_code"])
    prog.progress(20)
    return state

# ------------- LangGraph step 2: classify -------------------
def classify(state: VBAState) -> VBAState:
    user_prompt = (
        "Classify the following VBA into formulas, pivot_table, pivot_chart, "
        "user_form, normal_operations. Return only the word.\n\n" + state["vba_code"]
    )
    cat = "".join(stream_claude(user_prompt)).strip().lower()
    state["category"] = cat if cat in PROMPTS else "normal_operations"
    st.expander("Step 2 – Category").markdown(f"`{state['category']}`")
    prog.progress(40)
    return state

# ------------- LangGraph step 3: build prompt ---------------
def build_prompt(state: VBAState) -> VBAState:
    ref = best_reference(state["vba_code"], state["category"])
    if ref:
        ref_id, ref_vba, ref_py, *_ = ref
        state["ref_id"] = ref_id
        state["final_prompt"] = (
            "Here is a high-quality example of converting a similar macro:\n\n"
            f"Macro:\n{ref_vba}\n\nPython:\n{ref_py}\n\n"
            "Now convert the following macro to Python, following the same style. "
            "Do NOT repeat the old macro.\n\n"
            + state["vba_code"]
        )
        st.info("📎 Using reference example")
    else:
        state["ref_id"] = None
        state["final_prompt"] = PROMPTS[state["category"]].format(vba_code=state["vba_code"])
    st.expander("Step 3 – Prompt").code(state["final_prompt"])
    prog.progress(60)
    return state

# ------------- LangGraph step 4: generate -------------------
def generate(state: VBAState) -> VBAState:
    with st.spinner("Claude generating…"):
        full = "".join(stream_claude(state["final_prompt"]))
    code = full
    if "```python" in full:
        s = full.find("```python") + 9
        e = full.find("```", s)
        code = full[s:e].strip()
    state["generated_code"] = code
    st.expander("Step 4 – Python").code(code, language="python")

    # feedback buttons
    good = st.button("✅ Correct")
    bad  = st.button("❌ Incorrect")
    if good or bad:
        if state["ref_id"]:
            vote(state["ref_id"], good)
        if good:
            save_reference(state["category"], state["vba_code"], code)
        st.experimental_rerun()

    # save file
    py_path = os.path.splitext(st.session_state["xlsx_path"])[0] + ".py"
    with open(py_path, "w") as f: f.write(code)
    st.success(f"Saved → `{py_path}`")
    prog.progress(85)
    return state

# ------------- Build LangGraph ------------------------------
def graph():
    g = StateGraph(VBAState)
    for fn in (extract_vba, classify, build_prompt, generate):
        g.add_node(fn.__name__, fn)
    g.set_entry_point("extract_vba")
    g.add_edge("extract_vba", "classify")
    g.add_edge("classify", "build_prompt")
    g.add_edge("build_prompt", "generate")
    g.add_edge("generate", END)
    return g.compile()

# ----------------- Streamlit UI -----------------------------
st.set_page_config("VBA2PyGen", layout="wide")
st.title("🧩 VBA → Python Generator (Bedrock + Titan v2)")

prog = st.progress(0)
uploaded = st.file_uploader("Upload .xlsm / .xlsb / .xls workbook", type=["xlsm","xlsb","xls"])
if not uploaded: st.stop()

xlsx_path = save_file(uploaded)
st.session_state["xlsx_path"] = xlsx_path
st.caption(f"Macro-stripped copy saved at `{xlsx_path}`")

init_db()  # ensure table exists

workflow = graph()
for _state in workflow.stream({"file_path": xlsx_path}):
    pass  # LangGraph yields intermediate states but we only need end
st.success("✅ Done!  You can upload another file when ready.")
