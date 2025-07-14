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
    "pivot_table": """I have the following VBA code that creates a Pivot Table in Excel:\n{vba_code}
    Please write equivalent Python code that:
    Produces the same summarized data the pivot table would show (e.g., group by fields, aggregation like SUM, COUNT, AVERAGE).
    Uses pandas to perform the summary using pivot_table() or groupby().
    Saves the resulting table into a sheet where it is suppose to be in the same Excel file using pandas.ExcelWriter or openpyxl.
    Does not create a real Excel PivotTable, and does not use any fake or unsupported APIs like openpyxl.worksheet.table.tables.Table.
    Make sure all Python libraries used are valid and the code runs end-to-end.
    """,

    "pivot_chart": """I have the following VBA code that creates a Pivot Chart in Excel:\n{vba_code}
    Uses pandas to perform the same data summarization (as done by the PivotTable feeding the chart).
    Generates a chart that visually represents the same data, using a real Python charting library like matplotlib, seaborn, or plotly.
    The chart type should match what’s used in the VBA (e.g., column chart, line chart, pie chart, etc.).
    Saves the resulting chart to an image file (PNG/JPG) or embeds it into a new sheet of the same Excel workbook using openpyxl or xlsxwriter (if possible).
    Avoid using any non-existent Excel chart APIs in openpyxl or other libraries.
    Make sure all code is real, valid, and executable with standard Python libraries. Do not use functions like (ws.clear_rows())
    """
,

    "user_form": """
    I have the following VBA code that creates and manages a UserForm in Excel, including Private Subs for buttons, form initialization, validation, and possibly creating charts.

Please generate equivalent Python code that meets all of these requirements:

1. **UI Framework:**
   - Use tkinter (with ttk widgets if needed) or PyQt for the UI—whichever is most appropriate.
   - Recreate all form controls, including windows/forms, labels, textboxes, dropdowns, and buttons.

2. **Excel Operations:**
   - Use openpyxl for all reading and writing of Excel data.
   - Include the use of `openpyxl.chart` to create a standard chart (for example, a BarChart or LineChart) reflecting any relevant data, even if the VBA code references a pivot chart.
   - **DO NOT attempt to create pivot charts or pivot tables**, since these are not supported by openpyxl.
   - **DO NOT create any extra sheets or placeholder cells explaining unsupported features.**
   - **DO NOT write any explanatory text (e.g., "Note: This sheet would contain...") into the workbook.**
   - Only create actual charts that are supported by openpyxl.chart, and link them to real data ranges.
   - If any chart feature is not supported, simply omit that part—do not include TODO comments or extra remarks in code or Excel.

3. **Database Access:**
   - If the VBA code includes database operations, use pyodbc.
   - Use parameterized queries rather than f-strings or .format() to build SQL statements.

4. **Custom Modules:**
   - Do not import or rely on any custom or undefined helper modules (e.g., insert_form, delete_form, excel_operations).
   - All code must be self-contained in a single script.

5. **Event Logic:**
   - Convert each VBA Private Sub to an equivalent Python method or event handler.
   - Preserve all logic and behavior.

6. **Messages:**
   - Convert MsgBox or other VBA dialogs to tkinter.messagebox or PyQt dialogs.

7. **Comments:**
   - Only include comments that were present in the original VBA code.
   - **DO NOT add any TODO comments, explanatory remarks, or placeholders.**

8. **Code Validity and Imports:**
   - Use only valid, installable Python libraries (e.g., openpyxl, pyodbc, tkinter, PyQt).
   - Ensure all imports are real and correct.

The Python script should faithfully replicate the VBA UserForm functionality using openpyxl.chart for supported chart creation, without any unnecessary comments, TODOs, or explanatory cells in the workbook.

Here is the VBA code to convert: \n{vba_code} """,

    "formula": """I have the following VBA or Excel formula-based code:\n{vba_code}
    Please generate equivalent Python code that:
    Replicates the same logic and calculations performed by the formulas.
    Uses real Python libraries like pandas, numpy, or openpyxl to evaluate the logic.
    If the formulas are row-wise, apply them using pandas.apply() or vectorized operations.
    If they reference Excel ranges, load the file using pandas.read_excel() or openpyxl, and apply the logic accordingly.
    The goal is to get the same results as Excel would, but entirely in Python.
    Do not embed Excel formulas into the cells, but instead compute the result in Python and write the final value back to Excel.
    Make sure all code uses valid Python syntax and libraries that actually exist.
    """,


    "normal_operations": """I have the following VBA code that performs normal Excel operations (like inserting rows, copying values, deleting columns, formatting cells, renaming sheets, etc.):\n{vba_code}
    Please write equivalent Python code that:
    Performs the same operations using valid Python libraries like openpyxl or pandas.
    If the VBA modifies sheet structure (insert/delete rows/columns, rename sheets, copy data), implement the same logic using openpyxl.
    If the VBA performs value-level operations (like replacing text, copying cell values), use openpyxl or pandas appropriately.
    If any formatting is applied (e.g. bold text, cell colors, borders), replicate that formatting using openpyxl.styles.
    Do not use any unsupported or fake APIs — only use functions and methods that exist in real Python libraries.
    The final code should be fully executable and equivalent in logic to the original VBA.
    """
}

# === DB Init ===
def init_db():
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
    conn = sqlite3.connect(DB_PATH)
    best, score = None, -1.0
    for row in conn.execute("SELECT id,name,vba_macro,embedding,generated_code FROM macro_matches"):
        old = json.loads(row[3])
        sim = cosine_similarity(emb, old)
        if sim > score:
            best, score = row, sim
    conn.close()
    if best and score >= threshold:
        id_, name, vba, embstr, code = best[0], best[1], best[2], best[3], best[4]
        return {"id": id_, "name": name, "vba_macro": vba, "generated_code": code, "score": score}
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
st.title("🧠 VBA2PyGen + Titan Matching")

# Upload widget
uploaded_file = st.file_uploader("Upload Excel file (.xlsm/.xls/.xlsb)")
if not uploaded_file:
    st.stop()
file_id = uploaded_file.name

# Session flags
if "state" not in st.session_state:
    st.session_state["state"] = None
if "voted" not in st.session_state:
    st.session_state["voted"] = False
if "processed_file_id" not in st.session_state:
    st.session_state["processed_file_id"] = None

# Determine if we process
do_process = (st.session_state["processed_file_id"] != file_id) and not st.session_state["voted"]

# UI placeholders
progress = st.progress(0)

if do_process:
    # Step 1: Extract VBA
    with st.spinner("Step 1: Extracting VBA..."):
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(file_id)[1])
        tmp.write(uploaded_file.getbuffer()); tmp.flush(); tmp_path = tmp.name; tmp.close()
        vba_code = extract_vba(tmp_path)
    progress.progress(20)

    # Step 2: Embed & match
    with st.spinner("Step 2: Embedding & matching..."):
        emb = get_embedding(vba_code)
        match = find_best_match(emb)
    progress.progress(40)

    # Show code and match
    with st.expander("Extracted VBA Code"):
        st.code(vba_code, language="vb")
    if match:
        st.markdown(f"**Reference Found:** `{match['name']}` — `{match['score']*100:.1f}%`")
        with st.expander("Matched VBA Macro"):
            st.code(match['vba_macro'], language="vb")
        with st.expander("Matched Python Code"):
            st.code(match['generated_code'], language="python")

    # Step 3: Categorize
    with st.spinner("Step 3: Categorizing VBA..."):
        category = classify_vba(vba_code)
        st.markdown(f"**Detected Category:** `{category}`")
    progress.progress(60)

    # Step 4: Build Prompt
    with st.spinner("Step 4: Building prompt..."):
        prompt_text = PROMPTS[category].format(vba_code=vba_code) + "\n\nUse This Code Python Code:\n" + match['generated_code'] + f"\n\nThis code is `{match['score']*100:.1f}%` Accurate of what we want."
    with st.expander("Prompt Used"):
        st.code(prompt_text, language="text")
    progress.progress(80)

    # Step 5: Generate Python code
    with st.spinner("Step 5: Generating Python code..."):
        full = "".join(stream_claude(prompt_text))
        py_code = (full.split("```python", 1)[1].split("```", 1)[0].strip() if "```python" in full else full.strip())
    with st.expander("Generated Python Code"):
        st.code(py_code, language="python")
    progress.progress(100)

    # Save state
    st.session_state["state"] = VBAState(vba_code=vba_code, category=category, embedding=emb, match=match, py_code=py_code)
    st.session_state["processed_file_id"] = file_id

# Display existing state if not processing
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

# Voting callbacks

def upvote():
    # Insert and update feedback on upvote
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
    # Update feedback on downvote only
    if state.get('match'):
        update_feedback(state['match']['id'], -1)
    st.session_state['voted'] = True

# Render buttons and disable after vote
col1, col2 = st.columns(2)
col1.button(
    "👍 Helpful",
    on_click=upvote,
    disabled=st.session_state['voted']
)
col2.button(
    "👎 Not Helpful",
    on_click=downvote,
    disabled=st.session_state['voted']
)










TypeError: 'NoneType' object is not subscriptable
Traceback:
File "C:\Users\arpapate\Desktop\Generate_macro_prompt\.venv\Lib\site-packages\streamlit\runtime\scriptrunner\exec_code.py", line 121, in exec_func_with_error_handling
    result = func()
             ^^^^^^
File "C:\Users\arpapate\Desktop\Generate_macro_prompt\.venv\Lib\site-packages\streamlit\runtime\scriptrunner\script_runner.py", line 640, in code_to_exec
    exec(code, module.__dict__)
File "C:\Users\arpapate\Desktop\Generate_macro_prompt\test.py", line 265, in <module>
    prompt_text = PROMPTS[category].format(vba_code=vba_code) + "\n\nUse This Code Python Code:\n" + match['generated_code'] + f"\n\nThis code is `{match['score']*100:.1f}%` Accurate of what we want."
