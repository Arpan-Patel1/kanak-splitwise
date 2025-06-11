import os
import json
import tempfile
from typing import TypedDict, Optional
import importlib.util

import openpyxl
import pandas as pd
import streamlit as st
import boto3
from oletools.olevba import VBA_Parser
from langgraph.graph import StateGraph, END

# =====================
# Define prompts dictionary directly
# =====================
PROMPTS = {
    "pivot_table": """I have the following VBA code that creates a Pivot Table in Excel:\n{vba_code}
Please write equivalent Python code that:
- Produces the same summarized data the pivot table would show (e.g., group by fields, aggregation like SUM, COUNT, AVERAGE).
- Uses pandas to perform the summary using pivot_table() or groupby().
- Saves the resulting table into a sheet where it is supposed to be in the same Excel file using pandas.ExcelWriter or openpyxl.
- Does not create a real Excel PivotTable, and does not use any fake or unsupported APIs like openpyxl.worksheet.table.tables.Table.
- Make sure all Python libraries used are valid and the code runs end-to-end.
""",

    "pivot_chart": """I have the following VBA code that creates a Pivot Chart in Excel:\n{vba_code}
Uses pandas to perform the same data summarization (as done by the PivotTable feeding the chart).
Generates a chart that visually represents the same data, using a real Python charting library like matplotlib, seaborn, or plotly.
The chart type should match what’s used in the VBA (e.g., column chart, line chart, pie chart, etc.).
Saves the resulting chart to an image file (PNG/JPG) or embeds it into a new sheet of the same Excel workbook using openpyxl or xlsxwriter (if possible).
Avoid using any non-existent Excel chart APIs in openpyxl or other libraries.
Make sure all code is real, valid, and executable with standard Python libraries. Do not use functions like (ws.clear_rows())
""",

    "user_form": """I have the following VBA code that creates and handles a UserForm in Excel:\n{vba_code}
Please generate equivalent Python code that:
- Replicates the logic and UI flow of the UserForm.
- If the VBA code uses form fields like textboxes, dropdowns, buttons, etc., map them to similar components in a Python GUI using Tkinter or PyQt5.
- If the UserForm is used for data entry into Excel, make sure the Python version captures user input and writes it to the Excel file using pandas or openpyxl.
Do not use fake or unsupported libraries or UI frameworks like (from openpyxl.pivot.table import PivotTable).
Ensure the code uses only real, valid Python functions and libraries that exist.
The Python script should be self-contained and executable, and replicate the VBA UserForm’s functionality as closely as possible.
""",

    "formula": """I have the following VBA or Excel formula-based code:\n{vba_code}
Please generate equivalent Python code that:
- Replicates the same logic and calculations performed by the formulas.
- Uses real Python libraries like pandas, numpy, or openpyxl to evaluate the logic.
- If the formulas are row-wise, apply them using pandas.apply() or vectorized operations.
- If they reference Excel ranges, load the file using pandas.read_excel() or openpyxl, and apply the logic accordingly.
The goal is to get the same results as Excel would, but entirely in Python.
Do not embed Excel formulas into the cells, but instead compute the result in Python and write the final value back to Excel.
Make sure all code uses valid Python syntax and libraries that actually exist.
""",

    "normal_operations": """I have the following VBA code that performs normal Excel operations (like inserting rows, copying values, deleting columns, formatting cells, renaming sheets, etc.):\n{vba_code}
Please write equivalent Python code that:
- Performs the same operations using valid Python libraries like openpyxl or pandas.
- If the VBA modifies sheet structure (insert/delete rows/columns, rename sheets, copy data), implement the same logic using openpyxl.
- If the VBA performs value-level operations (like replacing text, copying cell values), use openpyxl or pandas appropriately.
- If any formatting is applied (e.g. bold text, cell colors, borders), replicate that formatting using openpyxl.styles.
Do not use any unsupported or fake APIs — only use functions and methods that exist in real Python libraries.
The final code should be fully executable and equivalent in logic to the original VBA.
"""
}

# =====================
# AWS Bedrock client
# =====================
bedrock = boto3.client("bedrock-runtime")

# =====================
# Stream Claude API (token-by-token)
# =====================
def stream_claude(prompt: str):
    try:
        payload = {
            "anthropic_version": "bedrock-2023-05-31",
            "messages": [{"role": "user", "content": prompt}],
            "max_tokens": 4000,
            "temperature": 0,
            "top_p": 0.9,
            "top_k": 250,
        }
        resp = bedrock.invoke_model_with_response_stream(
            modelId=(
                "arn:aws:bedrock:us-east-1:137360334857:"
                "inference-profile/us.anthropic.claude-3-7-sonnet-20250219-v1:0"
            ),
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

# =====================
# Type definitions
# =====================
class VBAState(TypedDict):
    file_path: Optional[str]
    vba_code: Optional[str]
    category: Optional[str]
    final_prompt: Optional[str]
    generated_code: Optional[str]

# =====================
# Utilities
# =====================
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

def verify_local_code(code: str) -> list[str]:
    errs = []
    try:
        compile(code, '<string>', 'exec')
    except Exception as e:
        errs.append(f"Syntax error: {e}")
    for line in code.splitlines():
        if line.strip().startswith(("import ", "from ")):
            mod = line.split()[1]
            root = mod.split(".")[0]
            if importlib.util.find_spec(root) is None:
                errs.append(f"Missing module: {root}")
    return errs

def fix_code_with_ai(original: str, errors: list[str]) -> str:
    prompt = (
        "The following Python code has errors: " + ", ".join(errors) +
        ". Please provide corrected code that runs without errors, uses only real libraries, "
        "and implements the intended functionality.\n```python\n"
        + original + "\n```"
    )
    acc = ""
    for chunk in stream_claude(prompt):
        acc += chunk
    # extract code block
    if "```python" in acc:
        s = acc.find("```python") + len("```python")
        e = acc.find("```", s)
        return acc[s:e].strip()
    return acc.strip()

# =====================
# Step functions
# =====================
def extract_vba(state: VBAState) -> VBAState:
    st.subheader("Step 1: Extracting VBA code")
    try:
        parser = VBA_Parser(state["file_path"])
    except Exception as e:
        st.error(f"Parse error: {e}")
        st.stop()
    modules = [code.strip() for _, _, _, code in parser.extract_macros() if code.strip()]
    if not modules:
        st.error("No VBA macros found.")
        st.stop()
    state["vba_code"] = "\n\n".join(modules)
    with st.expander("VBA Code"):
        st.text_area("VBA macros", state["vba_code"], height=300)
    progress.progress(20)
    return state

def categorize_vba(state: VBAState) -> VBAState:
    st.subheader("Step 2: Categorizing VBA code")
    prompt = (
        "Classify the following VBA code into: formulas, pivot_table, pivot_chart, user_form, normal_operations. "
        "Return only the category.\n\n" + state["vba_code"]
    )
    cat = "".join(stream_claude(prompt)).strip().lower()
    state["category"] = cat if cat in PROMPTS else "normal_operations"
    st.success(f"Category: {state['category']}")
    progress.progress(40)
    return state

def build_prompt(state: VBAState) -> VBAState:
    st.subheader("Step 3: Building AI prompt")
    state["final_prompt"] = PROMPTS[state["category"]].format(vba_code=state["vba_code"])
    with st.expander("AI Prompt"):
        st.text_area("Prompt", state["final_prompt"], height=200)
    progress.progress(60)
    return state

def generate_python_code(state: VBAState) -> VBAState:
    st.subheader("Step 4: Generating Python code")
    full = "".join(stream_claude(state["final_prompt"]))
    code = full
    if "```python" in full:
        s = full.find("```python") + len("```python")
        e = full.find("```", s)
        code = full[s:e].strip()
    with st.expander("Generated Code"):
        st.code(code, language="python")
    progress.progress(80)
    state["generated_code"] = code
    return state

def verify_code(state: VBAState) -> VBAState:
    st.subheader("Step 5: Verifying & fixing code")
    code = state.get("generated_code", "")
    errors = verify_local_code(code)
    if errors:
        for err in errors:
            st.error(err)
        fixed = fix_code_with_ai(code, errors)
        state["generated_code"] = fixed
        st.subheader("Corrected Code")
        st.code(fixed, language="python")
        errs2 = verify_local_code(fixed)
        if errs2:
            for err in errs2:
                st.error(err)
            st.stop()
        else:
            st.success("✅ Corrected code passes local checks.")
    else:
        st.success("✅ Local checks passed.")
    st.subheader("AI Cross-Verification")
    verify_prompt = (
        "You are a code reviewer. Analyze the following Python code and confirm it is valid, uses only real importable libraries, "
        "and fulfills the requested task. List any issues.\n```python\n"
        + state["generated_code"] + "\n```"
    )
    feedback = ""
    for chunk in stream_claude(verify_prompt):
        feedback += chunk
    st.markdown(feedback)
    progress.progress(100)
    return state

# =====================
# Build StateGraph
# =====================
def build_graph():
    g = StateGraph(VBAState)
    for name in ["extract_vba", "categorize_vba", "build_prompt", "generate_python_code", "verify_code"]:
        g.add_node(name, globals()[name])
    g.set_entry_point("extract_vba")
    g.add_edge("extract_vba", "categorize_vba")
    g.add_edge("categorize_vba", "build_prompt")
    g.add_edge("build_prompt", "generate_python_code")
    g.add_edge("generate_python_code", "verify_code")
    g.add_edge("verify_code", END)
    return g.compile()

# =====================
# Streamlit App
# =====================
st.set_page_config(page_title="VBA2PyGen with Auto-Fix", layout="wide")
st.markdown("""
<style>
  body {background:#0e1117; color:#c7d5e0}
  .stTextArea textarea, .stTextInput input {background:#1e222d; color:#c7d5e0}
</style>
""", unsafe_allow_html=True)
st.title("VBA2PyGen with Auto-Fix")
st.markdown("Upload your Excel (xlsm/xlsb/xls) and let AI convert & auto-correct the Python code.")
progress = st.progress(0)

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsm","xlsb","xls"])
if not uploaded_file:
    st.session_state.pop("generated_code", None)
    st.session_state.pop("base_name", None)
    st.info("Please upload a file to continue.")
    st.stop()

if "generated_code" not in st.session_state:
    st.session_state["generated_code"] = None
    st.session_state["base_name"] = None

if st.session_state["generated_code"] is None:
    suffix = os.path.splitext(uploaded_file.name)[1]
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded_file.getbuffer())
        tmp_path = tmp.name
    st.session_state["base_name"] = os.path.splitext(uploaded_file.name)[0]
    graph = build_graph()
    for state in graph.stream({"file_path": tmp_path}):
        final = state
    st.success("✅ Conversion & Auto-Fix completed!")
else:
    st.success("✅ Already processed.")
