import os
import json
import tempfile
from typing import TypedDict, Optional

import openpyxl
import pandas as pd
import streamlit as st
import boto3
from oletools.olevba import VBA_Parser
from langgraph.graph import StateGraph, END

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
    Avoid using any non-existent Excel chart APIs in openpyxl or other libraries.
    Make sure all code is real, valid, and executable with standard Python libraries. Do not use functions like (ws.clear_rows())
    """,

    "user_form": """I have the following VBA code that creates and handles a UserForm in Excel: \n{vba_code}
    Please generate equivalent Python code that:
    Replicates the logic and UI flow of the UserForm.
    If the VBA code uses form fields like textboxes, dropdowns, buttons, etc., map them to similar components in a Python GUI using Tkinter or PyQt5.
    If the UserForm is used for data entry into Excel, make sure the Python version captures user input and writes it to the Excel file using pandas or openpyxl.
    Do not use fake or unsupported libraries or UI frameworks like (from openpyxl.pivot.table import PivotTable).
    Ensure the code uses only real, valid Python functions and libraries that exist.
    The Python script should be self-contained and executable, and replicate the VBA UserForm’s functionality as closely as possible.
    """,

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

bedrock = boto3.client("bedrock-runtime")

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

class VBAState(TypedDict):
    file_path: Optional[str]
    vba_code: Optional[str]
    category: Optional[str]
    final_prompt: Optional[str]
    generated_code: Optional[str]

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
        prompt = (
            "Classify the following VBA code into: formulas, pivot_table, pivot_chart, user_form, normal_operations."  
            "Return only the category." + state["vba_code"]
        )
        cat = "".join(stream_claude(prompt)).strip().lower()
        state["category"] = cat if cat in PROMPTS else "normal_operations"
    with st.expander("Step 2: Detected Category"):
        st.markdown(f"**Category detected:** `{state['category']}`")
    progress.progress(40)
    return state

def build_prompt(state: VBAState) -> VBAState:
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
    with st.expander("Step 4: Generated Code before fix"):
        st.code(code, language="python")
    progress.progress(80)
    return state

def verify_and_fix_code(state: VBAState) -> VBAState:
    code = state.get("generated_code", "")
    xlsx_file = os.path.basename(st.session_state['xlsx_path'])
    vba_context = state["vba_code"]

    fix_prompt = (
        "you are a Python Code validation agent."
        "You are given:\n"
        f"VBA Macros:\n{vba_context}\n\n"
        f"Generated Python Code:\n```python\n{code}\n```"
        "Only fix the python code if it is broken or does NOT match the macro logic."
        "If the code is already correct, return it as-is."
        "Do NOT optimize, reformat, remane, or alter logic that already works."
        f"Use only real libraries and replace file paths with '{xlsx_file}'.\n\n"
        "Only return corrected code - no explaination."
    )

    acc = ""
    with st.spinner("Applying direct code fixes..."):
        for chunk in stream_claude(fix_prompt):
            acc += chunk

    if "```python" in acc:
        s = acc.find("```python") + len("```python")
        e = acc.find("```", s)
        final_code = acc[s:e].strip()
    else:
        final_code = acc.strip()

    state["generated_code"] = final_code
    with st.expander("Step 5: Final Corrected Code"):
        st.code(final_code, language="python")

    py_path = os.path.splitext(st.session_state['xlsx_path'])[0] + ".py"
    with open(py_path, "w") as f:
        f.write(final_code)

    st.markdown(f"**Saved Python at:** `{py_path}`")
    progress.progress(100)
    return state

def build_graph():
    steps = [extract_vba, categorize_vba, build_prompt, generate_python_code, verify_and_fix_code]
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
    st.success("✅ Conversion & AI Auto-Fix completed!")
else:
    st.success("✅ Already processed.")
