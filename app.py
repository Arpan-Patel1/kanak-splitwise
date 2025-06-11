```python
import os
import json
import tempfile
import re
from typing import TypedDict, Optional

import openpyxl
import pandas as pd
import streamlit as st
import boto3
from oletools.olevba import VBA_Parser
from langgraph.graph import StateGraph, END

# =====================
# Define prompts dictionary
# =====================
PROMPTS = {
    "pivot_table": """I have the following VBA code that creates a Pivot Table in Excel:\n{vba_code}
Please write equivalent Python code that:
- Produces the same summarized data the pivot table would show.
- Uses pandas to perform the summary.
- Saves the resulting table into a sheet in the same Excel file using pandas.ExcelWriter.
- Uses only real, supported APIs.
Ensure the code runs end-to-end and keep the basic structure unchanged.""",
    # Add other prompt entries similarly...
    "normal_operations": """I have the following VBA code that performs normal Excel operations:\n{vba_code}
Please write equivalent Python code using openpyxl or pandas to replicate the logic.
Ensure the code runs end-to-end and keep the basic structure unchanged."""
}

# =====================
# AWS Bedrock client
# =====================
bedrock = boto3.client("bedrock-runtime")

# =====================
# Stream Claude API
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
# State type
# =====================
class VBAState(TypedDict):
    file_path: Optional[str]
    vba_code: Optional[str]
    category: Optional[str]
    final_prompt: Optional[str]
    generated_code: Optional[str]

# =====================
# Save & strip macros
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

# =====================
# Step functions
# =====================
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
            "Classify the following VBA code into: formulas, pivot_table, pivot_chart, user_form, normal_operations.
"  
            "Return only the category.

" + state["vba_code"]
        )
        cat = "".join(stream_claude(prompt)).strip().lower()
        state["category"] = cat if cat in PROMPTS else "normal_operations"
    with st.expander("Step 2: Detected Category"):
        st.markdown(f"**Category detected:** {state['category']}")
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
    # AI lists fixes
    summary_prompt = (
        "Given the VBA macros and Python code, list pointwise what fixes are needed to ensure valid imports, syntax correctness, functionality match, and replace any Excel path with '" + xlsx_file + "'. Do not change code structure.\n\n"
        + "VBA:\n" + state["vba_code"] + "\n\nPython:\n```python\n" + code + "\n```"
    )
    fixes = ""
    with st.spinner("Summarizing fixes..."):
        for chunk in stream_claude(summary_prompt): fixes += chunk
    with st.expander("Step 5: Planned Fixes"):
        st.markdown(fixes)
    # AI applies fixes
    fix_prompt = (
        "Now apply those fixes to the Python code, preserving structure and functionality, using only real libraries and replacing paths with '" + xlsx_file + "'. Respond with only the code block.\n" + fixes
    )
    acc = ""
    with st.spinner("Applying fixes..."):
        for chunk in stream_claude(fix_prompt): acc += chunk
    if "```python" in acc:
        s = acc.find("```python") + len("```python")
        e = acc.find("```", s)
        final_code = acc[s:e].strip()
    else:
        final_code = acc.strip()
    state["generated_code"] = final_code
    with st.expander("Step 6: Final Corrected Code"):
        st.code(final_code, language="python")
    # Save
    py_path = os.path.splitext(st.session_state['xlsx_path'])[0] + ".py"
    with open(py_path, "w") as f:
        f.write(final_code)
    st.markdown(f"**Saved Python at:** `{py_path}`")
    progress.progress(100)
    return state

# =====================
# Build StateGraph
# =====================

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

# =====================
# Streamlit App
# =====================

st.set_page_config(page_title="VBA2PyGen", layout="wide")
st.markdown("""
<style>
  body {background:#0e1117; color:#c7d5e0}
  .stTextArea textarea, .stTextInput input {background:#1e222d; color:#c7d5e0}
</style>
""", unsafe_allow_html=True)
st.title("VBA2PyGen with AI Auto-Fix")
st.markdown("Upload your Excel file; AI will convert, list fixes, apply them, and replace paths, all in dropdowns.")
progress = st.progress(0)

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsm","xlsb","xls"])
if not uploaded_file:
    st.session_state.pop("generated_code", None)
    st.session_state.pop("xlsx_path", None)
    st.info("Please upload a file to continue.")
    st.stop()

# Save & strip macros
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
```
