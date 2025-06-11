import os
import json
import tempfile
import re
import ast
from typing import TypedDict, Optional
import importlib

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
    # ... include other prompts here ...
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
    errors = []
    # Syntax check
    try:
        compile(code, '<string>', 'exec')
    except Exception as e:
        errors.append(f"Syntax error: {e}")
    # Import checks via AST
    try:
        tree = ast.parse(code)
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                for alias in node.names:
                    mod = alias.name.split('.')[0]
                    try:
                        importlib.import_module(mod)
                    except ModuleNotFoundError:
                        errors.append(f"Missing module: {mod}")
            elif isinstance(node, ast.ImportFrom) and node.module:
                mod = node.module.split('.')[0]
                try:
                    importlib.import_module(mod)
                except ModuleNotFoundError:
                    errors.append(f"Missing module: {mod}")
    except Exception as e:
        errors.append(f"AST error: {e}")
    return errors

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
        "Classify the following VBA code into: formulas, pivot_table, pivot_chart, user_form, normal_operations. Return only the category.\n\n"
        + state["vba_code"]
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
    st.subheader("Step 5: Verifying, fixing & finalizing code")
    code = state.get("generated_code", "")
    errors = verify_local_code(code)
    if errors:
        for err in errors:
            st.error(err)
        fixed = fix_code_with_ai(code, errors)
        # Replace any hardcoded .xls paths with our stripped xlsx filename
        xlsx_file = os.path.basename(st.session_state["xlsx_path"])
        fixed = re.sub(r"(['\"]).+?\.xls[xm]?\\1", f"'{xlsx_file}'", fixed)
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
        code = fixed
    else:
        st.success("✅ Local checks passed.")
    # Final AI cross-verification & auto-fix
    verify_prompt = (
        "You are a code reviewer. The following Python code is final: verify it is valid, uses only real importable libraries, fulfills the requested task, " 
        "and if there are any issues, provide a fully corrected version. Respond with only the corrected code in a code block.\n```python\n"
        + code + "\n```"
    )
    acc = ""
    for chunk in stream_claude(verify_prompt):
        acc += chunk
    if "```python" in acc:
        s = acc.find("```python") + len("```python")
        e = acc.find("```", s)
        final_code = acc[s:e].strip()
    else:
        final_code = acc.strip()
    # Replace file references again
    final_code = re.sub(r"(['\"]).+?\.xls[xm]?\\1", f"'{xlsx_file}'", final_code)
    state["generated_code"] = final_code
    st.subheader("Final Corrected Code")
    st.code(final_code, language="python")
    # Save final Python alongside the xlsx
    py_path = os.path.splitext(st.session_state["xlsx_path"])[0] + ".py"
    with open(py_path, "w") as f:
        f.write(final_code)
    st.markdown(f"**Macro-stripped copy at:** `{st.session_state['xlsx_path']}`  \n**Final Python at:** `{py_path}`")
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
st.markdown("Upload your Excel (xlsm/xlsb/xls) and let AI convert & auto-correct the Python code. The final code and stripped workbook will be saved.")

progress = st.progress(0)
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsm", "xlsb", "xls"])
if not uploaded_file:
    st.session_state.pop("generated_code", None)
    st.session_state.pop("xlsx_path", None)
    st.info("Please upload a file to continue.")
    st.stop()

# Save and strip macros to .xlsx
xlsx_path, base_name = save_uploaded_file(uploaded_file)
st.session_state["xlsx_path"] = xlsx_path
st.markdown(f"**Macro-stripped copy:** `{xlsx_path}`")

if "generated_code" not in st.session_state:
    st.session_state["generated_code"] = None

if st.session_state["generated_code"] is None:
    # Preserve raw .xlsm for parsing
    suffix = os.path.splitext(uploaded_file.name)[1]
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded_file.getbuffer())
        tmp_path = tmp.name

    graph = build_graph()
    for state in graph.stream({"file_path": tmp_path}):
        final = state

    st.success("✅ Conversion & Auto-Fix completed!")
else:
    st.success("✅ Already processed.")
