import os
import json
import tempfile
from typing import Optional, TypedDict

import openpyxl
import streamlit as st
import boto3
from oletools.olevba import VBA_Parser
from langgraph.graph import StateGraph, END

# =====================
# Configuration & Constants
# =====================
PROMPTS = {
    "pivot_table": (
        """
I have the following VBA code that creates a Pivot Table in Excel:
{vba_code}
Please write equivalent Python code that:
- Produces the same summarized data the pivot table would show (e.g., group by fields, aggregation like SUM, COUNT, AVERAGE).
- Uses pandas to perform the summary using pivot_table() or groupby().
- Saves the resulting table into a sheet where it is supposed to be in the same Excel file using pandas.ExcelWriter or openpyxl.
- Does not create a real Excel PivotTable, and does not use any fake or unsupported APIs like openpyxl.worksheet.table.tables.Table.
- Make sure all Python libraries used are valid and the code runs end-to-end.
"""
    ),
    # ... other prompts unchanged ...
}

# AWS Bedrock client (Anthropic Claude)
bedrock = boto3.client("bedrock-runtime")

# =====================
# Types
# =====================
class VBAState(TypedDict):
    file_path: Optional[str]
    vba_code: Optional[str]
    category: Optional[str]
    final_prompt: Optional[str]
    generated_code: Optional[str]

# =====================
# Utility Functions
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
        response = bedrock.invoke_model_with_response_stream(
            modelId=(
                "arn:aws:bedrock:us-east-1:137360334857:inference-profile/"
                "us.anthropic.claude-3-7-sonnet-20250219-v1:0"
            ),
            body=json.dumps(payload),
        )
        for event in response.get("body", []):
            chunk = json.loads(event["chunk"]["bytes"] or "{}").get("delta", {})
            if text := chunk.get("text"):
                yield text
    except Exception as err:
        st.error(f"Error streaming Claude response: {err}")
        st.stop()


def save_uploaded_file(uploaded_file) -> tuple[str, str]:
    fname = uploaded_file.name
    target_path = os.path.abspath(fname)
    with open(target_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    if target_path.lower().endswith('.xlsm'):
        wb = openpyxl.load_workbook(target_path, keep_vba=False)
        xlsx_path = os.path.splitext(target_path)[0] + '.xlsx'
        wb.save(xlsx_path)
        os.remove(target_path)
        return xlsx_path, os.path.splitext(fname)[0]
    return target_path, os.path.splitext(fname)[0]

# =====================
# Step Functions
# =====================

def extract_vba(state: VBAState) -> VBAState:
    st.subheader("Step 1: Extracting VBA code")
    try:
        parser = VBA_Parser(state["file_path"])
    except Exception as err:
        st.error(f"Failed to parse VBA macros: {err}")
        st.stop()
    modules = []
    if parser.detect_vba_macros():
        for _, _, _, code in parser.extract_macros():
            if code.strip():
                modules.append(code.strip())
    if not modules:
        st.error("No VBA macros found in the file!")
        st.stop()
    state["vba_code"] = "\n\n".join(modules)
    with st.expander("Extracted VBA Code"):
        st.text_area("Extracted VBA Code", state["vba_code"], height=300)
    progress.progress(20)
    return state


def categorize_vba(state: VBAState) -> VBAState:
    st.subheader("Step 2: Categorizing VBA code")
    prompt = (
        "Classify the following VBA code into one of these categories: "
        "formulas, pivot_table, pivot_chart, user_form, normal_operations. "
        "Only return the category name without any explanation.\n\n"
        f"{state['vba_code']}"
    )
    response = ''.join(stream_claude(prompt)).strip().lower()
    state['category'] = response if response in PROMPTS else 'normal_operations'
    st.success(f"Detected category: `{state['category']}`")
    progress.progress(40)
    return state


def build_prompt(state: VBAState) -> VBAState:
    st.subheader("Step 3: Building prompt for AI")
    state['final_prompt'] = PROMPTS[state['category']].format(vba_code=state['vba_code'])
    with st.expander("AI Prompt"):
        st.text_area("AI Prompt", state['final_prompt'], height=200)
    progress.progress(60)
    return state


def generate_python_code(state: VBAState) -> VBAState:
    st.subheader("Step 4: Generating Python code")
    full_output = ''.join(stream_claude(state['final_prompt']))
    code = full_output
    if '```python' in full_output:
        start = full_output.find('```python') + len('```python')
        end = full_output.find('```', start)
        code = full_output[start:end].strip()
    with st.expander("Generated Python Code"):
        st.text_area("Generated Python Code", code, height=300)
    progress.progress(100)
    out_path = f"{st.session_state['base_name']}.py"
    with open(out_path, 'w') as f:
        f.write(code)
    state['generated_code'] = code
    return state

# =====================
# Streamlit UI & Graph Setup
# =====================

st.set_page_config(
    page_title="VBA2PyGen",
    layout="wide",
    initial_sidebar_state="expanded",
)
st.markdown(
    """
    <style>
        body { background-color: #0e1117; color: #c7d5e0; }
        .stTextArea textarea, .stTextInput input { background-color: #1e222d; color: #c7d5e0; }
    </style>
    """,
    unsafe_allow_html=True,
)
st.title("VBA2PyGen")
st.markdown("Upload your Excel file with VBA macros, and let the AI convert them to Python step by step! 🚀")
progress = st.progress(0)

# File uploader and reset logic
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsm", "xlsb", "xls"]):
if not uploaded_file:
    # Reset state when file is removed
    for k in ["generated_code", "base_name"]:
        if k in st.session_state:
            del st.session_state[k]
    st.info("Please upload an Excel file to start conversion.")
    st.stop()

# Initialize session state
if "generated_code" not in st.session_state:
    st.session_state["generated_code"] = None

# Process upload
file_path, base_name = save_uploaded_file(uploaded_file)
st.session_state["base_name"] = base_name

# Build and run the state graph
graph = StateGraph(VBAState)
graph.add_node("extract_vba", extract_vba)
graph.add_node("categorize_vba", categorize_vba)
graph.add_node("build_prompt", build_prompt)
graph.add_node("generate_python_code", generate_python_code)
graph.set_entry_point("extract_vba")
graph.add_edge("extract_vba", "categorize_vba")
graph.add_edge("categorize_vba", "build_prompt")
graph.add_edge("build_prompt", "generate_python_code")
graph.add_edge("generate_python_code", END)
compiled_graph = graph.compile()

initial_state: VBAState = {"file_path": file_path}
for state in compiled_graph.stream(initial_state):
    final_state = state

if result := final_state.get("generate_python_code", {}).get("generated_code"):
    st.session_state["generated_code"] = result
    st.success("✅ Conversion completed!")
else:
    st.error("No generated code found. Please check the previous steps.")

elif st.session_state.get("generated_code"):
    st.success("✅ Conversion completed!")
