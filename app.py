import streamlit as st
import boto3
import tempfile
import json
from typing import TypedDict, Optional
from langgraph.graph import StateGraph, END
from oletools.olevba import VBA_Parser
import pandas as pd
import os
import openpyxl

# =====================
# Define prompts dictionary directly
# =====================

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

# =====================
# AWS Bedrock client
# =====================

bedrock = boto3.client("bedrock-runtime")

# =====================
# Stream Claude API (Token by token)
# =====================

def stream_claude(prompt):
    try:
        body = {
            "anthropic_version": "bedrock-2023-05-31",
            "messages": [{"role": "user", "content": prompt}],
            "max_tokens": 4000,
            "temperature": 0,
            "top_p": 0.9,
            "top_k": 250,
        }

        response = bedrock.invoke_model_with_response_stream(
            modelId="arn:aws:bedrock:us-east-1:137360334857:inference-profile/us.anthropic.claude-3-7-sonnet-20250219-v1:0",
            body=json.dumps(body)
        )

        collected_chunks = []

        for event in response["body"]:
            chunk = json.loads(event["chunk"]["bytes"])
            if chunk and "delta" in chunk and chunk["delta"].get("text"):
                text = chunk["delta"]["text"]
                collected_chunks.append(text)
                yield text

        return "".join(collected_chunks)

    except Exception as e:
        st.error(f"Error streaming Claude response: {e}")
        st.stop()

# =====================
# LangGraph state
# =====================

class VBAState(TypedDict):
    file_path: Optional[str]
    vba_code: Optional[str]
    category: Optional[str]
    final_prompt: Optional[str]
    generated_code: Optional[str]


# Function to save uploaded file and remove macros if needed
def save_uploaded_file(uploaded_file):
    original_file_path = os.path.join(os.getcwd(), uploaded_file.name)
    
    with open(original_file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    new_file_path = os.path.splitext(original_file_path)[0] + '.xlsx'
    
    if original_file_path.endswith('.xlsm'):
        wb_macro_enabled = openpyxl.load_workbook(original_file_path, keep_vba=False)
        wb_macro_enabled.save(new_file_path)
        os.remove(original_file_path)
    else:
        new_file_path = original_file_path
    
    return new_file_path, os.path.splitext(uploaded_file.name)[0]

# Streamlit UI Setup
st.set_page_config(page_title="VBA to Python AI Converter", layout="wide", initial_sidebar_state="expanded")
st.markdown(
    """
    <style>
        body {
            background-color: #0e1117;
            color: #c7d5e0;
        }
        .stTextArea textarea {
            background-color: #1e222d;
            color: #c7d5e0;
        }
        .stTextInput input {
            background-color: #1e222d;
            color: #c7d5e0;
        }
        .css-1aumxhk {
            color: #c7d5e0;
        }
    </style>
    """,
    unsafe_allow_html=True
)

st.title("VBA2PyGen")
st.markdown("Upload your Excel file with VBA macros, and let the AI convert them to Python step by step! 🚀")

progress = st.progress(0)
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsm", "xlsb", "xls"])

# Initialize session state
if "generated_code" not in st.session_state:
    st.session_state["generated_code"] = None
if "base_name" not in st.session_state:
    st.session_state["base_name"] = None

if uploaded_file:
    new_file_path, base_name = save_uploaded_file(uploaded_file)
    st.session_state["base_name"] = base_name

# =====================
# Node 1: VBA Extraction
# =====================

def extract_vba(state: VBAState):
    st.subheader("Step 1: Extracting VBA code")

    try:
        vba_parser = VBA_Parser(state["file_path"])
    except Exception as e:
        st.error(f"Failed to parse VBA macros: {e}")
        st.stop()

    vba_modules = []

    if vba_parser.detect_vba_macros():
        for (_, _, _, vba_code) in vba_parser.extract_macros():
            if vba_code.strip():
                vba_modules.append(vba_code.strip())

    if not vba_modules:
        st.error("No VBA macros found in the file!")
        st.stop()

    full_vba_code = "\n\n".join(vba_modules)
    state["vba_code"] = full_vba_code

    with st.expander("Extracted VBA Code"):
        st.text_area("Extracted VBA Code", value=full_vba_code, height=300)
    progress.progress(20)

    return state

# =====================
# Node 2: Categorization
# =====================

def categorize_vba(state: VBAState):
    st.subheader("Step 2: Categorizing VBA code")

    categorization_prompt = (
        "Classify the following VBA code into one of these categories: "
        "formulas, pivot_table, pivot_chart, user_form, normal_operations. "
        "Only return the category name without any explanation.\n\n"
        f"{state['vba_code']}"
    )

    full_response = ""
    for chunk in stream_claude(categorization_prompt):
        full_response += chunk

    selected_category = full_response.strip().lower()
    if selected_category not in PROMPTS:
        selected_category = "normal_operations"

    st.success(f"Detected category: `{selected_category}`")
    state["category"] = selected_category
    progress.progress(40)

    return state

# =====================
# Node 3: Prompt Builder
# =====================

def build_prompt(state: VBAState):
    st.subheader("Step 3: Building prompt for AI")

    template = PROMPTS.get(state["category"])
    final_prompt = template.format(vba_code=state["vba_code"])
    state["final_prompt"] = final_prompt

    with st.expander("AI Prompt"):
        st.text_area("AI Prompt", value=final_prompt, height=200)
    progress.progress(60)

    return state

# =====================
# Node 4: Code Generator
# =====================

def generate_python_code(state: VBAState):
    st.subheader("Step 4: Generating Python code")

    generated_code = ""

    for chunk in stream_claude(state["final_prompt"]):
        generated_code += chunk

    # Extract the text between ```python and ```
    start_marker = "```python"
    end_marker = "```"
    start_index = generated_code.find(start_marker)
    end_index = generated_code.find(end_marker, start_index + len(start_marker))

    if start_index != -1 and end_index != -1:
        formatted_code = generated_code[start_index + len(start_marker):end_index].strip()
    else:
        formatted_code = generated_code.strip()

    with st.expander("Generated Python Code"):
        st.text_area("Generated Python Code", value=formatted_code, height=300)

    progress.progress(100)

    # Save the generated Python code to a file
    python_file_path = os.path.join(os.getcwd(), f"{st.session_state['base_name']}.py")
    with open(python_file_path, "w") as f:
        f.write(formatted_code)

    # Return the state with the generated code
    state["generated_code"] = formatted_code
    return state


# =====================
# Graph definition
# =====================

def build_graph():
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

    return graph.compile()


# =====================
# Run the pipeline
# =====================

if uploaded_file and st.session_state["generated_code"] is None:
    with tempfile.NamedTemporaryFile(delete=False) as tmp_file:
        tmp_file.write(uploaded_file.read())
        tmp_file_path = tmp_file.name

    initial_state = VBAState(file_path=tmp_file_path)

    graph = build_graph()

    for state in graph.stream(initial_state):
        final_state = state

    # Ensure the final state contains the generated code
    if "generated_code" in final_state["generate_python_code"]:
        st.session_state["generated_code"] = final_state["generate_python_code"]["generated_code"]
        st.success("✅ Conversion completed!")
    else:
        st.error("No generated code found. Please check the previous steps.")
elif st.session_state["generated_code"] is not None:
    st.success("✅ Conversion completed!")
