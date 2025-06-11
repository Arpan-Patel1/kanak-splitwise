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

# --- FILE UPLOADER + RESET ---
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsm", "xlsb", "xls"])

if uploaded_file is None:
    # If they clear the uploader, reset state and stop
    for key in ("generated_code", "base_name"):
        if key in st.session_state:
            del st.session_state[key]
    st.info("Please upload an Excel file to start conversion.")
    st.stop()

# At this point we have a valid uploaded_file
if "generated_code" not in st.session_state:
    st.session_state["generated_code"] = None

# Save and process only if we haven’t converted yet
if st.session_state["generated_code"] is None:
    file_path, base_name = save_uploaded_file(uploaded_file)
    st.session_state["base_name"] = base_name

    # Build state graph as before...
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
    compiled = graph.compile()

    initial: VBAState = {"file_path": file_path}
    for state in compiled.stream(initial):
        final = state

    result = final.get("generate_python_code", {}).get("generated_code")
    if result:
        st.session_state["generated_code"] = result
        st.success("✅ Conversion completed!")
    else:
        st.error("No generated code found. Please check the previous steps.")
else:
    st.success("✅ Conversion already completed!")
