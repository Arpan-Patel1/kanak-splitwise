state["vba_code"] = "\n\n".join(modules)
with st.expander("Step 1: Extracted VBA code"):
    st.code(state["vba_code"], language="vb")



with st.spinner("Step 4: Generating Python Code..."):
    full = "".join(stream_claude(state["final_prompt"]))
    if "```python" in full:
        s = full.find("```python") + len("```python")
        e = full.find("```", s)
        code = full[s:e].strip()
    else:
        code = full.strip()
    state["generated_code"] = code
with st.expander("Step 4: Generated Python code"):
    st.code(code, language="python")
