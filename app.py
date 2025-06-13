col1, col2 = st.columns(2)

if col1.button("👍 Save this result (Helpful)"):
    store_result(
        name=uploaded_file.name,
        macro=state["vba_code"],
        category=state["category"],
        embedding=state["embedding"],
        pycode=state["generated_code"],
        feedback=1
    )
    st.success("Stored with 👍 feedback")

if col2.button("👎 Save this result (Not Helpful)"):
    store_result(
        name=uploaded_file.name,
        macro=state["vba_code"],
        category=state["category"],
        embedding=state["embedding"],
        pycode=state["generated_code"],
        feedback=-1
    )
    st.success("Stored with 👎 feedback")
