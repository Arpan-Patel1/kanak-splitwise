AttributeError: 'NoneType' object has no attribute 'get'
Traceback:
File "C:\Users\arpapate\Desktop\Generate_macro_prompt\.venv\Lib\site-packages\streamlit\runtime\scriptrunner\exec_code.py", line 121, in exec_func_with_error_handling
    result = func()
             ^^^^^^
File "C:\Users\arpapate\Desktop\Generate_macro_prompt\.venv\Lib\site-packages\streamlit\runtime\scriptrunner\script_runner.py", line 640, in code_to_exec
    exec(code, module.__dict__)
File "C:\Users\arpapate\Desktop\Generate_macro_prompt\test.py", line 203, in <module>
    cat = "".join(stream_claude(cat_prompt)).strip().lower()
          ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
File "C:\Users\arpapate\Desktop\Generate_macro_prompt\test.py", line 141, in stream_claude
    if delta := chunk.get("delta") and delta.get("text"):
                                       ^^^^^^^^^
