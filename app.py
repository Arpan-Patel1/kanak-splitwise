sqlite3.OperationalError: near "references": syntax error
Traceback:
File "C:\Users\arpapate\Desktop\Generate_macro_prompt\.venv\Lib\site-packages\streamlit\runtime\scriptrunner\exec_code.py", line 121, in exec_func_with_error_handling
    result = func()
             ^^^^^^
File "C:\Users\arpapate\Desktop\Generate_macro_prompt\.venv\Lib\site-packages\streamlit\runtime\scriptrunner\script_runner.py", line 640, in code_to_exec
    exec(code, module.__dict__)
File "C:\Users\arpapate\Desktop\Generate_macro_prompt\test.py", line 30, in <module>
    init_db()
File "C:\Users\arpapate\Desktop\Generate_macro_prompt\test.py", line 20, in init_db
    conn.execute("""
