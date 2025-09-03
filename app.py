🔧 Current Tech Stack
Programming Language

Python 3.10+ (recommended; works with 3.9–3.12)

Core Framework

Streamlit (for UI, file upload, step-by-step workflow)

LangGraph (StateGraph, END) → for orchestrating the workflow steps

AI & Cloud

AWS Bedrock (for LLM code generation)

Anthropic Claude 3.7 Sonnet (via Bedrock inference profile) → for VBA → Python code generation and categorization

Excel & VBA Handling

oletools (olevba) → for extracting VBA macros from .xlsm, .xls, .xlsb

openpyxl → for Excel reading/writing and creating macro-free .xlsx replicas

pandas → used for data handling in generated Python code (e.g., formulas, pivot logic)

Database / Storage

None in this version (all outputs saved as local files: .xlsx and .py)

Math & Utility Libraries

json, os, tempfile, typing → file handling, serialization, type hints

📦 Python Module Versions (recommended stable)

streamlit >= 1.32

boto3 >= 1.34

openpyxl >= 3.1

pandas >= 2.2

oletools >= 0.60

langgraph >= 0.0.23

🖥️ Runtime Environment

Runs locally or on a server with Python

Requires AWS credentials with Bedrock access (region: us-east-1)

No external DB required (outputs stored in the working directory)

⚙️ Workflow Summary

User uploads Excel file (.xlsm, .xlsb, .xls)

App saves a macro-free .xlsx copy (via openpyxl)

VBA macros are extracted (olevba)

VBA code categorized (Claude via Bedrock) → formulas, pivot_table, pivot_chart, user_form, normal_operations

Prompt is built from category templates (PROMPTS)

Claude (Bedrock) generates equivalent Python code → streamed into app

Generated Python code is saved as a .py file (same name as workbook)

Results shown step by step in Streamlit expanders + progress bar
