🔧 Current Tech Stack
Programming Language

Python 3.10+ (recommended; works with 3.9–3.12)

Core Framework

Streamlit (for UI / file upload / workflow)

AI & Cloud

AWS Bedrock (for embeddings + LLM code generation)

Amazon Titan Embed Text v2 (amazon.titan-embed-text-v2:0) → for vector embeddings

Anthropic Claude 3.7 Sonnet (via Bedrock inference profile) → for VBA → Python code generation

Excel & VBA Handling

oletools (olevba) → for extracting VBA macros from .xlsm, .xls, .xlsb

openpyxl → for Excel reading/writing, sheet manipulation, formatting, and creating macro-free .xlsx replicas

Database / Storage

SQLite3 (local embedded DB)

Stores VBA macros, embeddings, generated Python code, and user feedback (upvote/downvote)

Math & ML Utils

NumPy → for cosine similarity on embeddings

Other Standard Libraries

hashlib → for fingerprinting (unique hash of VBA macros)

re (regex), json, tempfile, os, datetime → for parsing, serialization, file handling

📦 Python Module Versions (recommended stable)

streamlit >= 1.32

boto3 >= 1.34

oletools >= 0.60

openpyxl >= 3.1

numpy >= 1.26

🖥️ Runtime Environment

Runs locally or on a server with Python

No external DB required (SQLite file stored alongside app)

Requires AWS credentials with Bedrock access (region: us-east-1)

⚙️ Workflow Summary

User uploads Excel file (.xlsm/.xls/.xlsb/.xlsx)

App extracts VBA macros (olevba)

Normalizes + embeds macros (Titan Embed)

Looks for best match in SQLite DB (cosine similarity)

Builds prompt and sends to Claude (Bedrock)

Saves generated Python code (.py) and macro-free Excel replica (.xlsx)

User can upvote/downvote → feedback stored in DB to improve future matches
