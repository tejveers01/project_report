# checklist/cl_main.py
import os
import sys
import runpy
import streamlit as st

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))

if CURRENT_DIR not in sys.path:
    sys.path.insert(0, CURRENT_DIR)

CHECKLIST_PAGES = {
    "EDEN": "eden.py",
    "EWS Checklist": "checklistews.py",
    "Wave City Checklist": "Wave City.py",
    "Eligo Checklist": "CheckEligo.py",
    "Veridia": "veridia.py",
}

st.sidebar.title("Projects")

selected_page = st.sidebar.radio(
    "Select Checklist Project",
    list(CHECKLIST_PAGES.keys()),
    key="checklist_project_selector"
)

file_name = CHECKLIST_PAGES[selected_page]
file_path = os.path.join(CURRENT_DIR, file_name)

if not os.path.exists(file_path):
    st.error(f"File not found: {file_path}")
else:
    runpy.run_path(file_path, run_name="__main__")