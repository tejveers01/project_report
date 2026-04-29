from pathlib import Path

import streamlit as st

CURRENT_DIR = Path(__file__).resolve().parent

pages = {
    "Projects": [
        st.Page(str(CURRENT_DIR / "eden.py"), title="EDEN"),
        st.Page(str(CURRENT_DIR / "checklistews.py"), title="EWS Checklist"),
        st.Page(str(CURRENT_DIR / "Wave City.py"), title="Wave City Checklist"),
        st.Page(str(CURRENT_DIR / "CheckEligo.py"), title="Eligo Checklist"),
        st.Page(str(CURRENT_DIR / "veridia.py"), title="Veridia"),
    ]
}

pg = st.navigation(pages)
pg.run()
