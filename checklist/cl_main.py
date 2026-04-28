import streamlit as st


pages = {
    "Projects": [
        st.Page("checklist/eden.py", title="EDEN"),
        st.Page("checklist/checklistews.py", title="EWS Checklist"),
        st.Page("checklist/Wave City.py", title="Wave City Checklist"),
        st.Page("checklist/CheckEligo.py", title="Eligo Checklist"),
        st.Page("checklist/veridia.py", title="Veridia"),
    ]
}


pg = st.navigation(pages)

pg.run()

