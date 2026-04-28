# app.py
import os
import sys
import runpy
import streamlit as st

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))

PROJECTS = {
    "checklist": {
        "title": "Checklist Report",
        "icon": "📝",
        "folder": "checklist",
        "file": "cl_main.py",
    },
    "milestone": {
        "title": "Milestone Report",
        "icon": "📈",
        "folder": "Milestone",
        "file": "ml_main.py",
    },
    "ncr": {
        "title": "NCR Report",
        "icon": "📊",
        "folder": "NCR",
        "file": "ncr_main.py",
    },
    "overall": {
        "title": "Overall Report",
        "icon": "📑",
        "folder": "Overall",
        "file": "ol_main.py",
    },
}

st.set_page_config(
    page_title="Reports Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Keep sidebar visible, only hide menu/footer
st.markdown("""
<style>
#MainMenu {
    visibility: hidden;
}
footer {
    visibility: hidden;
}
</style>
""", unsafe_allow_html=True)


def run_project(project_key):
    project = PROJECTS[project_key]
    app_dir = os.path.join(ROOT_DIR, project["folder"])
    app_file = os.path.join(app_dir, project["file"])

    if ROOT_DIR not in sys.path:
        sys.path.insert(0, ROOT_DIR)

    if app_dir not in sys.path:
        sys.path.insert(0, app_dir)

    if not os.path.exists(app_file):
        st.error(f"File not found: {app_file}")
        return

    runpy.run_path(app_file, run_name="__main__")


project = st.query_params.get("project")

if project in PROJECTS:
    run_project(project)

else:
    st.title("📊 Reports Dashboard")
    st.write("Open reports in new tab")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("""
        <a href="/?project=checklist" target="_blank">
            <div style="padding:20px;border:1px solid #ddd;border-radius:12px;margin-bottom:15px;">
                📝 <b>Checklist Report</b>
            </div>
        </a>
        """, unsafe_allow_html=True)

        st.markdown("""
        <a href="/?project=ncr" target="_blank">
            <div style="padding:20px;border:1px solid #ddd;border-radius:12px;">
                📊 <b>NCR Report</b>
            </div>
        </a>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown("""
        <a href="/?project=milestone" target="_blank">
            <div style="padding:20px;border:1px solid #ddd;border-radius:12px;margin-bottom:15px;">
                📈 <b>Milestone Report</b>
            </div>
        </a>
        """, unsafe_allow_html=True)

        st.markdown("""
        <a href="/?project=overall" target="_blank">
            <div style="padding:20px;border:1px solid #ddd;border-radius:12px;">
                📑 <b>Overall Report</b>
            </div>
        </a>
        """, unsafe_allow_html=True)