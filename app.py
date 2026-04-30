# app.py
import os
import sys
import runpy
import streamlit as st

from env_loader import load_root_env


ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
load_root_env()

PROJECTS = {
    "checklist": {
        "title": "Checklist Report",
        "icon": "📝",
        "folder": "checklist",
        "file": "cl_main.py",
        "description": "Review project checklists with a focused workflow for site-wise tracking and quick report generation.",
    },
    "milestone": {
        "title": "Milestone Report",
        "icon": "📈",
        "folder": "Milestone",
        "file": "ml_main.py",
        "description": "Generate quarterly milestone summaries with KRA and tracker-driven progress insights.",
    },
    "ncr": {
        "title": "NCR Report",
        "icon": "📊",
        "folder": "NCR",
        "file": "ncr_main.py",
        "description": "Track non-conformance issues, summarize observations, and prepare action-oriented NCR outputs.",
    },
    "overall": {
        "title": "Overall Report",
        "icon": "📑",
        "folder": "Overall",
        "file": "ol_main.py",
        "description": "Bring multiple project signals together into a broader consolidated report view for leadership review.",
    },
}

st.set_page_config(
    page_title="Reports Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Manrope:wght@500;700;800&display=swap');

:root {
    --bg-start: #eef8ff;
    --bg-end: #d5eaff;
    --panel: rgba(255, 255, 255, 0.72);
    --panel-border: rgba(90, 136, 187, 0.18);
    --text-main: #1f5f99;
    --text-soft: #57728b;
    --accent: #2c7ec7;
    --accent-deep: #154f83;
    --shadow: 0 24px 60px rgba(33, 81, 126, 0.14);
}

html, body, [class*="css"] {
    font-family: 'Manrope', sans-serif;
}

body {
    background: linear-gradient(180deg, var(--bg-start) 0%, var(--bg-end) 100%);
}

.stApp {
    background:
        radial-gradient(circle at top left, rgba(255, 255, 255, 0.95), transparent 32%),
        radial-gradient(circle at top right, rgba(131, 198, 255, 0.35), transparent 30%),
        linear-gradient(180deg, var(--bg-start) 0%, var(--bg-end) 100%);
    color: var(--text-main);
}

#MainMenu, footer {
    visibility: hidden;
}

div[data-testid="stAppViewContainer"] > .main {
    padding-top: 2rem;
}

.dashboard-shell {
    max-width: 1120px;
    margin: 0 auto;
    padding-bottom: 2rem;
}

.hero-panel {
    background: var(--panel);
    backdrop-filter: blur(16px);
    border: 1px solid var(--panel-border);
    border-radius: 28px;
    padding: 2.3rem 2.2rem 1.9rem;
    box-shadow: var(--shadow);
    margin-bottom: 1.25rem;
}

.hero-kicker {
    display: inline-block;
    padding: 0.45rem 0.85rem;
    border-radius: 999px;
    background: rgba(44, 126, 199, 0.12);
    color: var(--accent-deep);
    font-size: 0.82rem;
    font-weight: 800;
    letter-spacing: 0.05em;
    text-transform: uppercase;
    margin-bottom: 1rem;
}

.hero-title {
    font-size: clamp(2.2rem, 4vw, 3.6rem);
    line-height: 1;
    margin: 0;
    color: var(--text-main);
    font-weight: 800;
}

.hero-subtitle {
    margin: 0.95rem 0 0;
    max-width: 760px;
    font-size: 1rem;
    line-height: 1.7;
    color: var(--text-soft);
}

.section-label {
    font-size: 0.82rem;
    font-weight: 800;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: var(--accent-deep);
    margin: 0.5rem 0 1rem;
    padding-left: 0.2rem;
}

.report-link {
    text-decoration: none;
    display: block;
}

.report-card {
    position: relative;
    overflow: hidden;
    min-height: 198px;
    padding: 1.5rem;
    border-radius: 24px;
    border: 1px solid rgba(84, 114, 143, 0.16);
    background: linear-gradient(145deg, rgba(255, 255, 255, 0.96), rgba(233, 245, 255, 0.82));
    box-shadow: 0 16px 34px rgba(23, 72, 116, 0.10);
    transition: transform 0.2s ease, box-shadow 0.2s ease, border-color 0.2s ease;
    margin-bottom: 1rem;
}

.report-card:hover {
    transform: translateY(-4px);
    box-shadow: 0 22px 42px rgba(23, 72, 116, 0.16);
    border-color: rgba(44, 126, 199, 0.32);
}

.report-card::after {
    content: "";
    position: absolute;
    right: -36px;
    bottom: -42px;
    width: 136px;
    height: 136px;
    border-radius: 50%;
    background: radial-gradient(circle, rgba(44, 126, 199, 0.18), transparent 68%);
}

.report-icon {
    font-size: 2rem;
    margin-bottom: 0.9rem;
}

.report-title {
    font-size: 1.28rem;
    font-weight: 800;
    color: var(--text-main);
    margin-bottom: 0.45rem;
}

.report-copy {
    color: var(--text-soft);
    line-height: 1.6;
    font-size: 0.96rem;
    margin-bottom: 1rem;
}

.report-cta {
    display: inline-flex;
    align-items: center;
    gap: 0.45rem;
    padding: 0.5rem 0.85rem;
    border-radius: 999px;
    background: rgba(44, 126, 199, 0.10);
    color: var(--accent-deep);
    font-size: 0.88rem;
    font-weight: 800;
}
</style>
""",
    unsafe_allow_html=True,
)


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
    st.markdown(
        """
        <div class="dashboard-shell">
            <div class="hero-panel">
                <div class="hero-kicker">Project Control Center</div>
                <h1 class="hero-title">Reports Dashboard</h1>
                <p class="hero-subtitle">
                    Launch checklist, milestone, NCR, and overall reports from one clean workspace.
                    Each module opens in a new tab so you can move between teams and outputs without losing context.
                </p>
            </div>
            <div class="section-label">Available Reports</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col1, col2 = st.columns(2)

    with col1:
        for key in ("checklist", "ncr"):
            project_info = PROJECTS[key]
            st.markdown(
                f"""
                <a class="report-link" href="/?project={key}" target="_blank">
                    <div class="report-card">
                        <div class="report-icon">{project_info["icon"]}</div>
                        <div class="report-title">{project_info["title"]}</div>
                        <div class="report-copy">{project_info["description"]}</div>
                        <div class="report-cta">Open module →</div>
                    </div>
                </a>
                """,
                unsafe_allow_html=True,
            )

    with col2:
        for key in ("milestone", "overall"):
            project_info = PROJECTS[key]
            st.markdown(
                f"""
                <a class="report-link" href="/?project={key}" target="_blank">
                    <div class="report-card">
                        <div class="report-icon">{project_info["icon"]}</div>
                        <div class="report-title">{project_info["title"]}</div>
                        <div class="report-copy">{project_info["description"]}</div>
                        <div class="report-cta">Open module →</div>
                    </div>
                </a>
                """,
                unsafe_allow_html=True,
            )
