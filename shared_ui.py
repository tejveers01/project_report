import streamlit as st


def inject_shared_ui() -> None:
    st.markdown(
        """
<style>
@import url('https://fonts.googleapis.com/css2?family=Manrope:wght@500;700;800&display=swap');

:root {
    --bg-start: #eef8ff;
    --bg-end: #d5eaff;
    --panel: rgba(255, 255, 255, 0.76);
    --panel-border: rgba(90, 136, 187, 0.18);
    --text-main: #12314f;
    --text-soft: #57728b;
    --accent: #2c7ec7;
    --accent-deep: #154f83;
    --success: #1d8f63;
    --warning: #d49116;
    --danger: #d84b5d;
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
        radial-gradient(circle at top left, rgba(255, 255, 255, 0.96), transparent 32%),
        radial-gradient(circle at top right, rgba(131, 198, 255, 0.33), transparent 30%),
        linear-gradient(180deg, var(--bg-start) 0%, var(--bg-end) 100%);
    color: var(--text-main);
}

header[data-testid="stHeader"] {
    background: transparent !important;
    height: 0 !important;
}

#MainMenu, footer, div[data-testid="stToolbar"], .stDeployButton, div[data-testid="stDecoration"] {
    visibility: hidden;
    display: none;
}

div[data-testid="stAppViewContainer"] > .main {
    padding-top: 1.8rem;
}

section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, rgba(255, 255, 255, 0.92), rgba(233, 245, 255, 0.9));
    border-right: 1px solid rgba(90, 136, 187, 0.14);
}

section[data-testid="stSidebar"] .stMarkdown,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] .stSelectbox,
section[data-testid="stSidebar"] .stTextInput {
    color: var(--text-main);
}

section[data-testid="stSidebar"] h1,
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3,
section[data-testid="stSidebar"] p {
    color: #000000 !important;
}

section[data-testid="stSidebar"] [data-testid="stRadio"] label,
section[data-testid="stSidebar"] [data-testid="stRadio"] div[role="radiogroup"] label,
section[data-testid="stSidebar"] [data-testid="stRadio"] p {
    color: var(--text-main) !important;
    background: transparent !important;
}

section[data-testid="stSidebar"] [data-testid="stRadio"] label > div {
    background: transparent !important;
}

section[data-testid="stSidebar"] [data-testid="stRadio"] label:hover {
    background: rgba(44, 126, 199, 0.08) !important;
    border-radius: 10px;
}

.app-shell, .main-container {
    max-width: 1180px;
    margin: 0 auto 2rem;
    background: rgba(255, 255, 255, 0.78);
    backdrop-filter: blur(16px);
    border: 1px solid var(--panel-border);
    border-radius: 28px;
    padding: 1.8rem;
    box-shadow: var(--shadow);
}

.hero-panel {
    background: linear-gradient(145deg, rgba(255, 255, 255, 0.96), rgba(233, 245, 255, 0.82));
    border: 1px solid rgba(90, 136, 187, 0.14);
    border-radius: 24px;
    padding: 1.7rem 1.8rem;
    margin-bottom: 1.25rem;
    box-shadow: 0 16px 34px rgba(23, 72, 116, 0.08);
}

.hero-kicker {
    display: inline-block;
    padding: 0.42rem 0.82rem;
    border-radius: 999px;
    background: rgba(44, 126, 199, 0.11);
    color: var(--accent-deep);
    font-size: 0.8rem;
    font-weight: 800;
    letter-spacing: 0.05em;
    text-transform: uppercase;
    margin-bottom: 0.95rem;
}

.hero-title, .main-title {
    margin: 0;
    color: var(--text-main) !important;
    font-size: clamp(2rem, 3.8vw, 3.2rem);
    line-height: 1.05;
    font-weight: 800;
}

.hero-subtitle, .subtitle {
    margin: 0.9rem 0 0;
    max-width: 760px;
    color: var(--text-soft) !important;
    font-size: 1rem;
    line-height: 1.7;
    font-weight: 600;
}

.section-card, .project-selection {
    background: linear-gradient(145deg, rgba(255, 255, 255, 0.98), rgba(239, 247, 255, 0.92));
    border: 1px solid rgba(90, 136, 187, 0.12);
    border-radius: 22px;
    padding: 1.35rem 1.45rem;
    margin: 1rem 0 1.25rem;
    box-shadow: 0 12px 28px rgba(23, 72, 116, 0.06);
}

.section-card h3, .project-selection h3 {
    margin: 0 0 0.55rem;
    color: var(--text-main) !important;
    font-size: 1.2rem;
    font-weight: 800;
}

.section-card p {
    margin: 0;
    color: var(--text-soft);
    line-height: 1.6;
}

.status-container, .success-container, .error-container, .system-info {
    border-radius: 22px;
    padding: 1.35rem 1.45rem;
    margin: 1rem 0 1.2rem;
    box-shadow: 0 12px 28px rgba(23, 72, 116, 0.08);
}

.status-container {
    background: linear-gradient(145deg, #fff8df, #fff0b9);
    border: 1px solid rgba(212, 145, 22, 0.3);
}

.success-container, .system-info {
    background: linear-gradient(145deg, #e6fbf2, #d3f5e7);
    border: 1px solid rgba(29, 143, 99, 0.24);
}

.error-container {
    background: linear-gradient(145deg, #ffecee, #ffd9df);
    border: 1px solid rgba(216, 75, 93, 0.24);
}

.status-container h4, .success-container h4, .error-container h4 {
    margin: 0 0 0.45rem;
    font-size: 1.15rem;
    font-weight: 800;
}

.status-container h4 { color: #9c6804; }
.success-container h4, .system-info { color: #136b49; }
.error-container h4 { color: #ac2f42; }

.status-container p, .success-container p, .error-container p {
    margin: 0;
    color: var(--text-main);
}

.chat-message {
    padding: 1.25rem;
    border-radius: 18px;
    margin-bottom: 1rem;
    display: flex;
    align-items: flex-start;
    box-shadow: 0 10px 24px rgba(23, 72, 116, 0.08);
}

.chat-message.bot {
    background: linear-gradient(145deg, #f5fbff, #e6f2ff);
    border-left: 4px solid var(--accent);
}

.chat-message.user {
    background: linear-gradient(145deg, #edf8ff, #e8f7ef);
    border-right: 4px solid var(--success);
}

.chat-message .avatar {
    width: 3rem;
    height: 3rem;
    border-radius: 50%;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 1.25rem;
    margin: 0 0.9rem;
    background: linear-gradient(145deg, var(--accent), #69a9de);
    color: white;
}

.chat-message.user .avatar {
    background: linear-gradient(145deg, var(--success), #4bbd8c);
}

.chat-message .message {
    flex: 1;
    color: var(--text-main) !important;
    font-weight: 600;
    line-height: 1.65;
}

.footer {
    text-align: center;
    color: var(--text-soft);
    margin-top: 2rem;
    padding: 1.25rem;
    border-radius: 20px;
    background: rgba(255, 255, 255, 0.5);
    border: 1px solid rgba(90, 136, 187, 0.12);
}

.stButton > button,
.stDownloadButton > button {
    background: linear-gradient(145deg, var(--accent), #5aa7eb) !important;
    color: white !important;
    border: none !important;
    border-radius: 14px !important;
    font-weight: 800 !important;
    box-shadow: 0 10px 24px rgba(44, 126, 199, 0.24) !important;
}

.stDownloadButton > button {
    min-height: 3.2rem !important;
}

.stButton > button:hover,
.stDownloadButton > button:hover {
    background: linear-gradient(145deg, #1e6dad, var(--accent)) !important;
    border: none !important;
}

.stProgress > div > div > div > div {
    background: linear-gradient(145deg, var(--accent), #5aa7eb) !important;
}

[data-testid="stDataFrame"], .stAlert, [data-testid="stExpander"] {
    border-radius: 18px;
}

hr {
    border: none;
    height: 1px;
    background: linear-gradient(90deg, transparent, rgba(44, 126, 199, 0.32), transparent);
    margin: 1.6rem 0;
}

@media (max-width: 768px) {
    .app-shell, .main-container {
        padding: 1rem;
        margin: 0 auto 1.2rem;
        border-radius: 20px;
    }

    .hero-panel {
        padding: 1.2rem;
    }
}
</style>
""",
        unsafe_allow_html=True,
    )


def render_app_header(title: str, subtitle: str, badge: str) -> None:
    st.markdown(
        f"""
<div class="hero-panel">
    <div class="hero-kicker">{badge}</div>
    <h1 class="hero-title">{title}</h1>
    <p class="hero-subtitle">{subtitle}</p>
</div>
""",
        unsafe_allow_html=True,
    )
