from html import escape
from itertools import combinations
from pathlib import Path
from textwrap import dedent

import pandas as pd
import streamlit as st

try:
    import plotly.express as px
    import plotly.graph_objects as go

    HAS_PLOTLY = True
except ModuleNotFoundError:
    HAS_PLOTLY = False

from novaclose_analysis import (
    AREA_LABELS,
    build_scenario_actions,
    calculate_kpis,
    calculate_readiness_score,
    generate_commentary,
    generate_copilot_response,
    load_data,
    priority_engine,
    simulate_close_scenario,
)

st.set_page_config(
    page_title="NovaClose",
    layout="wide",
    initial_sidebar_state="expanded",
)

DEFAULT_DATASET = "AI_Finance_Hackathon_Month_End_Dataset.xlsx"
AREA_ORDER = ["gl", "bank", "intercompany", "accruals", "ap", "checklist", "tb", "fa"]
DETAIL_MAP = {
    "gl": "pending_journals",
    "bank": "bank_exceptions",
    "intercompany": "ic_exceptions",
    "accruals": "accrual_actions",
    "ap": "ap_exceptions",
    "tb": "tb_variances",
    "fa": "fa_risks",
    "checklist": "checklist_bottlenecks",
}
QUICK_PROMPTS = [
    "What issues will delay the close?",
    "What should the GL approval agent do next?",
    "What can auto-post to ERP?",
    "What should the bank agent do next?",
    "What should the IC agent do next?",
    "What should the journal agent post?",
    "What gets us below 4 days?",
    "What checklist tasks need to be unblocked?",
    "What does audit need to review before posting?",
    "What should the controller do today?",
    "Which entity is riskiest?",
    "Explain MoM / YoY variances.",
]


@st.cache_data(show_spinner=False)
def get_data(path):
    return load_data(path)


def prepare(path):
    data = get_data(path)
    kpis = calculate_kpis(data)
    score = calculate_readiness_score(kpis)
    priorities = priority_engine(kpis)
    commentary = generate_commentary(kpis, score)
    return data, kpis, score, priorities, commentary


def safe_rerun():
    rerun_fn = getattr(st, "rerun", None)
    if callable(rerun_fn):
        rerun_fn()
        return
    rerun_fn = getattr(st, "experimental_rerun", None)
    if callable(rerun_fn):
        rerun_fn()


def apply_theme():
    st.markdown(
        """
        <style>
        :root {
            --ink: #0f1f2c;
            --muted: #5b6b78;
            --surface: #ffffff;
            --surface-soft: rgba(255, 255, 255, 0.82);
            --line: rgba(15, 31, 44, 0.10);
            --brand: #143642;
            --brand-2: #256d85;
            --accent: #d77a2d;
            --good: #2a9d8f;
            --warn: #e9a03b;
            --risk: #c75c5c;
        }

        [data-testid="stAppViewContainer"] {
            background:
                radial-gradient(circle at top right, rgba(215, 122, 45, 0.12), transparent 30%),
                radial-gradient(circle at top left, rgba(37, 109, 133, 0.10), transparent 28%),
                linear-gradient(180deg, #eef3f6 0%, #f8fbfc 100%);
        }

        header[data-testid="stHeader"],
        [data-testid="stToolbar"],
        [data-testid="stDecoration"],
        .stAppDeployButton,
        #MainMenu {
            display: none !important;
        }

        [data-testid="stSidebar"] {
            background: linear-gradient(180deg, #143642 0%, #193f4d 100%);
            color: #f5f7fa;
        }

        [data-testid="stSidebar"] * {
            color: #f5f7fa;
        }

        [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] {
            background: rgba(248, 251, 252, 0.97);
            border: 1px solid rgba(20, 54, 66, 0.16);
        }

        [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] * {
            color: #143642 !important;
        }

        [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] small {
            color: #5b6b78 !important;
        }

        [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] svg {
            fill: #256d85 !important;
        }

        [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button {
            background: #ffffff !important;
            color: #143642 !important;
            border: 1px solid rgba(20, 54, 66, 0.18) !important;
        }

        [data-testid="stSidebar"] [data-testid="stFileUploaderDropzone"] button:hover {
            border-color: #256d85 !important;
            color: #0f1f2c !important;
        }

        .block-container {
            padding-top: 0.9rem;
            padding-bottom: 3rem;
            max-width: 1420px;
        }

        [data-testid="stTabs"] {
            position: relative;
            z-index: 5;
        }

        [data-testid="stTabs"] [role="tablist"] {
            gap: 0.25rem;
            padding: 0.32rem;
            margin-bottom: 1rem;
            border: 1px solid rgba(20, 54, 66, 0.12);
            border-radius: 22px;
            background: rgba(255, 255, 255, 0.74);
            box-shadow: 0 16px 40px rgba(15, 31, 44, 0.06);
            backdrop-filter: blur(10px);
            flex-wrap: nowrap;
            justify-content: space-between;
        }

        [data-testid="stTabs"] [data-baseweb="tab-highlight"] {
            background: transparent !important;
        }

        [data-testid="stTabs"] [role="tab"] {
            height: auto;
            flex: 1 1 0;
            min-width: 0;
            padding: 0.56rem 0.62rem;
            border-radius: 14px;
            border: none !important;
            background: transparent;
            color: var(--muted) !important;
            font-weight: 700;
            transition: background 0.18s ease, color 0.18s ease, box-shadow 0.18s ease;
        }

        [data-testid="stTabs"] [role="tab"] p {
            margin: 0;
            font-size: 0.87rem;
            font-weight: 700;
            color: inherit !important;
            white-space: nowrap;
        }

        [data-testid="stTabs"] [role="tab"]:hover {
            background: rgba(20, 54, 66, 0.07);
            color: var(--brand) !important;
        }

        [data-testid="stTabs"] [role="tab"][aria-selected="true"] {
            background: linear-gradient(135deg, #143642 0%, #256d85 100%);
            color: #f8fbfc !important;
            box-shadow: 0 12px 24px rgba(20, 54, 66, 0.22);
        }


        .hero {
            background: linear-gradient(135deg, rgba(20, 54, 66, 0.97) 0%, rgba(37, 109, 133, 0.95) 58%, rgba(236, 242, 246, 0.95) 58%);
            border: 1px solid rgba(20, 54, 66, 0.12);
            border-radius: 28px;
            padding: 2rem 2.1rem;
            margin-bottom: 1.2rem;
            box-shadow: 0 24px 60px rgba(15, 31, 44, 0.10);
        }

        .hero-eyebrow {
            display: inline-flex;
            align-items: center;
            gap: 0.45rem;
            padding: 0.3rem 0.7rem;
            border-radius: 999px;
            background: rgba(255, 255, 255, 0.14);
            color: #f8fbfc;
            font-size: 0.78rem;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            font-weight: 700;
        }

        .hero-grid {
            display: grid;
            grid-template-columns: 1.4fr 0.8fr;
            gap: 1.2rem;
            margin-top: 1rem;
        }

        .hero-title {
            color: #f8fbfc;
            font-size: 2.35rem;
            line-height: 1.05;
            font-weight: 800;
            margin: 0 0 0.8rem 0;
        }

        .hero-copy {
            max-width: 56rem;
            color: rgba(248, 251, 252, 0.88);
            font-size: 1rem;
            line-height: 1.5;
            margin: 0;
        }

        .hero-side {
            background: rgba(248, 251, 252, 0.92);
            border-radius: 22px;
            padding: 1rem 1.1rem;
            color: var(--ink);
        }

        .hero-side-label {
            text-transform: uppercase;
            letter-spacing: 0.08em;
            font-size: 0.72rem;
            font-weight: 700;
            color: var(--muted);
        }

        .hero-side-value {
            margin-top: 0.35rem;
            font-size: 1.8rem;
            font-weight: 800;
            color: var(--brand);
        }

        .app-header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 1rem;
            padding: 1rem 1.15rem;
            margin-bottom: 1rem;
            border-radius: 22px;
            background: linear-gradient(135deg, rgba(255, 255, 255, 0.92) 0%, rgba(244, 249, 252, 0.96) 100%);
            border: 1px solid rgba(15, 31, 44, 0.08);
            box-shadow: 0 14px 34px rgba(15, 31, 44, 0.06);
        }

        .app-header-kicker {
            font-size: 0.72rem;
            text-transform: uppercase;
            letter-spacing: 0.10em;
            font-weight: 700;
            color: var(--brand-2);
            margin-bottom: 0.2rem;
        }

        .app-header-title {
            margin: 0;
            color: var(--ink);
            font-size: 1.28rem;
            font-weight: 800;
            line-height: 1.2;
        }

        .app-header-copy {
            margin-top: 0.2rem;
            color: var(--muted);
            font-size: 0.9rem;
            line-height: 1.4;
        }

        .app-header-right {
            display: flex;
            align-items: center;
            justify-content: flex-end;
            gap: 0.65rem;
            flex-wrap: wrap;
        }

        .period-pill {
            display: inline-flex;
            align-items: center;
            gap: 0.35rem;
            padding: 0.45rem 0.8rem;
            border-radius: 999px;
            border: 1px solid rgba(37, 109, 133, 0.18);
            background: rgba(37, 109, 133, 0.10);
            color: var(--brand);
            font-size: 0.78rem;
            font-weight: 800;
            letter-spacing: 0.04em;
            text-transform: uppercase;
            white-space: nowrap;
        }

        .status-pill {
            display: inline-flex;
            align-items: center;
            gap: 0.35rem;
            padding: 0.4rem 0.75rem;
            border-radius: 999px;
            border: 1px solid rgba(42, 157, 143, 0.25);
            background: rgba(42, 157, 143, 0.12);
            color: #1f7f75;
            font-size: 0.72rem;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.06em;
        }

        .status-dot {
            width: 7px;
            height: 7px;
            border-radius: 50%;
            background: #2a9d8f;
            display: inline-block;
        }

        .app-header .stButton > button {
            width: 100%;
            padding: 0.5rem 0.85rem;
            border-radius: 12px;
            border: 1px solid rgba(20, 54, 66, 0.2);
            background: linear-gradient(135deg, #143642 0%, #256d85 100%);
            color: #f8fbfc;
            font-weight: 800;
            font-size: 0.82rem;
            white-space: nowrap;
        }

        .hero-side-copy {
            margin-top: 0.55rem;
            color: var(--muted);
            line-height: 1.45;
        }

        .kpi-strip {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
            gap: 0.8rem;
            margin-bottom: 1.1rem;
        }

        .kpi-tile {
            background: rgba(255, 255, 255, 0.92);
            border-radius: 18px;
            border: 1px solid rgba(15, 31, 44, 0.08);
            padding: 0.75rem 0.9rem;
            box-shadow: 0 10px 26px rgba(15, 31, 44, 0.05);
        }

        .kpi-tile-label {
            text-transform: uppercase;
            letter-spacing: 0.08em;
            font-size: 0.68rem;
            font-weight: 700;
            color: var(--muted);
        }

        .kpi-tile-value {
            font-size: 1.45rem;
            font-weight: 800;
            margin: 0.3rem 0 0.2rem 0;
            color: var(--brand);
        }

        .kpi-tile-copy {
            font-size: 0.78rem;
            color: var(--muted);
        }

        .intel-band {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 1rem;
            margin-bottom: 1.25rem;
        }

        .intel-card {
            background: rgba(255, 255, 255, 0.94);
            border-radius: 20px;
            border: 1px solid rgba(20, 54, 66, 0.12);
            padding: 1rem 1.05rem;
            box-shadow: 0 12px 30px rgba(15, 31, 44, 0.06);
        }

        .intel-title {
            font-weight: 800;
            color: var(--ink);
            margin-bottom: 0.35rem;
        }

        .intel-copy {
            font-size: 0.9rem;
            color: var(--muted);
            line-height: 1.45;
        }

        .timeline {
            background: rgba(255, 255, 255, 0.94);
            border-radius: 22px;
            border: 1px solid rgba(15, 31, 44, 0.08);
            padding: 1rem 1.1rem;
            margin-bottom: 1.4rem;
            box-shadow: 0 12px 30px rgba(15, 31, 44, 0.05);
        }

        .timeline-title {
            font-weight: 800;
            color: var(--ink);
            margin-bottom: 0.6rem;
        }

        .timeline-track {
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 0.6rem;
        }

        .timeline-step {
            background: rgba(20, 54, 66, 0.08);
            border-radius: 14px;
            padding: 0.6rem;
            text-align: center;
            font-size: 0.78rem;
            color: var(--muted);
        }

        .timeline-step.active {
            background: rgba(20, 54, 66, 0.15);
            color: var(--brand);
            font-weight: 800;
        }

        .opportunity-card {
            background: rgba(20, 54, 66, 0.08);
            border: 1px solid rgba(20, 54, 66, 0.12);
            border-radius: 18px;
            padding: 0.85rem 1rem;
            margin-top: 1rem;
            color: var(--ink);
        }

        .opportunity-title {
            font-weight: 800;
            margin-bottom: 0.25rem;
        }

        .focus-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
            gap: 0.9rem;
            margin: 1rem 0 1.3rem 0;
        }

        .focus-card {
            background: rgba(255, 255, 255, 0.95);
            border-radius: 20px;
            border: 1px solid rgba(20, 54, 66, 0.12);
            padding: 1rem 1.1rem;
            box-shadow: 0 12px 28px rgba(15, 31, 44, 0.06);
        }

        .focus-title {
            font-weight: 800;
            color: var(--ink);
            margin-bottom: 0.5rem;
        }

        .focus-list {
            display: grid;
            gap: 0.45rem;
            font-size: 0.9rem;
            color: var(--muted);
        }

        .focus-item strong {
            color: var(--ink);
        }

        .focus-chip {
            display: inline-flex;
            align-items: center;
            gap: 0.35rem;
            padding: 0.25rem 0.6rem;
            border-radius: 999px;
            border: 1px solid rgba(20, 54, 66, 0.12);
            background: rgba(20, 54, 66, 0.08);
            font-size: 0.72rem;
            font-weight: 700;
            color: var(--brand);
            margin-top: 0.4rem;
        }

        .tracker-table-wrap {
            max-height: 720px;
            overflow: auto;
        }

        .tracker-table thead th {
            position: sticky;
            top: 0;
            background: rgba(248, 251, 252, 0.95);
            z-index: 2;
        }

        .tracker-table tbody tr:nth-child(even) {
            background: rgba(15, 31, 44, 0.03);
        }

        .section-label {
            color: var(--muted);
            text-transform: uppercase;
            letter-spacing: 0.08em;
            font-size: 0.74rem;
            font-weight: 700;
            margin-bottom: 0.6rem;
        }

        .tight-top {
            margin-top: 0.45rem !important;
        }

        .section-label.tight-top {
            margin-bottom: 0.35rem;
        }

        .metric-card, .area-card, .priority-card, .entity-callout, .copilot-shell {
            background: var(--surface-soft);
            border: 1px solid var(--line);
            border-radius: 24px;
            box-shadow: 0 16px 44px rgba(15, 31, 44, 0.06);
            backdrop-filter: blur(8px);
        }

        .metric-card {
            padding: 1rem 1.05rem;
            min-height: 152px;
        }

        .metric-label {
            color: var(--muted);
            font-size: 0.8rem;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            font-weight: 700;
            margin-bottom: 0.6rem;
        }

        .metric-value {
            color: var(--ink);
            font-size: 2rem;
            font-weight: 800;
            line-height: 1;
            margin-bottom: 0.5rem;
        }

        .metric-subtitle {
            color: var(--muted);
            line-height: 1.45;
            font-size: 0.92rem;
        }

        .metric-card.high .metric-value,
        .priority-card.high .priority-score {
            color: var(--risk);
        }

        .metric-card.medium .metric-value {
            color: var(--accent);
        }

        .metric-card.low .metric-value {
            color: var(--good);
        }

        .priority-card.medium .priority-score {
            color: var(--accent);
        }

        .priority-card.low .priority-score {
            color: var(--brand-2);
        }

        .chip-row {
            display: flex;
            flex-wrap: wrap;
            gap: 0.45rem;
            margin-top: 0.75rem;
            margin-bottom: 0.1rem;
            width: 100%;
            align-items: flex-start;
        }

        .chip {
            display: inline-flex;
            align-items: center;
            border-radius: 999px;
            background: rgba(20, 54, 66, 0.08);
            border: 1px solid rgba(20, 54, 66, 0.10);
            color: var(--brand);
            font-size: 0.74rem;
            font-weight: 700;
            line-height: 1.2;
            padding: 0.32rem 0.58rem;
            max-width: 100%;
            white-space: normal;
        }

        .area-card {
            padding: 1rem 1rem 0.95rem 1rem;
            min-height: 250px;
        }

        .area-header, .priority-top {
            display: flex;
            justify-content: space-between;
            align-items: baseline;
            gap: 1rem;
        }

        .area-title {
            color: var(--ink);
            font-size: 1rem;
            font-weight: 800;
        }

        .area-score {
            color: var(--brand);
            font-weight: 800;
            font-size: 1.5rem;
        }

        .area-headline {
            color: var(--ink);
            font-size: 1.1rem;
            font-weight: 750;
            margin-top: 0.8rem;
            line-height: 1.3;
        }

        .area-subheadline {
            color: var(--muted);
            line-height: 1.4;
            margin-top: 0.55rem;
            min-height: 2.7rem;
        }

        .area-insight {
            color: var(--brand);
            font-size: 0.9rem;
            line-height: 1.45;
            margin-top: 0.75rem;
        }

        .priority-card {
            padding: 1rem 1.05rem;
            min-height: 220px;
        }

        .priority-area {
            color: var(--muted);
            font-size: 0.78rem;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            font-weight: 700;
        }

        .priority-title {
            color: var(--ink);
            font-size: 1.06rem;
            font-weight: 800;
            line-height: 1.35;
            margin-top: 0.55rem;
        }

        .priority-score {
            font-size: 1.3rem;
            font-weight: 800;
            color: var(--brand);
        }

        .priority-copy {
            color: var(--muted);
            line-height: 1.45;
            margin-top: 0.6rem;
            font-size: 0.92rem;
        }

        .entity-callout {
            padding: 1rem 1.05rem;
        }

        .entity-name {
            font-size: 1.25rem;
            font-weight: 800;
            color: var(--ink);
        }

        .entity-copy {
            color: var(--muted);
            line-height: 1.5;
            margin-top: 0.45rem;
        }

        .copilot-shell {
            padding: 0.85rem 0.95rem;
            margin-bottom: 0.25rem;
        }

        .copilot-shell.tight-top {
            margin-top: 0.45rem;
        }

        .copilot-title {
            color: var(--ink);
            font-size: 1.1rem;
            font-weight: 800;
            margin-bottom: 0.1rem;
        }

        .copilot-copy {
            color: var(--muted);
            line-height: 1.45;
            margin-top: 0.2rem;
            margin-bottom: 0.15rem;
        }

        .small-note {
            color: var(--muted);
            font-size: 0.84rem;
        }

        div[data-testid="stButton"] > button {
            min-height: 0;
            padding: 0.72rem 0.85rem;
            border-radius: 18px;
            line-height: 1.35;
            font-weight: 600;
        }

        div[data-testid="stChatMessage"] {
            background: rgba(255, 255, 255, 0.62);
            border: 1px solid rgba(15, 31, 44, 0.08);
            border-radius: 18px;
            padding: 0.75rem 0.85rem;
            margin-bottom: 0.55rem;
            overflow: hidden;
        }

        div[data-testid="stChatMessageContent"] {
            flex: 1 1 auto;
            width: auto;
            min-width: 0;
        }

        div[data-testid="stChatMessage"] [data-testid="stMarkdownContainer"] {
            margin-bottom: 0 !important;
        }

        .chat-answer {
            color: var(--ink);
            line-height: 1.65;
        }

        .chat-answer p {
            margin-bottom: 0.55rem;
        }

        .tracker-board {
            background: linear-gradient(180deg, rgba(255, 255, 255, 0.94) 0%, rgba(245, 249, 252, 0.96) 100%);
            border: 1px solid rgba(15, 31, 44, 0.10);
            border-radius: 24px;
            box-shadow: 0 18px 42px rgba(15, 31, 44, 0.08);
            overflow: hidden;
            margin-top: 0.35rem;
            margin-bottom: 1.2rem;
        }

        .tracker-header {
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            gap: 1rem;
            padding: 1.1rem 1.2rem 0.95rem 1.2rem;
            border-bottom: 1px solid rgba(15, 31, 44, 0.08);
            background: linear-gradient(180deg, rgba(20, 54, 66, 0.05) 0%, rgba(37, 109, 133, 0.03) 100%);
        }

        .tracker-kicker {
            font-size: 0.74rem;
            text-transform: uppercase;
            letter-spacing: 0.12em;
            font-weight: 700;
            color: var(--brand-2);
            margin-bottom: 0.45rem;
        }

        .tracker-title {
            color: var(--ink);
            font-size: 1.18rem;
            font-weight: 800;
            margin: 0;
        }

        .tracker-copy {
            margin-top: 0.35rem;
            color: var(--muted);
            font-size: 0.94rem;
            line-height: 1.45;
            max-width: 52rem;
        }

        .tracker-badges {
            display: flex;
            gap: 0.45rem;
            flex-wrap: wrap;
            justify-content: flex-end;
        }

        .tracker-badge {
            display: inline-flex;
            align-items: center;
            gap: 0.3rem;
            padding: 0.32rem 0.72rem;
            border-radius: 999px;
            font-size: 0.74rem;
            line-height: 1;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.06em;
            border: 1px solid transparent;
            white-space: nowrap;
        }

        .tracker-badge.done {
            color: #1b7f70;
            background: rgba(42, 157, 143, 0.12);
            border-color: rgba(42, 157, 143, 0.18);
        }

        .tracker-badge.active {
            color: #256d85;
            background: rgba(37, 109, 133, 0.12);
            border-color: rgba(37, 109, 133, 0.16);
        }

        .tracker-badge.blocked {
            color: #b34f5a;
            background: rgba(199, 92, 92, 0.12);
            border-color: rgba(199, 92, 92, 0.18);
        }

        .tracker-badge.waiting {
            color: #b57a1b;
            background: rgba(233, 160, 59, 0.14);
            border-color: rgba(233, 160, 59, 0.18);
        }

        .tracker-table-wrap {
            overflow-x: auto;
        }

        .tracker-table {
            width: 100%;
            border-collapse: collapse;
            min-width: 1040px;
        }

        .tracker-table thead th {
            text-align: left;
            padding: 0.78rem 0.7rem;
            font-size: 0.74rem;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            color: var(--muted);
            border-bottom: 1px solid rgba(15, 31, 44, 0.10);
            background: rgba(20, 54, 66, 0.04);
        }

        .tracker-table tbody td {
            padding: 0.82rem 0.7rem;
            border-bottom: 1px solid rgba(15, 31, 44, 0.08);
            color: var(--ink);
            font-size: 0.92rem;
            vertical-align: top;
        }

        .tracker-table tbody tr:nth-child(even) {
            background: rgba(20, 54, 66, 0.02);
        }

        .tracker-table tbody tr:hover {
            background: rgba(37, 109, 133, 0.06);
        }

        .tracker-step {
            color: var(--brand-2);
            font-weight: 700;
        }

        .tracker-task {
            color: var(--ink);
            font-weight: 700;
            line-height: 1.35;
            margin-bottom: 0.22rem;
        }

        .tracker-subtask {
            color: var(--muted);
            font-size: 0.8rem;
            line-height: 1.35;
        }

        .tracker-owner,
        .tracker-date,
        .tracker-hours {
            color: var(--ink);
            white-space: nowrap;
        }

        .tracker-muted {
            color: var(--muted);
            font-size: 0.8rem;
            display: block;
            margin-top: 0.2rem;
        }

        .status-pill,
        .auto-pill {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            padding: 0.28rem 0.64rem;
            border-radius: 999px;
            font-size: 0.72rem;
            line-height: 1;
            font-weight: 800;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            border: 1px solid transparent;
            white-space: nowrap;
        }

        .status-pill.completed {
            color: #1b7f70;
            background: rgba(42, 157, 143, 0.12);
            border-color: rgba(42, 157, 143, 0.18);
        }

        .status-pill.in-progress {
            color: #256d85;
            background: rgba(37, 109, 133, 0.12);
            border-color: rgba(37, 109, 133, 0.18);
        }

        .status-pill.not-started {
            color: #6a7a88;
            background: rgba(91, 107, 120, 0.10);
            border-color: rgba(91, 107, 120, 0.14);
        }

        .status-pill.waiting-on-input {
            color: #b57a1b;
            background: rgba(233, 160, 59, 0.14);
            border-color: rgba(233, 160, 59, 0.20);
        }

        .status-pill.blocked {
            color: #b34f5a;
            background: rgba(199, 92, 92, 0.12);
            border-color: rgba(199, 92, 92, 0.18);
        }

        .auto-pill.auto,
        .auto-pill.already-automated {
            color: #1b7f70;
            background: rgba(42, 157, 143, 0.12);
            border-color: rgba(42, 157, 143, 0.18);
        }

        .auto-pill.rule,
        .auto-pill.high-rule-based {
            color: #256d85;
            background: rgba(37, 109, 133, 0.12);
            border-color: rgba(37, 109, 133, 0.18);
        }

        .auto-pill.ai,
        .auto-pill.high-ai-candidate {
            color: #2f5f8d;
            background: rgba(47, 95, 141, 0.12);
            border-color: rgba(47, 95, 141, 0.18);
        }

        .auto-pill.medium {
            color: #5b6b78;
            background: rgba(91, 107, 120, 0.10);
            border-color: rgba(91, 107, 120, 0.14);
        }

        .auto-pill.low-judgment-required,
        .auto-pill.low {
            color: #6a7a88;
            background: rgba(91, 107, 120, 0.08);
            border-color: rgba(91, 107, 120, 0.12);
        }

        .tracker-notes {
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 0.8rem;
            margin-top: 1rem;
        }

        .tracker-note {
            background: rgba(255, 255, 255, 0.92);
            border: 1px solid rgba(15, 31, 44, 0.08);
            border-radius: 18px;
            padding: 0.95rem 1rem;
            box-shadow: 0 12px 28px rgba(15, 31, 44, 0.06);
        }

        .tracker-note-top {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 0.75rem;
            margin-bottom: 0.45rem;
        }

        .tracker-note-title {
            color: var(--ink);
            font-size: 0.95rem;
            font-weight: 800;
            line-height: 1.35;
        }

        .tracker-note-copy {
            color: var(--muted);
            font-size: 0.84rem;
            line-height: 1.45;
        }

        @media (max-width: 1200px) {
            .tracker-notes {
                grid-template-columns: 1fr;
            }
        }

        @media (max-width: 1200px) {
            .hero-grid {
                grid-template-columns: 1fr;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def tone_from_score(score):
    if score >= 80:
        return "low"
    if score >= 65:
        return "medium"
    return "high"


def priority_tone(priority_score):
    if priority_score >= 9:
        return "high"
    if priority_score >= 7:
        return "medium"
    return "low"


def format_chip_row(values):
    chips = "".join(f"<span class='chip'>{value}</span>" for value in values)
    return f"<div class='chip-row'>{chips}</div>"


def format_chat_answer(answer, source_metrics):
    formatted_answer = answer.replace("\n", "<br>")
    return (
        f"<div class='chat-answer'>{formatted_answer}</div>"
        f"{format_chip_row(source_metrics)}"
    )


def slugify_label(value):
    text = str(value).strip().lower()
    for old, new in {
        " – ": "-",
        " / ": "-",
        "/": "-",
        "&": "and",
        " ": "-",
        "_": "-",
    }.items():
        text = text.replace(old, new)
    return "".join(ch for ch in text if ch.isalnum() or ch == "-").strip("-")


def format_status_pill(value):
    label = escape(str(value))
    return f"<span class='status-pill {slugify_label(value)}'>{label}</span>"


def format_auto_pill(value):
    raw = str(value).strip()
    lower = raw.lower()
    if "already automated" in lower:
        label = "Auto"
        css = "auto"
    elif "high" in lower and "rule" in lower:
        label = "Rule"
        css = "rule"
    elif "high" in lower and "ai" in lower:
        label = "AI"
        css = "ai"
    elif lower.startswith("medium"):
        label = "Medium"
        css = "medium"
    elif lower.startswith("low"):
        label = "Judgment"
        css = "low-judgment-required"
    else:
        label = raw or "Manual"
        css = slugify_label(label)
    return f"<span class='auto-pill {css}'>{escape(label)}</span>"


def format_hours_short(value):
    if pd.isna(value):
        return ""
    value = float(value)
    text = f"{value:.2f}".rstrip("0").rstrip(".")
    return f"{text}h"


def build_tracker_board_html(worklist, summary):
    rows = []
    display = worklist.loc[
        :,
        [
            "Step",
            "Task",
            "Category",
            "Dependency Task",
            "Dependency Status",
            "Dependency Owner",
            "Owner",
            "Deadline",
            "Status",
            "Automation Potential",
            "Estimated Hours",
            "Next Action",
        ],
    ].copy()

    for _, row in display.iterrows():
        rows.append(
            dedent(
                f"""
                <tr>
                    <td class="tracker-step">{int(row['Step'])}</td>
                    <td>
                        <div class="tracker-task">{escape(str(row['Task']))}</div>
                        <div class="tracker-subtask">{escape(str(row['Next Action']))}</div>
                    </td>
                    <td>{escape(str(row['Category']))}</td>
                    <td>
                        <div class="tracker-task">{escape(str(row['Dependency Task']))}</div>
                        <div class="tracker-subtask">{escape(str(row['Dependency Status']))} · {escape(str(row['Dependency Owner']))}</div>
                    </td>
                    <td><span class="tracker-owner">{escape(str(row['Owner']))}</span></td>
                    <td><span class="tracker-date">{escape(str(row['Deadline']))}</span></td>
                    <td>{format_status_pill(row['Status'])}</td>
                    <td>{format_auto_pill(row['Automation Potential'])}</td>
                    <td><span class="tracker-hours">{format_hours_short(row['Estimated Hours'])}</span></td>
                </tr>
                """
            ).strip()
        )

    return dedent(
        f"""
        <div class="tracker-board">
            <div class="tracker-header">
                <div>
                    <div class="tracker-kicker">Close Checklist Tracker</div>
                    <div class="tracker-title">Open checklist handoffs and critical path tasks</div>
                    <div class="tracker-copy">
                        The board below strips the checklist down to the open queue so the team can see who owns the next move,
                        which dependency is holding it back, and where automation is most realistic.
                    </div>
                </div>
                <div class="tracker-badges">
                    <span class="tracker-badge done">{summary['completed_tasks']} done</span>
                    <span class="tracker-badge active">{summary['in_progress_tasks']} active</span>
                    <span class="tracker-badge waiting">{summary['waiting_tasks']} waiting</span>
                    <span class="tracker-badge blocked">{summary['blocked_tasks']} blocked</span>
                </div>
            </div>
            <div class="tracker-table-wrap">
                <table class="tracker-table">
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Task</th>
                            <th>Category</th>
                            <th>Dependency</th>
                            <th>Owner</th>
                            <th>Deadline</th>
                            <th>Status</th>
                            <th>Auto Potential</th>
                            <th>Est. Hrs</th>
                        </tr>
                    </thead>
                    <tbody>{''.join(rows)}</tbody>
                </table>
            </div>
        </div>
        """
    ).strip()


def build_dependency_summary_html(dependency_summary):
    cards = [
        ("External Inputs", dependency_summary["external_inputs"], "Tasks waiting on input outside the checklist chain."),
        ("Upstream In Progress", dependency_summary["upstream_in_progress"], "Tasks depending on upstream work already underway."),
        ("Upstream Not Started", dependency_summary["upstream_not_started"], "Tasks blocked by work that has not started yet."),
        ("Completed Handoffs", dependency_summary["completed_handoffs_open"], "Tasks whose dependency is complete but the downstream step is still open."),
    ]
    blocks = []
    for title, value, copy in cards:
        blocks.append(
            dedent(
                f"""
                <div class="tracker-note">
                    <div class="tracker-note-top">
                        <div class="tracker-note-title">{escape(str(title))}</div>
                        <span class="tracker-badge active">{value}</span>
                    </div>
                    <div class="tracker-note-copy">{escape(str(copy))}</div>
                </div>
                """
            ).strip()
        )

    return f"<div class='tracker-notes'>{''.join(blocks)}</div>"


def build_tracker_notes_html(handoff_queue):
    cards = []
    for _, row in handoff_queue.head(3).iterrows():
        cards.append(
            dedent(
                f"""
                <div class="tracker-note">
                    <div class="tracker-note-top">
                        <div class="tracker-note-title">Step {int(row['Step'])}: {escape(str(row['Task']))}</div>
                        {format_status_pill(row['Status'])}
                    </div>
                    <div class="tracker-note-copy">
                        Owner: {escape(str(row['Owner']))}<br>
                        Dependency: {escape(str(row['Dependency Task']))} ({escape(str(row['Dependency Status']))})<br>
                        Next action: {escape(str(row['Next Action']))}
                    </div>
                </div>
                """
            ).strip()
        )

    return f"<div class='tracker-notes'>{''.join(cards)}</div>"


def format_display_frame(df):
    frame = df.copy()
    currency_label_overrides = {"Unresolved Difference", "Open AR Balance"}
    currency_columns = []

    for column in frame.columns:
        if not pd.api.types.is_numeric_dtype(frame[column]):
            continue

        if "(EUR)" in str(column) or column in currency_label_overrides:
            currency_columns.append(column)

    if not currency_columns:
        return frame

    return (
        frame.style.format({column: "{:,.2f}" for column in currency_columns}, na_rep="")
        .set_properties(subset=currency_columns, **{"text-align": "right"})
    )


def render_metric_card(title, value, subtitle, tone="medium"):
    st.markdown(
        f"""
        <div class="metric-card {tone}">
            <div class="metric-label">{title}</div>
            <div class="metric-value">{value}</div>
            <div class="metric-subtitle">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_kpi_strip(summary):
    pending_actions = summary["not_started_tasks"] + summary["blocked_tasks"]
    high_risk_items = summary["accrual_risks"] + summary["ap_3way_exceptions"]
    match_rate = summary["bank_match_rate"]
    bank_matched = summary["bank_matched"]
    bank_total = summary["bank_total"]
    ai_candidates = summary["checklist_automation_candidates"]

    st.markdown(
        f"""
        <div class="kpi-strip">
            <div class="kpi-tile">
                <div class="kpi-tile-label">Close Day Target</div>
                <div class="kpi-tile-value">4</div>
                <div class="kpi-tile-copy">From 7 business days</div>
            </div>
            <div class="kpi-tile">
                <div class="kpi-tile-label">Pending Actions</div>
                <div class="kpi-tile-value">{pending_actions}</div>
                <div class="kpi-tile-copy">{summary['not_started_tasks']} not started · {summary['blocked_tasks']} blocked</div>
            </div>
            <div class="kpi-tile">
                <div class="kpi-tile-label">High Risk Items</div>
                <div class="kpi-tile-value">{high_risk_items}</div>
                <div class="kpi-tile-copy">{summary['accrual_risks']} accruals · {summary['ap_3way_exceptions']} AP</div>
            </div>
            <div class="kpi-tile">
                <div class="kpi-tile-label">Auto-Reconciled</div>
                <div class="kpi-tile-value">{match_rate}%</div>
                <div class="kpi-tile-copy">{bank_matched} of {bank_total} bank items</div>
            </div>
            <div class="kpi-tile">
                <div class="kpi-tile-label">AI Candidates</div>
                <div class="kpi-tile-value">{ai_candidates}</div>
                <div class="kpi-tile-copy">{summary['bank_journal_candidates']} bank drafts · {summary['journal_agent_ready_drafts']} journals</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_close_intelligence(summary, kpis):
    variance_rows = kpis["details"]["tb_variances"].head(2)
    variance_text = ", ".join(
        f"{row['Account Name']} ({row['Variance %']})" for _, row in variance_rows.iterrows()
    )
    variance_text = variance_text or "No major variances flagged"

    st.markdown(
        f"""
        <div class="intel-band">
            <div class="intel-card">
                <div class="intel-title">GL Approval Bottleneck</div>
                <div class="intel-copy">
                    <strong>{summary['pending_gl']}</strong> entries pending review or approval and
                    <strong> {summary['manual_jes']} </strong> manual journals flagged. Straight-through batching can remove most of this drag.
                </div>
            </div>
            <div class="intel-card">
                <div class="intel-title">Variance Hotspots</div>
                <div class="intel-copy">
                    <strong>{summary['large_variances']}</strong> accounts moved &gt;10%. Largest movements: {variance_text}.
                    AI variance commentary can cut review time.
                </div>
            </div>
            <div class="intel-card">
                <div class="intel-title">Close Effort at Risk</div>
                <div class="intel-copy">
                    <strong>{summary['checklist_recoverable_hours']} hours</strong> are recoverable by clearing critical checklist handoffs.
                    Bank automation can clear <strong>{summary['bank_auto_clear_candidates']}</strong> timing items.
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_close_timeline(summary):
    total_tasks = sum(summary["checklist_status_counts"].values())
    progress_ratio = summary["completed_tasks"] / total_tasks if total_tasks else 0
    day_marker = max(1, min(4, int(round(progress_ratio * 4 + 0.5))))
    labels = [
        "Day 1 · Pre-close",
        "Day 2 · Recons & accruals",
        "Day 3 · IC elim & TB review",
        "Day 4 · Pack & sign-off",
    ]
    steps = []
    for index, label in enumerate(labels, start=1):
        active = " active" if index == day_marker else ""
        steps.append(f"<div class='timeline-step{active}'>{label}</div>")

    st.markdown(
        f"""
        <div class="timeline">
            <div class="timeline-title">Close Day Timeline — Target: 4 Days · Day {day_marker} of 4</div>
            <div class="timeline-track">{''.join(steps)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_ai_opportunity(title, copy):
    st.markdown(
        f"""
        <div class="opportunity-card">
            <div class="opportunity-title">{title}</div>
            <div class="intel-copy">{copy}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_focus_board(kpis, score, priorities):
    summary = kpis["summary"]
    posting = kpis["agents"]["erp_posting"]["summary"]
    checklist_agent = kpis["agents"]["checklist"]
    top_priorities = priorities[:3]
    recs = build_below_four_recommendations(kpis)

    if recs.empty:
        bundle_text = "No tested bundle gets below 4.0 days yet."
        bundle_chip = f"{score['continuous_close_days']} day forecast"
    else:
        top = recs.iloc[0]
        bundle_text = f"{top['Bundle']} → {top['Continuous Forecast']} day forecast"
        bundle_chip = f"{top['Hours Saved']} hrs recovered"

    checklist_focus = checklist_agent["handoff_queue"].head(2)
    if checklist_focus.empty:
        checklist_text = "No critical checklist handoffs are currently waiting."
    else:
        checklist_text = " · ".join(
            f"{row['Owner']} owns {row['Task']}" for _, row in checklist_focus.iterrows()
        )

    priority_lines = "".join(
        f"<div class='focus-item'><strong>{escape(item['priority_item'])}</strong> — {escape(item['downstream_unlock'])}</div>"
        for item in top_priorities
    )

    st.markdown(
        f"""
        <div class="focus-grid">
            <div class="focus-card">
                <div class="focus-title">What to do next</div>
                <div class="focus-list">{priority_lines}</div>
                <span class="focus-chip">{summary['pending_gl']} GL approvals · {summary['bank_exceptions']} bank exceptions</span>
            </div>
            <div class="focus-card">
                <div class="focus-title">Fastest path below 4 days</div>
                <div class="focus-list">
                    <div class="focus-item">{escape(bundle_text)}</div>
                </div>
                <span class="focus-chip">{escape(str(bundle_chip))}</span>
            </div>
            <div class="focus-card">
                <div class="focus-title">Posting & handoffs</div>
                <div class="focus-list">
                    <div class="focus-item"><strong>{posting['auto_post']}</strong> auto-post · <strong>{posting['ready_to_post']}</strong> ready · <strong>{posting['manual_hold']}</strong> hold</div>
                    <div class="focus-item">{escape(checklist_text)}</div>
                </div>
                <span class="focus-chip">{summary['checklist_recoverable_hours']} hrs recoverable</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_area_card(area):
    metrics = [f"{label}: {value}" for label, value in area["metrics"].items()]
    st.markdown(
        f"""
        <div class="area-card">
            <div class="area-header">
                <div class="area-title">{area['label']}</div>
                <div class="area-score">{area['severity_score']}</div>
            </div>
            <div class="small-note">{area['sheet']}</div>
            <div class="area-headline">{area['headline']}</div>
            <div class="area-subheadline">{area['subheadline']}</div>
            {format_chip_row(metrics)}
            <div class="area-insight">{area['insight']}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_priority_card(priority):
    st.markdown(
        f"""
        <div class="priority-card {priority_tone(priority['priority_score'])}">
            <div class="priority-top">
                <div class="priority-area">{priority['area']}</div>
                <div class="priority-score">{priority['priority_score']}</div>
            </div>
            <div class="priority-title">{priority['priority_item']}</div>
            <div class="priority-copy">{priority['why_it_matters']}</div>
            <div class="priority-copy"><strong>Impact:</strong> {priority['impact']}</div>
            {format_chip_row([f"{priority['hours_saved_est']} hrs recoverable", priority['downstream_unlock']])}
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_entity_callout(entity):
    st.markdown(
        f"""
        <div class="entity-callout">
            <div class="section-label">Highest Entity Risk</div>
            <div class="entity-name">{entity['entity']}</div>
            <div class="entity-copy">
                Risk score {entity['risk_score']}/100. Primary blocker: {entity['primary_blocker']}.
                Current driver: {entity['driver_summary']}.
            </div>
            {format_chip_row([
                f"{entity['pending_gl']} pending GL",
                f"{entity['ap_exceptions']} AP issues",
                f"{entity['tb_large_variances']} large TB variances",
            ])}
        </div>
        """,
        unsafe_allow_html=True,
    )


def build_readiness_gauge(score):
    if not HAS_PLOTLY:
        return None

    figure = go.Figure(
        go.Indicator(
            mode="gauge+number",
            value=score["readiness_score"],
            number={"suffix": "/100", "font": {"size": 34, "color": "#143642"}},
            title={"text": "Close Readiness", "font": {"size": 16, "color": "#5b6b78"}},
            gauge={
                "axis": {"range": [0, 100], "tickwidth": 0, "tickcolor": "#5b6b78"},
                "bar": {"color": "#143642"},
                "steps": [
                    {"range": [0, 65], "color": "#f2d6d6"},
                    {"range": [65, 80], "color": "#f8e3c1"},
                    {"range": [80, 100], "color": "#d8efe7"},
                ],
                "threshold": {
                    "line": {"color": "#d77a2d", "width": 6},
                    "thickness": 0.9,
                    "value": 80,
                },
            },
        )
    )
    figure.update_layout(
        height=300,
        margin=dict(l=20, r=20, t=40, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        font={"family": "Arial", "color": "#0f1f2c"},
    )
    return figure


def build_area_chart(kpis):
    area_rows = []
    for key in AREA_ORDER:
        area = kpis["areas"][key]
        area_rows.append({"Area": area["label"], "Severity": area["severity_score"]})
    chart_data = pd.DataFrame(area_rows).sort_values("Severity")

    if HAS_PLOTLY:
        figure = px.bar(
            chart_data,
            x="Severity",
            y="Area",
            orientation="h",
            text="Severity",
            color="Severity",
            color_continuous_scale=["#2a9d8f", "#e9a03b", "#c75c5c"],
        )
        figure.update_layout(
            height=360,
            margin=dict(l=10, r=10, t=20, b=10),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            coloraxis_showscale=False,
        )
        figure.update_traces(marker_line_width=0, textposition="outside")
        return figure

    return chart_data.set_index("Area")


def build_checklist_chart(summary):
    chart_data = pd.DataFrame(
        {
            "Status": list(summary["checklist_status_counts"].keys()),
            "Count": list(summary["checklist_status_counts"].values()),
        }
    )

    if HAS_PLOTLY:
        figure = px.pie(
            chart_data,
            names="Status",
            values="Count",
            hole=0.58,
            color="Status",
            color_discrete_map={
                "Completed": "#2a9d8f",
                "In Progress": "#256d85",
                "Not Started": "#d6dde2",
                "Waiting on Input": "#e9a03b",
                "Blocked": "#c75c5c",
            },
        )
        figure.update_layout(
            height=320,
            margin=dict(l=10, r=10, t=10, b=10),
            showlegend=True,
            paper_bgcolor="rgba(0,0,0,0)",
        )
        return figure

    return chart_data.set_index("Status")


def build_gl_lane_chart(gl_agent):
    chart_data = gl_agent["lane_breakdown"].copy()
    if HAS_PLOTLY:
        figure = px.bar(
            chart_data,
            x="Approval Lane",
            y="Count",
            text="Count",
            color="Approval Lane",
            color_discrete_map={
                "Straight-through": "#2a9d8f",
                "Manager queue": "#256d85",
                "Controller queue": "#d77a2d",
                "CFO queue": "#c75c5c",
            },
        )
        figure.update_layout(
            height=320,
            margin=dict(l=10, r=10, t=20, b=10),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            xaxis_title="",
            yaxis_title="",
            showlegend=False,
        )
        figure.update_traces(marker_line_width=0, textposition="outside")
        return figure

    return chart_data.set_index("Approval Lane")


def build_gl_entity_chart(gl_agent):
    chart_data = gl_agent["entity_breakdown"].copy().sort_values("Pending Items")
    if HAS_PLOTLY:
        figure = px.bar(
            chart_data,
            x="Pending Items",
            y="Entity",
            orientation="h",
            text="Pending Items",
            color="Pending Items",
            color_continuous_scale=["#9fc5d1", "#d77a2d", "#c75c5c"],
        )
        figure.update_layout(
            height=320,
            margin=dict(l=10, r=10, t=20, b=10),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            coloraxis_showscale=False,
            xaxis_title="",
            yaxis_title="",
        )
        figure.update_traces(marker_line_width=0, textposition="outside")
        return figure

    return chart_data.set_index("Entity")


def build_bank_status_chart(bank_agent):
    chart_data = bank_agent["status_breakdown"].copy()
    if HAS_PLOTLY:
        figure = px.bar(
            chart_data,
            x="Match Status",
            y="Count",
            text="Count",
            color="Count",
            color_continuous_scale=["#9fc5d1", "#d77a2d", "#c75c5c"],
        )
        figure.update_layout(
            height=320,
            margin=dict(l=10, r=10, t=20, b=10),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            coloraxis_showscale=False,
            xaxis_title="",
            yaxis_title="",
        )
        figure.update_traces(marker_line_width=0, textposition="outside")
        return figure

    return chart_data.set_index("Match Status")


def build_bank_entity_chart(bank_agent):
    chart_data = bank_agent["entity_breakdown"].copy().sort_values("Open Items")
    if HAS_PLOTLY:
        figure = px.bar(
            chart_data,
            x="Open Items",
            y="Entity",
            orientation="h",
            text="Open Items",
            color="Open Items",
            color_continuous_scale=["#9fc5d1", "#d77a2d", "#c75c5c"],
        )
        figure.update_layout(
            height=320,
            margin=dict(l=10, r=10, t=20, b=10),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            coloraxis_showscale=False,
            xaxis_title="",
            yaxis_title="",
        )
        figure.update_traces(marker_line_width=0, textposition="outside")
        return figure

    return chart_data.set_index("Entity")


def build_ic_issue_chart(ic_agent):
    chart_data = ic_agent["issue_breakdown"].copy()
    if HAS_PLOTLY:
        figure = px.bar(
            chart_data,
            x="Issue Type",
            y="Count",
            text="Count",
            color="Issue Type",
            color_discrete_map={
                "FX mismatch": "#c75c5c",
                "Elimination variance": "#d77a2d",
                "Transfer pricing review": "#256d85",
                "Agreement pending": "#5b6b78",
            },
        )
        figure.update_layout(
            height=320,
            margin=dict(l=10, r=10, t=20, b=10),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            xaxis_title="",
            yaxis_title="",
            showlegend=False,
        )
        figure.update_traces(marker_line_width=0, textposition="outside")
        return figure

    return chart_data.set_index("Issue Type")


def build_ic_pair_chart(ic_agent):
    chart_data = ic_agent["pair_breakdown"].copy().sort_values("Unresolved Difference (EUR)")
    if HAS_PLOTLY:
        figure = px.bar(
            chart_data,
            x="Unresolved Difference (EUR)",
            y="Pair",
            orientation="h",
            text="Open Items",
            color="Unresolved Difference (EUR)",
            color_continuous_scale=["#9fc5d1", "#d77a2d", "#c75c5c"],
            hover_data={"FX Flags": True, "TP Flags": True, "Open Items": True},
        )
        figure.update_layout(
            height=320,
            margin=dict(l=10, r=10, t=20, b=10),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            coloraxis_showscale=False,
            xaxis_title="",
            yaxis_title="",
        )
        figure.update_traces(texttemplate="%{text} open", textposition="outside", marker_line_width=0)
        return figure

    return chart_data.set_index("Pair")


def build_journal_type_chart(journal_agent):
    chart_data = journal_agent["type_breakdown"].copy()
    if HAS_PLOTLY:
        figure = px.pie(
            chart_data,
            names="JE Type",
            values="Count",
            hole=0.55,
            color="JE Type",
            color_discrete_sequence=["#143642", "#256d85", "#d77a2d", "#c75c5c"],
        )
        figure.update_layout(
            height=320,
            margin=dict(l=10, r=10, t=10, b=10),
            paper_bgcolor="rgba(0,0,0,0)",
            showlegend=True,
        )
        return figure

    return chart_data.set_index("JE Type")


def build_journal_entity_chart(journal_agent):
    chart_data = journal_agent["entity_breakdown"].copy().sort_values("Drafts")
    if HAS_PLOTLY:
        figure = px.bar(
            chart_data,
            x="Drafts",
            y="Entity",
            orientation="h",
            text="Drafts",
            color="Drafts",
            color_continuous_scale=["#9fc5d1", "#d77a2d", "#c75c5c"],
        )
        figure.update_layout(
            height=320,
            margin=dict(l=10, r=10, t=20, b=10),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            coloraxis_showscale=False,
            xaxis_title="",
            yaxis_title="",
        )
        figure.update_traces(marker_line_width=0, textposition="outside")
        return figure

    return chart_data.set_index("Entity")


def build_audit_status_chart(audit_agent):
    chart_data = audit_agent["status_breakdown"].copy()
    if HAS_PLOTLY:
        figure = px.bar(
            chart_data,
            x="Posting Recommendation",
            y="Count",
            text="Count",
            color="Posting Recommendation",
            color_discrete_map={
                "Ready to Post": "#2a9d8f",
                "Conditional Approval": "#e9a03b",
                "Blocked for Review": "#c75c5c",
            },
        )
        figure.update_layout(
            height=320,
            margin=dict(l=10, r=10, t=20, b=10),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            xaxis_title="",
            yaxis_title="",
            showlegend=False,
        )
        figure.update_traces(marker_line_width=0, textposition="outside")
        return figure

    return chart_data.set_index("Posting Recommendation")


def build_audit_source_chart(audit_agent):
    chart_data = audit_agent["source_breakdown"].copy()
    if HAS_PLOTLY:
        figure = px.bar(
            chart_data,
            x="Drafts",
            y="Source Agent",
            orientation="h",
            text="Avg Control Score",
            color="Avg Control Score",
            color_continuous_scale=["#9fc5d1", "#e9a03b", "#c75c5c"],
            hover_data={"Drafts": True, "Avg Control Score": True},
        )
        figure.update_layout(
            height=320,
            margin=dict(l=10, r=10, t=20, b=10),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            coloraxis_showscale=False,
            xaxis_title="",
            yaxis_title="",
        )
        figure.update_traces(texttemplate="%{text}/100", textposition="outside", marker_line_width=0)
        return figure

    return chart_data.set_index("Source Agent")


def build_posting_status_chart(posting_simulator):
    chart_data = posting_simulator["status_breakdown"].copy()
    if HAS_PLOTLY:
        figure = px.bar(
            chart_data,
            x="Posting Outcome",
            y="Count",
            text="Count",
            color="Posting Outcome",
            color_discrete_map={
                "Auto-Post": "#2a9d8f",
                "Ready to Post": "#256d85",
                "Manual Hold": "#c75c5c",
            },
        )
        figure.update_layout(
            height=320,
            margin=dict(l=10, r=10, t=20, b=10),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            xaxis_title="",
            yaxis_title="",
            showlegend=False,
        )
        figure.update_traces(marker_line_width=0, textposition="outside")
        return figure

    return chart_data.set_index("Posting Outcome")


def build_posting_source_chart(posting_simulator):
    chart_data = posting_simulator["source_breakdown"].copy()
    if HAS_PLOTLY:
        figure = px.bar(
            chart_data,
            x="Count",
            y="Source Agent",
            color="Posting Outcome",
            orientation="h",
            text="Count",
            color_discrete_map={
                "Auto-Post": "#2a9d8f",
                "Ready to Post": "#256d85",
                "Manual Hold": "#c75c5c",
            },
            hover_data={"Amount (EUR)": True},
        )
        figure.update_layout(
            barmode="stack",
            height=320,
            margin=dict(l=10, r=10, t=20, b=10),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            xaxis_title="",
            yaxis_title="",
        )
        figure.update_traces(marker_line_width=0, textposition="outside")
        return figure

    return chart_data.set_index(["Source Agent", "Posting Outcome"])[["Count", "Amount (EUR)"]]


def build_checklist_action_chart(checklist_agent):
    chart_data = checklist_agent["action_breakdown"].copy()
    if HAS_PLOTLY:
        figure = px.bar(
            chart_data,
            x="Action Lane",
            y="Count",
            text="Count",
            color="Count",
            color_continuous_scale=["#9fc5d1", "#d77a2d", "#c75c5c"],
        )
        figure.update_layout(
            height=320,
            margin=dict(l=10, r=10, t=20, b=10),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            coloraxis_showscale=False,
            xaxis_title="",
            yaxis_title="",
        )
        figure.update_traces(marker_line_width=0, textposition="outside")
        return figure

    return chart_data.set_index("Action Lane")


def build_checklist_category_chart(checklist_agent):
    chart_data = checklist_agent["category_breakdown"].copy().sort_values("Hours at Risk")
    if HAS_PLOTLY:
        figure = px.bar(
            chart_data,
            x="Hours at Risk",
            y="Category",
            orientation="h",
            text="Open Tasks",
            color="Hours at Risk",
            color_continuous_scale=["#9fc5d1", "#e9a03b", "#c75c5c"],
            hover_data={"Hours at Risk": True, "Open Tasks": True},
        )
        figure.update_layout(
            height=320,
            margin=dict(l=10, r=10, t=20, b=10),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            coloraxis_showscale=False,
            xaxis_title="",
            yaxis_title="",
        )
        figure.update_traces(texttemplate="%{text} open", textposition="outside", marker_line_width=0)
        return figure

    return chart_data.set_index("Category")


def build_entity_chart(entities):
    entity_data = pd.DataFrame(entities).sort_values("risk_score")
    if HAS_PLOTLY:
        figure = px.bar(
            entity_data,
            x="risk_score",
            y="entity",
            orientation="h",
            text="risk_score",
            color="risk_score",
            color_continuous_scale=["#2a9d8f", "#e9a03b", "#c75c5c"],
            hover_data={
                "primary_blocker": True,
                "driver_summary": True,
                "risk_score": False,
            },
        )
        figure.update_layout(
            height=360,
            margin=dict(l=10, r=10, t=20, b=10),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            coloraxis_showscale=False,
        )
        figure.update_traces(textposition="outside", marker_line_width=0)
        return figure

    return entity_data.set_index("entity")[["risk_score"]]


def build_priority_chart(priorities):
    chart_data = pd.DataFrame(priorities[:6])
    if HAS_PLOTLY:
        figure = px.bar(
            chart_data,
            x="priority_score",
            y="priority_item",
            orientation="h",
            text="hours_saved_est",
            color="priority_score",
            color_continuous_scale=["#9fc5d1", "#d77a2d", "#c75c5c"],
        )
        figure.update_layout(
            height=360,
            margin=dict(l=10, r=10, t=20, b=10),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            coloraxis_showscale=False,
        )
        figure.update_traces(texttemplate="%{text} hrs", textposition="outside", marker_line_width=0)
        return figure

    return chart_data.set_index("priority_item")[["priority_score", "hours_saved_est"]]


def build_below_four_recommendations(kpis):
    actions = build_scenario_actions(kpis)
    label_map = {
        "gl_straight_through": "GL auto-approve",
        "gl_fast_track_controller": "GL fast-track",
        "bank_auto_clear": "Bank clear/post",
        "ic_post_drafts": "IC elimination",
        "journal_post_ready": "Journal auto-post",
        "checklist_unblock": "Checklist unblock",
        "ap_exception_sweep": "AP sweep",
    }
    recommendations = []

    for size in (2, 3, 4):
        for combo in combinations(actions, size):
            scenario = simulate_close_scenario(kpis, [item["id"] for item in combo])
            if scenario["gets_below_four"]:
                recommendations.append(
                    {
                        "Bundle": " + ".join(label_map[item["id"]] for item in combo),
                        "Actions": ", ".join(item["title"] for item in combo),
                        "Continuous Forecast": scenario["score"]["continuous_close_days"],
                        "Readiness Score": scenario["score"]["readiness_score"],
                        "Hours Saved": scenario["total_hours_saved"],
                    }
                )

    if not recommendations:
        return pd.DataFrame(columns=["Bundle", "Actions", "Continuous Forecast", "Readiness Score", "Hours Saved"])

    frame = pd.DataFrame(recommendations).drop_duplicates(subset=["Bundle"])
    return frame.sort_values(
        ["Continuous Forecast", "Readiness Score", "Hours Saved"],
        ascending=[True, False, False],
    ).head(6)


def dataset_key_for(source):
    if isinstance(source, str):
        return source
    return f"{getattr(source, 'name', 'uploaded')}::{getattr(source, 'size', 'na')}"


def reset_copilot_if_needed(source_key, kpis, score, priorities):
    if st.session_state.get("copilot_dataset_key") != source_key:
        initial = generate_copilot_response("Give me a CFO summary.", kpis, score, priorities)
        st.session_state["copilot_messages"] = [
            {
                "role": "assistant",
                "answer": initial["answer"],
                "source_metrics": initial["source_metrics"],
                "intent": initial["intent"],
                "suggested_prompts": initial["suggested_prompts"],
            }
        ]
        st.session_state["copilot_dataset_key"] = source_key


def submit_prompt(prompt, kpis, score, priorities):
    response = generate_copilot_response(prompt, kpis, score, priorities)
    st.session_state["copilot_messages"].append({"role": "user", "content": prompt})
    st.session_state["copilot_messages"].append(
        {
            "role": "assistant",
            "answer": response["answer"],
            "source_metrics": response["source_metrics"],
            "intent": response["intent"],
            "suggested_prompts": response["suggested_prompts"],
        }
    )


def render_chat(kpis, score, priorities):
    st.markdown(
        """
        <div class="copilot-shell tight-top">
            <div class="copilot-title">NovaClose Copilot</div>
            <div class="copilot-copy">
                Ask in plain English. The copilot routes prompts to the local close-intelligence engine, so the demo remains reliable even without internet access.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    quick_cols = st.columns(2, gap="small")
    for index, prompt in enumerate(QUICK_PROMPTS):
        if quick_cols[index % 2].button(prompt, key=f"quick_prompt_{index}", use_container_width=True):
            submit_prompt(prompt, kpis, score, priorities)
            safe_rerun()

    for message in st.session_state["copilot_messages"]:
        with st.chat_message(message["role"]):
            if message["role"] == "user":
                st.markdown(message["content"])
            else:
                st.markdown(
                    format_chat_answer(message["answer"], message["source_metrics"]),
                    unsafe_allow_html=True,
                )

    prompt = st.chat_input("Ask NovaClose Copilot about blockers, entities, or variances.")
    if prompt:
        submit_prompt(prompt, kpis, score, priorities)
        safe_rerun()


def render_command_center(kpis, score, priorities, commentary):
    summary = kpis["summary"]
    riskiest_entity = kpis["entities"][0]

    render_kpi_strip(summary)
    render_focus_board(kpis, score, priorities)

    st.markdown(
        f"""
        <div class="hero">
            <div class="hero-eyebrow">Predictive Close Intelligence</div>
            <div class="hero-grid">
                <div>
                    <div class="hero-title">NovaClose Command Centre</div>
                    <p class="hero-copy">
                        NovaTech is not missing its target because finance cannot process transactions.
                        It is missing because exceptions, approvals, and dependencies are surfacing too late in the close.
                        This demo turns that operational noise into a ranked action plan.
                    </p>
                </div>
                <div class="hero-side">
                    <div class="hero-side-label">Live Readout</div>
                    <div class="hero-side-value">{score['predicted_close_days']} day projected close</div>
                    <div class="hero-side-copy">
                        {score['gap_to_target_days']} days above the 4-day CFO target.
                        Highest current entity risk: {summary['riskiest_entity']}.
                    </div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    render_close_intelligence(summary, kpis)
    render_close_timeline(summary)

    card_cols = st.columns(4)
    with card_cols[0]:
        render_metric_card(
            "Close Readiness",
            f"{score['readiness_score']}/100",
            f"{score['risk_level']} risk of missing the 4-day target",
            tone_from_score(score["readiness_score"]),
        )
    with card_cols[1]:
        render_metric_card(
            "Predicted Close",
            f"{score['predicted_close_days']} days",
            f"Gap to target: {score['gap_to_target_days']} days",
            tone_from_score(100 - int(score["gap_to_target_days"] * 20)),
        )
    with card_cols[2]:
        render_metric_card(
            "Top Operational Blocker",
            f"{summary['pending_gl']}",
            "Pending GL approvals and reviews are the largest close drag.",
            "high",
        )
    with card_cols[3]:
        render_metric_card(
            "Highest-Risk Entity",
            riskiest_entity["entity"].replace("NovaTech ", ""),
            f"{riskiest_entity['driver_summary']} is driving its score of {riskiest_entity['risk_score']}/100.",
            tone_from_score(100 - riskiest_entity["risk_score"]),
        )

    left, right = st.columns([1.12, 0.88], gap="large")
    with left:
        st.markdown("<div class='section-label tight-top'>Close Signal</div>", unsafe_allow_html=True)
        gauge = build_readiness_gauge(score)
        if gauge is not None:
            st.plotly_chart(
                gauge,
                use_container_width=True,
                config={"displayModeBar": False},
                key="command_center_readiness_gauge",
            )
        else:
            st.info("Install `plotly` to unlock the full readiness gauge.")
            st.progress(score["readiness_score"] / 100)

        st.markdown("<div class='section-label'>Executive Narrative</div>", unsafe_allow_html=True)
        st.info(commentary)

        with st.expander("Risk by area", expanded=False):
            area_chart = build_area_chart(kpis)
            if HAS_PLOTLY:
                st.plotly_chart(
                    area_chart,
                    use_container_width=True,
                    config={"displayModeBar": False},
                    key="command_center_area_chart",
                )
            else:
                st.bar_chart(area_chart)

    with right:
        st.markdown(
            """
            <div class="entity-callout">
                <div class="section-label">AI Copilot</div>
                <div class="entity-name">Open the AI Copilot tab for live Q&A</div>
                <div class="entity-copy">
                    The full copilot experience is now centralized to reduce clutter across tabs.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    footer_cols = st.columns([0.92, 1.08], gap="large")
    with footer_cols[0]:
        render_entity_callout(riskiest_entity)
    with footer_cols[1]:
        st.markdown("<div class='section-label'>Checklist Composition</div>", unsafe_allow_html=True)
        checklist_chart = build_checklist_chart(summary)
        if HAS_PLOTLY:
            st.plotly_chart(
                checklist_chart,
                use_container_width=True,
                config={"displayModeBar": False},
                key="command_center_checklist_chart",
            )
        else:
            st.bar_chart(checklist_chart)


def render_risk_atlas(kpis):
    st.markdown("<div class='section-label'>Eight-Sheet Risk Atlas</div>", unsafe_allow_html=True)
    area_chart = build_area_chart(kpis)
    if HAS_PLOTLY:
        st.plotly_chart(
            area_chart,
            use_container_width=True,
            config={"displayModeBar": False},
            key="risk_atlas_area_chart",
        )
    else:
        st.bar_chart(area_chart)

    for start in range(0, len(AREA_ORDER), 4):
        cols = st.columns(4, gap="medium")
        for offset, area_key in enumerate(AREA_ORDER[start : start + 4]):
            with cols[offset]:
                render_area_card(kpis["areas"][area_key])

    options = {AREA_LABELS[key]: key for key in AREA_ORDER}
    selected_label = st.selectbox("Inspect detailed evidence", list(options.keys()), index=0)
    selected_key = options[selected_label]
    detail_key = DETAIL_MAP[selected_key]
    detail_frame = format_display_frame(kpis["details"][detail_key])
    st.dataframe(detail_frame, use_container_width=True, hide_index=True)


def render_gl_agent(kpis):
    gl_agent = kpis["agents"]["gl"]
    summary = gl_agent["summary"]

    st.markdown(
        f"""
        <div class="entity-callout">
            <div class="section-label">GL Approval Agent</div>
            <div class="entity-name">{summary['straight_through_candidates']} journals ready for straight-through approval</div>
            <div class="entity-copy">
                The GL approval agent triages the pending journal queue into straight-through, manager, controller, and CFO lanes
                so the team can bulk-approve low-risk items and isolate the real escalation work.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    card_cols = st.columns(6, gap="medium")
    with card_cols[0]:
        render_metric_card("Pending GL", str(summary["pending_items"]), "Journals still sitting in review or approval.", "high")
    with card_cols[1]:
        render_metric_card("Straight-Through", str(summary["straight_through_candidates"]), "Low-risk journals that can move through an auto-approval batch.", "low")
    with card_cols[2]:
        render_metric_card("Manager Queue", str(summary["manager_queue"]), "Low-to-medium risk journals that can be packeted for manager sign-off.", "medium")
    with card_cols[3]:
        render_metric_card("Controller Queue", str(summary["controller_queue"]), "Manual, intercompany, revenue, or high-value items requiring controller review.", "high")
    with card_cols[4]:
        render_metric_card("CFO Queue", str(summary["cfo_queue"]), "Executive sign-off items based on value or equity impact.", "high")
    with card_cols[5]:
        render_metric_card("Hours Recoverable", f"{summary['recoverable_hours']} hrs", "Approval time the agent can recover through batching and straight-through routing.", "low")

    left, right = st.columns(2, gap="large")
    with left:
        st.markdown("<div class='section-label'>Approval Lane Mix</div>", unsafe_allow_html=True)
        lane_chart = build_gl_lane_chart(gl_agent)
        if HAS_PLOTLY:
            st.plotly_chart(
                lane_chart,
                use_container_width=True,
                config={"displayModeBar": False},
                key="gl_agent_lane_chart",
            )
        else:
            st.bar_chart(lane_chart)

    with right:
        st.markdown("<div class='section-label'>Pending GL by Entity</div>", unsafe_allow_html=True)
        entity_chart = build_gl_entity_chart(gl_agent)
        if HAS_PLOTLY:
            st.plotly_chart(
                entity_chart,
                use_container_width=True,
                config={"displayModeBar": False},
                key="gl_agent_entity_chart",
            )
        else:
            st.bar_chart(entity_chart)

    upper_left, upper_right = st.columns(2, gap="large")
    with upper_left:
        st.markdown("<div class='section-label'>Straight-Through Approval Batch</div>", unsafe_allow_html=True)
        st.dataframe(
            format_display_frame(gl_agent["approval_ready"]),
            use_container_width=True,
            hide_index=True,
        )

    with upper_right:
        st.markdown("<div class='section-label'>Approval Packets</div>", unsafe_allow_html=True)
        st.dataframe(
            format_display_frame(gl_agent["approval_packets"]),
            use_container_width=True,
            hide_index=True,
        )

    with st.expander("Full GL approval worklist", expanded=False):
        st.dataframe(
            format_display_frame(gl_agent["worklist"]),
            use_container_width=True,
            hide_index=True,
        )

    render_ai_opportunity(
        "AI Opportunity",
        f"Auto-approve {summary['straight_through_candidates']} straight-through journals and batch the manager/controller queue to recover ~{summary['recoverable_hours']} hours of approval chase time.",
    )


def render_bank_agent(kpis):
    bank_agent = kpis["agents"]["bank"]
    summary = bank_agent["summary"]

    st.markdown(
        f"""
        <div class="entity-callout">
            <div class="section-label">Bank Reconciliation Agent</div>
            <div class="entity-name">{summary['open_items']} open cash exceptions triaged</div>
            <div class="entity-copy">
                This agent groups bank exceptions by action bucket, proposes next steps, highlights draft journal candidates,
                and isolates statement-level breaks that should block close sign-off.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    top_card_cols = st.columns(3, gap="medium")
    with top_card_cols[0]:
        render_metric_card(
            "Open Exceptions",
            str(summary["open_items"]),
            "All non-matched bank items currently in the reconciliation queue.",
            "high",
        )
    with top_card_cols[1]:
        render_metric_card(
            "Auto-Clear Candidates",
            str(summary["auto_clear_candidates"]),
            "Timing items the agent expects to clear next month without a journal.",
            "medium",
        )
    with top_card_cols[2]:
        render_metric_card(
            "Journal Candidates",
            str(summary["journal_candidates"]),
            "Items that likely need an FX, rounding, or residual clearing entry.",
            "medium",
        )

    lower_card_cols = st.columns(3, gap="medium")
    with lower_card_cols[0]:
        render_metric_card(
            "Escalations",
            str(summary["escalations"]),
            "Items the agent wants escalated to treasury, vendor, or bank counterparties.",
            "high",
        )
    with lower_card_cols[1]:
        render_metric_card(
            "Manual Investigation",
            str(summary["manual_investigations"]),
            "Items that still need analyst review before a clean reconciliation decision.",
            "high",
        )
    with lower_card_cols[2]:
        render_metric_card(
            "Statement Break",
            f"EUR {summary['statement_break_value']:,.2f}",
            "Residual statement-level break that still needs treasury investigation.",
            "high",
        )

    left, right = st.columns(2, gap="large")
    with left:
        st.markdown("<div class='section-label'>Exception Mix</div>", unsafe_allow_html=True)
        status_chart = build_bank_status_chart(bank_agent)
        if HAS_PLOTLY:
            st.plotly_chart(
                status_chart,
                use_container_width=True,
                config={"displayModeBar": False},
                key="bank_agent_status_chart",
            )
        else:
            st.bar_chart(status_chart)

    with right:
        st.markdown("<div class='section-label'>Entity Workload</div>", unsafe_allow_html=True)
        entity_chart = build_bank_entity_chart(bank_agent)
        if HAS_PLOTLY:
            st.plotly_chart(
                entity_chart,
                use_container_width=True,
                config={"displayModeBar": False},
                key="bank_agent_entity_chart",
            )
        else:
            st.bar_chart(entity_chart)

    st.markdown("<div class='section-label'>Agent Worklist</div>", unsafe_allow_html=True)
    st.dataframe(
        format_display_frame(bank_agent["worklist"]),
        use_container_width=True,
        hide_index=True,
    )

    lower_left, lower_right = st.columns(2, gap="large")
    with lower_left:
        st.markdown("<div class='section-label'>ERP-Ready JE Drafts</div>", unsafe_allow_html=True)
        st.dataframe(
            format_display_frame(bank_agent["erp_journal_drafts"]),
            use_container_width=True,
            hide_index=True,
        )

    with lower_right:
        st.markdown("<div class='section-label'>Subledger Break Analysis</div>", unsafe_allow_html=True)
        st.dataframe(
            format_display_frame(bank_agent["subledger_breaks"]),
            use_container_width=True,
            hide_index=True,
        )

    render_ai_opportunity(
        "AI Opportunity",
        f"Auto-clear {summary['auto_clear_candidates']} timing items and post {summary['journal_candidates']} bank drafts once audit clears them.",
    )


def render_ic_agent(kpis):
    ic_agent = kpis["agents"]["ic"]
    summary = ic_agent["summary"]

    st.markdown(
        f"""
        <div class="entity-callout">
            <div class="section-label">Intercompany Agent</div>
            <div class="entity-name">{summary['auto_matched']} IC pairs auto-matched, {summary['open_exceptions']} still need elimination support</div>
            <div class="entity-copy">
                The IC agent auto-clears matched counterparties, isolates FX-driven mismatches, drafts elimination entries,
                and pushes transfer-pricing or agreement issues into a separate compliance watchlist.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    card_cols = st.columns(6, gap="medium")
    with card_cols[0]:
        render_metric_card(
            "Auto-Matched",
            str(summary["auto_matched"]),
            "IC pairs where sending and receiving values align without manual intervention.",
            "low",
        )
    with card_cols[1]:
        render_metric_card(
            "Open Exceptions",
            str(summary["open_exceptions"]),
            "Pairs with unresolved amount differences still blocking clean elimination.",
            "high",
        )
    with card_cols[2]:
        render_metric_card(
            "FX Mismatches",
            str(summary["fx_mismatches"]),
            "Exceptions where the mismatch appears to be FX-driven.",
            "medium",
        )
    with card_cols[3]:
        render_metric_card(
            "Elimination Drafts",
            str(summary["elimination_drafts"]),
            "ERP-ready elimination entries generated from open IC mismatches.",
            "medium",
        )
    with card_cols[4]:
        render_metric_card(
            "TP Flags",
            str(summary["tp_flags"]),
            "Transactions that still need transfer-pricing or agreement review.",
            "medium",
        )
    with card_cols[5]:
        render_metric_card(
            "Unresolved Diff",
            f"EUR {summary['unresolved_difference']:,.2f}",
            "Absolute unresolved difference across the open IC exception set.",
            "high",
        )

    left, right = st.columns(2, gap="large")
    with left:
        st.markdown("<div class='section-label'>Issue Mix</div>", unsafe_allow_html=True)
        issue_chart = build_ic_issue_chart(ic_agent)
        if HAS_PLOTLY:
            st.plotly_chart(
                issue_chart,
                use_container_width=True,
                config={"displayModeBar": False},
                key="ic_agent_issue_chart",
            )
        else:
            st.bar_chart(issue_chart)

    with right:
        st.markdown("<div class='section-label'>Open Exposure by Pair</div>", unsafe_allow_html=True)
        pair_chart = build_ic_pair_chart(ic_agent)
        if HAS_PLOTLY:
            st.plotly_chart(
                pair_chart,
                use_container_width=True,
                config={"displayModeBar": False},
                key="ic_agent_pair_chart",
            )
        else:
            st.bar_chart(pair_chart)

    upper_left, upper_right = st.columns(2, gap="large")
    with upper_left:
        st.markdown("<div class='section-label'>IC Exception Worklist</div>", unsafe_allow_html=True)
        st.dataframe(
            format_display_frame(ic_agent["worklist"]),
            use_container_width=True,
            hide_index=True,
        )

    with upper_right:
        st.markdown("<div class='section-label'>ERP-Ready Elimination Entries</div>", unsafe_allow_html=True)
        st.dataframe(
            format_display_frame(ic_agent["erp_elimination_drafts"]),
            use_container_width=True,
            hide_index=True,
        )

    lower_left, lower_right = st.columns(2, gap="large")
    with lower_left:
        st.markdown("<div class='section-label'>Transfer Pricing Watchlist</div>", unsafe_allow_html=True)
        st.dataframe(
            format_display_frame(ic_agent["tp_watchlist"]),
            use_container_width=True,
            hide_index=True,
        )

    with lower_right:
        with st.expander("Auto-matched IC pairs", expanded=False):
            st.dataframe(
                format_display_frame(ic_agent["auto_matches"]),
                use_container_width=True,
                hide_index=True,
            )

    render_ai_opportunity(
        "AI Opportunity",
        f"Post {summary['elimination_drafts']} elimination drafts and auto-clear {summary['auto_matched']} matched pairs to remove FX-driven consolidation noise.",
    )


def render_journal_agent(kpis):
    journal_agent = kpis["agents"]["journal"]
    summary = journal_agent["summary"]

    st.markdown(
        f"""
        <div class="entity-callout">
            <div class="section-label">Journal Entry Agent</div>
            <div class="entity-name">{summary['erp_ready_drafts']} draft JEs ready for ERP posting</div>
            <div class="entity-copy">
                The journal agent reads accrual support, applies standard reclassification rules, and appends an AI-style
                audit trail to every candidate before deciding whether the draft is ERP-ready or still needs controller review.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    card_cols = st.columns(5, gap="medium")
    with card_cols[0]:
        render_metric_card(
            "ERP-Ready Drafts",
            str(summary["erp_ready_drafts"]),
            "Draft journal entries that are ready to move into the ERP posting workflow.",
            "high",
        )
    with card_cols[1]:
        render_metric_card(
            "Contract-Backed",
            str(summary["contract_backed_drafts"]),
            "Drafts supported by contract references, POs, or email confirmations.",
            "medium",
        )
    with card_cols[2]:
        render_metric_card(
            "Std. Reclasses",
            str(summary["standard_reclasses"]),
            "Rule-based reclassification drafts for recurring close patterns.",
            "medium",
        )
    with card_cols[3]:
        render_metric_card(
            "Audit Trails",
            str(summary["audit_trails_attached"]),
            "Every candidate carries an AI-generated audit trail narrative.",
            "medium",
        )
    with card_cols[4]:
        render_metric_card(
            "Review Needed",
            str(summary["review_needed"]),
            "Candidates that still need support or controller review before ERP posting.",
            "high",
        )

    left, right = st.columns(2, gap="large")
    with left:
        st.markdown("<div class='section-label'>JE Type Mix</div>", unsafe_allow_html=True)
        type_chart = build_journal_type_chart(journal_agent)
        if HAS_PLOTLY:
            st.plotly_chart(
                type_chart,
                use_container_width=True,
                config={"displayModeBar": False},
                key="journal_agent_type_chart",
            )
        else:
            st.bar_chart(type_chart)

    with right:
        st.markdown("<div class='section-label'>ERP-Ready Drafts by Entity</div>", unsafe_allow_html=True)
        entity_chart = build_journal_entity_chart(journal_agent)
        if HAS_PLOTLY:
            st.plotly_chart(
                entity_chart,
                use_container_width=True,
                config={"displayModeBar": False},
                key="journal_agent_entity_chart",
            )
        else:
            st.bar_chart(entity_chart)

    upper_left, upper_right = st.columns(2, gap="large")
    with upper_left:
        st.markdown("<div class='section-label'>ERP-Ready JE Drafts</div>", unsafe_allow_html=True)
        st.dataframe(
            format_display_frame(journal_agent["erp_journal_drafts"]),
            use_container_width=True,
            hide_index=True,
        )

    with upper_right:
        st.markdown("<div class='section-label'>Candidate Worklist</div>", unsafe_allow_html=True)
        st.dataframe(
            format_display_frame(journal_agent["worklist"]),
            use_container_width=True,
            hide_index=True,
        )

    st.markdown("<div class='section-label'>AI Audit Trail Pack</div>", unsafe_allow_html=True)
    st.dataframe(
        format_display_frame(journal_agent["audit_trails"]),
        use_container_width=True,
        hide_index=True,
    )

    render_ai_opportunity(
        "AI Opportunity",
        f"Route {summary['erp_ready_drafts']} ERP-ready drafts into posting after audit, leaving only {summary['review_needed']} items for manual review.",
    )


def render_audit_agent(kpis):
    audit_agent = kpis["agents"]["audit"]
    summary = audit_agent["summary"]

    st.markdown(
        f"""
        <div class="entity-callout">
            <div class="section-label">Audit &amp; Compliance Agent</div>
            <div class="entity-name">{summary['ready_to_post']} drafts cleared for posting, {summary['blocked_for_review']} still blocked</div>
            <div class="entity-copy">
                This control gate reviews every ERP-ready draft JE from the Bank and Journal agents, scores control completeness,
                assigns the required approver, and separates straight-through entries from drafts that still need audit or controller attention.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    card_cols = st.columns(6, gap="medium")
    with card_cols[0]:
        render_metric_card(
            "Drafts Reviewed",
            str(summary["drafts_reviewed"]),
            "All ERP-ready drafts that passed through the audit gate.",
            "medium",
        )
    with card_cols[1]:
        render_metric_card(
            "Ready to Post",
            str(summary["ready_to_post"]),
            "Drafts with no material control gaps remaining.",
            "low",
        )
    with card_cols[2]:
        render_metric_card(
            "Conditional",
            str(summary["conditional_approval"]),
            "Drafts that need a stronger approver or light evidence follow-up.",
            "medium",
        )
    with card_cols[3]:
        render_metric_card(
            "Blocked",
            str(summary["blocked_for_review"]),
            "Drafts that should not be posted until control gaps are resolved.",
            "high",
        )
    with card_cols[4]:
        render_metric_card(
            "Avg Score",
            f"{summary['average_control_score']}/100",
            "Average control completeness across all draft JEs.",
            "medium",
        )
    with card_cols[5]:
        render_metric_card(
            "High-Value JEs",
            str(summary["high_value_items"]),
            "Entries requiring secondary approval because of material value.",
            "medium",
        )

    left, right = st.columns(2, gap="large")
    with left:
        st.markdown("<div class='section-label'>Posting Recommendation Mix</div>", unsafe_allow_html=True)
        status_chart = build_audit_status_chart(audit_agent)
        if HAS_PLOTLY:
            st.plotly_chart(
                status_chart,
                use_container_width=True,
                config={"displayModeBar": False},
                key="audit_agent_status_chart",
            )
        else:
            st.bar_chart(status_chart)

    with right:
        st.markdown("<div class='section-label'>Control Coverage by Source Agent</div>", unsafe_allow_html=True)
        source_chart = build_audit_source_chart(audit_agent)
        if HAS_PLOTLY:
            st.plotly_chart(
                source_chart,
                use_container_width=True,
                config={"displayModeBar": False},
                key="audit_agent_source_chart",
            )
        else:
            st.bar_chart(source_chart)

    upper_left, upper_right = st.columns(2, gap="large")
    with upper_left:
        st.markdown("<div class='section-label'>Control Review Pack</div>", unsafe_allow_html=True)
        st.dataframe(
            format_display_frame(audit_agent["review_pack"]),
            use_container_width=True,
            hide_index=True,
        )

    with upper_right:
        st.markdown("<div class='section-label'>Exception Queue</div>", unsafe_allow_html=True)
        st.dataframe(
            format_display_frame(audit_agent["exception_queue"]),
            use_container_width=True,
            hide_index=True,
        )

    st.markdown("<div class='section-label'>Control Memos</div>", unsafe_allow_html=True)
    st.dataframe(
        format_display_frame(audit_agent["control_memos"]),
        use_container_width=True,
        hide_index=True,
    )

    render_ai_opportunity(
        "AI Opportunity",
        f"Straight-through post {summary['ready_to_post']} drafts once the control score threshold is met and auto-hold the {summary['conditional_approval']} conditional items.",
    )


def render_flux_agent(kpis):
    flux_agent = kpis["agents"]["flux"]
    summary = flux_agent["summary"]

    st.markdown(
        f"""
        <div class="entity-callout">
            <div class="section-label">Flux Analysis Agent</div>
            <div class="entity-name">{summary['mom_anomalies']} MoM anomalies · {summary['yoy_anomalies']} YoY anomalies</div>
            <div class="entity-copy">
                This agent computes MoM and YoY movements, scans context cues from GL, accruals, and AP/AR notes, and drafts executive-ready commentary.
                YoY comparisons are derived from a proxy baseline when prior-year data is not available.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    card_cols = st.columns(4, gap="medium")
    with card_cols[0]:
        render_metric_card(
            "MoM Anomalies",
            str(summary["mom_anomalies"]),
            "Accounts moving >10% MoM or flagged for review.",
            "high",
        )
    with card_cols[1]:
        render_metric_card(
            "YoY Anomalies",
            str(summary["yoy_anomalies"]),
            "Accounts moving >15% vs derived YoY baseline.",
            "medium",
        )
    with card_cols[2]:
        render_metric_card(
            "Unreconciled",
            str(summary["unreconciled"]),
            "Variance lines still marked unreconciled.",
            "high",
        )
    with card_cols[3]:
        render_metric_card(
            "Top Variance",
            f"EUR {summary['top_variance_amount']:,.2f}",
            f"Peak variance in {summary['top_variance_account']}.",
            "medium",
        )

    st.markdown("<div class='section-label'>Executive Commentary</div>", unsafe_allow_html=True)
    if flux_agent["commentary"]:
        for line in flux_agent["commentary"]:
            st.info(line)
    else:
        st.info("No significant anomalies were detected for this period.")

    st.markdown("<div class='section-label'>Anomalies</div>", unsafe_allow_html=True)
    st.dataframe(
        format_display_frame(flux_agent["anomalies"]),
        use_container_width=True,
        hide_index=True,
    )

    st.markdown("<div class='section-label'>Flux Worklist</div>", unsafe_allow_html=True)
    st.dataframe(
        format_display_frame(flux_agent["worklist"]),
        use_container_width=True,
        hide_index=True,
    )

    render_ai_opportunity(
        "AI Opportunity",
        "Auto-generate MoM/YoY variance commentary for all flagged accounts and route the top three anomalies directly into the controller review pack.",
    )


def render_checklist_agent(kpis):
    checklist_agent = kpis["agents"]["checklist"]
    summary = checklist_agent["summary"]
    base_summary = kpis["summary"]
    dependency_summary = checklist_agent["dependency_summary"]

    st.markdown(
        f"""
        <div class="entity-callout">
            <div class="section-label">Close Tracker</div>
            <div class="entity-name">{summary['blocked_tasks']} blocked, {summary['waiting_tasks']} waiting, {summary['critical_open_tasks']} critical still open</div>
            <div class="entity-copy">
                The close tracker turns the checklist into an unblock queue. It traces handoffs across reconciliation,
                journals, and reporting so the team can remove dependency bottlenecks before they delay consolidation and reporting.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    card_cols = st.columns(6, gap="medium")
    with card_cols[0]:
        render_metric_card(
            "Blocked",
            str(summary["blocked_tasks"]),
            "Tasks that cannot move until a prior dependency is cleared.",
            "high",
        )
    with card_cols[1]:
        render_metric_card(
            "Waiting on Input",
            str(summary["waiting_tasks"]),
            "Tasks stalled because the next owner is waiting for an upstream handoff.",
            "medium",
        )
    with card_cols[2]:
        render_metric_card(
            "Critical Open",
            str(summary["critical_open_tasks"]),
            "Critical tasks still open on the close path.",
            "high",
        )
    with card_cols[3]:
        render_metric_card(
            "Handoffs at Risk",
            str(summary["handoffs_at_risk"]),
            "Open tasks with dependencies that can still break the close chain.",
            "medium",
        )
    with card_cols[4]:
        render_metric_card(
            "Hours Recoverable",
            f"{summary['recoverable_hours']} hrs",
            "Time that can be recovered by clearing the top bottleneck first.",
            "low",
        )
    with card_cols[5]:
        render_metric_card(
            "Automation Candidates",
            str(summary["automation_candidates"]),
            "Open checklist steps already showing high automation potential.",
            "medium",
        )

    st.markdown(
        build_tracker_board_html(
            checklist_agent["worklist"],
            {
                "completed_tasks": base_summary["completed_tasks"],
                "in_progress_tasks": base_summary["in_progress_tasks"],
                "waiting_tasks": summary["waiting_tasks"],
                "blocked_tasks": summary["blocked_tasks"],
            },
        ),
        unsafe_allow_html=True,
    )

    st.markdown("<div class='section-label'>Dependency Summary</div>", unsafe_allow_html=True)
    st.markdown(build_dependency_summary_html(dependency_summary), unsafe_allow_html=True)

    st.markdown("<div class='section-label'>Top Handoffs To Clear Today</div>", unsafe_allow_html=True)
    st.markdown(build_tracker_notes_html(checklist_agent["handoff_queue"]), unsafe_allow_html=True)

    left, right = st.columns([1.05, 0.95], gap="large")
    with left:
        st.markdown("<div class='section-label'>Hours at Risk by Category</div>", unsafe_allow_html=True)
        category_chart = build_checklist_category_chart(checklist_agent)
        if HAS_PLOTLY:
            st.plotly_chart(
                category_chart,
                use_container_width=True,
                config={"displayModeBar": False},
                key="checklist_agent_category_chart",
            )
        else:
            st.bar_chart(category_chart)

    with right:
        st.markdown("<div class='section-label'>Unblock Action Mix</div>", unsafe_allow_html=True)
        action_chart = build_checklist_action_chart(checklist_agent)
        if HAS_PLOTLY:
            st.plotly_chart(
                action_chart,
                use_container_width=True,
                config={"displayModeBar": False},
                key="checklist_agent_action_chart",
            )
        else:
            st.bar_chart(action_chart)

    with st.expander("Detailed checklist agent tables", expanded=False):
        upper_left, upper_right = st.columns(2, gap="large")
        with upper_left:
            st.markdown("<div class='section-label'>Critical Path</div>", unsafe_allow_html=True)
            st.dataframe(
                format_display_frame(checklist_agent["critical_path"]),
                use_container_width=True,
                hide_index=True,
            )

        with upper_right:
            st.markdown("<div class='section-label'>Dependency Handoffs</div>", unsafe_allow_html=True)
            st.dataframe(
                format_display_frame(checklist_agent["handoff_queue"]),
                use_container_width=True,
                hide_index=True,
            )


def render_agents_hub(kpis):
    agent_options = {
        "GL Approval Agent": render_gl_agent,
        "Bank Reconciliation Agent": render_bank_agent,
        "IC Reconciliation Agent": render_ic_agent,
        "Journal Agent (JE)": render_journal_agent,
        "Audit & Compliance Agent": render_audit_agent,
        "Flux Analysis Agent": render_flux_agent,
    }

    st.markdown(
        """
        <div class="entity-callout">
            <div class="section-label">Agents</div>
            <div class="entity-name">Operational agents for approvals, reconciliations, journals, and controls</div>
            <div class="entity-copy">
                Use the selector below to move between the specialist agents without leaving the main navigation.
                This keeps the top bar compact while preserving the full workflow across GL, bank, IC, journals, and audit.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    selected_agent = st.selectbox(
        "Choose an agent view",
        list(agent_options.keys()),
        index=0,
    )
    agent_options[selected_agent](kpis)


def render_scenario_lab(kpis, score):
    actions = build_scenario_actions(kpis)
    recommendations = build_below_four_recommendations(kpis)
    base_posting = kpis["agents"]["erp_posting"]

    st.markdown(
        f"""
        <div class="entity-callout">
            <div class="section-label">Scenario Lab</div>
            <div class="entity-name">{score['continuous_close_days']} day continuous forecast vs. {score['predicted_close_days']} day dashboard forecast</div>
            <div class="entity-copy">
                This simulator applies the agent actions already identified in the app and recomputes a continuous close-day forecast
                so the team can see which bundles actually get NovaTech below the 4-day target.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    selector_cols = st.columns(2, gap="medium")
    selected_ids = []
    for index, action in enumerate(actions):
        checked = selector_cols[index % 2].checkbox(
            f"{action['area']}: {action['title']}",
            value=False,
            key=f"scenario_action_{action['id']}",
            help=action["description"],
        )
        if checked:
            selected_ids.append(action["id"])

    scenario = simulate_close_scenario(kpis, selected_ids)
    delta_days = round(score["continuous_close_days"] - scenario["score"]["continuous_close_days"], 1)
    posting = scenario["posting_simulator"] if selected_ids else base_posting

    metric_cols = st.columns(5, gap="medium")
    with metric_cols[0]:
        render_metric_card("Base Forecast", f"{score['continuous_close_days']} days", "Continuous current-state close forecast.", "medium")
    with metric_cols[1]:
        render_metric_card("Scenario Forecast", f"{scenario['score']['continuous_close_days']} days", "Forecast after selected actions are applied.", "low" if scenario["gets_below_four"] else "medium")
    with metric_cols[2]:
        render_metric_card("Delta", f"{delta_days} days", "Reduction from the current continuous close forecast.", "low" if delta_days > 0 else "medium")
    with metric_cols[3]:
        render_metric_card("Scenario Score", f"{scenario['score']['readiness_score']}/100", "Readiness score after the selected actions.", "low" if scenario["score"]["readiness_score"] >= 80 else "medium")
    with metric_cols[4]:
        render_metric_card("Hours Saved", f"{scenario['total_hours_saved']} hrs", "Recovered effort from the selected automation bundle.", "low")

    if selected_ids:
        status_text = "Below 4 days" if scenario["gets_below_four"] else "Still above 4 days"
        status_tone = "low" if scenario["gets_below_four"] else "high"
        render_metric_card(
            "Scenario Status",
            status_text,
            f"Gap to target: {scenario['score']['continuous_gap_to_target_days']} days",
            status_tone,
        )

    posting_cols = st.columns(4, gap="medium")
    with posting_cols[0]:
        render_metric_card(
            "Auto-Post",
            str(posting["summary"]["auto_post"]),
            f"EUR {posting['summary']['auto_post_amount']:,.2f} can move straight into ERP.",
            "low",
        )
    with posting_cols[1]:
        render_metric_card(
            "Ready to Post",
            str(posting["summary"]["ready_to_post"]),
            f"EUR {posting['summary']['ready_amount']:,.2f} is cleared but still awaiting release.",
            "medium",
        )
    with posting_cols[2]:
        render_metric_card(
            "Manual Hold",
            str(posting["summary"]["manual_hold"]),
            f"EUR {posting['summary']['hold_amount']:,.2f} still needs approval or control work.",
            "high",
        )
    with posting_cols[3]:
        render_metric_card(
            "Eligible for Auto-Post",
            str(posting["summary"]["eligible_for_auto_post"]),
            "Items that can move to auto-post when the relevant agent action is selected.",
            "medium",
        )

    left, right = st.columns(2, gap="large")
    with left:
        st.markdown("<div class='section-label'>ERP Posting Outcome Mix</div>", unsafe_allow_html=True)
        posting_status_chart = build_posting_status_chart(posting)
        if HAS_PLOTLY:
            st.plotly_chart(
                posting_status_chart,
                use_container_width=True,
                config={"displayModeBar": False},
                key="scenario_lab_posting_status_chart",
            )
        else:
            st.bar_chart(posting_status_chart)

    with right:
        st.markdown("<div class='section-label'>Posting Outcome by Source Agent</div>", unsafe_allow_html=True)
        posting_source_chart = build_posting_source_chart(posting)
        if HAS_PLOTLY:
            st.plotly_chart(
                posting_source_chart,
                use_container_width=True,
                config={"displayModeBar": False},
                key="scenario_lab_posting_source_chart",
            )
        else:
            st.bar_chart(posting_source_chart)

    queue_left, queue_right = st.columns(2, gap="large")
    with queue_left:
        st.markdown("<div class='section-label'>Auto-Post and Ready Queue</div>", unsafe_allow_html=True)
        ready_display = pd.concat(
            [posting["auto_post_queue"], posting["ready_queue"]],
            ignore_index=True,
        )
        st.dataframe(
            format_display_frame(ready_display),
            use_container_width=True,
            hide_index=True,
        )

    with queue_right:
        st.markdown("<div class='section-label'>Manual Hold Queue</div>", unsafe_allow_html=True)
        st.dataframe(
            format_display_frame(posting["manual_hold_queue"]),
            use_container_width=True,
            hide_index=True,
        )

    st.markdown("<div class='section-label'>Recommended Bundles Below 4 Days</div>", unsafe_allow_html=True)
    st.dataframe(
        format_display_frame(recommendations),
        use_container_width=True,
        hide_index=True,
    )


def render_priority_engine(priorities):
    recoverable_hours = sum(priority["hours_saved_est"] for priority in priorities[:4])

    st.markdown(
        f"""
        <div class="entity-callout">
            <div class="section-label">Priority Engine</div>
            <div class="entity-name">{recoverable_hours} hours recoverable in the top four moves</div>
            <div class="entity-copy">
                The engine ranks work by operational impact, downstream unlock, and closeness to the CFO target.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    left, right = st.columns([0.9, 1.1], gap="large")
    with left:
        priority_chart = build_priority_chart(priorities)
        if HAS_PLOTLY:
            st.plotly_chart(
                priority_chart,
                use_container_width=True,
                config={"displayModeBar": False},
                key="priority_engine_chart",
            )
        else:
            st.bar_chart(priority_chart)

    with right:
        top_cols = st.columns(2, gap="medium")
        for index, priority in enumerate(priorities[:4]):
            with top_cols[index % 2]:
                render_priority_card(priority)


def render_automation_plays():
    st.markdown(
        """
        <div class="entity-callout">
            <div class="section-label">Automation Plays</div>
            <div class="entity-name">High-ROI automation ideas for the close</div>
            <div class="entity-copy">
                These plays sit below the immediate priority queue. They are the next layer of improvement once the
                team has stabilized the current close and wants to reduce repeat manual effort month after month.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("<div class='section-label'>Automation Plays</div>", unsafe_allow_html=True)
    auto_cols = st.columns(4, gap="medium")
    automation_cards = [
        ("Journal routing", "Escalate manual journals and pending approvals before day 1."),
        ("Bank matching", "Auto-match recurring timing differences and classify stale items."),
        ("Accrual drafting", "Pre-fill recurring accruals and flag missing support earlier."),
        ("Variance commentary", "Generate finance-ready narratives for the largest TB movements."),
    ]
    for col, (title, copy) in zip(auto_cols, automation_cards):
        with col:
            render_metric_card(title, "AI", copy, "medium")

    roadmap = pd.DataFrame(
        [
            {
                "Play": "Journal routing",
                "Primary Trigger": "Pending review or approval journals",
                "Expected Outcome": "Less approval chasing and faster sign-off",
                "Implementation Style": "Workflow and approval escalation",
            },
            {
                "Play": "Bank matching",
                "Primary Trigger": "Timing differences and recurring bank exceptions",
                "Expected Outcome": "Higher match rate with fewer manual touches",
                "Implementation Style": "Rule engine plus reconciliation agent",
            },
            {
                "Play": "Accrual drafting",
                "Primary Trigger": "Recurring month-end accrual patterns",
                "Expected Outcome": "Earlier accrual preparation and fewer misses",
                "Implementation Style": "Template engine plus reviewer approval",
            },
            {
                "Play": "Variance commentary",
                "Primary Trigger": "Large trial balance movements",
                "Expected Outcome": "Controller-ready narrative in minutes",
                "Implementation Style": "Narrative generation with finance checks",
            },
        ]
    )
    st.markdown("<div class='section-label'>Automation Roadmap</div>", unsafe_allow_html=True)
    st.dataframe(roadmap, use_container_width=True, hide_index=True)


def render_entity_view(kpis):
    entities = kpis["entities"]
    entity_chart = build_entity_chart(entities)

    left, right = st.columns([0.95, 1.05], gap="large")
    with left:
        st.markdown("<div class='section-label'>Entity Risk Comparison</div>", unsafe_allow_html=True)
        if HAS_PLOTLY:
            st.plotly_chart(
                entity_chart,
                use_container_width=True,
                config={"displayModeBar": False},
                key="entity_view_chart",
            )
        else:
            st.bar_chart(entity_chart)

    with right:
        entity_table = pd.DataFrame(entities)[
            [
                "entity",
                "risk_score",
                "risk_level",
                "primary_blocker",
                "pending_gl",
                "bank_exceptions",
                "ap_exceptions",
                "tb_large_variances",
            ]
        ].rename(
            columns={
                "entity": "Entity",
                "risk_score": "Risk Score",
                "risk_level": "Risk Level",
                "primary_blocker": "Primary Blocker",
                "pending_gl": "Pending GL",
                "bank_exceptions": "Bank Exceptions",
                "ap_exceptions": "AP Exceptions",
                "tb_large_variances": "TB Variances",
            }
        )
        st.dataframe(format_display_frame(entity_table), use_container_width=True, hide_index=True)

    spotlight_cols = st.columns(3, gap="medium")
    for col, entity in zip(spotlight_cols, entities[:3]):
        with col:
            render_metric_card(
                entity["entity"].replace("NovaTech ", ""),
                f"{entity['risk_score']}/100",
                f"{entity['primary_blocker']} | {entity['driver_summary']}",
                tone_from_score(100 - entity["risk_score"]),
            )


def render_app_header(kpis):
    summary = kpis["summary"]
    st.markdown("<div class='app-header'>", unsafe_allow_html=True)
    left, right = st.columns([0.62, 0.38], gap="small")
    with left:
        st.markdown(
            """
            <div>
                <div class="app-header-kicker">Month-End Workspace</div>
                <div class="app-header-title">NovaClose</div>
                <div class="app-header-copy">
                    Current close dataset loaded and ready for review across risk, agents, and actions.
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with right:
        status_col, period_col, ask_col = st.columns([0.28, 0.52, 0.2], gap="small")
        with status_col:
            st.markdown(
                "<div class='status-pill'><span class='status-dot'></span>Live</div>",
                unsafe_allow_html=True,
            )
        with period_col:
            st.markdown(
                f"<div class='period-pill'>Period: {summary['accounting_period_label']}</div>",
                unsafe_allow_html=True,
            )
        with ask_col:
            if st.button("Ask AI", key="ask_ai_button", use_container_width=True):
                st.session_state["nav_radio"] = "AI Copilot"
                safe_rerun()
    st.markdown("</div>", unsafe_allow_html=True)


apply_theme()

with st.sidebar:
    st.title("NovaClose")
    st.caption("Pitch-ready demo for the Month-End Challenge")
    uploaded_file = st.file_uploader("Upload an Excel workbook", type=["xlsx"])
    data_source = uploaded_file if uploaded_file is not None else DEFAULT_DATASET
    st.markdown("### Default Demo")
    st.write("If no file is uploaded, the bundled NovaTech workbook is loaded automatically.")
    st.markdown("### Copilot Starters")
    for prompt in QUICK_PROMPTS:
        st.write(f"- {prompt}")
    if not HAS_PLOTLY:
        st.warning("`plotly` is not installed in this environment. The app will fall back to native Streamlit charts.")

data, kpis, score, priorities, commentary = prepare(data_source)
reset_copilot_if_needed(dataset_key_for(data_source), kpis, score, priorities)

with st.sidebar:
    st.markdown("### Accounting Period")
    st.write(kpis["summary"]["accounting_period_label"])

render_app_header(kpis)

nav_items = [
    "Command Centre",
    "Close Tracker",
    "Agents",
    "Risks",
    "Scenarios",
    "Priorities",
    "Automation",
    "AI Copilot",
]

tab_order = nav_items.copy()
if st.session_state.get("nav_target") == "AI Copilot":
    tab_order = ["AI Copilot"] + [tab for tab in nav_items if tab != "AI Copilot"]

tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs(tab_order)
tab_map = dict(zip(tab_order, [tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8]))

with tab_map["Command Centre"]:
    render_command_center(kpis, score, priorities, commentary)

with tab_map["Close Tracker"]:
    render_checklist_agent(kpis)

with tab_map["Agents"]:
    render_agents_hub(kpis)

with tab_map["Risks"]:
    render_risk_atlas(kpis)
    render_entity_view(kpis)

with tab_map["Scenarios"]:
    render_scenario_lab(kpis, score)

with tab_map["Priorities"]:
    render_priority_engine(priorities)

with tab_map["Automation"]:
    render_automation_plays()

with tab_map["AI Copilot"]:
    render_chat(kpis, score, priorities)

if st.session_state.get("nav_target") == "AI Copilot":
    st.session_state["nav_target"] = None

st.divider()
st.caption(
    "NovaClose turns NovaTech’s close from reactive exception hunting into predictive close intelligence."
)
