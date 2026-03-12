from html import escape
from itertools import combinations
from pathlib import Path
from textwrap import dedent
import json
import re

import numpy as np
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


def planning_signature(kpis):
    summary = kpis["summary"]
    return json.dumps(summary, sort_keys=True, default=str)


def get_planning_cache():
    return st.session_state.setdefault("_planning_cache", {})


def get_cached_actions(kpis):
    cache = get_planning_cache()
    key = ("actions", planning_signature(kpis))
    if key not in cache:
        cache[key] = build_scenario_actions(kpis)
    return cache[key]


def get_cached_plan_scenario(kpis, action_ids):
    cache = get_planning_cache()
    key = ("scenario", planning_signature(kpis), tuple(sorted(action_ids)))
    if key not in cache:
        cache[key] = simulate_close_scenario(kpis, action_ids)
    return cache[key]


def get_cached_recommendations(kpis):
    cache = get_planning_cache()
    key = ("recommendations", planning_signature(kpis))
    if key not in cache:
        cache[key] = build_below_four_recommendations(kpis)
    return cache[key]


def format_number(value, decimals=0):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    try:
        number = float(value)
    except (TypeError, ValueError):
        return str(value)
    formatted = f"{abs(number):,.{decimals}f}"
    if decimals == 0:
        formatted = formatted.split(".")[0]
    return f"({formatted})" if number < 0 else formatted


def format_percent(value, decimals=2):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    try:
        number = float(value)
    except (TypeError, ValueError):
        return str(value)
    formatted = f"{abs(number):,.{decimals}f}%"
    return f"({formatted})" if number < 0 else formatted


def format_metric_value(value):
    if isinstance(value, (int, float, np.integer, np.floating)):
        decimals = 2 if isinstance(value, (float, np.floating)) and not float(value).is_integer() else 0
        return format_number(value, decimals)

    if not isinstance(value, str):
        return value

    text = value.strip()
    if not text:
        return value

    percent_match = re.search(r"-?\d+(?:\.\d+)?", text)
    if "%" in text and percent_match:
        number = float(percent_match.group())
        formatted = format_percent(number, 2).replace("%", "")
        return re.sub(r"-?\d+(?:\.\d+)?", formatted, text, count=1)

    currency_match = re.match(r"^([A-Za-z]{3})\s+(-?\d+(?:\.\d+)?)(.*)$", text)
    if currency_match:
        prefix, number, suffix = currency_match.groups()
        decimals = 2 if "." in number else 0
        return f"{prefix} {format_number(float(number), decimals)}{suffix}"

    number_match = re.match(r"^(-?\d+(?:\.\d+)?)(.*)$", text)
    if number_match:
        number, suffix = number_match.groups()
        decimals = 2 if "." in number else 0
        formatted = format_number(float(number), decimals)
        suffix = suffix.strip()
        return f"{formatted} {suffix}".strip()

    return value


def run_all_automations(kpis, actions):
    action_ids = [action["id"] for action in actions]
    scenario = simulate_close_scenario(kpis, action_ids)
    posting = scenario["posting_simulator"]["summary"]
    log_entries = []
    for action in actions:
        log_entries.append(
            {
                "Timestamp": pd.Timestamp.utcnow().strftime("%Y-%m-%d %H:%M UTC"),
                "Action": action["title"],
                "Area": action["area"],
                "Hours Saved": action["hours_saved"],
                "Forecast After": scenario["score"]["predicted_close_days"],
                "Auto-Post": posting["auto_post"],
                "Ready to Post": posting["ready_to_post"],
                "Manual Hold": posting["manual_hold"],
            }
        )
    st.session_state["automation_run"] = True
    st.session_state["applied_actions"] = set(action_ids)
    st.session_state["automation_log"] = log_entries
    st.session_state["automation_summary"] = {
        "forecast": scenario["score"]["predicted_close_days"],
        "score": scenario["score"]["readiness_score"],
        "auto_post": posting["auto_post"],
        "ready_to_post": posting["ready_to_post"],
        "manual_hold": posting["manual_hold"],
    }


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
            --ai: #7353e5;
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

        .app-header-main,
        .app-header-meta {
            padding: 1rem 1.15rem;
            margin-bottom: 1rem;
            border-radius: 22px;
            background: linear-gradient(135deg, rgba(255, 255, 255, 0.92) 0%, rgba(244, 249, 252, 0.96) 100%);
            border: 1px solid rgba(15, 31, 44, 0.08);
            box-shadow: 0 14px 34px rgba(15, 31, 44, 0.06);
            min-height: 118px;
        }

        .app-header-main {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 1rem;
            flex-wrap: wrap;
        }

        .app-header-main-copy {
            flex: 1 1 360px;
            min-width: 0;
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

        .header-ask-ai div[data-testid="stButton"] > button {
            min-height: 52px;
            border-radius: 18px;
            border: none !important;
            background: linear-gradient(135deg, #143642 0%, #256d85 100%) !important;
            color: #f8fbfc !important;
            font-weight: 800;
            box-shadow: 0 12px 24px rgba(20, 54, 66, 0.22);
        }

        .header-ask-ai div[data-testid="stButton"] > button:hover {
            background: linear-gradient(135deg, #143642 0%, #256d85 100%) !important;
            color: #f8fbfc !important;
            filter: brightness(0.98);
        }

        .app-header-meta-stack {
            display: flex;
            align-items: center;
            justify-content: flex-end;
            gap: 0.65rem;
            flex-wrap: wrap;
            flex: 0 0 auto;
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
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            text-align: center;
            min-height: 132px;
        }

        .kpi-tile-label {
            text-transform: uppercase;
            letter-spacing: 0.08em;
            font-size: 0.68rem;
            font-weight: 700;
            color: var(--muted);
            text-align: center;
        }

        .kpi-tile-value {
            font-size: 1.45rem;
            font-weight: 800;
            margin: 0.3rem 0 0.2rem 0;
            color: var(--brand);
            text-align: center;
            display: block;
            width: 100%;
        }

        .kpi-tile-copy {
            font-size: 0.78rem;
            color: var(--muted);
            text-align: center;
        }

        .kpi-tile-value.tone-risk {
            color: var(--risk);
        }

        .kpi-tile-value.tone-warn {
            color: var(--accent);
        }

        .kpi-tile-value.tone-good {
            color: var(--good);
        }

        .kpi-tile-value.tone-ai {
            color: var(--ai);
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
            margin: 0;
        }

        .timeline-head {
            display: flex;
            align-items: center;
            justify-content: space-between;
            gap: 1rem;
            margin-bottom: 0.6rem;
        }

        .timeline-status {
            display: inline-flex;
            align-items: center;
            justify-content: center;
            padding: 0.35rem 0.75rem;
            border-radius: 999px;
            border: 1px solid rgba(20, 54, 66, 0.12);
            background: rgba(20, 54, 66, 0.06);
            color: var(--brand);
            font-size: 0.78rem;
            font-weight: 800;
            white-space: nowrap;
        }

        .timeline-track {
            position: relative;
            display: grid;
            grid-template-columns: repeat(4, 1fr);
            gap: 0.8rem;
            align-items: start;
            padding-top: 0.5rem;
        }

        .timeline-track::before {
            content: "";
            position: absolute;
            left: 8%;
            right: 8%;
            top: 1.05rem;
            height: 3px;
            background: rgba(20, 54, 66, 0.12);
            border-radius: 999px;
            z-index: 0;
        }

        .timeline-step {
            position: relative;
            z-index: 1;
            display: flex;
            flex-direction: column;
            align-items: center;
            text-align: center;
            color: var(--muted);
        }

        .timeline-node {
            width: 24px;
            height: 24px;
            border-radius: 50%;
            background: #f8fbfc;
            border: 3px solid rgba(20, 54, 66, 0.16);
            box-shadow: 0 0 0 6px rgba(255, 255, 255, 0.9);
            display: flex;
            align-items: center;
            justify-content: center;
            font-size: 0.72rem;
            font-weight: 800;
            color: var(--muted);
        }

        .timeline-card {
            margin-top: 0.7rem;
            width: 100%;
            background: rgba(20, 54, 66, 0.05);
            border: 1px solid rgba(20, 54, 66, 0.08);
            border-radius: 16px;
            padding: 0.75rem 0.7rem 0.8rem 0.7rem;
            min-height: 86px;
        }

        .timeline-day {
            font-size: 0.78rem;
            font-weight: 800;
            color: var(--ink);
            margin-bottom: 0.2rem;
        }

        .timeline-copy {
            font-size: 0.8rem;
            line-height: 1.35;
            color: var(--muted);
        }

        .timeline-step.complete .timeline-node {
            background: rgba(42, 157, 143, 0.16);
            border-color: rgba(42, 157, 143, 0.9);
            color: var(--good);
        }

        .timeline-step.complete .timeline-card {
            background: rgba(42, 157, 143, 0.08);
            border-color: rgba(42, 157, 143, 0.18);
        }

        .timeline-step.current .timeline-node {
            background: rgba(215, 122, 45, 0.14);
            border-color: rgba(215, 122, 45, 0.9);
            color: var(--accent);
        }

        .timeline-step.current .timeline-card {
            background: rgba(215, 122, 45, 0.08);
            border-color: rgba(215, 122, 45, 0.2);
        }

        .timeline-step.future .timeline-node {
            background: #f8fbfc;
            border-color: rgba(20, 54, 66, 0.16);
        }

        .timeline-step.future .timeline-card {
            background: rgba(20, 54, 66, 0.04);
            border-color: rgba(20, 54, 66, 0.08);
        }

        .timeline-step.current .timeline-day,
        .timeline-step.current .timeline-copy {
            color: var(--ink);
        }

        .timeline-step.complete .timeline-day {
            color: var(--good);
        }

        .timeline-step.future .timeline-day {
            color: var(--ink);
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

        .focus-card.me-step-card {
            height: 232px;
            display: flex;
            flex-direction: column;
        }

        .focus-card.me-step-card .focus-list {
            flex: 1 1 auto;
            align-content: start;
        }

        .focus-card.me-step-card .focus-chip {
            margin-top: auto;
            align-self: flex-start;
        }

        .me-step-panel {
            display: grid;
            gap: 0.9rem;
        }

        .focus-card.me-step-card.panel {
            height: auto;
            min-height: 168px;
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

        .focus-chip.done {
            border-color: rgba(42, 157, 143, 0.24);
            background: rgba(42, 157, 143, 0.14);
            color: var(--good);
        }

        .me-step-footer {
            margin-top: 0.6rem;
            min-height: 52px;
            display: flex;
            align-items: stretch;
        }

        .me-step-pill {
            width: 100%;
            min-height: 52px;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            border-radius: 18px;
            border: 1px solid rgba(15, 31, 44, 0.10);
            background: rgba(15, 31, 44, 0.06);
            color: var(--muted);
            font-size: 0.96rem;
            font-weight: 700;
            text-align: center;
        }

        .me-step-pill.done {
            border-color: rgba(42, 157, 143, 0.24);
            background: rgba(42, 157, 143, 0.14);
            color: var(--good);
        }

        .checklist-filter-label {
            color: var(--muted);
            font-size: 0.78rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            margin-bottom: 0.4rem;
        }

        .checklist-status-shell {
            display: flex;
            flex-direction: column;
            align-items: flex-end;
        }

        .checklist-status-shell div[data-testid="stPills"] {
            width: 100%;
        }

        .checklist-status-shell div[data-testid="stPills"] > div,
        .checklist-status-shell div[data-testid="stPills"] [role="radiogroup"] {
            width: 100%;
            display: flex;
            justify-content: flex-end;
            flex-wrap: wrap;
            gap: 0.55rem;
        }

        .checklist-status-shell div[data-testid="stPills"] button {
            min-height: 42px;
            padding: 0.46rem 0.88rem;
            border-radius: 999px;
            font-weight: 700;
            border: 1px solid rgba(15, 31, 44, 0.10);
            box-shadow: none;
            color: var(--ink) !important;
        }

        .checklist-status-shell div[data-testid="stPills"] button[aria-checked="true"],
        .checklist-status-shell div[data-testid="stPills"] button[aria-pressed="true"] {
            box-shadow: inset 0 0 0 1px currentColor !important;
        }

        .checklist-status-shell div[data-testid="stPills"] button:nth-of-type(1),
        .checklist-status-shell div[data-testid="stPills"] button:nth-of-type(1)[aria-checked="true"],
        .checklist-status-shell div[data-testid="stPills"] button:nth-of-type(1)[aria-pressed="true"] {
            color: var(--brand) !important;
            background: rgba(20, 54, 66, 0.06) !important;
            border-color: rgba(20, 54, 66, 0.12) !important;
        }

        .checklist-status-shell div[data-testid="stPills"] button:nth-of-type(2),
        .checklist-status-shell div[data-testid="stPills"] button:nth-of-type(2)[aria-checked="true"],
        .checklist-status-shell div[data-testid="stPills"] button:nth-of-type(2)[aria-pressed="true"] {
            color: #1b7f70 !important;
            background: rgba(42, 157, 143, 0.12) !important;
            border-color: rgba(42, 157, 143, 0.18) !important;
        }

        .checklist-status-shell div[data-testid="stPills"] button:nth-of-type(3),
        .checklist-status-shell div[data-testid="stPills"] button:nth-of-type(3)[aria-checked="true"],
        .checklist-status-shell div[data-testid="stPills"] button:nth-of-type(3)[aria-pressed="true"] {
            color: #256d85 !important;
            background: rgba(37, 109, 133, 0.12) !important;
            border-color: rgba(37, 109, 133, 0.18) !important;
        }

        .checklist-status-shell div[data-testid="stPills"] button:nth-of-type(4),
        .checklist-status-shell div[data-testid="stPills"] button:nth-of-type(4)[aria-checked="true"],
        .checklist-status-shell div[data-testid="stPills"] button:nth-of-type(4)[aria-pressed="true"] {
            color: #6a7a88 !important;
            background: rgba(91, 107, 120, 0.10) !important;
            border-color: rgba(91, 107, 120, 0.14) !important;
        }

        .checklist-status-shell div[data-testid="stPills"] button:nth-of-type(5),
        .checklist-status-shell div[data-testid="stPills"] button:nth-of-type(5)[aria-checked="true"],
        .checklist-status-shell div[data-testid="stPills"] button:nth-of-type(5)[aria-pressed="true"] {
            color: #b57a1b !important;
            background: rgba(233, 160, 59, 0.14) !important;
            border-color: rgba(233, 160, 59, 0.20) !important;
        }

        .checklist-status-shell div[data-testid="stPills"] button:nth-of-type(6),
        .checklist-status-shell div[data-testid="stPills"] button:nth-of-type(6)[aria-checked="true"],
        .checklist-status-shell div[data-testid="stPills"] button:nth-of-type(6)[aria-pressed="true"] {
            color: #b34f5a !important;
            background: rgba(199, 92, 92, 0.12) !important;
            border-color: rgba(199, 92, 92, 0.18) !important;
        }

        div[data-testid="stSelectbox"] label,
        div[data-testid="stSelectbox"] label p,
        div[data-testid="stSelectbox"] label div {
            color: var(--ink) !important;
            opacity: 1 !important;
            font-weight: 700 !important;
        }

        div[data-testid="stSelectbox"] [data-baseweb="select"] > div {
            color: var(--ink) !important;
            background: rgba(15, 31, 44, 0.96) !important;
            border-color: rgba(15, 31, 44, 0.18) !important;
        }

        div[data-testid="stSelectbox"] [data-baseweb="select"] svg,
        div[data-testid="stSelectbox"] [data-baseweb="select"] span,
        div[data-testid="stSelectbox"] [data-baseweb="select"] input,
        div[data-testid="stSelectbox"] [data-baseweb="select"] div {
            color: #f4f7fb !important;
            fill: #f4f7fb !important;
            opacity: 1 !important;
        }

        .queue-row {
            background: rgba(255, 255, 255, 0.95);
            border-radius: 20px;
            border: 1px solid rgba(20, 54, 66, 0.12);
            padding: 0.9rem 1.1rem;
            margin-bottom: 1.2rem;
            box-shadow: 0 12px 28px rgba(15, 31, 44, 0.06);
        }

        .queue-title {
            font-weight: 800;
            color: var(--ink);
            margin-bottom: 0.6rem;
        }

        .area-card.compact {
            padding: 0.7rem 0.85rem;
            border-radius: 16px;
        }

        .area-card.compact .area-headline {
            font-size: 0.88rem;
        }

        .area-card.compact .area-subheadline,
        .area-card.compact .area-insight {
            display: none;
        }

        .snapshot-block {
            background: rgba(255, 255, 255, 0.96);
            border-radius: 20px;
            border: 1px solid rgba(20, 54, 66, 0.12);
            padding: 1rem 1.1rem;
            box-shadow: 0 12px 28px rgba(15, 31, 44, 0.06);
            margin-bottom: 1.2rem;
        }

        .snapshot-title {
            font-weight: 800;
            color: var(--ink);
            margin-bottom: 0.6rem;
        }

        .tracker-table-wrap {
            max-height: 720px;
            overflow: auto;
        }

        .tracker-table thead th {
            position: sticky;
            top: 0;
            background: #f8fbfc;
            z-index: 3;
            box-shadow: 0 2px 0 rgba(15, 31, 44, 0.08);
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

        .section-label.centered {
            text-align: center;
        }

        .section-label.section-heading {
            text-align: center;
            font-size: 0.96rem;
            letter-spacing: 0.1em;
            margin-bottom: 0.85rem;
        }

        .priority-sticky-header {
            position: sticky;
            top: 7.2rem;
            z-index: 15;
            background: rgba(244, 247, 251, 0.96);
            border: 1px solid rgba(15, 31, 44, 0.08);
            border-radius: 16px;
            padding: 0.55rem 0.85rem 0.5rem;
            margin: 0.5rem 0 0.85rem;
            box-shadow: 0 8px 24px rgba(15, 31, 44, 0.06);
            backdrop-filter: blur(8px);
        }

        .priority-sticky-header .section-label {
            margin-bottom: 0.18rem;
        }

        .priority-sticky-header .priority-chart-heading {
            font-size: 0.82rem;
            letter-spacing: 0.08em;
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
            min-height: 200px;
            height: 100%;
            display: flex;
            flex-direction: column;
        }

        .metric-card.split-card {
            min-height: 264px;
            text-align: center;
            display: flex;
            flex-direction: column;
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
            text-align: center;
            display: block;
            width: 100%;
        }

        .metric-subtitle {
            color: var(--muted);
            line-height: 1.45;
            font-size: 0.92rem;
            margin-top: auto;
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

        .split-metric-grid {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 0.7rem;
            margin-bottom: 0.6rem;
        }

        .split-metric-panel {
            background: rgba(20, 54, 66, 0.04);
            border: 1px solid rgba(20, 54, 66, 0.08);
            border-radius: 16px;
            padding: 0.6rem 0.45rem 0.55rem 0.45rem;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            text-align: center;
            min-height: 124px;
        }

        .split-metric-tag {
            text-align: center;
            color: var(--muted);
            font-size: 0.68rem;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            font-weight: 700;
            margin-bottom: 0.35rem;
        }

        .split-card .metric-value {
            margin-bottom: 0;
            font-size: 1.75rem;
        }

        .split-card .metric-label,
        .split-card .metric-subtitle {
            text-align: center;
            width: 100%;
        }

        .split-card .metric-subtitle {
            min-height: 4.3rem;
            display: flex;
            align-items: flex-start;
            justify-content: center;
            margin-top: auto;
        }

        .metric-value.tone-neutral {
            color: var(--ink) !important;
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
            min-height: 360px;
            height: 100%;
            display: flex;
            flex-direction: column;
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

        .area-score.tone-risk {
            color: var(--risk);
        }

        .area-score.tone-warn {
            color: var(--accent);
        }

        .area-score.tone-good {
            color: var(--good);
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
            min-height: 3.2rem;
        }

        .area-insight {
            color: var(--brand);
            font-size: 0.9rem;
            line-height: 1.45;
            margin-top: auto;
            padding-top: 0.75rem;
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

        .tracker-list {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(360px, 1fr));
            gap: 0.9rem;
            padding: 1rem 1.1rem 1.15rem 1.1rem;
            background: #f8fbfc;
        }

        .tracker-row-card {
            background: rgba(255, 255, 255, 0.96);
            border: 1px solid rgba(20, 54, 66, 0.08);
            border-radius: 18px;
            padding: 0.95rem 1rem;
            box-shadow: 0 10px 24px rgba(15, 31, 44, 0.05);
        }

        .detail-card-grid {
            display: grid;
            gap: 0.8rem;
        }

        .detail-card {
            background: rgba(255, 255, 255, 0.96);
            border: 1px solid rgba(20, 54, 66, 0.08);
            border-radius: 18px;
            padding: 0.95rem 1rem;
            box-shadow: 0 10px 24px rgba(15, 31, 44, 0.05);
        }

        .detail-card-top {
            display: flex;
            align-items: flex-start;
            justify-content: space-between;
            gap: 0.8rem;
            margin-bottom: 0.55rem;
        }

        .detail-card-heading {
            display: flex;
            align-items: flex-start;
            gap: 0.75rem;
            min-width: 0;
        }

        .detail-card-body {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 0.55rem 0.85rem;
        }

        .detail-card-item {
            background: rgba(20, 54, 66, 0.04);
            border-radius: 12px;
            padding: 0.5rem 0.6rem;
        }

        .detail-card-label {
            display: block;
            color: var(--muted);
            font-size: 0.68rem;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            font-weight: 700;
            margin-bottom: 0.18rem;
        }

        .detail-card-value {
            color: var(--ink);
            font-size: 0.88rem;
            line-height: 1.35;
            font-weight: 700;
        }

        @media (max-width: 900px) {
            .detail-card-body {
                grid-template-columns: 1fr;
            }
        }

        .tracker-row-top {
            display: flex;
            align-items: flex-start;
            justify-content: space-between;
            gap: 0.8rem;
            margin-bottom: 0.55rem;
        }

        .tracker-row-heading {
            display: flex;
            align-items: flex-start;
            gap: 0.75rem;
            min-width: 0;
        }

        .tracker-step-badge {
            flex: 0 0 auto;
            min-width: 48px;
            height: 48px;
            border-radius: 14px;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            background: rgba(20, 54, 66, 0.08);
            color: var(--brand);
            font-size: 0.8rem;
            font-weight: 800;
            line-height: 1.1;
            text-align: center;
        }

        .tracker-card-title {
            color: var(--ink);
            font-weight: 800;
            line-height: 1.35;
            font-size: 1rem;
            margin-bottom: 0.22rem;
        }

        .tracker-card-meta {
            display: grid;
            grid-template-columns: repeat(2, minmax(0, 1fr));
            gap: 0.55rem 0.85rem;
            margin-top: 0.75rem;
        }

        .tracker-card-meta-item {
            background: rgba(20, 54, 66, 0.04);
            border-radius: 12px;
            padding: 0.5rem 0.6rem;
        }

        .tracker-card-meta-label {
            display: block;
            color: var(--muted);
            font-size: 0.68rem;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            font-weight: 700;
            margin-bottom: 0.18rem;
        }

        .tracker-card-meta-value {
            color: var(--ink);
            font-size: 0.88rem;
            line-height: 1.35;
            font-weight: 700;
        }

        .tracker-card-action {
            margin-top: 0.75rem;
            padding-top: 0.75rem;
            border-top: 1px solid rgba(20, 54, 66, 0.08);
        }

        .tracker-card-action .tracker-card-meta-label {
            margin-bottom: 0.28rem;
        }

        @media (max-width: 900px) {
            .tracker-card-meta {
                grid-template-columns: 1fr;
            }
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
            align-items: flex-start;
            justify-content: space-between;
            gap: 0.75rem;
            margin-bottom: 0.45rem;
            flex-wrap: wrap;
        }

        .tracker-note-title {
            color: var(--ink);
            font-size: 0.95rem;
            font-weight: 800;
            line-height: 1.35;
            flex: 1 1 180px;
            min-width: 0;
        }

        .tracker-note-top .status-pill {
            flex: 0 0 auto;
            max-width: 100%;
            white-space: normal;
            text-align: center;
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
    display["Deadline Sort"] = pd.to_datetime(display["Deadline"], errors="coerce")
    display = display.sort_values(
        ["Step", "Deadline Sort"], ascending=[True, True], na_position="last"
    ).drop(columns=["Deadline Sort"])

    cards = []
    if display.empty:
        cards.append(
            dedent(
                """
                <div class="tracker-row-card">
                    <div class="tracker-card-title">No tasks match the selected filters</div>
                    <div class="tracker-subtask">Change the status, owner, or deadline pills to broaden the checklist view.</div>
                </div>
                """
            ).strip()
        )

    for _, row in display.iterrows():
        cards.append(
            dedent(
                f"""
                <div class="tracker-row-card">
                    <div class="tracker-row-top">
                        <div class="tracker-row-heading">
                            <div class="tracker-step-badge">Step<br>{int(row['Step'])}</div>
                            <div>
                                <div class="tracker-card-title">{escape(str(row['Task']))}</div>
                                <div class="tracker-subtask">{escape(str(row['Category']))} · {escape(str(row['Owner']))} · {escape(str(row['Deadline']))}</div>
                            </div>
                        </div>
                        {format_status_pill(row['Status'])}
                    </div>
                    <div class="tracker-card-meta">
                        <div class="tracker-card-meta-item">
                            <span class="tracker-card-meta-label">Dependency</span>
                            <div class="tracker-card-meta-value">{escape(str(row['Dependency Task']))}</div>
                            <div class="tracker-subtask">{escape(str(row['Dependency Status']))} · {escape(str(row['Dependency Owner']))}</div>
                        </div>
                        <div class="tracker-card-meta-item">
                            <span class="tracker-card-meta-label">Automation Potential</span>
                            <div class="tracker-card-meta-value">{format_auto_pill(row['Automation Potential'])}</div>
                        </div>
                        <div class="tracker-card-meta-item">
                            <span class="tracker-card-meta-label">Owner</span>
                            <div class="tracker-card-meta-value">{escape(str(row['Owner']))}</div>
                        </div>
                        <div class="tracker-card-meta-item">
                            <span class="tracker-card-meta-label">Estimated Hours</span>
                            <div class="tracker-card-meta-value">{format_hours_short(row['Estimated Hours'])}</div>
                        </div>
                    </div>
                    <div class="tracker-card-action">
                        <span class="tracker-card-meta-label">Next Action</span>
                        <div class="tracker-subtask">{escape(str(row['Next Action']))}</div>
                    </div>
                </div>
                """
            ).strip()
        )

    return dedent(
        f"""
        <div class="tracker-board">
            <div class="tracker-header">
                <div>
                    <div class="tracker-kicker">Close Checklist Tracker</div>
                    <div class="tracker-title">Checklist task tracker</div>
                    <div class="tracker-copy">
                        Use the filters to focus the close by status, owner, and deadline without dropping into a spreadsheet view.
                    </div>
                </div>
            </div>
            <div class="tracker-list">{''.join(cards)}</div>
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
    if handoff_queue.empty:
        return (
            "<div class='tracker-notes'>"
            "<div class='tracker-note'>"
            "<div class='tracker-note-top'>"
            "<div class='tracker-note-title'>No handoffs currently queued</div>"
            "</div>"
            "<div class='tracker-note-copy'>The current checklist state does not have any blocked or waiting handoffs that need immediate escalation.</div>"
            "</div>"
            "</div>"
        )

    sorted_queue = handoff_queue.sort_values("Step", ascending=True)
    cards = []
    for _, row in sorted_queue.head(3).iterrows():
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


def build_checklist_detail_cards_html(frame, mode):
    if frame.empty:
        return (
            "<div class='detail-card-grid'>"
            "<div class='detail-card'>"
            "<div class='tracker-card-title'>Nothing queued</div>"
            "<div class='tracker-subtask'>There are no checklist items to show in this section right now.</div>"
            "</div>"
            "</div>"
        )

    display = frame.copy()
    if "Step" in display.columns:
        display = display.sort_values("Step", ascending=True)

    cards = []
    for _, row in display.iterrows():
        if mode == "critical":
            detail_items = [
                ("Category", row["Category"]),
                ("Owner", row["Owner"]),
                ("Status", row["Status"]),
                ("Dependency", row["Dependency Task"]),
                ("Downstream Open", format_number(row["Downstream Open Tasks"], 0)),
                ("Estimated Hours", format_hours_short(row["Estimated Hours"])),
                ("Action Lane", row["Action Lane"]),
                ("Unblock Score", format_number(row["Unblock Score"], 0)),
            ]
        else:
            detail_items = [
                ("Category", row["Category"]),
                ("Owner", row["Owner"]),
                ("Priority", row["Priority"]),
                ("Dependency", row["Dependency Task"]),
                ("Dependency Status", row["Dependency Status"]),
                ("Dependency Owner", row["Dependency Owner"]),
                ("Estimated Hours", format_hours_short(row["Estimated Hours"])),
                ("Action Lane", row["Action Lane"]),
            ]

        item_html = "".join(
            dedent(
                f"""
                <div class="detail-card-item">
                    <span class="detail-card-label">{escape(str(label))}</span>
                    <div class="detail-card-value">{escape(str(value))}</div>
                </div>
                """
            ).strip()
            for label, value in detail_items
        )

        cards.append(
            dedent(
                f"""
                <div class="detail-card">
                    <div class="detail-card-top">
                        <div class="detail-card-heading">
                            <div class="tracker-step-badge">Step<br>{int(row['Step'])}</div>
                            <div>
                                <div class="tracker-card-title">{escape(str(row['Task']))}</div>
                                <div class="tracker-subtask">{escape(str(row['Next Action']))}</div>
                            </div>
                        </div>
                        {format_status_pill(row['Status'])}
                    </div>
                    <div class="detail-card-body">{item_html}</div>
                </div>
                """
            ).strip()
        )

    return f"<div class='detail-card-grid'>{''.join(cards)}</div>"


def format_display_frame(df):
    frame = df.copy()
    currency_label_overrides = {"Unresolved Difference", "Open AR Balance"}
    centered_label_columns = {
        "Journal Entry ID",
        "Entity",
        "Match Status",
        "Source",
        "Status",
        "Posting Date",
        "Days Outstanding",
        "Sending Entity",
        "Receiving Entity",
        "Elimination Status",
        "Notes",
        "Accrual ID",
        "Risk Flag",
        "Document ID",
        "3-Way Match",
        "Approval Status",
        "Category",
        "Owner",
        "Deadline",
        "Priority",
        "Estimated Hours",
        "Variance %",
        "Reconciled",
        "Disposal Status",
        "Last Physical Verification",
    }
    percent_columns = []
    currency_columns = []
    hours_columns = []
    numeric_columns = []

    for column in frame.columns:
        if not pd.api.types.is_numeric_dtype(frame[column]):
            continue
        col_name = str(column)
        if "%" in col_name or "Rate" in col_name:
            percent_columns.append(column)
        elif "(EUR)" in col_name or column in currency_label_overrides:
            currency_columns.append(column)
        elif "Hours" in col_name or "hrs" in col_name:
            hours_columns.append(column)
        else:
            numeric_columns.append(column)

    centered_columns = [column for column in frame.columns if str(column) in centered_label_columns]
    right_aligned_columns = percent_columns + currency_columns + hours_columns + numeric_columns
    left_aligned_columns = [
        column for column in frame.columns if column not in centered_columns and column not in right_aligned_columns
    ]

    if not (percent_columns or currency_columns or hours_columns or numeric_columns or centered_columns):
        return frame

    formatters = {}
    for column in percent_columns:
        formatters[column] = lambda x: format_percent(x, 2)
    for column in currency_columns:
        formatters[column] = lambda x: format_number(x, 2)
    for column in hours_columns:
        formatters[column] = lambda x: format_number(x, 2)
    for column in numeric_columns:
        formatters[column] = lambda x: format_number(x, 0)

    styler = frame.style.format(formatters, na_rep="")

    if left_aligned_columns:
        styler = styler.set_properties(subset=left_aligned_columns, **{"text-align": "left"})
    if right_aligned_columns:
        styler = styler.set_properties(subset=right_aligned_columns, **{"text-align": "right"})
    if centered_columns:
        styler = styler.set_properties(subset=centered_columns, **{"text-align": "center"})

        header_styles = []
        for column in centered_columns:
            try:
                column_index = frame.columns.get_loc(column)
            except KeyError:
                continue
            header_styles.append(
                {
                    "selector": f"th.col_heading.level0.col{column_index}",
                    "props": [("text-align", "center")],
                }
            )
        if header_styles:
            styler = styler.set_table_styles(header_styles, overwrite=False)

    return styler


def render_metric_card(title, value, subtitle, tone="medium"):
    display_value = value
    if isinstance(value, (int, float, np.integer, np.floating)):
        decimals = 2 if isinstance(value, (float, np.floating)) and not float(value).is_integer() else 0
        display_value = format_number(value, decimals)
    st.markdown(
        f"""
        <div class="metric-card {tone}">
            <div class="metric-label">{title}</div>
            <div class="metric-value">{display_value}</div>
            <div class="metric-subtitle">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_before_after_metric_card(
    title,
    before_value,
    after_value,
    subtitle,
    before_tone="medium",
    after_tone="low",
    before_color=True,
    after_color=True,
):
    before_display = format_metric_value(before_value)
    after_display = format_metric_value(after_value)
    before_class = before_tone if before_color else "tone-neutral"
    after_class = after_tone if after_color else "tone-neutral"
    st.markdown(
        f"""
        <div class="metric-card split-card">
            <div class="metric-label">{title}</div>
            <div class="split-metric-grid">
                <div class="split-metric-panel">
                    <div class="split-metric-tag">Before</div>
                    <div class="metric-value {before_class}">{before_display}</div>
                </div>
                <div class="split-metric-panel">
                    <div class="split-metric-tag">After</div>
                    <div class="metric-value {after_class}">{after_display}</div>
                </div>
            </div>
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

    def tone_class(value, warn_threshold=1, high_threshold=5, good_when_zero=True):
        if good_when_zero and value == 0:
            return "tone-good"
        if value >= high_threshold:
            return "tone-risk"
        if value >= warn_threshold:
            return "tone-warn"
        return "tone-good"

    pending_tone = tone_class(pending_actions, warn_threshold=1, high_threshold=4, good_when_zero=True)
    risk_tone = tone_class(high_risk_items, warn_threshold=1, high_threshold=3, good_when_zero=True)
    ai_tone = "tone-ai" if ai_candidates > 0 else "tone-good"
    if match_rate >= 90:
        match_tone = "tone-good"
    elif match_rate >= 80:
        match_tone = "tone-warn"
    else:
        match_tone = "tone-risk"

    st.markdown(
        f"""
        <div class="kpi-strip">
            <div class="kpi-tile">
                <div class="kpi-tile-label">Close Day Target</div>
                <div class="kpi-tile-value tone-good">4</div>
                <div class="kpi-tile-copy">From 7 business days</div>
            </div>
            <div class="kpi-tile">
                <div class="kpi-tile-label">Pending Actions</div>
                <div class="kpi-tile-value {pending_tone}">{format_number(pending_actions, 0)}</div>
                <div class="kpi-tile-copy">{format_number(summary['not_started_tasks'], 0)} not started · {format_number(summary['blocked_tasks'], 0)} blocked</div>
            </div>
            <div class="kpi-tile">
                <div class="kpi-tile-label">High Risk Items</div>
                <div class="kpi-tile-value {risk_tone}">{format_number(high_risk_items, 0)}</div>
                <div class="kpi-tile-copy">{format_number(summary['accrual_risks'], 0)} accruals · {format_number(summary['ap_3way_exceptions'], 0)} AP</div>
            </div>
            <div class="kpi-tile">
                <div class="kpi-tile-label">Auto-Reconciled</div>
                <div class="kpi-tile-value {match_tone}">{format_percent(match_rate, 2)}</div>
                <div class="kpi-tile-copy">{format_number(bank_matched, 0)} of {format_number(bank_total, 0)} bank items</div>
            </div>
            <div class="kpi-tile">
                <div class="kpi-tile-label">AI Candidates</div>
                <div class="kpi-tile-value {ai_tone}">{format_number(ai_candidates, 0)}</div>
                <div class="kpi-tile-copy">{format_number(summary['bank_journal_candidates'], 0)} bank drafts · {format_number(summary['journal_agent_ready_drafts'], 0)} journals</div>
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
        ("Day 1", "Pre-close", "Cut-off, source confirmations, and period readiness."),
        ("Day 2", "Recons & accruals", "Cash, subledgers, and recurring month-end entries."),
        ("Day 3", "IC elim & TB review", "Intercompany cleanup, elimination, and variance review."),
        ("Day 4", "Pack & sign-off", "Management pack, controller review, and final close."),
    ]
    steps = []
    for index, (day_label, title, copy) in enumerate(labels, start=1):
        if index < day_marker:
            state = "complete"
            node_label = "✓"
        elif index == day_marker:
            state = "current"
            node_label = str(index)
        else:
            state = "future"
            node_label = str(index)
        steps.append(
            dedent(
                f"""
                <div class="timeline-step {state}">
                    <div class="timeline-node">{node_label}</div>
                    <div class="timeline-card">
                        <div class="timeline-day">{day_label}</div>
                        <div class="timeline-copy"><strong>{title}</strong><br>{copy}</div>
                    </div>
                </div>
                """
            ).strip()
        )

    st.markdown(
        f"""
        <div class="timeline">
            <div class="timeline-head">
                <div class="timeline-title">Close Day Timeline — Target: 4 Days</div>
                <div class="timeline-status">Day {day_marker} of 4</div>
            </div>
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


def render_focus_board(kpis, score, priorities, recommendations):
    summary = kpis["summary"]
    posting = kpis["agents"]["erp_posting"]["summary"]
    checklist_agent = kpis["agents"]["checklist"]
    top_priorities = priorities[:3]
    recs = recommendations

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


def render_me_steps(kpis, actions, panel=False):
    automation_run = st.session_state.get("automation_run", False)
    automation_summary = st.session_state.get("automation_summary", {})
    primary_color = "var(--good)" if automation_run else "#ff4d4f"

    st.markdown("<div class='section-label'>Month-End Steps</div>", unsafe_allow_html=True)
    st.markdown(
        f"""
        <style>
        div[data-testid="stButton"] > button[kind="primary"] {{
            background: {primary_color};
            border-color: {primary_color};
            color: #ffffff;
            box-shadow: none;
        }}

        div[data-testid="stButton"] > button[kind="primary"]:hover {{
            background: {primary_color};
            border-color: {primary_color};
            color: #ffffff;
            filter: brightness(0.97);
        }}

        div[data-testid="stButton"] > button[kind="primary"]:disabled {{
            background: rgba(15, 31, 44, 0.08);
            border-color: rgba(15, 31, 44, 0.10);
            color: #7b8893;
            filter: none;
            cursor: not-allowed;
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )
    step_card_class = "focus-card me-step-card panel" if panel else "focus-card me-step-card"

    def render_step_one():
        st.markdown(
            f"""
            <div class="{step_card_class}">
                <div class="focus-title">Step 1 · Load the data</div>
                <div class="focus-list">
                    <div class="focus-item"><strong>Status:</strong> Data loaded and validated.</div>
                    <div class="focus-item">Dataset ready for close review.</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown(
            "<div class='me-step-footer'><span class='me-step-pill done'>Complete</span></div>",
            unsafe_allow_html=True,
        )

    def render_step_two():
        st.markdown(
            f"""
            <div class="{step_card_class}">
                <div class="focus-title">Step 2 · Run automations</div>
                <div class="focus-list">
                    <div class="focus-item"><strong>Status:</strong> Execute the automation bundle.</div>
                    <div class="focus-item">Applies GL, bank, IC, journal, AP, and checklist automations.</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if automation_run:
            st.markdown(
                "<div class='me-step-footer'><span class='me-step-pill done'>Completed</span></div>",
                unsafe_allow_html=True,
            )
        else:
            if st.button("Run all automations", key="me_run_all", use_container_width=True, type="primary"):
                run_all_automations(kpis, actions)
                safe_rerun()

    def render_step_three():
        status_line = "Waiting for Step 2" if not automation_run else "Work remaining priorities"
        forecast = automation_summary.get("forecast", "--")
        st.markdown(
            f"""
            <div class="{step_card_class}">
                <div class="focus-title">Step 3 · Work the remaining priorities</div>
                <div class="focus-list">
                    <div class="focus-item"><strong>Status:</strong> {status_line}</div>
                    <div class="focus-item">Use the Priorities tab to review the next owner actions and close what automation did not finish.</div>
                </div>
                <span class="focus-chip">Forecast: {forecast} days</span>
            </div>
            """,
            unsafe_allow_html=True,
        )
        open_priorities = st.button(
            "Open Priorities",
            key="me_open_priorities",
            use_container_width=True,
            type="primary",
            disabled=not automation_run,
        )
        if open_priorities:
            st.session_state["nav_target"] = "Priorities"
            safe_rerun()

    if panel:
        st.markdown("<div class='me-step-panel'>", unsafe_allow_html=True)
        render_step_one()
        render_step_two()
        render_step_three()
        st.markdown("</div>", unsafe_allow_html=True)
        return

    step_cols = st.columns(3, gap="medium")
    with step_cols[0]:
        render_step_one()
    with step_cols[1]:
        render_step_two()
    with step_cols[2]:
        render_step_three()


def render_action_queue(kpis):
    queue = kpis["agents"]["action_queue"].copy()
    if queue.empty:
        return

    if "Deadline" in queue.columns:
        queue["__deadline_sort"] = pd.to_datetime(queue["Deadline"], errors="coerce")
        queue = queue.sort_values("__deadline_sort", ascending=True, na_position="last").drop(columns="__deadline_sort")

    st.markdown(
        """
        <div class="queue-row">
            <div class="queue-title">Team Action Queue (next owner step)</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.dataframe(
        format_display_frame(queue),
        use_container_width=True,
        hide_index=True,
    )


def area_score_tone(severity_score):
    if severity_score >= 85:
        return "tone-risk"
    if severity_score >= 50:
        return "tone-warn"
    return "tone-good"


def render_area_card(area):
    metrics = [f"{label}: {value}" for label, value in area["metrics"].items()]
    tone_class = area_score_tone(area["severity_score"])
    st.markdown(
        f"""
        <div class="area-card">
            <div class="area-header">
                <div class="area-title">{area['label']}</div>
                <div class="area-score {tone_class}">{area['severity_score']}</div>
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


def render_area_card_compact(area):
    metrics = [f"{label}: {value}" for label, value in area["metrics"].items()]
    tone_class = area_score_tone(area["severity_score"])
    st.markdown(
        f"""
        <div class="area-card compact">
            <div class="area-header">
                <div class="area-title">{area['label']}</div>
                <div class="area-score {tone_class}">{area['severity_score']}</div>
            </div>
            <div class="area-headline">{area['headline']}</div>
            {format_chip_row(metrics)}
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_agent_snapshot_tables(kpis):
    cols = st.columns(2, gap="medium")
    with cols[0]:
        st.markdown("<div class='section-label'>GL approvals</div>", unsafe_allow_html=True)
        st.dataframe(
            format_display_frame(kpis["agents"]["gl"]["worklist"].head(5)),
            use_container_width=True,
            hide_index=True,
        )
        st.markdown("<div class='section-label'>Bank exceptions</div>", unsafe_allow_html=True)
        st.dataframe(
            format_display_frame(kpis["agents"]["bank"]["worklist"].head(5)),
            use_container_width=True,
            hide_index=True,
        )
        st.markdown("<div class='section-label'>IC exceptions</div>", unsafe_allow_html=True)
        st.dataframe(
            format_display_frame(kpis["agents"]["ic"]["worklist"].head(5)),
            use_container_width=True,
            hide_index=True,
        )
    with cols[1]:
        st.markdown("<div class='section-label'>Journal drafts</div>", unsafe_allow_html=True)
        st.dataframe(
            format_display_frame(kpis["agents"]["journal"]["erp_journal_drafts"].head(5)),
            use_container_width=True,
            hide_index=True,
        )
        st.markdown("<div class='section-label'>Checklist handoffs</div>", unsafe_allow_html=True)
        st.dataframe(
            format_display_frame(kpis["agents"]["checklist"]["handoff_queue"].head(5)),
            use_container_width=True,
            hide_index=True,
        )


def render_before_after_snapshot(base_kpis, after_kpis, after_label):
    focus_areas = ["gl", "bank", "intercompany", "accruals", "ap", "checklist"]
    before_cols = st.columns(2, gap="large")

    with before_cols[0]:
        st.markdown("<div class='snapshot-block'>", unsafe_allow_html=True)
        st.markdown("<div class='snapshot-title'>Before automation</div>", unsafe_allow_html=True)
        area_cols = st.columns(2, gap="small")
        for idx, key in enumerate(focus_areas):
            with area_cols[idx % 2]:
                render_area_card_compact(base_kpis["areas"][key])
        st.markdown("<div class='section-label'>Agent tables</div>", unsafe_allow_html=True)
        render_agent_snapshot_tables(base_kpis)
        st.markdown("</div>", unsafe_allow_html=True)

    with before_cols[1]:
        st.markdown("<div class='snapshot-block'>", unsafe_allow_html=True)
        st.markdown(f"<div class='snapshot-title'>After automation · {escape(after_label)}</div>", unsafe_allow_html=True)
        area_cols = st.columns(2, gap="small")
        for idx, key in enumerate(focus_areas):
            with area_cols[idx % 2]:
                render_area_card_compact(after_kpis["areas"][key])
        st.markdown("<div class='section-label'>Agent tables</div>", unsafe_allow_html=True)
        render_agent_snapshot_tables(after_kpis)
        st.markdown("</div>", unsafe_allow_html=True)


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


def build_readiness_gauge(score, title="Close Readiness"):
    if not HAS_PLOTLY:
        return None

    figure = go.Figure(
        go.Indicator(
            mode="gauge+number",
            value=score["readiness_score"],
            number={"suffix": "/100", "font": {"size": 34, "color": "#143642"}},
            title={"text": title, "font": {"size": 16, "color": "#5b6b78"}},
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


def style_bar_chart(
    figure,
    *,
    height=320,
    left=48,
    right=48,
    top=24,
    bottom=56,
    showlegend=None,
    x_tickangle=None,
):
    layout_kwargs = {
        "height": height,
        "margin": dict(l=left, r=right, t=top, b=bottom),
        "paper_bgcolor": "rgba(0,0,0,0)",
        "plot_bgcolor": "rgba(0,0,0,0)",
        "font": dict(color="#4a5966", size=13),
        "uniformtext_minsize": 10,
        "uniformtext_mode": "hide",
    }
    if showlegend is not None:
        layout_kwargs["showlegend"] = showlegend
    figure.update_layout(**layout_kwargs)
    figure.update_xaxes(
        automargin=True,
        tickfont=dict(color="#7f8d98", size=11),
        title_font=dict(color="#4a5966"),
        gridcolor="rgba(15, 31, 44, 0.08)",
        zerolinecolor="rgba(15, 31, 44, 0.12)",
    )
    figure.update_yaxes(
        automargin=True,
        tickfont=dict(color="#7f8d98", size=11),
        title_font=dict(color="#4a5966"),
        gridcolor="rgba(15, 31, 44, 0.08)",
        zerolinecolor="rgba(15, 31, 44, 0.12)",
    )
    if x_tickangle is not None:
        figure.update_xaxes(tickangle=x_tickangle)
    figure.update_traces(
        selector=dict(type="bar"),
        marker_line_width=0,
        textfont=dict(color="#6c7a86", size=14),
        cliponaxis=False,
    )


def style_donut_chart(figure, *, height=320, left=18, right=120, top=12, bottom=12):
    figure.update_layout(
        height=height,
        margin=dict(l=left, r=right, t=top, b=bottom),
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#3b4a57", size=14),
        uniformtext_minsize=10,
        uniformtext_mode="hide",
        legend=dict(
            font=dict(color="#3b4a57", size=14),
            bgcolor="rgba(255,255,255,0.0)",
            x=1.02,
            xanchor="left",
            y=0.98,
            yanchor="top",
        ),
    )


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
        figure.update_layout(coloraxis_showscale=False)
        style_bar_chart(figure, height=360, left=165, right=70, top=20, bottom=28)
        figure.update_xaxes(title="Severity")
        figure.update_yaxes(title="Area")
        figure.update_traces(
            marker_line_width=0,
            textposition="outside",
            textfont=dict(color="#6c7a86", size=14),
            cliponaxis=False,
        )
        return figure

    return chart_data.set_index("Area")


def build_checklist_chart(summary):
    status_counts = summary.get("checklist_status_counts")
    if not status_counts:
        status_counts = {
            "Completed": int(summary.get("completed_tasks", 0)),
            "In Progress": int(summary.get("in_progress_tasks", 0)),
            "Not Started": int(summary.get("not_started_tasks", 0)),
            "Waiting on Input": int(summary.get("waiting_tasks", 0)),
            "Blocked": int(summary.get("blocked_tasks", 0)),
        }

    chart_data = pd.DataFrame(
        {
            "Status": list(status_counts.keys()),
            "Count": list(status_counts.values()),
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
                "Not Started": "#a8b5bf",
                "Waiting on Input": "#e9a03b",
                "Blocked": "#c75c5c",
            },
        )
        figure.update_layout(showlegend=True)
        style_donut_chart(figure, height=320, left=18, right=150, top=12, bottom=12)
        figure.update_traces(
            marker=dict(line=dict(color="rgba(255,255,255,0.92)", width=1.5)),
            textfont=dict(color="#0f1f2c", size=16),
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
        style_bar_chart(figure, height=320, left=40, right=36, top=20, bottom=62, showlegend=False)
        figure.update_layout(xaxis_title="", yaxis_title="")
        figure.update_traces(
            marker_line_width=0,
            textposition="outside",
            textfont=dict(color="#6c7a86", size=14),
        )
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
        figure.update_layout(coloraxis_showscale=False, xaxis_title="", yaxis_title="")
        style_bar_chart(figure, height=320, left=150, right=60, top=20, bottom=24)
        figure.update_traces(
            marker_line_width=0,
            textposition="outside",
            textfont=dict(color="#6c7a86", size=14),
            cliponaxis=False,
        )
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
        figure.update_layout(coloraxis_showscale=False, xaxis_title="", yaxis_title="")
        style_bar_chart(figure, height=320, left=40, right=36, top=20, bottom=62, x_tickangle=0)
        figure.update_traces(
            marker_line_width=0,
            textposition="outside",
            textfont=dict(color="#6c7a86", size=14),
        )
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
        figure.update_layout(coloraxis_showscale=False, xaxis_title="", yaxis_title="")
        style_bar_chart(figure, height=320, left=150, right=60, top=20, bottom=24)
        figure.update_traces(
            marker_line_width=0,
            textposition="outside",
            textfont=dict(color="#6c7a86", size=14),
            cliponaxis=False,
        )
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
        style_bar_chart(figure, height=320, left=40, right=40, top=20, bottom=76, showlegend=False, x_tickangle=-18)
        figure.update_layout(xaxis_title="", yaxis_title="")
        figure.update_traces(
            marker_line_width=0,
            textposition="outside",
            textfont=dict(color="#5f6c78", size=14),
            cliponaxis=False,
        )
        figure.update_xaxes(tickfont=dict(color="#6f7d89", size=12))
        figure.update_yaxes(tickfont=dict(color="#6f7d89", size=12))
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
        figure.update_layout(coloraxis_showscale=False, xaxis_title="", yaxis_title="")
        style_bar_chart(figure, height=320, left=220, right=70, top=20, bottom=24)
        figure.update_traces(
            texttemplate="%{text} open",
            textposition="outside",
            marker_line_width=0,
            textfont=dict(color="#5f6c78", size=14),
            cliponaxis=False,
        )
        figure.update_xaxes(tickfont=dict(color="#6f7d89", size=12))
        figure.update_yaxes(tickfont=dict(color="#6f7d89", size=12))
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
        figure.update_layout(showlegend=True)
        style_donut_chart(figure, height=320, left=18, right=170, top=12, bottom=12)
        figure.update_traces(
            marker=dict(line=dict(color="rgba(255,255,255,0.92)", width=1.5)),
            textfont=dict(color="#ffffff", size=16),
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
        figure.update_layout(coloraxis_showscale=False, xaxis_title="", yaxis_title="")
        style_bar_chart(figure, height=320, left=150, right=60, top=20, bottom=24)
        figure.update_traces(
            marker_line_width=0,
            textposition="outside",
            textfont=dict(color="#6c7a86", size=14),
            cliponaxis=False,
        )
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
        style_bar_chart(figure, height=320, left=40, right=40, top=20, bottom=72, showlegend=False, x_tickangle=-15)
        figure.update_layout(xaxis_title="", yaxis_title="")
        figure.update_traces(
            marker_line_width=0,
            textposition="outside",
            textfont=dict(color="#6c7a86", size=14),
        )
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
        figure.update_layout(coloraxis_showscale=False, xaxis_title="", yaxis_title="")
        style_bar_chart(figure, height=320, left=150, right=70, top=20, bottom=24)
        figure.update_traces(
            texttemplate="%{text}/100",
            textposition="outside",
            marker_line_width=0,
            textfont=dict(color="#6c7a86", size=14),
            cliponaxis=False,
        )
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
        style_bar_chart(figure, height=320, left=40, right=40, top=20, bottom=72, showlegend=False, x_tickangle=-15)
        figure.update_layout(xaxis_title="", yaxis_title="")
        figure.update_traces(
            marker_line_width=0,
            textposition="outside",
            textfont=dict(color="#6c7a86", size=14),
        )
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
            xaxis_title="",
            yaxis_title="",
            legend=dict(
                font=dict(color="#4a5966", size=13),
                bgcolor="rgba(255,255,255,0.0)",
                title_font=dict(color="#4a5966", size=13),
                x=1.02,
                xanchor="left",
                y=1,
                yanchor="top",
            ),
        )
        style_bar_chart(figure, height=320, left=160, right=110, top=20, bottom=24)
        figure.update_traces(
            marker_line_width=0,
            textposition="outside",
            textfont=dict(color="#6c7a86", size=14),
            cliponaxis=False,
        )
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
        figure.update_layout(coloraxis_showscale=False, xaxis_title="", yaxis_title="")
        style_bar_chart(figure, height=320, left=40, right=40, top=20, bottom=84, x_tickangle=-24)
        figure.update_traces(
            marker_line_width=0,
            textposition="outside",
            textfont=dict(color="#6c7a86", size=14),
        )
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
        figure.update_layout(coloraxis_showscale=False, xaxis_title="", yaxis_title="")
        style_bar_chart(figure, height=320, left=145, right=65, top=20, bottom=24)
        figure.update_traces(
            texttemplate="%{text} open",
            textposition="outside",
            marker_line_width=0,
            textfont=dict(color="#6c7a86", size=14),
            cliponaxis=False,
        )
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
        figure.update_layout(coloraxis_showscale=False)
        style_bar_chart(figure, height=360, left=175, right=70, top=20, bottom=28)
        figure.update_xaxes(title="Risk Score")
        figure.update_yaxes(title="Entity")
        figure.update_traces(
            textposition="outside",
            marker_line_width=0,
            textfont=dict(color="#6c7a86", size=14),
            cliponaxis=False,
        )
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
        figure.update_layout(coloraxis_showscale=False)
        style_bar_chart(figure, height=400, left=210, right=85, top=20, bottom=28)
        figure.update_xaxes(title="Priority Score")
        figure.update_yaxes(title="Priority Item")
        figure.update_xaxes(automargin=True, range=[0, max(chart_data["priority_score"].max() + 1.5, 10)])
        figure.update_yaxes(automargin=True)
        figure.update_traces(
            texttemplate="%{text} hrs",
            textposition="outside",
            marker_line_width=0,
            cliponaxis=False,
            textfont=dict(color="#6c7a86", size=14),
        )
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


def render_command_center(
    kpis,
    score,
    priorities,
    commentary,
    actions,
    recommendations,
    baseline_score,
    automation_score,
    automation_run,
):
    summary = kpis["summary"]
    baseline_summary = baseline_score.get("summary_snapshot", kpis["summary"])
    after_summary = automation_score.get("summary_snapshot", kpis["summary"])
    riskiest_entity = kpis["entities"][0]
    baseline_entity_label = baseline_summary["riskiest_entity"].replace("NovaTech ", "")
    after_entity_label = after_summary["riskiest_entity"].replace("NovaTech ", "")

    def forecast_tone(days):
        return "low" if days <= 4 else "high"

    def blocker_tone(value):
        if value == 0:
            return "low"
        if value <= 50:
            return "medium"
        return "high"

    render_kpi_strip(summary)
    st.markdown("<div class='section-label section-heading tight-top'>Close Readiness</div>", unsafe_allow_html=True)
    before_col, after_col = st.columns(2, gap="medium")
    with before_col:
        st.markdown("<div class='section-label centered'>Before Automation</div>", unsafe_allow_html=True)
        gauge = build_readiness_gauge(baseline_score, "")
        if gauge is not None:
            st.plotly_chart(
                gauge,
                use_container_width=True,
                config={"displayModeBar": False},
                key="command_center_readiness_gauge_before",
            )
        else:
            st.info("Install `plotly` to unlock the full readiness gauge.")
            st.progress(baseline_score["readiness_score"] / 100)

    with after_col:
        st.markdown("<div class='section-label centered'>Projected After Automation</div>", unsafe_allow_html=True)
        gauge = build_readiness_gauge(automation_score, "")
        if gauge is not None:
            st.plotly_chart(
                gauge,
                use_container_width=True,
                config={"displayModeBar": False},
                key="command_center_readiness_gauge_after",
            )
        else:
            st.info("Install `plotly` to unlock the full readiness gauge.")
            st.progress(automation_score["readiness_score"] / 100)

    card_cols = st.columns(4)
    with card_cols[0]:
        render_before_after_metric_card(
            "Close Readiness",
            f"{baseline_score['readiness_score']}/100",
            f"{automation_score['readiness_score']}/100",
            "Readiness score before and after the automation bundle.",
            "high",
            "low",
        )
    with card_cols[1]:
        render_before_after_metric_card(
            "Predicted Close",
            f"{baseline_score['predicted_close_days']} days",
            f"{automation_score['predicted_close_days']} days",
            "Dashboard close forecast before and after automation.",
            "high",
            "low",
        )
    with card_cols[2]:
        render_before_after_metric_card(
            "Top Operational Blocker",
            f"{baseline_summary['pending_gl']}",
            f"{after_summary['pending_gl']}",
            "Pending GL approvals and reviews before and after automation.",
            "high",
            "low",
        )
    with card_cols[3]:
        render_before_after_metric_card(
            "Highest-Risk Entity",
            baseline_entity_label,
            after_entity_label,
            "Highest-risk entity before and after the automation bundle.",
            "high",
            "low",
        )

    top_left, top_right = st.columns([1.45, 0.9], gap="large")
    with top_left:
        render_focus_board(kpis, score, priorities, recommendations)
    with top_right:
        render_me_steps(kpis, actions, panel=True)

    render_entity_callout(riskiest_entity)

    render_action_queue(kpis)

    render_close_intelligence(summary, kpis)


def render_risk_atlas(kpis):
    area_chart = build_area_chart(kpis)
    sorted_area_keys = sorted(
        AREA_ORDER,
        key=lambda key: kpis["areas"][key]["severity_score"],
        reverse=True,
    )
    if HAS_PLOTLY:
        st.plotly_chart(
            area_chart,
            use_container_width=True,
            config={"displayModeBar": False},
            key="risk_atlas_area_chart",
        )
    else:
        st.bar_chart(area_chart)

    for start in range(0, len(sorted_area_keys), 4):
        cols = st.columns(4, gap="medium")
        for offset, area_key in enumerate(sorted_area_keys[start : start + 4]):
            with cols[offset]:
                render_area_card(kpis["areas"][area_key])


def render_risk_details(kpis):
    options = {AREA_LABELS[key]: key for key in AREA_ORDER}
    selected_label = st.selectbox("Inspect detailed evidence", list(options.keys()), index=0)
    selected_key = options[selected_label]
    detail_key = DETAIL_MAP[selected_key]
    detail_frame = format_display_frame(kpis["details"][detail_key])
    st.dataframe(detail_frame, use_container_width=True, hide_index=True)


def render_entity_details(kpis):
    entities = kpis["entities"]
    entity_options = [entity["entity"] for entity in entities]
    selected_entity = st.selectbox("Inspect entity evidence", entity_options, index=0)
    entity = next(item for item in entities if item["entity"] == selected_entity)

    detail_cols = st.columns(4, gap="medium")
    with detail_cols[0]:
        render_metric_card(
            "Pending GL",
            str(entity["pending_gl"]),
            "Approval and review items still open in this legal entity.",
            "high" if entity["pending_gl"] > 0 else "low",
        )
    with detail_cols[1]:
        render_metric_card(
            "Bank Exceptions",
            str(entity["bank_exceptions"]),
            "Cash items still unresolved in this entity's bank reconciliation.",
            "high" if entity["bank_exceptions"] > 0 else "low",
        )
    with detail_cols[2]:
        render_metric_card(
            "TB Variances",
            str(entity["tb_large_variances"]),
            "Large trial balance movements needing controller review.",
            "high" if entity["tb_large_variances"] > 0 else "low",
        )
    with detail_cols[3]:
        render_metric_card(
            "AP Exceptions",
            str(entity["ap_exceptions"]),
            "3-way match or approval exceptions still open.",
            "medium" if entity["ap_exceptions"] > 0 else "low",
        )

    chip_values = [
        f"Primary blocker: {entity['primary_blocker']}",
        f"Risk level: {entity['risk_level']}",
        f"Manual JEs: {format_number(entity['manual_jes'], 0)}",
        f"Accrual risks: {format_number(entity['accrual_risks'], 0)}",
        f"IC exceptions: {format_number(entity['ic_exceptions'], 0)}",
        f"IC unresolved diff: EUR {format_number(entity['ic_total_diff'], 2)}",
        f"AR overdue: {format_number(entity['ar_overdue'], 0)}",
        f"FA flags: {format_number(entity['fa_impairments'] + entity['fa_review_flags'] + entity['fa_unverified'], 0)}",
    ]
    chips = "".join(f"<span class='chip'>{escape(value)}</span>" for value in chip_values)
    st.markdown(
        f"""
        <div class="entity-callout">
            <div class="entity-name">{escape(entity['entity'])}</div>
            <div class="entity-copy">
                Risk score {format_number(entity['risk_score'], 0)}/100. Current driver: {escape(entity['driver_summary'])}.
            </div>
            <div class="chip-row">{chips}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_risks_tab(kpis):
    st.markdown(
        """
        <div class="entity-callout">
            <div class="section-label">Risks</div>
            <div class="entity-name">Review close risk in two passes</div>
            <div class="entity-copy">Use the sub-tabs to separate process-level close blockers from entity-level concentration risk.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    process_tab, entity_tab = st.tabs(["Close Process Risk", "Entity Risk"])

    with process_tab:
        st.markdown("<div class='section-label'>Close Process Risk</div>", unsafe_allow_html=True)
        render_risk_atlas(kpis)
        st.markdown("<div class='section-label'>Detailed Evidence</div>", unsafe_allow_html=True)
        render_risk_details(kpis)

    with entity_tab:
        st.markdown("<div class='section-label'>Entity Risk</div>", unsafe_allow_html=True)
        render_entity_view(kpis)
        st.markdown("<div class='section-label'>Detailed Evidence</div>", unsafe_allow_html=True)
        render_entity_details(kpis)


def render_gl_agent(kpis):
    gl_agent = kpis["agents"]["gl"]
    summary = gl_agent["summary"]

    st.markdown(
        f"""
        <div class="entity-callout">
            <div class="section-label">GL Approval Agent</div>
            <div class="entity-name">{summary['straight_through_candidates']} journals ready for straight-through approval</div>
            <div class="entity-copy">Approval lanes split low-risk batchable journals from real escalation items.</div>
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

def render_bank_agent(kpis):
    bank_agent = kpis["agents"]["bank"]
    summary = bank_agent["summary"]

    st.markdown(
        f"""
        <div class="entity-callout">
            <div class="section-label">Bank Reconciliation Agent</div>
            <div class="entity-name">{summary['open_items']} open cash exceptions triaged</div>
            <div class="entity-copy">Exception routing, ERP-ready drafts, and statement breaks in one queue.</div>
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

def render_ic_agent(kpis):
    ic_agent = kpis["agents"]["ic"]
    summary = ic_agent["summary"]

    st.markdown(
        f"""
        <div class="entity-callout">
            <div class="section-label">Intercompany Agent</div>
            <div class="entity-name">{summary['auto_matched']} IC pairs auto-matched, {summary['open_exceptions']} still need elimination support</div>
            <div class="entity-copy">Auto-matching, FX mismatch resolution, elimination drafts, and TP watchlist coverage.</div>
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

def render_journal_agent(kpis):
    journal_agent = kpis["agents"]["journal"]
    summary = journal_agent["summary"]

    st.markdown(
        f"""
        <div class="entity-callout">
            <div class="section-label">Journal Entry Agent</div>
            <div class="entity-name">{summary['erp_ready_drafts']} draft JEs ready for ERP posting</div>
            <div class="entity-copy">Accrual drafting, standard reclasses, and audit-ready ERP posting drafts.</div>
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

def render_audit_agent(kpis):
    audit_agent = kpis["agents"]["audit"]
    summary = audit_agent["summary"]

    st.markdown(
        f"""
        <div class="entity-callout">
            <div class="section-label">Audit &amp; Compliance Agent</div>
            <div class="entity-name">{summary['ready_to_post']} drafts cleared for posting, {summary['blocked_for_review']} still blocked</div>
            <div class="entity-copy">Control scoring, posting readiness, and exception gating before ERP release.</div>
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

def render_flux_agent(kpis):
    flux_agent = kpis["agents"]["flux"]
    summary = flux_agent["summary"]

    st.markdown(
        f"""
        <div class="entity-callout">
            <div class="section-label">Flux Analysis Agent</div>
            <div class="entity-name">{summary['mom_anomalies']} MoM anomalies · {summary['yoy_anomalies']} YoY anomalies</div>
            <div class="entity-copy">Variance signals, contextual drivers, and executive-ready commentary.</div>
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

def render_checklist_agent(kpis):
    checklist_agent = kpis["agents"]["checklist"]
    summary = checklist_agent["summary"]
    base_summary = kpis["summary"]
    dependency_summary = checklist_agent["dependency_summary"]
    all_tasks = checklist_agent.get("all_tasks", checklist_agent["worklist"]).copy()

    st.markdown(
        f"""
        <div class="entity-callout">
            <div class="section-label">Close Tracker</div>
            <div class="entity-name">{summary['blocked_tasks']} blocked, {summary['waiting_tasks']} waiting, {summary['critical_open_tasks']} critical still open</div>
            <div class="entity-copy">Critical-path handoffs, dependency blockers, and the next owner move.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    card_cols = st.columns(6, gap="medium")
    blocked_tone = "low" if summary["blocked_tasks"] == 0 else "high"
    waiting_tone = "low" if summary["waiting_tasks"] == 0 else "medium"
    critical_tone = "low" if summary["critical_open_tasks"] == 0 else "high"
    handoff_tone = "low" if summary["handoffs_at_risk"] == 0 else "medium"
    automation_tone = "low" if summary["automation_candidates"] == 0 else "medium"
    with card_cols[0]:
        render_metric_card(
            "Blocked",
            str(summary["blocked_tasks"]),
            "Tasks that cannot move until a prior dependency is cleared.",
            blocked_tone,
        )
    with card_cols[1]:
        render_metric_card(
            "Waiting on Input",
            str(summary["waiting_tasks"]),
            "Tasks stalled because the next owner is waiting for an upstream handoff.",
            waiting_tone,
        )
    with card_cols[2]:
        render_metric_card(
            "Critical Open",
            str(summary["critical_open_tasks"]),
            "Critical tasks still open on the close path.",
            critical_tone,
        )
    with card_cols[3]:
        render_metric_card(
            "Handoffs at Risk",
            str(summary["handoffs_at_risk"]),
            "Open tasks with dependencies that can still break the close chain.",
            handoff_tone,
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
            automation_tone,
        )

    st.markdown("<div class='section-label section-heading tight-top'>Checklist Composition</div>", unsafe_allow_html=True)
    checklist_chart = build_checklist_chart(base_summary)
    if HAS_PLOTLY:
        st.plotly_chart(
            checklist_chart,
            use_container_width=True,
            config={"displayModeBar": False},
            key="checklist_agent_composition_chart",
        )
    else:
        st.bar_chart(checklist_chart)

    st.markdown("<div class='section-label'>Filter Checklist</div>", unsafe_allow_html=True)
    filter_left, filter_right = st.columns([0.44, 0.56], gap="large")

    status_counts = {
        "All": int(all_tasks.shape[0]),
        "Completed": int(all_tasks["Status"].eq("Completed").sum()),
        "In Progress": int(all_tasks["Status"].eq("In Progress").sum()),
        "Not Started": int(all_tasks["Status"].eq("Not Started").sum()),
        "Waiting on Input": int(all_tasks["Status"].eq("Waiting on Input").sum()),
        "Blocked": int(all_tasks["Status"].eq("Blocked").sum()),
    }
    status_labels = {
        "All": "All",
        "Completed": "Done",
        "In Progress": "Active",
        "Not Started": "Not started",
        "Waiting on Input": "Waiting",
        "Blocked": "Blocked",
    }
    status_options = ["All", "Completed", "In Progress", "Not Started", "Waiting on Input", "Blocked"]

    owner_counts = {"All": int(all_tasks.shape[0])}
    for owner, count in all_tasks["Owner"].value_counts().items():
        owner_counts[str(owner)] = int(count)
    owner_options = ["All"] + sorted([option for option in owner_counts if option != "All"])

    deadline_series = pd.to_datetime(all_tasks["Deadline"], errors="coerce")
    all_tasks_with_deadline = all_tasks.assign(_deadline_sort=deadline_series)
    deadline_options = ["All"] + [
        str(value)
        for value in all_tasks_with_deadline.sort_values("_deadline_sort")["Deadline"].dropna().astype(str).drop_duplicates().tolist()
    ]
    deadline_counts = {"All": int(all_tasks.shape[0])}
    for deadline in deadline_options[1:]:
        deadline_counts[deadline] = int(all_tasks["Deadline"].astype(str).eq(deadline).sum())

    with filter_left:
        owner_col, deadline_col = st.columns(2, gap="small")
        with owner_col:
            owner_filter = st.selectbox(
                "Owner",
                owner_options,
                index=0,
                format_func=lambda option: "All owners" if option == "All" else str(option),
                key="checklist_owner_filter",
            )
        with deadline_col:
            deadline_filter = st.selectbox(
                "Deadline",
                deadline_options,
                index=0,
                format_func=lambda option: "All dates" if option == "All" else str(option),
                key="checklist_deadline_filter",
            )

    with filter_right:
        st.markdown("<div class='checklist-status-shell'><div class='checklist-filter-label'>Status</div>", unsafe_allow_html=True)
        status_filter = st.pills(
            "Status",
            status_options,
            selection_mode="single",
            default="All",
            format_func=lambda option: f"{status_labels[option]} ({status_counts[option]})",
            key="checklist_status_filter",
            width="stretch",
            label_visibility="collapsed",
        )
        st.markdown("</div>", unsafe_allow_html=True)

    filtered_tasks = all_tasks.copy()
    if status_filter and status_filter != "All":
        filtered_tasks = filtered_tasks.loc[filtered_tasks["Status"].eq(status_filter)].copy()
    if owner_filter and owner_filter != "All":
        filtered_tasks = filtered_tasks.loc[filtered_tasks["Owner"].eq(owner_filter)].copy()
    if deadline_filter and deadline_filter != "All":
        filtered_tasks = filtered_tasks.loc[filtered_tasks["Deadline"].astype(str).eq(deadline_filter)].copy()

    st.markdown(
        build_tracker_board_html(
            filtered_tasks,
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
            st.markdown(
                build_checklist_detail_cards_html(checklist_agent["critical_path"], "critical"),
                unsafe_allow_html=True,
            )

        with upper_right:
            st.markdown("<div class='section-label'>Dependency Handoffs</div>", unsafe_allow_html=True)
            st.markdown(
                build_checklist_detail_cards_html(checklist_agent["handoff_queue"], "handoff"),
                unsafe_allow_html=True,
            )


def render_agents_hub(kpis):
    st.markdown(
        """
        <div class="entity-callout">
            <div class="section-label">Agents</div>
            <div class="entity-name">Operational agents for approvals, reconciliations, journals, and controls</div>
            <div class="entity-copy">Specialist views for the close areas that move forecast, readiness, and posting risk.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    agent_tabs = st.tabs(
        [
            "GL Approval",
            "Bank",
            "Intercompany",
            "Journals",
            "Audit",
            "Variance Analysis",
        ]
    )

    with agent_tabs[0]:
        render_gl_agent(kpis)
    with agent_tabs[1]:
        render_bank_agent(kpis)
    with agent_tabs[2]:
        render_ic_agent(kpis)
    with agent_tabs[3]:
        render_journal_agent(kpis)
    with agent_tabs[4]:
        render_audit_agent(kpis)
    with agent_tabs[5]:
        render_flux_agent(kpis)


def render_scenario_lab(kpis, score, actions, recommendations, plan_scenario):
    base_posting = kpis["agents"]["erp_posting"]
    if "applied_actions" not in st.session_state:
        st.session_state["applied_actions"] = set()
    if "automation_log" not in st.session_state:
        st.session_state["automation_log"] = []
    if "automation_run" not in st.session_state:
        st.session_state["automation_run"] = False

    st.markdown(
        f"""
        <div class="entity-callout">
            <div class="section-label">Automation Audit Trail</div>
            <div class="entity-name">{score['continuous_close_days']} day continuous forecast vs. {score['predicted_close_days']} day dashboard forecast</div>
            <div class="entity-copy">Test the automation bundle and see which actions move the close below the 4-day target.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    plan_ids = [action["id"] for action in actions]
    posting = plan_scenario["posting_simulator"]
    automation_run = st.session_state.get("automation_run", False)

    metric_cols = st.columns(4, gap="medium")
    with metric_cols[0]:
        render_metric_card(
            "Base Dashboard Forecast",
            f"{score['predicted_close_days']} days",
            "Bucketed dashboard forecast before automation.",
            "medium",
        )
    with metric_cols[1]:
        render_metric_card(
            "After Automation",
            f"{plan_scenario['score']['predicted_close_days']} days",
            "Bucketed dashboard forecast after running the automation bundle.",
            "low" if plan_scenario["gets_below_four"] else "medium",
        )
    with metric_cols[2]:
        render_metric_card("Automation Score", f"{plan_scenario['score']['readiness_score']}/100", "Readiness score after the automation bundle.", "low" if plan_scenario["score"]["readiness_score"] >= 80 else "medium")
    with metric_cols[3]:
        render_metric_card(
            "Hours Saved",
            f"{plan_scenario['total_hours_saved']} hrs",
            "Recovered effort from the automation bundle.",
            "low",
        )

    status_text = "Below 4 days" if plan_scenario["gets_below_four"] else "Still above 4 days"
    status_tone = "low" if plan_scenario["gets_below_four"] else "high"
    render_metric_card(
        "Automation Status",
        status_text,
        f"Gap to target: {plan_scenario['score']['continuous_gap_to_target_days']} days",
        status_tone,
    )

    st.markdown("<div class='section-label'>Automation Execution</div>", unsafe_allow_html=True)
    applied_panel = (
        f"{plan_scenario['score']['predicted_close_days']} day dashboard forecast · "
        f"{plan_scenario['score']['continuous_close_days']} day continuous model · "
        f"{len(plan_ids)} actions applied"
        if automation_run
        else "Automation bundle not yet executed. Run the bundle from Command Centre Step 2."
    )
    st.info(f"Applied automation state: {applied_panel}")

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

    if st.session_state["automation_log"]:
        st.markdown("<div class='section-label'>Automation Run Sheet</div>", unsafe_allow_html=True)
        st.dataframe(
            format_display_frame(pd.DataFrame(st.session_state["automation_log"])),
            use_container_width=True,
            hide_index=True,
        )

    st.markdown("<div class='section-label'>Automation Bundle</div>", unsafe_allow_html=True)
    st.dataframe(
        format_display_frame(pd.DataFrame(actions)),
        use_container_width=True,
        hide_index=True,
    )

    st.markdown("<div class='section-label'>Recommended Bundles Below 4 Days</div>", unsafe_allow_html=True)
    st.dataframe(
        format_display_frame(recommendations),
        use_container_width=True,
        hide_index=True,
    )


def render_priority_engine(before_priorities, after_priorities, automation_run, before_kpis, after_kpis):
    before_hours = sum(priority["hours_saved_est"] for priority in before_priorities[:4])
    after_hours = sum(priority["hours_saved_est"] for priority in after_priorities[:4])
    after_label = "After automation" if automation_run else "Projected after automation"
    before_handoffs = before_kpis["agents"]["checklist"]["handoff_queue"]
    after_handoffs = after_kpis["agents"]["checklist"]["handoff_queue"]

    st.markdown(
        f"""
        <div class="entity-callout">
            <div class="section-label">Priority Engine</div>
            <div class="entity-name">{before_hours} hrs before -> {after_hours} hrs {after_label.lower()}</div>
            <div class="entity-copy">Compare what finance should focus on before the bundle runs and what still matters after automation.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    before_col, after_col = st.columns(2, gap="large")

    with before_col:
        st.markdown(
            f"""
            <div class="entity-callout">
                <div class="entity-name">{before_hours} hours recoverable in the top four moves</div>
                <div class="entity-copy">Ordered by impact, downstream unlock, and closeness to the 4-day target.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown(
            """
            <div class="priority-sticky-header">
                <div class="section-label centered">Before Automation</div>
                <div class="section-label centered priority-chart-heading">Priority Score Ranking</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        before_chart = build_priority_chart(before_priorities)
        if HAS_PLOTLY:
            st.plotly_chart(
                before_chart,
                use_container_width=True,
                config={"displayModeBar": False},
                key="priority_engine_before_chart",
            )
        else:
            st.bar_chart(before_chart)

        before_cards = st.columns(2, gap="medium")
        for index, priority in enumerate(before_priorities[:4]):
            with before_cards[index % 2]:
                render_priority_card(priority)

        st.markdown(
            "<div class='section-label section-heading centered'>Top Handoffs To Clear Today</div>",
            unsafe_allow_html=True,
        )
        st.markdown(build_tracker_notes_html(before_handoffs), unsafe_allow_html=True)

    with after_col:
        st.markdown(
            f"""
            <div class="entity-callout">
                <div class="entity-name">{after_hours} hours recoverable in the top four moves</div>
                <div class="entity-copy">What still needs human focus after the automation bundle has run.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown(
            f"""
            <div class="priority-sticky-header">
                <div class="section-label centered">{after_label}</div>
                <div class="section-label centered priority-chart-heading">Priority Score Ranking</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        after_chart = build_priority_chart(after_priorities)
        if HAS_PLOTLY:
            st.plotly_chart(
                after_chart,
                use_container_width=True,
                config={"displayModeBar": False},
                key="priority_engine_after_chart",
            )
        else:
            st.bar_chart(after_chart)

        after_cards = st.columns(2, gap="medium")
        for index, priority in enumerate(after_priorities[:4]):
            with after_cards[index % 2]:
                render_priority_card(priority)

        st.markdown(
            "<div class='section-label section-heading centered'>Top Handoffs After Automation</div>",
            unsafe_allow_html=True,
        )
        st.markdown(build_tracker_notes_html(after_handoffs), unsafe_allow_html=True)


def render_automation_plays():
    st.markdown(
        """
        <div class="entity-callout">
            <div class="section-label">Automation Plays</div>
            <div class="entity-name">High-ROI automation ideas for the close</div>
            <div class="entity-copy">Next-wave opportunities once the current close is stable.</div>
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
    left, right = st.columns([0.82, 0.18], gap="small")
    with left:
        st.markdown(
            f"""
            <div class="app-header-main">
                <div class="app-header-main-copy">
                    <div class="app-header-kicker">Month-End Workspace</div>
                    <div class="app-header-title">NovaClose</div>
                    <div class="app-header-copy">
                        AI month-end close command centre
                    </div>
                </div>
                <div class="app-header-meta-stack">
                    <div class='status-pill'><span class='status-dot'></span>Live</div>
                    <div class='period-pill'>Period: {summary['accounting_period_label']}</div>
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with right:
        st.markdown("<div class='header-ask-ai'>", unsafe_allow_html=True)
        if st.button("Ask AI", key="ask_ai_button", use_container_width=True):
            st.session_state["nav_target"] = "AI Copilot"
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

data, base_kpis, base_score, base_priorities, base_commentary = prepare(data_source)
reset_copilot_if_needed(dataset_key_for(data_source), base_kpis, base_score, base_priorities)
base_actions = get_cached_actions(base_kpis)
base_recommendations = get_cached_recommendations(base_kpis)
base_plan_scenario = get_cached_plan_scenario(base_kpis, [action["id"] for action in base_actions])
base_plan_priorities = priority_engine(base_plan_scenario["kpis"])

with st.sidebar:
    st.markdown("### Accounting Period")
    st.write(base_kpis["summary"]["accounting_period_label"])

if st.session_state.get("automation_run"):
    applied = get_cached_plan_scenario(base_kpis, [action["id"] for action in base_actions])
    kpis = applied["kpis"]
    score = applied["score"]
    priorities = priority_engine(kpis)
    commentary = generate_commentary(kpis, score)
else:
    kpis = base_kpis
    score = base_score
    priorities = base_priorities
    commentary = base_commentary

render_app_header(kpis)
render_close_timeline(kpis["summary"])
current_actions = get_cached_actions(kpis)
current_recommendations = get_cached_recommendations(kpis)

nav_items = [
    "Command Centre",
    "Close Tracker",
    "Risks",
    "Priorities",
    "Agents",
    "Automation Audit Trail",
    "AI Copilot",
]

tab_order = nav_items.copy()
nav_target = st.session_state.get("nav_target")
if nav_target in nav_items:
    tab_order = [nav_target] + [tab for tab in nav_items if tab != nav_target]

tabs = st.tabs(tab_order)
tab_map = dict(zip(tab_order, tabs))

with tab_map["Command Centre"]:
    render_command_center(
        kpis,
        score,
        priorities,
        commentary,
        current_actions,
        current_recommendations,
        base_score,
        base_plan_scenario["score"],
        st.session_state.get("automation_run", False),
    )

with tab_map["Close Tracker"]:
    render_checklist_agent(kpis)

with tab_map["Agents"]:
    render_agents_hub(kpis)

with tab_map["Risks"]:
    render_risks_tab(kpis)

with tab_map["Automation Audit Trail"]:
    render_scenario_lab(base_kpis, base_score, base_actions, base_recommendations, base_plan_scenario)

with tab_map["Priorities"]:
    render_priority_engine(
        base_priorities,
        base_plan_priorities,
        st.session_state.get("automation_run", False),
        base_kpis,
        base_plan_scenario["kpis"],
    )

with tab_map["AI Copilot"]:
    render_chat(kpis, score, priorities)

if nav_target in nav_items:
    st.session_state["nav_target"] = None
