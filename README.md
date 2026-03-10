# NovaClose AI

Pitch-ready month-end close intelligence demo for the NovaTech finance hackathon dataset.

## What It Does

NovaClose AI turns a slow, exception-heavy close into a ranked action plan.

The app scans the NovaTech workbook, detects what is blocking the close, predicts how far the team is from the 4-day CFO target, and lets the user query the results through a copilot-style interface.

Core outcomes:

- detect close blockers across all 8 sheets
- compute a Close Readiness Score and predicted close days
- rank the highest-impact actions for the controller
- compare risk across the 5 legal entities
- surface finance-ready commentary without requiring internet or API keys

## Current Demo Output

Running the bundled workbook currently produces:

```text
Readiness Score: 67/100
Risk Level: Medium
Predicted Close Days: 4.4
Gap To 4-Day Target: 0.4
```

Headline KPI checks:

- Pending GL approvals/reviews: `171`
- Manual journal entries: `66`
- Bank exceptions: `10`
- Intercompany exceptions: `2`
- AP 3-way exceptions: `8`
- Blocked checklist tasks: `3`

Current top priorities:

1. Clear pending GL reviews and approvals
2. Resolve bank reconciliation exceptions
3. Unblock critical close checklist tasks
4. Fix intercompany elimination gaps
5. Review high-risk accruals and missing support
6. Clear AP 3-way match exceptions

## Why This Matters

The key insight from the dataset is that NovaTech is not delayed by transaction processing volume. It is delayed by exception management, approval friction, and close dependencies surfacing too late.

The current workbook shows:

- a large GL approval backlog
- unresolved cash exceptions
- intercompany elimination gaps
- accrual documentation and approval issues
- AP exception handling friction
- blocked critical checklist tasks
- significant trial balance variance movement
- fixed-asset control gaps that are smaller, but still visible

## App Structure

The Streamlit app is organized into four presentation-ready views:

### Command Center

- branded hero section
- readiness gauge
- predicted close days and gap to target
- top blocker callouts
- chat-style NovaClose Copilot
- checklist status composition

### Risk Atlas

- card-based scan of GL, bank rec, intercompany, accruals, AP, checklist, trial balance, and fixed assets
- risk-by-area visual
- drill-down tables for evidence behind each risk area

### Priority Engine

- ranked action board
- hours recoverable estimate
- downstream unlock explanation
- automation opportunity ideas

### Entity View

- 5-entity risk comparison
- primary blocker by entity
- side-by-side risk table for Germany, Netherlands, UK, USA, and France

## Copilot Design

The demo is intentionally LLM-first in user experience, but deterministic in implementation.

The app uses a local intent router and finance response builder for prompts such as:

- What issues will delay the close?
- What should the controller do today?
- Which entity is riskiest?
- Explain the biggest variances.
- Which processes should be automated first?

This keeps the demo reliable in a hackathon setting while still feeling like a close copilot.

## Analysis Logic

The backend groups outputs into:

- `summary`: headline KPI values
- `areas`: per-domain risk blocks and severity scores
- `entities`: risk rollups for each legal entity
- `details`: drill-down tables used by the UI

Finance logic corrections included in this version:

- AP 3-way match exceptions count only true AP exception rows
- AR rows no longer inflate AP exception metrics
- trial balance and fixed-assets signals are visible as secondary risk areas

## Repository Structure

```text
AI Finance Hackathon Mar 2026/
├── app.py
├── novaclose_analysis.py
├── demo_script.md
├── requirements.txt
├── AI_Finance_Hackathon_Month_End_Dataset.xlsx
└── README.md
```

## Running The Demo

Install dependencies:

```bash
pip3 install -r requirements.txt
```

Launch the app:

```bash
python3 -m streamlit run app.py
```

Run the CLI analysis check:

```bash
python3 novaclose_analysis.py
```

The Streamlit app opens at:

```text
http://localhost:8501
```

## Hackathon Narrative

Suggested one-line pitch:

```text
NovaClose AI identifies the exceptions that delay month-end close, predicts close risk early, and tells finance exactly what to fix first.
```

Suggested before/after framing:

- Before: 7-day close, low visibility, manual exception hunting
- After: blockers surfaced immediately, priorities ranked automatically, realistic path to a 4-day close

## Future Extensions

This demo is designed to be practical first, but it leaves a clean path for:

- live LLM integration behind the same copilot interface
- ERP/API ingestion instead of Excel upload
- auto-generated journal routing workflows
- bank auto-match confidence scoring
- recurring accrual drafting
- richer scenario forecasting for the path to target
