# NovaClose AI Demo MVP Plan

## Summary

Upgrade the current prototype into a pitch-ready Streamlit demo that is visibly smarter, more polished, and more reliable for judges. The app should feel LLM-first through a chat-style copilot experience, but use a deterministic local response engine so the demo works without internet or API keys. The main goal is a stronger live demo, not a broader product build.

## Key Changes

### Analysis and metric layer

- Refactor [novaclose_analysis.py](/Users/michelle/Downloads/AI Finance Hackathon Mar 2026/novaclose_analysis.py) so KPI output is grouped into `summary`, `areas`, `entities`, and `details` instead of a flat dict only.
- Correct finance logic where current metrics are misleading:
  - Count `AP 3-way exceptions` as non-null non-`Matched` values only, so the demo shows `8` rather than including AR blanks.
  - Split AP and AR signals so blank `3-Way Match` values do not inflate exception counts.
- Keep the main blocker areas as primary score drivers: GL, bank rec, intercompany, accruals, AP, checklist.
- Add secondary signals for hybrid 8-sheet coverage:
  - Trial Balance: unreconciled accounts, large variance count, top variance accounts.
  - Fixed Assets: non-zero impairments, assets under review or disposal, assets not physically verified.
- Add entity rollups for the 5 legal entities so the UI can show which entity is riskiest and why.
- Recalibrate the readiness score after KPI fixes so it remains CFO-friendly and still outputs `readiness_score`, `risk_level`, `predicted_close_days`, `gap_to_target_days`, and penalty drivers.

### Copilot and interaction model

- Replace the current selectbox advisor in [app.py](/Users/michelle/Downloads/AI Finance Hackathon Mar 2026/app.py) with an LLM-style chat interface using `st.chat_input` and `st.chat_message`.
- Implement a local intent router plus response builder instead of a live API:
  - Supported intents: close blockers, controller actions today, automation opportunities, CFO summary, riskiest entity, trial balance variance summary.
  - Free-text prompts map to the nearest supported intent using keyword rules.
  - Unknown prompts fall back to a general executive answer plus suggested follow-up prompts.
- Return copilot responses with finance-style prose and visible source chips such as `171 pending GL`, `10 bank exceptions`, `2 IC exceptions`.
- Keep the design provider-ready internally by isolating copilot generation behind one helper function, but do not require `openai` or network access in this MVP.

### Showpiece UI

- Rebuild the app as four presentation-grade views:
  - `Command Center`: branded hero, readiness gauge, predicted close days, target gap, top blockers, chat-first copilot panel.
  - `Risk Atlas`: card-based scanner for GL, bank, IC, accruals, AP, checklist, plus lighter TB and fixed-assets panels.
  - `Priority Engine`: ranked action board with impact, hours saved, and downstream unlock explanation.
  - `Entity View`: 5-entity comparison with a heatmap or ranked bars plus each entity’s primary blocker.
- Add custom CSS for a pitch-ready visual system:
  - Industrial finance palette, bright background, strong header, score cards, risk badges, consistent spacing.
  - Avoid default Streamlit look; use custom containers and typographic hierarchy.
- Add Plotly visuals:
  - Readiness gauge.
  - Risk-by-area bar chart.
  - Entity risk comparison.
  - Checklist status composition.
- Surface one visible insight from all 8 sheets, but keep the narrative centered on the highest-impact blockers.

### Demo assets and docs

- Update [README.md](/Users/michelle/Downloads/AI Finance Hackathon Mar 2026/README.md) so the documented KPIs and app behavior match corrected finance logic.
- Update [demo_script.md](/Users/michelle/Downloads/AI Finance Hackathon Mar 2026/demo_script.md) to match the new flow:
  - lead with the copilot,
  - show the readiness gauge,
  - then drill into risks and priorities.
- Add a simple dependency file or documented install list including `plotly`.

## Public Interfaces

- Keep `load_data(path)` unchanged.
- Expand `calculate_kpis(data)` to return a nested structure with:
  - `summary` for headline metrics,
  - `areas` for per-domain risk blocks,
  - `entities` for per-entity rollups,
  - `details` for drill-down tables and top examples.
- Keep `calculate_readiness_score(kpis)` but extend its return shape with `gap_to_target_days`.
- Keep `priority_engine(kpis)` but ensure each item also includes `area` and a short `downstream_unlock`.
- Add a new helper such as `generate_copilot_response(prompt, kpis, score, priorities)` that returns `intent`, `answer`, and `source_metrics`.

## Test Plan

- Verify the bundled workbook produces finance-correct headline KPIs:
  - pending GL `171`
  - manual JEs `66`
  - bank exceptions `10`
  - IC exceptions `2`
  - AP 3-way exceptions `8`
  - blocked tasks `3`
- Verify secondary coverage is non-empty:
  - trial balance variances and unreconciled accounts populate,
  - fixed-asset review/verification signals populate.
- Verify score output is bounded, risk level is consistent, and predicted close remains above the 4-day target for the demo dataset.
- Verify copilot prompt routing for at least these prompts:
  - “What issues will delay the close?”
  - “What should the controller do today?”
  - “Which entity is riskiest?”
  - “Explain the biggest variances.”
- Run a manual Streamlit smoke test:
  - app loads with bundled dataset,
  - upload path works,
  - charts render,
  - chat returns answers with no API key,
  - no empty or contradictory KPI cards.

## Assumptions

- Upgrade the existing [app.py](/Users/michelle/Downloads/AI Finance Hackathon Mar 2026/app.py) instead of creating a second polished app to avoid confusion during the hackathon.
- The demo is intentionally LLM-first in UX but not dependent on a live model.
- Use the bundled dataset as the default demo path and keep file upload as a secondary option.
- No team logo or custom brand asset is required; the app uses NovaClose AI branding only unless assets are added later.
- `plotly` is the only new runtime dependency needed for the planned visuals.

