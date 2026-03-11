from copy import deepcopy
from pathlib import Path

import numpy as np
import pandas as pd

DATA_FILE = Path("AI_Finance_Hackathon_Month_End_Dataset.xlsx")
PRIMARY_AREA_KEYS = ("gl", "bank", "intercompany", "accruals", "ap", "checklist")
AREA_LABELS = {
    "gl": "GL & Journals",
    "bank": "Bank Reconciliation",
    "intercompany": "Intercompany",
    "accruals": "Accruals",
    "ap": "AP Exceptions",
    "ar": "AR Watchlist",
    "checklist": "Close Checklist",
    "tb": "Trial Balance",
    "fa": "Fixed Assets",
}
ENTITY_COLUMNS = {
    "gl": "Entity",
    "tb": "Entity",
    "accruals": "Entity",
    "bank": "Entity",
    "fa": "Entity",
    "apar": "Entity",
}


def load_data(path=DATA_FILE):
    source = path
    if hasattr(source, "seek"):
        source.seek(0)

    workbook = pd.ExcelFile(source)
    bank = pd.read_excel(workbook, sheet_name="Bank Reconciliation")
    bank = bank.loc[
        bank["Bank Account"].notna()
        & ~bank["Bank Account"].astype(str).str.strip().eq("")
    ].copy()

    return {
        "gl": pd.read_excel(workbook, sheet_name="General Ledger"),
        "tb": pd.read_excel(workbook, sheet_name="Trial Balance"),
        "accruals": pd.read_excel(workbook, sheet_name="Accruals & Provisions"),
        "bank": bank,
        "ic": pd.read_excel(workbook, sheet_name="Intercompany"),
        "fa": pd.read_excel(workbook, sheet_name="Fixed Assets"),
        "apar": pd.read_excel(workbook, sheet_name="AP & AR Aging"),
        "checklist": pd.read_excel(workbook, sheet_name="Close Checklist"),
    }


def safe_pct(num, den):
    return round((num / den) * 100, 1) if den else 0.0


def _clamp(value, low=0, high=100):
    return max(low, min(high, int(round(value))))


def _severity_score(value, warn, critical, floor=0):
    if critical <= warn:
        return floor

    normalized = (value - warn) / (critical - warn)
    return _clamp(floor + (normalized * (100 - floor)))


def _top_records(df, columns, limit=6, sort_by=None, ascending=False):
    if sort_by is not None:
        df = df.sort_values(sort_by, ascending=ascending)
    frame = df.loc[:, columns].copy()
    return frame.head(limit)


def _format_currency(value):
    return f"EUR {value:,.0f}"


def _parse_variance_pct(value):
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None
    try:
        return float(text.rstrip("%"))
    except ValueError:
        return None


def _derive_accounting_period(data):
    period_candidates = []
    for sheet_name in ("gl", "accruals", "ic"):
        frame = data[sheet_name]
        if "Period" in frame.columns:
            period_candidates.extend(
                frame["Period"].dropna().astype(str).str.strip().tolist()
            )

    normalized_periods = [value for value in period_candidates if value and value.lower() != "nan"]
    if normalized_periods:
        primary = pd.Series(normalized_periods).mode().iloc[0]
        try:
            period = pd.Period(primary, freq="M")
            return {
                "raw": primary,
                "label": period.strftime("%b %Y"),
            }
        except Exception:
            return {"raw": primary, "label": primary}

    gl = data["gl"]
    if "Posting Date" in gl.columns:
        posting_dates = pd.to_datetime(gl["Posting Date"], errors="coerce").dropna()
        if not posting_dates.empty:
            latest = posting_dates.max().to_period("M")
            return {
                "raw": str(latest),
                "label": latest.strftime("%b %Y"),
            }

    return {"raw": "Unknown", "label": "Unknown"}


def _bank_offset_account(description, likely_cause):
    text = f"{description} {likely_cause}".lower()
    if "fx" in text or "foreign exchange" in text or "rounding" in text:
        return (8100, "Foreign Exchange Gain/Loss")
    if "loan" in text:
        return (3300, "Short-Term Loan Payable")
    if "tax" in text:
        return (3600, "Income Tax Payable")
    if "payroll" in text:
        return (3020, "Accrued Salaries & Wages")
    if "customer receipt" in text:
        return (1100, "Accounts Receivable")
    if "supplier payment" in text:
        return (3000, "Accounts Payable")
    return (3010, "Accrued Expenses")


def _period_to_posting_date(period_value):
    period_str = str(period_value)
    try:
        return pd.Period(period_str, freq="M").end_time.normalize()
    except Exception:
        return pd.NaT


def _journal_ready_status(confidence, mapped):
    return "Draft - Ready for ERP" if mapped and confidence >= 75 else "Draft - Review Needed"


def _journal_agent_mapping(row):
    description = str(row["Description"]).lower()
    account_code = int(row["Account Code"])
    amount = float(row["Accrual Amount (EUR)"])
    notes = str(row["Notes"]).lower()

    if "deferred revenue release" in description or account_code == 3200:
        return {
            "je_type": "Standard Reclassification",
            "debit_code": 3200,
            "debit_name": "Deferred Revenue",
            "credit_code": 6020,
            "credit_name": "Revenue – SaaS Subscriptions",
            "rationale": "Release deferred revenue into the SaaS revenue line for the current period.",
            "mapped": True,
        }
    if "insurance premium allocation" in description:
        return {
            "je_type": "Standard Reclassification",
            "debit_code": 7700,
            "debit_name": "Insurance Expense",
            "credit_code": 1300,
            "credit_name": "Prepaid Expenses",
            "rationale": "Allocate prepaid insurance into current-period expense.",
            "mapped": True,
        }
    if "ifrs 16" in description or account_code == 3500:
        return {
            "je_type": "Standard Reclassification",
            "debit_code": 2200,
            "debit_name": "Right-of-Use Assets (IFRS 16)",
            "credit_code": 3500,
            "credit_name": "Lease Liability – Current (IFRS 16)",
            "rationale": "Reclass the lease adjustment into the IFRS 16 balance sheet structure.",
            "mapped": True,
        }
    if account_code == 6200 or "intercompany" in description:
        return {
            "je_type": "Contract-Backed Accrual",
            "debit_code": 1500,
            "debit_name": "Intercompany Receivable",
            "credit_code": 6200,
            "credit_name": "Intercompany Revenue",
            "rationale": "Accrue the intercompany management fee against the intercompany receivable position.",
            "mapped": True,
        }
    if account_code == 3030 or "bonus" in description:
        return {
            "je_type": "Contract-Backed Accrual",
            "debit_code": 7100,
            "debit_name": "Salaries & Wages",
            "credit_code": 3030,
            "credit_name": "Accrued Bonus Provision",
            "rationale": "Accrue employee bonus expense into the bonus provision liability.",
            "mapped": True,
        }
    if account_code == 8100 or "foreign exchange" in description:
        if amount >= 0:
            debit_code, debit_name = 8100, "Foreign Exchange Gain/Loss"
            credit_code, credit_name = 3010, "Accrued Expenses"
        else:
            debit_code, debit_name = 3010, "Accrued Expenses"
            credit_code, credit_name = 8100, "Foreign Exchange Gain/Loss"
        return {
            "je_type": "Standard Reclassification",
            "debit_code": debit_code,
            "debit_name": debit_name,
            "credit_code": credit_code,
            "credit_name": credit_name,
            "rationale": "Book the FX revaluation adjustment between P&L and accrued expense staging.",
            "mapped": True,
        }
    if 7000 <= account_code < 8000 or account_code in {7400, 7410, 7420}:
        return {
            "je_type": "Contract-Backed Accrual",
            "debit_code": account_code,
            "debit_name": row["Account Name"],
            "credit_code": 3010,
            "credit_name": "Accrued Expenses",
            "rationale": "Accrue the current-period operating expense into accrued expenses.",
            "mapped": True,
        }

    if notes and "allocation" in notes:
        return {
            "je_type": "Standard Reclassification",
            "debit_code": account_code,
            "debit_name": row["Account Name"],
            "credit_code": 3010,
            "credit_name": "Accrued Expenses",
            "rationale": "Standardize the pending allocation through accrued expenses for close booking.",
            "mapped": True,
        }

    return {
        "je_type": "Review Required",
        "debit_code": account_code,
        "debit_name": row["Account Name"],
        "credit_code": 3010,
        "credit_name": "Accrued Expenses",
        "rationale": "No clean template match was found, so the draft requires controller review.",
        "mapped": False,
    }


def _journal_agent_confidence(row, mapping):
    support = str(row["Supporting Doc"])
    status = str(row["Status"])
    notes = str(row["Notes"]).lower()

    if support in {"Contract ref", "PO attached", "Email confirmation"}:
        confidence = 90
    elif support in {"Estimate from vendor", "Budget forecast"}:
        confidence = 76
    else:
        confidence = 62

    if status == "Needs Documentation":
        confidence -= 15
    elif status == "Pending Approval":
        confidence -= 4

    if "true-up" in notes:
        confidence -= 8
    if "verify" in notes:
        confidence -= 6
    if "cross-entity allocation pending" in notes:
        confidence -= 10

    if mapping["je_type"] == "Standard Reclassification":
        confidence += 2

    if not mapping["mapped"]:
        confidence -= 12

    return _clamp(confidence)


def _journal_agent_audit_trail(row, mapping, confidence):
    period = str(row["Period"])
    support = str(row["Supporting Doc"]) if pd.notna(row["Supporting Doc"]) else "No formal support attached"
    notes = str(row["Notes"]) if pd.notna(row["Notes"]) else "No additional notes"
    return (
        f"AI audit trail: Source {row['Accrual ID']} for {period}. "
        f"Support reviewed: {support}. "
        f"Business rationale: {mapping['rationale']} "
        f"Entry proposed as debit {mapping['debit_code']} {mapping['debit_name']} and credit "
        f"{mapping['credit_code']} {mapping['credit_name']} for EUR {abs(float(row['Accrual Amount (EUR)'])):,.2f}. "
        f"Status at extraction: {row['Status']}. Confidence score: {confidence}/100. "
        f"Reviewer note context: {notes}."
    )


def _is_blank(value):
    return pd.isna(value) or str(value).strip() == "" or str(value).strip().lower() == "nan"


def _audit_required_approver(source_agent, amount, recommendation):
    if source_agent == "Bank Agent":
        if recommendation == "Blocked for Review" or amount >= 50000:
            return "Treasury Manager + Controller"
        if amount >= 10000:
            return "Treasury Manager"
        return "Cash Controller"

    if recommendation == "Blocked for Review" or amount >= 50000:
        return "Controller + CFO"
    if amount >= 20000:
        return "Controller"
    return "Accounting Manager"


def _audit_agent_review(row):
    journal_id = str(row.get("Journal ID", "Unassigned"))
    source_agent = str(row.get("Source Agent", "Unknown"))
    amount = abs(float(pd.to_numeric(row.get("Amount (EUR)"), errors="coerce") or 0))
    confidence = int(pd.to_numeric(row.get("Confidence %"), errors="coerce") or 0)
    support = "" if _is_blank(row.get("Support Pack")) else str(row.get("Support Pack"))
    support_lower = support.lower()
    control_gaps = []
    blockers = []
    score = 100

    if _is_blank(row.get("Posting Date")):
        score -= 12
        blockers.append("Missing posting date")

    if amount <= 0:
        score -= 25
        blockers.append("Invalid journal amount")

    debit_account = row.get("Debit Account")
    credit_account = row.get("Credit Account")
    if _is_blank(debit_account) or _is_blank(credit_account):
        score -= 25
        blockers.append("Missing debit or credit account")
    elif str(debit_account).strip() == str(credit_account).strip():
        score -= 22
        blockers.append("Debit and credit accounts are identical")

    if _is_blank(row.get("Reference")):
        score -= 8
        control_gaps.append("Missing source reference")

    if _is_blank(row.get("Header Text")):
        score -= 5
        control_gaps.append("Missing header text")

    if _is_blank(row.get("Line Text")):
        score -= 5
        control_gaps.append("Missing line text")

    if _is_blank(support) or "missing" in support_lower:
        score -= 20
        blockers.append("Missing support pack")
    elif "estimate" in support_lower or "forecast" in support_lower:
        score -= 12
        control_gaps.append("Support is estimate-based")
    elif "email" in support_lower:
        score -= 6
        control_gaps.append("Support is email-only")
    elif "bank rec exception" in support_lower:
        score -= 4
        control_gaps.append("Support depends on reconciler notes")

    if confidence < 70:
        score -= 18
        blockers.append("Confidence below control threshold")
    elif confidence < 80:
        score -= 10
        control_gaps.append("Confidence below auto-post threshold")
    elif confidence < 90:
        score -= 4
        control_gaps.append("Confidence below straight-through threshold")

    if source_agent == "Journal Agent" and _is_blank(row.get("AI Audit Trail")):
        score -= 8
        control_gaps.append("Missing journal audit trail")

    if amount >= 50000:
        control_gaps.append("High-value JE requires secondary approval")

    score = _clamp(score)
    blocking_issues = blockers.copy()
    if blockers or score < 70:
        recommendation = "Blocked for Review"
    elif score < 85 or amount >= 50000:
        recommendation = "Conditional Approval"
    else:
        recommendation = "Ready to Post"

    required_approver = _audit_required_approver(source_agent, amount, recommendation)
    all_gaps = blockers + control_gaps
    primary_gap = all_gaps[0] if all_gaps else "No material control gaps"
    gaps_text = "; ".join(all_gaps) if all_gaps else "No material control gaps"
    control_memo = (
        f"Control review for {journal_id}: {source_agent} draft for {row.get('Entity', 'Unknown entity')} "
        f"scored {score}/100. Recommendation: {recommendation}. "
        f"Support basis: {support or 'Not provided'}. Required approver: {required_approver}. "
        f"Key control observations: {gaps_text}. "
        f"Reference {row.get('Reference', 'N/A')} for EUR {amount:,.2f}."
    )

    return {
        "Control Score": score,
        "Posting Recommendation": recommendation,
        "Required Approver": required_approver,
        "Primary Control Gap": primary_gap,
        "Control Gaps": gaps_text,
        "Blocking Issues": "; ".join(blocking_issues) if blocking_issues else "None",
        "Control Memo": control_memo,
    }


def _dependency_step(value):
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None


def _flux_yoy_factor(account_type):
    account_type = str(account_type).lower()
    if "revenue" in account_type:
        return 1.06
    if "expense" in account_type:
        return 1.05
    if "asset" in account_type:
        return 1.03
    if "liability" in account_type:
        return 1.03
    if "equity" in account_type:
        return 1.0
    return 1.04


def _collect_flux_context(account_code, account_name, gl, accruals, apar, reviewer_note):
    snippets = []
    if reviewer_note and str(reviewer_note).strip():
        snippets.append(str(reviewer_note).strip())

    gl_desc = (
        gl.loc[gl["Account Code"].eq(account_code), "Description"]
        .dropna()
        .astype(str)
        .head(2)
        .tolist()
    )
    snippets.extend([text for text in gl_desc if text and text not in snippets])

    accrual_notes = (
        accruals.loc[accruals["Account Code"].eq(account_code), "Notes"]
        .dropna()
        .astype(str)
        .head(2)
        .tolist()
    )
    snippets.extend([text for text in accrual_notes if text and text not in snippets])

    accrual_desc = (
        accruals.loc[accruals["Account Code"].eq(account_code), "Description"]
        .dropna()
        .astype(str)
        .head(1)
        .tolist()
    )
    snippets.extend([text for text in accrual_desc if text and text not in snippets])

    name_lower = str(account_name).lower()
    if "receivable" in name_lower:
        notes = (
            apar.loc[apar["Type"].eq("AR"), "Notes"]
            .dropna()
            .astype(str)
            .head(1)
            .tolist()
        )
        snippets.extend([text for text in notes if text and text not in snippets])
    if "payable" in name_lower or "accrued" in name_lower:
        notes = (
            apar.loc[apar["Type"].eq("AP"), "Notes"]
            .dropna()
            .astype(str)
            .head(1)
            .tolist()
        )
        snippets.extend([text for text in notes if text and text not in snippets])

    return snippets[:3]


def _build_flux_agent(tb, gl, accruals, apar):
    frame = tb.copy()
    frame["Closing Balance (EUR)"] = pd.to_numeric(frame["Closing Balance (EUR)"], errors="coerce").fillna(0)
    frame["Prior Month Balance (EUR)"] = pd.to_numeric(frame["Prior Month Balance (EUR)"], errors="coerce").fillna(0)
    frame["Variance (EUR)"] = pd.to_numeric(frame["Variance (EUR)"], errors="coerce").fillna(
        frame["Closing Balance (EUR)"] - frame["Prior Month Balance (EUR)"]
    )
    frame["MoM %"] = frame["Variance %"].apply(_parse_variance_pct)
    missing_mom = frame["MoM %"].isna()
    frame.loc[missing_mom, "MoM %"] = np.where(
        frame.loc[missing_mom, "Prior Month Balance (EUR)"].abs() > 0,
        (frame.loc[missing_mom, "Variance (EUR)"] / frame.loc[missing_mom, "Prior Month Balance (EUR)"].abs()) * 100,
        np.nan,
    )

    if "Prior Year Balance (EUR)" in frame.columns:
        frame["Prior Year Balance (EUR)"] = pd.to_numeric(
            frame["Prior Year Balance (EUR)"], errors="coerce"
        ).fillna(0)
        frame["YoY Source"] = "Reported"
        frame["YoY Variance (EUR)"] = frame["Closing Balance (EUR)"] - frame["Prior Year Balance (EUR)"]
        frame["YoY %"] = np.where(
            frame["Prior Year Balance (EUR)"].abs() > 0,
            (frame["YoY Variance (EUR)"] / frame["Prior Year Balance (EUR)"].abs()) * 100,
            np.nan,
        )
    else:
        frame["YoY Source"] = "Derived proxy"
        proxy_base = frame["Prior Month Balance (EUR)"] * frame["Account Type"].apply(_flux_yoy_factor)
        frame["YoY Variance (EUR)"] = frame["Closing Balance (EUR)"] - proxy_base
        frame["YoY %"] = np.where(
            proxy_base.abs() > 0,
            (frame["YoY Variance (EUR)"] / proxy_base.abs()) * 100,
            np.nan,
        )

    frame["MoM Flag"] = frame["MoM %"].abs() >= 10
    frame["YoY Flag"] = frame["YoY %"].abs() >= 15
    frame["Value Flag"] = frame["Variance (EUR)"].abs() >= 50000
    frame["Recon Flag"] = ~frame["Reconciled"].eq("Yes")

    anomalies = frame.loc[
        frame["MoM Flag"] | frame["YoY Flag"] | frame["Value Flag"] | frame["Recon Flag"]
    ].copy()

    context_rows = []
    for _, row in anomalies.iterrows():
        context = _collect_flux_context(
            row["Account Code"],
            row["Account Name"],
            gl,
            accruals,
            apar,
            row.get("Reviewer Notes"),
        )
        context_rows.append("; ".join(context))

    anomalies["Context Cues"] = context_rows
    anomalies["Anomaly Reason"] = anomalies.apply(
        lambda row: ", ".join(
            reason
            for reason, flag in [
                ("MoM swing", row["MoM Flag"]),
                ("YoY swing", row["YoY Flag"]),
                ("Large value move", row["Value Flag"]),
                ("Unreconciled", row["Recon Flag"]),
            ]
            if flag
        ),
        axis=1,
    )

    anomalies = anomalies.sort_values(
        by="Variance (EUR)", key=lambda col: col.abs(), ascending=False
    )

    top_anomalies = anomalies.head(6)
    commentary_blocks = []
    for _, row in top_anomalies.head(3).iterrows():
        context = row["Context Cues"] if row["Context Cues"] else "No supporting notes captured yet."
        commentary_blocks.append(
            f"{row['Entity']} {row['Account Name']} moved {row['Variance (EUR)']:.0f} EUR "
            f"({row['MoM %']:.1f}% MoM, {row['YoY %']:.1f}% YoY). Context: {context}"
        )

    summary = {
        "mom_anomalies": int(anomalies["MoM Flag"].sum()),
        "yoy_anomalies": int(anomalies["YoY Flag"].sum()),
        "unreconciled": int(anomalies["Recon Flag"].sum()),
        "top_variance_account": top_anomalies["Account Name"].iloc[0] if not top_anomalies.empty else "None",
        "top_variance_amount": round(float(top_anomalies["Variance (EUR)"].iloc[0]), 2)
        if not top_anomalies.empty
        else 0.0,
        "yoy_source": frame["YoY Source"].iloc[0] if not frame.empty else "Derived proxy",
    }

    worklist = anomalies.loc[
        :,
        [
            "Entity",
            "Account Code",
            "Account Name",
            "Account Type",
            "Closing Balance (EUR)",
            "Prior Month Balance (EUR)",
            "Variance (EUR)",
            "MoM %",
            "YoY %",
            "YoY Source",
            "Reconciled",
            "Reviewer Notes",
            "Context Cues",
            "Anomaly Reason",
        ],
    ].copy()

    return {
        "summary": summary,
        "worklist": worklist,
        "commentary": commentary_blocks,
        "anomalies": top_anomalies.loc[
            :,
            ["Entity", "Account Name", "Variance (EUR)", "MoM %", "YoY %", "Anomaly Reason", "Context Cues"],
        ].copy(),
    }
    if "Step " in text:
        text = text.split("Step ", 1)[1]
    try:
        return int(float(text))
    except ValueError:
        return None


def _count_open_downstream(step, children_by_step, open_steps, memo):
    if step in memo:
        return memo[step]

    total = 0
    for child in children_by_step.get(step, []):
        if child in open_steps:
            total += 1
        total += _count_open_downstream(child, children_by_step, open_steps, memo)

    memo[step] = total
    return total


def _ic_pair_label(row):
    return f"{row['Sending Entity']} -> {row['Receiving Entity']}"


def _ic_fx_flag(row):
    status = str(row["Elimination Status"]).lower()
    notes = str(row["Notes"]).lower() if pd.notna(row["Notes"]) else ""
    difference = abs(float(pd.to_numeric(row["Difference (EUR)"], errors="coerce") or 0))
    return difference > 0 and ("fx" in status or "fx" in notes)


def _ic_primary_issue(row):
    difference = float(pd.to_numeric(row["Difference (EUR)"], errors="coerce") or 0)
    abs_difference = abs(difference)
    tp_status = str(row["Transfer Pricing Compliant"])
    agreement = str(row["Supporting Agreement"])
    fx_flag = _ic_fx_flag(row)

    if abs_difference > 0 and fx_flag:
        return (
            "FX mismatch",
            "Generate FX elimination entry",
            "Counterpart values are out of line because of an FX timing or rate difference.",
            "Group Consolidation",
            88,
        )
    if abs_difference > 0:
        return (
            "Elimination variance",
            "Generate elimination entry",
            "Counterpart amounts are mismatched and need a consolidating elimination adjustment.",
            "Group Consolidation",
            82,
        )
    if tp_status != "Yes":
        return (
            "Transfer pricing review",
            "Escalate transfer pricing review",
            "The IC transaction is matched, but transfer pricing is still under review.",
            "Tax / Controller",
            72,
        )
    if agreement == "Pending":
        return (
            "Agreement pending",
            "Request supporting agreement",
            "The values are matched, but the supporting intercompany agreement is still pending.",
            "Legal / Controller",
            76,
        )
    return (
        "Auto-matched",
        "Auto-clear matched pair",
        "Sending and receiving amounts align, so the IC pair can clear without manual intervention.",
        "IC Reconciler",
        95,
    )


def _ic_elimination_draft(row, issue_type, owner, confidence):
    difference = float(pd.to_numeric(row["Difference (EUR)"], errors="coerce") or 0)
    abs_difference = abs(difference)
    posting_date = _period_to_posting_date(row["Period"])
    pair = _ic_pair_label(row)

    if issue_type == "FX mismatch":
        offset_code, offset_name = 8100, "Foreign Exchange Gain/Loss"
        je_type = "FX Elimination Adjustment"
    else:
        offset_code, offset_name = 8200, "Intercompany Elimination Variance"
        je_type = "IC Elimination Adjustment"

    clearing_code, clearing_name = 1590, "Intercompany Clearing"
    if difference >= 0:
        debit_code, debit_name = offset_code, offset_name
        credit_code, credit_name = clearing_code, clearing_name
    else:
        debit_code, debit_name = clearing_code, clearing_name
        credit_code, credit_name = offset_code, offset_name

    support_pack = (
        f"{row['Supporting Agreement']} agreement + IC reconciliation notes"
        if pd.notna(row["Supporting Agreement"])
        else "IC reconciliation notes"
    )

    return {
        "Journal ID": f"ICA-JE-{str(row['IC Transaction ID']).split('-')[-1]}",
        "Source ID": row["IC Transaction ID"],
        "Pair": pair,
        "JE Type": je_type,
        "Posting Date": posting_date,
        "Currency": "EUR",
        "Debit Account": debit_code,
        "Debit Account Name": debit_name,
        "Credit Account": credit_code,
        "Credit Account Name": credit_name,
        "Amount (EUR)": round(abs_difference, 2),
        "Reference": row["IC Transaction ID"],
        "Header Text": f"IC agent draft - {row['Transaction Type']} elimination",
        "Line Text": f"{row['Transaction Type']} | {pair}",
        "Support Pack": support_pack,
        "Approval Status": "Draft - Ready for ERP",
        "Confidence %": confidence,
        "Owner": owner,
    }


def _gl_approval_decision(row):
    amount = float(pd.to_numeric(row["Entry Amount (EUR)"], errors="coerce") or 0)
    source = str(row["Source"])
    account_type = str(row["Account Type"])
    status = str(row["Status"])
    is_intercompany = str(row["Intercompany Flag"]) == "Yes"

    if amount >= 200000 or account_type == "Equity":
        return (
            "CFO queue",
            "CFO",
            "Executive approval required",
            "High-value or equity-impact journal requires CFO sign-off before posting.",
            64,
        )

    if source == "Manual Entry" or is_intercompany or account_type == "Revenue" or amount >= 125000:
        return (
            "Controller queue",
            "Controller",
            "Controller review",
            "Manual, intercompany, revenue, or high-value content makes this a controller approval item.",
            72,
        )

    if (
        source in {"ERP Auto", "Recurring", "Payroll System", "Fixed Assets"}
        and not is_intercompany
        and account_type not in {"Revenue", "Equity"}
        and amount < 100000
    ):
        return (
            "Straight-through",
            "Auto-approval batch",
            "Ready for auto-approval",
            "System-generated recurring journal falls inside the straight-through tolerance band.",
            91 if status == "Pending Approval" else 87,
        )

    return (
        "Manager queue",
        "Accounting Manager",
        "Manager review",
        "Low-to-medium risk journal can be cleared through the manager approval lane.",
        80 if status == "Pending Approval" else 76,
    )


def _build_gl_agent(gl):
    frame = gl.copy()
    frame["Entry Amount (EUR)"] = frame[["Debit (EUR)", "Credit (EUR)"]].max(axis=1)
    pending = frame.loc[frame["Status"].isin(["Pending Review", "Pending Approval"])].copy()

    if pending.empty:
        return {
            "summary": {
                "pending_items": 0,
                "straight_through_candidates": 0,
                "manager_queue": 0,
                "controller_queue": 0,
                "cfo_queue": 0,
                "manual_high_risk": 0,
                "erp_ready_after_approval": 0,
                "recoverable_hours": 0,
            },
            "worklist": pd.DataFrame(),
            "approval_packets": pd.DataFrame(),
            "approval_ready": pd.DataFrame(),
            "lane_breakdown": pd.DataFrame(columns=["Approval Lane", "Count"]),
            "entity_breakdown": pd.DataFrame(columns=["Entity", "Pending Items"]),
        }

    records = []
    for _, row in pending.iterrows():
        lane, approver, approval_action, rationale, confidence = _gl_approval_decision(row)
        amount = round(float(row["Entry Amount (EUR)"]), 2)
        records.append(
            {
                "Journal Entry ID": row["Journal Entry ID"],
                "Entity": row["Entity"],
                "Account Name": row["Account Name"],
                "Account Type": row["Account Type"],
                "Source": row["Source"],
                "Status": row["Status"],
                "Entry Amount (EUR)": amount,
                "Approval Lane": lane,
                "Required Approver": approver,
                "Approval Action": approval_action,
                "Intercompany": row["Intercompany Flag"],
                "Created By": row["Created By"],
                "Approved By": row["Approved By"] if pd.notna(row["Approved By"]) else "Unassigned",
                "Confidence %": confidence,
                "Rationale": rationale,
            }
        )

    worklist = pd.DataFrame(records)
    lane_rank = {
        "CFO queue": 3,
        "Controller queue": 2,
        "Manager queue": 1,
        "Straight-through": 0,
    }
    worklist["Lane Rank"] = worklist["Approval Lane"].map(lane_rank).fillna(0)
    worklist = worklist.sort_values(
        ["Lane Rank", "Entry Amount (EUR)", "Confidence %"],
        ascending=[False, False, False],
    ).drop(columns=["Lane Rank"])

    approval_packets = worklist.loc[
        worklist["Approval Lane"].isin(["Manager queue", "Controller queue", "CFO queue"]),
        [
            "Journal Entry ID",
            "Entity",
            "Account Name",
            "Source",
            "Status",
            "Entry Amount (EUR)",
            "Approval Lane",
            "Required Approver",
            "Rationale",
        ],
    ].copy()

    approval_ready = worklist.loc[
        worklist["Approval Lane"].eq("Straight-through"),
        [
            "Journal Entry ID",
            "Entity",
            "Account Name",
            "Source",
            "Status",
            "Entry Amount (EUR)",
            "Approval Action",
            "Confidence %",
        ],
    ].copy()

    lane_breakdown = (
        worklist["Approval Lane"].value_counts()
        .rename_axis("Approval Lane")
        .reindex(["Straight-through", "Manager queue", "Controller queue", "CFO queue"], fill_value=0)
        .reset_index(name="Count")
    )
    entity_breakdown = (
        worklist["Entity"].value_counts().rename_axis("Entity").reset_index(name="Pending Items")
    )

    summary = {
        "pending_items": int(worklist.shape[0]),
        "straight_through_candidates": int(worklist["Approval Lane"].eq("Straight-through").sum()),
        "manager_queue": int(worklist["Approval Lane"].eq("Manager queue").sum()),
        "controller_queue": int(worklist["Approval Lane"].eq("Controller queue").sum()),
        "cfo_queue": int(worklist["Approval Lane"].eq("CFO queue").sum()),
        "manual_high_risk": int(
            worklist["Source"].eq("Manual Entry").sum()
            + worklist["Intercompany"].eq("Yes").sum()
        ),
        "erp_ready_after_approval": int(approval_ready.shape[0]),
        "recoverable_hours": int(min(12, round(approval_ready.shape[0] * 0.3 + 3))),
    }

    return {
        "summary": summary,
        "worklist": worklist,
        "approval_packets": approval_packets,
        "approval_ready": approval_ready,
        "lane_breakdown": lane_breakdown,
        "entity_breakdown": entity_breakdown,
    }


def _build_journal_agent(accruals):
    candidate_mask = accruals["Status"].isin(["Pending Review", "Pending Approval", "Needs Documentation"])
    candidates = accruals.loc[candidate_mask].copy()
    if candidates.empty:
        return {
            "summary": {
                "candidate_items": 0,
                "erp_ready_drafts": 0,
                "contract_backed_drafts": 0,
                "standard_reclasses": 0,
                "audit_trails_attached": 0,
                "review_needed": 0,
            },
            "worklist": pd.DataFrame(),
            "erp_journal_drafts": pd.DataFrame(),
            "audit_trails": pd.DataFrame(),
            "type_breakdown": pd.DataFrame(columns=["JE Type", "Count"]),
            "entity_breakdown": pd.DataFrame(columns=["Entity", "Drafts"]),
        }

    records = []
    draft_rows = []
    audit_rows = []

    for _, row in candidates.iterrows():
        mapping = _journal_agent_mapping(row)
        confidence = _journal_agent_confidence(row, mapping)
        posting_date = _period_to_posting_date(row["Period"])
        support = str(row["Supporting Doc"]) if pd.notna(row["Supporting Doc"]) else "Missing"
        ready_status = _journal_ready_status(confidence, mapping["mapped"])
        je_type = mapping["je_type"]
        source_mode = (
            "Contract-backed accrual"
            if support in {"Contract ref", "PO attached", "Email confirmation"}
            else "Estimate-based accrual"
        )
        if je_type == "Standard Reclassification":
            source_mode = "Standard reclassification"

        audit_trail = _journal_agent_audit_trail(row, mapping, confidence)
        amount = round(abs(float(row["Accrual Amount (EUR)"])), 2)
        journal_id = f"JEA-{str(row['Accrual ID']).split('-')[-1]}"

        records.append(
            {
                "Source ID": row["Accrual ID"],
                "Entity": row["Entity"],
                "Description": row["Description"],
                "JE Type": je_type,
                "Source Mode": source_mode,
                "Support": support,
                "Status": row["Status"],
                "Amount (EUR)": amount,
                "Confidence %": confidence,
                "Approval Status": ready_status,
                "Recommended Action": mapping["rationale"],
            }
        )

        if ready_status == "Draft - Ready for ERP":
            draft_rows.append(
                {
                    "Journal ID": journal_id,
                    "Source ID": row["Accrual ID"],
                    "Entity": row["Entity"],
                    "JE Type": je_type,
                    "Posting Date": posting_date,
                    "Currency": "EUR",
                    "Debit Account": mapping["debit_code"],
                    "Debit Account Name": mapping["debit_name"],
                    "Credit Account": mapping["credit_code"],
                    "Credit Account Name": mapping["credit_name"],
                    "Amount (EUR)": amount,
                    "Reference": row["Accrual ID"],
                    "Header Text": f"Journal agent draft - {row['Description']}",
                    "Line Text": f"{row['Description']} | {source_mode}",
                    "Support Pack": support,
                    "Approval Status": ready_status,
                    "Confidence %": confidence,
                    "AI Audit Trail": audit_trail,
                }
            )

        audit_rows.append(
            {
                "Journal ID": journal_id,
                "Source ID": row["Accrual ID"],
                "Entity": row["Entity"],
                "JE Type": je_type,
                "Approval Status": ready_status,
                "Confidence %": confidence,
                "AI Audit Trail": audit_trail,
            }
        )

    worklist = pd.DataFrame(records).sort_values(
        ["Approval Status", "Confidence %", "Amount (EUR)"],
        ascending=[True, False, False],
    )
    erp_journal_drafts = pd.DataFrame(draft_rows)
    audit_trails = pd.DataFrame(audit_rows)
    type_breakdown = (
        worklist["JE Type"].value_counts().rename_axis("JE Type").reset_index(name="Count")
    )
    entity_breakdown = (
        worklist.loc[worklist["Approval Status"].eq("Draft - Ready for ERP"), "Entity"]
        .value_counts()
        .rename_axis("Entity")
        .reset_index(name="Drafts")
    )

    summary = {
        "candidate_items": int(worklist.shape[0]),
        "erp_ready_drafts": int((worklist["Approval Status"] == "Draft - Ready for ERP").sum()),
        "contract_backed_drafts": int(
            worklist.loc[worklist["Source Mode"].eq("Contract-backed accrual") & worklist["Approval Status"].eq("Draft - Ready for ERP")].shape[0]
        ),
        "standard_reclasses": int(
            worklist.loc[worklist["JE Type"].eq("Standard Reclassification") & worklist["Approval Status"].eq("Draft - Ready for ERP")].shape[0]
        ),
        "audit_trails_attached": int(audit_trails.shape[0]),
        "review_needed": int((worklist["Approval Status"] == "Draft - Review Needed").sum()),
    }

    return {
        "summary": summary,
        "worklist": worklist,
        "erp_journal_drafts": erp_journal_drafts,
        "audit_trails": audit_trails,
        "type_breakdown": type_breakdown,
        "entity_breakdown": entity_breakdown,
    }


def _build_audit_agent(bank_agent, ic_agent, journal_agent):
    bank_drafts = bank_agent["erp_journal_drafts"].copy()
    ic_drafts = ic_agent["erp_elimination_drafts"].copy()
    journal_drafts = journal_agent["erp_journal_drafts"].copy()

    if not bank_drafts.empty:
        bank_drafts["Source Agent"] = "Bank Agent"
        bank_drafts["JE Type"] = bank_drafts.get("JE Type", "Cash Clearing JE")
        if "Source ID" not in bank_drafts.columns:
            bank_drafts["Source ID"] = bank_drafts["Reference"]
        if "AI Audit Trail" not in bank_drafts.columns:
            bank_drafts["AI Audit Trail"] = ""

    if not ic_drafts.empty:
        ic_drafts["Source Agent"] = "IC Agent"
        if "AI Audit Trail" not in ic_drafts.columns:
            ic_drafts["AI Audit Trail"] = ""

    if not journal_drafts.empty:
        journal_drafts["Source Agent"] = "Journal Agent"

    combined = pd.concat([bank_drafts, ic_drafts, journal_drafts], ignore_index=True, sort=False)
    if combined.empty:
        return {
            "summary": {
                "drafts_reviewed": 0,
                "ready_to_post": 0,
                "conditional_approval": 0,
                "blocked_for_review": 0,
                "average_control_score": 0,
                "high_value_items": 0,
                "control_memos_attached": 0,
            },
            "review_pack": pd.DataFrame(),
            "exception_queue": pd.DataFrame(),
            "status_breakdown": pd.DataFrame(columns=["Posting Recommendation", "Count"]),
            "source_breakdown": pd.DataFrame(columns=["Source Agent", "Drafts"]),
            "control_memos": pd.DataFrame(),
        }

    review_rows = []
    for _, row in combined.iterrows():
        review = _audit_agent_review(row)
        review_rows.append(
            {
                "Journal ID": row.get("Journal ID"),
                "Source Agent": row.get("Source Agent"),
                "Source ID": row.get("Source ID"),
                "Entity": row.get("Entity"),
                "JE Type": row.get("JE Type", "Journal Entry"),
                "Posting Date": row.get("Posting Date"),
                "Currency": row.get("Currency", "EUR"),
                "Amount (EUR)": float(pd.to_numeric(row.get("Amount (EUR)"), errors="coerce") or 0),
                "Reference": row.get("Reference"),
                "Support Pack": row.get("Support Pack"),
                "Confidence %": int(pd.to_numeric(row.get("Confidence %"), errors="coerce") or 0),
                "Debit Account": row.get("Debit Account"),
                "Credit Account": row.get("Credit Account"),
                "Header Text": row.get("Header Text"),
                "Line Text": row.get("Line Text"),
                "Approval Status": row.get("Approval Status"),
                "Control Score": review["Control Score"],
                "Posting Recommendation": review["Posting Recommendation"],
                "Required Approver": review["Required Approver"],
                "Primary Control Gap": review["Primary Control Gap"],
                "Control Gaps": review["Control Gaps"],
                "Blocking Issues": review["Blocking Issues"],
                "Control Memo": review["Control Memo"],
            }
        )

    review_pack = pd.DataFrame(review_rows)
    status_order = {
        "Blocked for Review": 2,
        "Conditional Approval": 1,
        "Ready to Post": 0,
    }
    review_pack["Status Rank"] = review_pack["Posting Recommendation"].map(status_order).fillna(0)
    review_pack = review_pack.sort_values(
        ["Status Rank", "Control Score", "Amount (EUR)"],
        ascending=[False, True, False],
    ).drop(columns=["Status Rank"])

    exception_queue = review_pack.loc[
        ~review_pack["Posting Recommendation"].eq("Ready to Post"),
        [
            "Journal ID",
            "Source Agent",
            "Entity",
            "JE Type",
            "Amount (EUR)",
            "Control Score",
            "Posting Recommendation",
            "Required Approver",
            "Primary Control Gap",
            "Blocking Issues",
        ],
    ].copy()

    status_breakdown = (
        review_pack["Posting Recommendation"]
        .value_counts()
        .reindex(["Ready to Post", "Conditional Approval", "Blocked for Review"], fill_value=0)
        .rename_axis("Posting Recommendation")
        .reset_index(name="Count")
    )
    source_breakdown = (
        review_pack.groupby("Source Agent")
        .agg(
            Drafts=("Journal ID", "count"),
            Avg_Control_Score=("Control Score", "mean"),
        )
        .reset_index()
    )
    source_breakdown["Avg Control Score"] = source_breakdown["Avg_Control_Score"].round(1)
    source_breakdown = source_breakdown.drop(columns=["Avg_Control_Score"])

    control_memos = review_pack.loc[
        :,
        [
            "Journal ID",
            "Source Agent",
            "Posting Recommendation",
            "Control Score",
            "Required Approver",
            "Control Memo",
        ],
    ].copy()

    summary = {
        "drafts_reviewed": int(review_pack.shape[0]),
        "ready_to_post": int(review_pack["Posting Recommendation"].eq("Ready to Post").sum()),
        "conditional_approval": int(review_pack["Posting Recommendation"].eq("Conditional Approval").sum()),
        "blocked_for_review": int(review_pack["Posting Recommendation"].eq("Blocked for Review").sum()),
        "average_control_score": round(float(review_pack["Control Score"].mean()), 1),
        "high_value_items": int(review_pack["Amount (EUR)"].ge(50000).sum()),
        "control_memos_attached": int(control_memos.shape[0]),
    }

    return {
        "summary": summary,
        "review_pack": review_pack,
        "exception_queue": exception_queue,
        "status_breakdown": status_breakdown,
        "source_breakdown": source_breakdown,
        "control_memos": control_memos,
    }


def _build_checklist_agent(checklist):
    frame = checklist.copy()
    frame["Dependency Step"] = frame["Dependencies"].apply(_dependency_step)
    open_mask = ~frame["Status"].eq("Completed")
    open_tasks = frame.loc[open_mask].copy()

    if open_tasks.empty:
        return {
            "summary": {
                "blocked_tasks": 0,
                "waiting_tasks": 0,
                "critical_open_tasks": 0,
                "handoffs_at_risk": 0,
                "recoverable_hours": 0,
                "automation_candidates": 0,
            },
            "worklist": pd.DataFrame(),
            "critical_path": pd.DataFrame(),
            "handoff_queue": pd.DataFrame(),
            "action_breakdown": pd.DataFrame(columns=["Action Lane", "Count"]),
            "category_breakdown": pd.DataFrame(columns=["Category", "Open Tasks", "Hours at Risk"]),
        }

    step_lookup = frame.set_index("Step")[
        ["Task", "Status", "Owner / Responsible", "Priority", "Category"]
    ].to_dict(orient="index")
    children_by_step = {}
    for _, row in frame.dropna(subset=["Dependency Step"]).iterrows():
        children_by_step.setdefault(int(row["Dependency Step"]), []).append(int(row["Step"]))

    open_steps = set(open_tasks["Step"].astype(int).tolist())
    memo = {}
    records = []

    for _, row in open_tasks.iterrows():
        step = int(row["Step"])
        dependency_step = row["Dependency Step"]
        dependency_record = step_lookup.get(int(dependency_step), {}) if pd.notna(dependency_step) else {}
        dependency_task = dependency_record.get("Task", "No upstream dependency")
        dependency_status = dependency_record.get("Status", "Ready")
        dependency_owner = dependency_record.get("Owner / Responsible", "N/A")
        downstream_open = _count_open_downstream(step, children_by_step, open_steps, memo)

        status = str(row["Status"])
        priority = str(row["Priority"])
        owner = str(row["Owner / Responsible"])
        estimated_hours = float(pd.to_numeric(row["Estimated Hours"], errors="coerce") or 0)
        automation_potential = str(row["Automation Potential"])
        notes = str(row["Notes"]) if pd.notna(row["Notes"]) else ""

        status_weight = {
            "Blocked": 5,
            "Waiting on Input": 4,
            "Not Started": 3,
            "In Progress": 2,
        }.get(status, 1)
        priority_weight = {"Critical": 5, "High": 3, "Medium": 2, "Low": 1}.get(priority, 1)
        unblock_score = _clamp((status_weight * 10) + (priority_weight * 8) + min(downstream_open * 6, 24) + min(estimated_hours, 8))

        if status == "Blocked":
            action_lane = "Escalate owner"
            next_action = f"Escalate {owner} and clear the dependency handoff from {dependency_owner}."
        elif status == "Waiting on Input":
            if dependency_task == "No upstream dependency":
                action_lane = "Chase external input"
                next_action = f"Obtain the missing external confirmation so {owner} can resume and release downstream steps."
            else:
                action_lane = "Chase dependency"
                next_action = f"Obtain the missing handoff from {dependency_owner} so {owner} can resume."
        elif priority == "Critical" and status == "Not Started":
            action_lane = "Start in parallel"
            next_action = f"Pre-stage support with {owner} and begin parallel prep before the upstream task fully closes."
        elif priority == "Critical" and status == "In Progress":
            action_lane = "Protect same-day completion"
            next_action = f"Keep {owner} on the critical path and hand off immediately to downstream teams."
        else:
            action_lane = "Monitor"
            next_action = f"Keep {owner} on schedule and confirm the next dependency stays on track."

        records.append(
            {
                "Step": step,
                "Category": row["Category"],
                "Task": row["Task"],
                "Owner": owner,
                "Deadline": row["Deadline"],
                "Status": status,
                "Priority": priority,
                "Dependency Task": dependency_task,
                "Dependency Status": dependency_status,
                "Dependency Owner": dependency_owner,
                "Downstream Open Tasks": downstream_open,
                "Estimated Hours": round(estimated_hours, 2),
                "Automation Potential": automation_potential,
                "Tool / System": row["Tool / System"],
                "Action Lane": action_lane,
                "Next Action": next_action,
                "Unblock Score": unblock_score,
                "Notes": notes or "No additional note",
            }
        )

    worklist = pd.DataFrame(records).sort_values(
        ["Unblock Score", "Estimated Hours"],
        ascending=[False, False],
    )
    critical_path = worklist.loc[
        worklist["Priority"].eq("Critical"),
        [
            "Step",
            "Category",
            "Task",
            "Status",
            "Owner",
            "Dependency Task",
            "Downstream Open Tasks",
            "Estimated Hours",
            "Action Lane",
            "Next Action",
            "Unblock Score",
        ],
    ].copy()
    handoff_queue = worklist.loc[
        worklist["Status"].isin(["Blocked", "Waiting on Input"]),
        [
            "Step",
            "Category",
            "Task",
            "Status",
            "Priority",
            "Owner",
            "Dependency Task",
            "Dependency Status",
            "Dependency Owner",
            "Estimated Hours",
            "Action Lane",
            "Next Action",
            "Unblock Score",
        ],
    ].copy()
    action_breakdown = (
        worklist["Action Lane"].value_counts().rename_axis("Action Lane").reset_index(name="Count")
    )
    category_breakdown = (
        worklist.groupby("Category")
        .agg(
            **{
                "Open Tasks": ("Task", "count"),
                "Hours at Risk": ("Estimated Hours", "sum"),
            }
        )
        .reset_index()
        .sort_values(["Hours at Risk", "Open Tasks"], ascending=[False, False])
    )
    category_breakdown["Hours at Risk"] = category_breakdown["Hours at Risk"].round(1)

    waiting_or_blocked = worklist["Status"].isin(["Blocked", "Waiting on Input"])
    critical_open = worklist["Priority"].eq("Critical")
    recoverable_hours = int(
        round(
            worklist.loc[waiting_or_blocked, "Estimated Hours"].max()
            if waiting_or_blocked.any()
            else worklist["Estimated Hours"].head(1).max()
        )
    )

    automation_candidates = int(
        worklist["Automation Potential"].astype(str).str.contains("High|AI candidate", case=False, na=False).sum()
    )
    dependency_summary = {
        "external_inputs": int(worklist["Dependency Task"].eq("No upstream dependency").sum()),
        "upstream_in_progress": int(worklist["Dependency Status"].eq("In Progress").sum()),
        "upstream_not_started": int(worklist["Dependency Status"].eq("Not Started").sum()),
        "upstream_waiting": int(worklist["Dependency Status"].eq("Waiting on Input").sum()),
        "completed_handoffs_open": int(worklist["Dependency Status"].eq("Completed").sum()),
    }

    summary = {
        "blocked_tasks": int(worklist["Status"].eq("Blocked").sum()),
        "waiting_tasks": int(worklist["Status"].eq("Waiting on Input").sum()),
        "critical_open_tasks": int(critical_open.sum()),
        "handoffs_at_risk": int(
            worklist.loc[waiting_or_blocked | critical_open, "Dependency Task"].ne("No upstream dependency").sum()
        ),
        "recoverable_hours": recoverable_hours,
        "automation_candidates": automation_candidates,
    }

    return {
        "summary": summary,
        "dependency_summary": dependency_summary,
        "worklist": worklist,
        "critical_path": critical_path,
        "handoff_queue": handoff_queue,
        "action_breakdown": action_breakdown,
        "category_breakdown": category_breakdown,
    }


def _build_ic_agent(ic):
    frame = ic.copy()
    if frame.empty:
        return {
            "summary": {
                "total_pairs": 0,
                "auto_matched": 0,
                "open_exceptions": 0,
                "fx_mismatches": 0,
                "elimination_drafts": 0,
                "tp_flags": 0,
                "pending_agreements": 0,
                "unresolved_difference": 0.0,
            },
            "worklist": pd.DataFrame(),
            "issue_breakdown": pd.DataFrame(columns=["Issue Type", "Count"]),
            "pair_breakdown": pd.DataFrame(columns=["Pair", "Open Items", "Unresolved Difference (EUR)"]),
            "erp_elimination_drafts": pd.DataFrame(),
            "tp_watchlist": pd.DataFrame(),
            "auto_matches": pd.DataFrame(),
        }

    records = []
    draft_rows = []

    for _, row in frame.iterrows():
        pair = _ic_pair_label(row)
        issue_type, recommended_action, likely_cause, owner, confidence = _ic_primary_issue(row)
        difference = float(pd.to_numeric(row["Difference (EUR)"], errors="coerce") or 0)
        abs_difference = abs(difference)
        tp_status = str(row["Transfer Pricing Compliant"])
        agreement = str(row["Supporting Agreement"])
        fx_flag = _ic_fx_flag(row)
        open_exception = abs_difference > 0
        tp_flag = tp_status != "Yes" or agreement == "Pending"
        match_outcome = "Auto-matched" if abs_difference == 0 else "Exception"

        records.append(
            {
                "IC Transaction ID": row["IC Transaction ID"],
                "Pair": pair,
                "Transaction Type": row["Transaction Type"],
                "Match Outcome": match_outcome,
                "Issue Type": issue_type,
                "Currency": row["Currency"],
                "Sending Amount (EUR)": float(pd.to_numeric(row["Sending Amount (EUR)"], errors="coerce") or 0),
                "Receiving Amount (EUR)": float(pd.to_numeric(row["Receiving Amount (EUR)"], errors="coerce") or 0),
                "Difference (EUR)": round(difference, 2),
                "Elimination Status": row["Elimination Status"],
                "FX Flag": "Yes" if fx_flag else "No",
                "Transfer Pricing": tp_status,
                "Supporting Agreement": agreement,
                "Reconciler": row["Reconciler"],
                "Recommended Action": recommended_action,
                "Likely Cause": likely_cause,
                "Owner": owner,
                "Confidence %": confidence,
            }
        )

        if open_exception:
            draft_rows.append(_ic_elimination_draft(row, issue_type, owner, confidence))

    all_rows = pd.DataFrame(records)
    worklist = all_rows.loc[
        all_rows["Issue Type"].ne("Auto-matched"),
        :,
    ].copy()
    issue_rank = {
        "FX mismatch": 4,
        "Elimination variance": 3,
        "Transfer pricing review": 2,
        "Agreement pending": 1,
    }
    if not worklist.empty:
        worklist["Issue Rank"] = worklist["Issue Type"].map(issue_rank).fillna(0)
        worklist = worklist.sort_values(
            ["Issue Rank", "Difference (EUR)", "Confidence %"],
            ascending=[False, False, False],
        ).drop(columns=["Issue Rank"])

    erp_elimination_drafts = pd.DataFrame(draft_rows)
    auto_matches = all_rows.loc[
        all_rows["Match Outcome"].eq("Auto-matched"),
        [
            "IC Transaction ID",
            "Pair",
            "Transaction Type",
            "Currency",
            "Transfer Pricing",
            "Supporting Agreement",
            "Recommended Action",
            "Confidence %",
        ],
    ].copy()
    issue_breakdown = (
        worklist["Issue Type"].value_counts().rename_axis("Issue Type").reset_index(name="Count")
        if not worklist.empty
        else pd.DataFrame(columns=["Issue Type", "Count"])
    )
    pair_breakdown = (
        worklist.groupby("Pair")
        .agg(
            **{
                "Open Items": ("IC Transaction ID", "count"),
                "Unresolved Difference (EUR)": ("Difference (EUR)", lambda s: round(s.abs().sum(), 2)),
                "FX Flags": ("FX Flag", lambda s: int((s == "Yes").sum())),
                "TP Flags": ("Transfer Pricing", lambda s: int((s != "Yes").sum())),
            }
        )
        .reset_index()
        .sort_values(["Unresolved Difference (EUR)", "Open Items"], ascending=[False, False])
        if not worklist.empty
        else pd.DataFrame(columns=["Pair", "Open Items", "Unresolved Difference (EUR)", "FX Flags", "TP Flags"])
    )
    tp_watchlist = all_rows.loc[
        all_rows["Transfer Pricing"].ne("Yes") | all_rows["Supporting Agreement"].eq("Pending"),
        [
            "IC Transaction ID",
            "Pair",
            "Transaction Type",
            "Transfer Pricing",
            "Supporting Agreement",
            "Recommended Action",
            "Owner",
            "Confidence %",
        ],
    ].copy()

    summary = {
        "total_pairs": int(frame.shape[0]),
        "auto_matched": int(all_rows["Match Outcome"].eq("Auto-matched").sum()),
        "open_exceptions": int(all_rows["Match Outcome"].eq("Exception").sum()),
        "fx_mismatches": int(all_rows["FX Flag"].eq("Yes").sum()),
        "elimination_drafts": int(erp_elimination_drafts.shape[0]),
        "tp_flags": int(tp_watchlist.shape[0]),
        "pending_agreements": int(all_rows["Supporting Agreement"].eq("Pending").sum()),
        "unresolved_difference": round(float(all_rows["Difference (EUR)"].abs().sum()), 2),
    }

    return {
        "summary": summary,
        "worklist": worklist,
        "issue_breakdown": issue_breakdown,
        "pair_breakdown": pair_breakdown,
        "erp_elimination_drafts": erp_elimination_drafts,
        "tp_watchlist": tp_watchlist,
        "auto_matches": auto_matches,
    }


def _build_bank_agent(bank):
    open_items = bank.loc[~bank["Match Status"].eq("Matched")].copy()
    if open_items.empty:
        return {
            "summary": {
                "open_items": 0,
                "auto_clear_candidates": 0,
                "journal_candidates": 0,
                "escalations": 0,
                "manual_investigations": 0,
                "statement_break_value": 0.0,
                "largest_difference": 0.0,
            },
            "worklist": pd.DataFrame(),
            "status_breakdown": pd.DataFrame(columns=["Match Status", "Count"]),
            "entity_breakdown": pd.DataFrame(columns=["Entity", "Open Items"]),
            "journal_candidates": pd.DataFrame(),
            "subledger_breaks": pd.DataFrame(),
        }

    open_items["Entity"] = open_items["Entity"].fillna("Group Statement")
    records = []

    for _, row in open_items.iterrows():
        status = row["Match Status"]
        description = str(row["Description"])
        notes = str(row["Reconciler Notes"])
        counterparty = str(row["Counterparty"])
        entity = row["Entity"]
        days = float(row["Days Outstanding"]) if pd.notna(row["Days Outstanding"]) else 0.0
        difference = float(row["Difference (EUR)"]) if pd.notna(row["Difference (EUR)"]) else 0.0
        abs_difference = abs(difference)

        notes_lower = notes.lower()
        description_lower = description.lower()

        if status == "Timing Difference":
            action_bucket = "Auto-clear next month"
            likely_cause = "Normal cut-off timing lag between bank value date and book posting."
            action = "Carry forward and auto-clear next month if the timing item reverses as expected."
            owner = "Cash accountant"
            confidence = 93 if days <= 14 else 79
            suggested_journal = "No journal; monitor the reversal window."
        elif status == "Stale Item":
            if "duplicate" in notes_lower:
                action_bucket = "Manual investigation"
                likely_cause = "Possible duplicate book-side posting."
                action = "Check the GL and subledger for duplicate postings and reverse the duplicate if confirmed."
                owner = "Cash accountant"
                confidence = 84
                suggested_journal = "Reverse duplicate entry if confirmed."
            elif "vendor confirmation" in notes_lower:
                action_bucket = "Escalate"
                likely_cause = "External confirmation is still missing on an aged bank item."
                action = "Escalate to the bank or vendor and collect support before close certification."
                owner = "Treasury / AP"
                confidence = 79
                suggested_journal = "No journal until support is received."
            elif "timing" in notes_lower:
                action_bucket = "Escalate"
                likely_cause = "A timing item has aged beyond the normal clearing window."
                action = "Escalate the aged timing difference and clear it only after confirming the underlying transaction."
                owner = "Cash accountant"
                confidence = 74
                suggested_journal = "No journal unless the aged item is validated."
            else:
                action_bucket = "Escalate"
                likely_cause = "Aged reconciling item still unresolved."
                action = "Escalate the stale item and tie it to the underlying bank or subledger evidence."
                owner = "Treasury controller"
                confidence = 72
                suggested_journal = "No journal until root cause is confirmed."
        elif status == "Unmatched":
            if "timing" in notes_lower:
                action_bucket = "Auto-clear next month"
                likely_cause = "Likely cut-off timing lag that should clear in the next statement cycle."
                action = "Hold the item open, monitor the next statement, and only journal if the timing difference does not reverse."
                owner = "Cash accountant"
                confidence = 81
                suggested_journal = "No journal; monitor the reversal window."
            elif "fx" in notes_lower or abs_difference <= 300:
                action_bucket = "Journal candidate"
                likely_cause = "FX rounding or tolerance-level mismatch between bank and books."
                action = "Draft a small FX true-up or rounding journal and attach reconciliation support."
                owner = "Cash accountant"
                confidence = 88
                suggested_journal = f"Draft FX rounding journal for EUR {difference:,.2f}."
            else:
                action_bucket = "Manual investigation"
                likely_cause = "Missing or misclassified book-side entry."
                action = "Trace the bank line to the subledger and book the missing entry if the transaction is valid."
                owner = "Cash accountant"
                confidence = 77
                suggested_journal = "Draft book-side entry only after trace is confirmed."
        elif status == "Partial Match":
            action_bucket = "Journal candidate"
            likely_cause = "Partial settlement or batched book-side clearing."
            action = "Split the bank line against subledger items and post a residual clearing entry for the remaining difference."
            owner = "Cash accountant"
            confidence = 82
            suggested_journal = f"Draft residual clearing journal for EUR {difference:,.2f}."
        else:
            action_bucket = "Statement break"
            likely_cause = "Statement-level reconciliation break still exists after line matching."
            action = "Tie the statement balance break to the open reconciling inventory before close sign-off."
            owner = "Treasury controller"
            confidence = 74
            suggested_journal = "No journal until reconciling inventory is tied out."

        if "charge" in description_lower and action_bucket == "Auto-clear next month":
            likely_cause = "Bank fee posted in a different accounting period."
            action = "Confirm fee cut-off and auto-clear next month if the bank charge posts on the expected date."

        records.append(
            {
                "Line": int(row["Line"]) if pd.notna(row["Line"]) else None,
                "Entity": entity,
                "Transaction Date": row["Transaction Date"],
                "Value Date": row["Value Date"],
                "Bank Reference": row["Bank Reference"],
                "GL Account": int(row["GL Account"]) if pd.notna(row["GL Account"]) else None,
                "Description": description,
                "Match Status": status,
                "Difference (EUR)": round(difference, 2),
                "Days Outstanding": int(days) if days else 0,
                "Counterparty": counterparty,
                "Likely Cause": likely_cause,
                "Recommended Action": action,
                "Action Bucket": action_bucket,
                "Owner": owner,
                "Confidence %": confidence,
                "Suggested Journal": suggested_journal,
            }
        )

    worklist = pd.DataFrame(records)
    priority_order = {
        "Statement break": 4,
        "Escalate": 3,
        "Manual investigation": 2,
        "Journal candidate": 1,
        "Auto-clear next month": 0,
    }
    worklist["Priority Rank"] = worklist["Action Bucket"].map(priority_order).fillna(0)
    worklist = worklist.sort_values(
        ["Priority Rank", "Days Outstanding", "Confidence %"],
        ascending=[False, False, False],
    ).drop(columns=["Priority Rank"])

    journal_candidates = worklist.loc[
        worklist["Action Bucket"].eq("Journal candidate"),
        [
            "Line",
            "Entity",
            "Value Date",
            "Bank Reference",
            "GL Account",
            "Description",
            "Difference (EUR)",
            "Likely Cause",
            "Suggested Journal",
            "Confidence %",
        ],
    ].copy()

    draft_rows = []
    for _, row in journal_candidates.iterrows():
        bank_account_code = int(row["GL Account"]) if pd.notna(row["GL Account"]) else 1000
        bank_account_name = "Cash & Cash Equivalents" if bank_account_code == 1000 else "Bank Account"
        offset_code, offset_name = _bank_offset_account(row["Description"], row["Likely Cause"])
        amount = abs(float(row["Difference (EUR)"]))

        if float(row["Difference (EUR)"]) > 0:
            debit_code, debit_name = bank_account_code, bank_account_name
            credit_code, credit_name = offset_code, offset_name
        else:
            debit_code, debit_name = offset_code, offset_name
            credit_code, credit_name = bank_account_code, bank_account_name

        draft_rows.append(
            {
                "Journal ID": f"BRA-JE-{int(row['Line']):03d}",
                "Entity": row["Entity"],
                "Posting Date": row["Value Date"],
                "Currency": "EUR",
                "Debit Account": debit_code,
                "Debit Account Name": debit_name,
                "Credit Account": credit_code,
                "Credit Account Name": credit_name,
                "Amount (EUR)": round(amount, 2),
                "Reference": row["Bank Reference"],
                "Header Text": f"Bank rec agent draft - {row['Description']}",
                "Line Text": f"{row['Description']} | {row['Likely Cause']}",
                "Support Pack": "Bank rec exception + reconciler notes attached",
                "Approval Status": "Draft - Ready for ERP",
                "Confidence %": row["Confidence %"],
            }
        )

    erp_journal_drafts = pd.DataFrame(draft_rows)

    subledger_breaks = (
        worklist.groupby(["Entity", "Counterparty"], dropna=False)
        .agg(
            **{
                "Open Items": ("Description", "count"),
                "Net Difference (EUR)": ("Difference (EUR)", "sum"),
                "Average Age": ("Days Outstanding", "mean"),
            }
        )
        .reset_index()
        .sort_values(["Open Items", "Net Difference (EUR)"], ascending=[False, False])
    )
    subledger_breaks["Average Age"] = subledger_breaks["Average Age"].round(1)
    subledger_breaks = subledger_breaks.head(10)

    status_breakdown = (
        worklist["Match Status"].value_counts().rename_axis("Match Status").reset_index(name="Count")
    )
    entity_breakdown = (
        worklist["Entity"].value_counts().rename_axis("Entity").reset_index(name="Open Items")
    )

    summary = {
        "open_items": int(worklist.shape[0]),
        "auto_clear_candidates": int(worklist["Action Bucket"].eq("Auto-clear next month").sum()),
        "journal_candidates": int(worklist["Action Bucket"].eq("Journal candidate").sum()),
        "escalations": int(worklist["Action Bucket"].eq("Escalate").sum()),
        "manual_investigations": int(worklist["Action Bucket"].eq("Manual investigation").sum()),
        "statement_break_value": float(
            worklist.loc[worklist["Action Bucket"].eq("Statement break"), "Difference (EUR)"].abs().sum()
        ),
        "largest_difference": float(worklist["Difference (EUR)"].abs().max()),
    }

    return {
        "summary": summary,
        "worklist": worklist,
        "status_breakdown": status_breakdown,
        "entity_breakdown": entity_breakdown,
        "journal_candidates": journal_candidates,
        "erp_journal_drafts": erp_journal_drafts,
        "subledger_breaks": subledger_breaks,
    }


def _build_entity_rollups(data):
    gl = data["gl"].copy()
    tb = data["tb"].copy()
    accruals = data["accruals"].copy()
    bank = data["bank"].copy()
    ic = data["ic"].copy()
    fa = data["fa"].copy()
    apar = data["apar"].copy()

    tb["variance_pct_num"] = pd.to_numeric(
        tb["Variance %"].astype(str).str.rstrip("%"),
        errors="coerce",
    )

    ap = apar.loc[apar["Type"].eq("AP")].copy()
    ap["has_3way_exception"] = ap["3-Way Match"].notna() & ~ap["3-Way Match"].eq("Matched")

    ar = apar.loc[apar["Type"].eq("AR")].copy()
    ar["overdue"] = ar["Aging Bucket"].isin(["61-90 Days", "90+ Days"]) & ar["Outstanding (EUR)"].gt(0)

    entities = sorted(
        set(gl["Entity"].dropna())
        | set(tb["Entity"].dropna())
        | set(accruals["Entity"].dropna())
        | set(bank["Entity"].dropna())
        | set(fa["Entity"].dropna())
        | set(apar["Entity"].dropna())
        | set(ic["Sending Entity"].dropna())
        | set(ic["Receiving Entity"].dropna())
    )

    records = []
    for entity in entities:
        pending_gl = int(
            gl.loc[
                gl["Entity"].eq(entity) & gl["Status"].isin(["Pending Review", "Pending Approval"])
            ].shape[0]
        )
        manual_jes = int(gl.loc[gl["Entity"].eq(entity) & gl["Source"].eq("Manual Entry")].shape[0])

        entity_bank = bank.loc[bank["Entity"].eq(entity)]
        bank_exceptions = int(entity_bank.loc[~entity_bank["Match Status"].eq("Matched")].shape[0])
        bank_stale = int(entity_bank.loc[entity_bank["Days Outstanding"].fillna(0).gt(30)].shape[0])

        entity_accruals = accruals.loc[accruals["Entity"].eq(entity)]
        accrual_risks = int(
            entity_accruals["Status"].isin(
                ["Pending Review", "Pending Approval", "Needs Documentation"]
            ).sum()
        )
        accrual_missing_docs = int(
            entity_accruals["Supporting Doc"]
            .astype(str)
            .str.contains("Missing", case=False, na=False)
            .sum()
        )

        entity_ap = ap.loc[ap["Entity"].eq(entity)]
        ap_exceptions = int(entity_ap["has_3way_exception"].sum())
        ap_disputed = int(
            entity_ap["Approval Status"].isin(["Disputed", "On Hold", "Pending Approval"]).sum()
        )

        entity_ar = ar.loc[ar["Entity"].eq(entity)]
        ar_overdue = int(entity_ar["overdue"].sum())

        entity_tb = tb.loc[tb["Entity"].eq(entity)]
        tb_large_variances = int(entity_tb["variance_pct_num"].abs().gt(10).sum())
        tb_unreconciled = int(entity_tb.loc[~entity_tb["Reconciled"].eq("Yes")].shape[0])

        entity_fa = fa.loc[fa["Entity"].eq(entity)]
        fa_impairments = int(entity_fa["Impairment (EUR)"].fillna(0).gt(0).sum())
        fa_review_flags = int(
            entity_fa["Disposal Status"].isin(["Under Review", "Scheduled for Disposal"]).sum()
        )
        fa_unverified = int(
            entity_fa["Last Physical Verification"].astype(str).str.contains(
                "Not Yet Verified",
                case=False,
                na=False,
            ).sum()
        )

        ic_mask = (
            (~ic["Elimination Status"].eq("Eliminated"))
            & (ic["Sending Entity"].eq(entity) | ic["Receiving Entity"].eq(entity))
        )
        ic_exceptions = int(ic.loc[ic_mask].shape[0])
        ic_total_diff = float(ic.loc[ic_mask, "Difference (EUR)"].fillna(0).abs().sum())

        components = {
            "GL backlog": pending_gl * 0.45 + manual_jes * 0.30,
            "Bank reconciliation": bank_exceptions * 4.00 + bank_stale * 2.50,
            "Intercompany": ic_exceptions * 6.00 + min(ic_total_diff / 500, 8),
            "Accruals": accrual_risks * 3.50 + accrual_missing_docs * 3.00,
            "AP exceptions": ap_exceptions * 3.00 + ap_disputed * 1.25,
            "AR collections": ar_overdue * 2.50,
            "Trial balance": tb_large_variances * 1.75 + tb_unreconciled * 2.25,
            "Fixed assets": fa_impairments * 5.00 + fa_review_flags * 2.00 + fa_unverified * 3.00,
        }
        raw_score = float(sum(components.values()))
        primary_blocker = max(components, key=components.get)

        records.append(
            {
                "entity": entity,
                "pending_gl": pending_gl,
                "manual_jes": manual_jes,
                "bank_exceptions": bank_exceptions,
                "bank_stale": bank_stale,
                "accrual_risks": accrual_risks,
                "accrual_missing_docs": accrual_missing_docs,
                "ap_exceptions": ap_exceptions,
                "ap_disputed": ap_disputed,
                "ar_overdue": ar_overdue,
                "tb_large_variances": tb_large_variances,
                "tb_unreconciled": tb_unreconciled,
                "fa_impairments": fa_impairments,
                "fa_review_flags": fa_review_flags,
                "fa_unverified": fa_unverified,
                "ic_exceptions": ic_exceptions,
                "ic_total_diff": round(ic_total_diff, 2),
                "raw_risk_score": raw_score,
                "primary_blocker": primary_blocker,
            }
        )

    if not records:
        return []

    max_raw = max(record["raw_risk_score"] for record in records) or 1
    for record in records:
        normalized = 35 + (65 * record["raw_risk_score"] / max_raw)
        record["risk_score"] = _clamp(normalized)
        if record["risk_score"] >= 78:
            record["risk_level"] = "High"
        elif record["risk_score"] >= 60:
            record["risk_level"] = "Medium"
        else:
            record["risk_level"] = "Low"

        record["driver_summary"] = {
            "GL backlog": f"{record['pending_gl']} pending GL",
            "Bank reconciliation": f"{record['bank_exceptions']} bank exceptions",
            "Intercompany": f"{record['ic_exceptions']} IC exceptions",
            "Accruals": f"{record['accrual_risks']} accrual actions",
            "AP exceptions": f"{record['ap_exceptions']} AP 3-way issues",
            "AR collections": f"{record['ar_overdue']} overdue AR items",
            "Trial balance": f"{record['tb_large_variances']} large TB variances",
            "Fixed assets": (
                f"{record['fa_impairments'] + record['fa_review_flags'] + record['fa_unverified']} FA flags"
            ),
        }[record["primary_blocker"]]

    return sorted(records, key=lambda item: item["risk_score"], reverse=True)


def calculate_kpis(data):
    gl = data["gl"].copy()
    tb = data["tb"].copy()
    accruals = data["accruals"].copy()
    bank = data["bank"].copy()
    ic = data["ic"].copy()
    fa = data["fa"].copy()
    apar = data["apar"].copy()
    checklist = data["checklist"].copy()
    accounting_period = _derive_accounting_period(data)

    gl["Entry Amount (EUR)"] = gl[["Debit (EUR)", "Credit (EUR)"]].max(axis=1)
    tb["variance_pct_num"] = pd.to_numeric(
        tb["Variance %"].astype(str).str.rstrip("%"),
        errors="coerce",
    )

    ap = apar.loc[apar["Type"].eq("AP")].copy()
    ap["has_3way_exception"] = ap["3-Way Match"].notna() & ~ap["3-Way Match"].eq("Matched")

    ar = apar.loc[apar["Type"].eq("AR")].copy()
    ar["overdue"] = ar["Aging Bucket"].isin(["61-90 Days", "90+ Days"]) & ar["Outstanding (EUR)"].gt(0)

    pending_mask = gl["Status"].isin(["Pending Review", "Pending Approval"])
    pending_gl = int(pending_mask.sum())
    manual_jes = int(gl["Source"].eq("Manual Entry").sum())
    posted_jes = int(gl["Status"].eq("Posted").sum())
    gl_agent = _build_gl_agent(gl)

    bank_exception_mask = ~bank["Match Status"].eq("Matched")
    bank_exceptions = int(bank_exception_mask.sum())
    bank_matched = int(bank["Match Status"].eq("Matched").sum())
    bank_total = int(bank.shape[0])
    bank_stale = int(bank["Days Outstanding"].fillna(0).gt(30).sum())
    bank_match_rate = float(safe_pct(bank_matched, bank_total))
    bank_exception_breakdown = (
        bank.loc[bank_exception_mask, "Match Status"].value_counts().sort_values(ascending=False).to_dict()
    )
    ic_agent = _build_ic_agent(ic)
    bank_agent = _build_bank_agent(bank)
    journal_agent = _build_journal_agent(accruals)
    audit_agent = _build_audit_agent(bank_agent, ic_agent, journal_agent)
    checklist_agent = _build_checklist_agent(checklist)
    flux_agent = _build_flux_agent(tb, gl, accruals, apar)

    ic_exception_mask = ~ic["Elimination Status"].eq("Eliminated")
    ic_exceptions = int(ic_exception_mask.sum())
    ic_total_diff = float(ic.loc[ic_exception_mask, "Difference (EUR)"].fillna(0).abs().sum())

    accrual_action_mask = accruals["Status"].isin(
        ["Pending Review", "Pending Approval", "Needs Documentation"]
    )
    accrual_risks = int(accrual_action_mask.sum())
    accrual_missing_docs = int(
        accruals["Supporting Doc"].astype(str).str.contains("Missing", case=False, na=False).sum()
    )
    high_risk_accruals = int(accruals["Risk Flag"].eq("High").sum())

    ap_3way_exceptions = int(ap["has_3way_exception"].sum())
    ap_disputed = int(ap["Approval Status"].isin(["Disputed", "On Hold", "Pending Approval"]).sum())
    ar_overdue = int(ar["overdue"].sum())
    ar_open_balance = float(ar.loc[ar["Outstanding (EUR)"].gt(0), "Outstanding (EUR)"].sum())

    blocked_tasks = int(checklist["Status"].eq("Blocked").sum())
    waiting_tasks = int(checklist["Status"].eq("Waiting on Input").sum())
    not_started_tasks = int(checklist["Status"].eq("Not Started").sum())
    in_progress_tasks = int(checklist["Status"].eq("In Progress").sum())
    completed_tasks = int(checklist["Status"].eq("Completed").sum())
    critical_open_tasks = int(
        checklist.loc[
            checklist["Priority"].eq("Critical")
            & checklist["Status"].isin(["Blocked", "Waiting on Input", "Not Started", "In Progress"])
        ].shape[0]
    )
    checklist_status_counts = (
        checklist["Status"]
        .value_counts()
        .reindex(["Completed", "In Progress", "Not Started", "Waiting on Input", "Blocked"], fill_value=0)
        .to_dict()
    )

    unreconciled_tb = int((~tb["Reconciled"].eq("Yes")).sum())
    large_variances = int(tb["variance_pct_num"].abs().gt(10).sum())
    max_tb_variance_pct = float(tb["variance_pct_num"].abs().max())

    fa_impairments = int(fa["Impairment (EUR)"].fillna(0).gt(0).sum())
    fa_review_flags = int(fa["Disposal Status"].isin(["Under Review", "Scheduled for Disposal"]).sum())
    fa_unverified = int(
        fa["Last Physical Verification"].astype(str).str.contains("Not Yet Verified", case=False, na=False).sum()
    )

    areas = {
        "gl": {
            "label": AREA_LABELS["gl"],
            "sheet": "General Ledger",
            "severity_score": _clamp((pending_gl / 175) * 70 + (manual_jes / 70) * 30),
            "headline": f"{pending_gl} journals pending review or approval",
            "subheadline": f"{manual_jes} manual journal entries still need extra scrutiny",
            "metrics": {
                "Pending GL": pending_gl,
                "Manual JEs": manual_jes,
                "Posted JEs": posted_jes,
            },
            "insight": "Approval backlog is the single biggest operational bottleneck in the close.",
        },
        "bank": {
            "label": AREA_LABELS["bank"],
            "sheet": "Bank Reconciliation",
            "severity_score": _clamp((bank_exceptions / 12) * 70 + (bank_stale / 4) * 30),
            "headline": f"{bank_exceptions} cash exceptions still need investigation",
            "subheadline": f"{bank_match_rate}% match rate with {bank_stale} stale items over 30 days",
            "metrics": {
                "Exceptions": bank_exceptions,
                "Stale Items": bank_stale,
                "Match Rate %": bank_match_rate,
            },
            "insight": "The open cash exceptions are concentrated in Germany, the Netherlands, and the UK.",
        },
        "intercompany": {
            "label": AREA_LABELS["intercompany"],
            "sheet": "Intercompany",
            "severity_score": _clamp((ic_exceptions / 2) * 70 + min(ic_total_diff / 6000, 1) * 30),
            "headline": f"{ic_exceptions} intercompany elimination gaps remain open",
            "subheadline": f"{_format_currency(ic_total_diff)} unresolved across elimination exceptions",
            "metrics": {
                "IC Exceptions": ic_exceptions,
                "Unresolved Difference": round(ic_total_diff, 2),
            },
            "insight": "The open IC items are split between one variance investigation and one FX mismatch.",
        },
        "accruals": {
            "label": AREA_LABELS["accruals"],
            "sheet": "Accruals & Provisions",
            "severity_score": _clamp(
                (accrual_risks / 10) * 70 + (accrual_missing_docs / 3) * 20 + (high_risk_accruals / 5) * 10
            ),
            "headline": f"{accrual_risks} accruals require action before close",
            "subheadline": f"{accrual_missing_docs} are missing support and {high_risk_accruals} are high risk",
            "metrics": {
                "Action Items": accrual_risks,
                "Missing Support": accrual_missing_docs,
                "High Risk": high_risk_accruals,
            },
            "insight": "US accruals carry most of the open approval and documentation workload.",
        },
        "ap": {
            "label": AREA_LABELS["ap"],
            "sheet": "AP & AR Aging",
            "severity_score": _clamp((ap_3way_exceptions / 10) * 70 + (ap_disputed / 15) * 30),
            "headline": f"{ap_3way_exceptions} AP items are outside clean 3-way match",
            "subheadline": f"{ap_disputed} AP documents are disputed, on hold, or pending approval",
            "metrics": {
                "3-Way Exceptions": ap_3way_exceptions,
                "AP Approval Friction": ap_disputed,
            },
            "insight": "These are true AP exceptions only; AR rows are excluded from the 3-way match metric.",
        },
        "ar": {
            "label": AREA_LABELS["ar"],
            "sheet": "AP & AR Aging",
            "severity_score": _clamp((ar_overdue / 4) * 60 + min(ar_open_balance / 500000, 1) * 40),
            "headline": f"{ar_overdue} overdue AR items are beyond 60 days",
            "subheadline": f"{_format_currency(ar_open_balance)} remains open across AR documents",
            "metrics": {
                "Overdue AR Items": ar_overdue,
                "Open AR Balance": round(ar_open_balance, 2),
            },
            "insight": "Collections risk is secondary to AP exceptions, but it still affects working capital visibility.",
        },
        "checklist": {
            "label": AREA_LABELS["checklist"],
            "sheet": "Close Checklist",
            "severity_score": _clamp((blocked_tasks / 3) * 50 + (waiting_tasks / 3) * 20 + (critical_open_tasks / 7) * 30),
            "headline": f"{blocked_tasks} checklist tasks are blocked and {waiting_tasks} are waiting on input",
            "subheadline": f"{critical_open_tasks} critical tasks are still open across the close workflow",
            "metrics": {
                "Blocked": blocked_tasks,
                "Waiting": waiting_tasks,
                "Critical Open": critical_open_tasks,
            },
            "insight": "The biggest workflow pressure is in reporting, reconciliation, and consolidation steps.",
        },
        "tb": {
            "label": AREA_LABELS["tb"],
            "sheet": "Trial Balance",
            "severity_score": _clamp((large_variances / 40) * 60 + (unreconciled_tb / 20) * 40),
            "headline": f"{large_variances} trial balance lines moved more than 10%",
            "subheadline": f"{unreconciled_tb} trial balance accounts are not reconciled",
            "metrics": {
                "Large Variances": large_variances,
                "Unreconciled": unreconciled_tb,
                "Max Variance %": round(max_tb_variance_pct, 1),
            },
            "insight": "US inventory and DE deferred revenue are among the most visible balance movements.",
        },
        "fa": {
            "label": AREA_LABELS["fa"],
            "sheet": "Fixed Assets",
            "severity_score": _clamp((fa_impairments / 2) * 40 + (fa_review_flags / 6) * 35 + (fa_unverified / 2) * 25),
            "headline": f"{fa_impairments} assets carry impairment risk and {fa_review_flags} need review",
            "subheadline": f"{fa_unverified} assets still show no physical verification",
            "metrics": {
                "Impairments": fa_impairments,
                "Under Review / Disposal": fa_review_flags,
                "Not Verified": fa_unverified,
            },
            "insight": "Fixed assets are not the main delay driver, but they surface control gaps for the demo.",
        },
    }

    entities = _build_entity_rollups(data)
    riskiest_entity = entities[0] if entities else None

    details = {
        "pending_journals": _top_records(
            gl.loc[pending_mask],
            [
                "Journal Entry ID",
                "Entity",
                "Account Name",
                "Entry Amount (EUR)",
                "Source",
                "Status",
                "Posting Date",
            ],
            limit=10,
            sort_by="Entry Amount (EUR)",
        ),
        "bank_exceptions": _top_records(
            bank.loc[bank_exception_mask],
            [
                "Entity",
                "Description",
                "Difference (EUR)",
                "Match Status",
                "Days Outstanding",
            ],
            limit=10,
            sort_by="Days Outstanding",
        ),
        "ic_exceptions": _top_records(
            ic.loc[ic_exception_mask],
            [
                "Sending Entity",
                "Receiving Entity",
                "Difference (EUR)",
                "Elimination Status",
                "Notes",
            ],
            limit=8,
            sort_by="Difference (EUR)",
        ),
        "accrual_actions": _top_records(
            accruals.loc[accrual_action_mask],
            ["Accrual ID", "Entity", "Description", "Accrual Amount (EUR)", "Status", "Risk Flag"],
            limit=10,
            sort_by="Accrual Amount (EUR)",
        ),
        "ap_exceptions": _top_records(
            ap.loc[ap["has_3way_exception"]],
            ["Document ID", "Entity", "Outstanding (EUR)", "3-Way Match", "Approval Status"],
            limit=10,
            sort_by="Outstanding (EUR)",
        ),
        "ar_watchlist": _top_records(
            ar.loc[ar["overdue"]],
            ["Document ID", "Entity", "Outstanding (EUR)", "Aging Bucket", "Approval Status"],
            limit=10,
            sort_by="Outstanding (EUR)",
        ),
        "tb_variances": _top_records(
            tb.assign(abs_variance_pct=tb["variance_pct_num"].abs()),
            ["Entity", "Account Name", "Variance (EUR)", "Variance %", "Reconciled"],
            limit=10,
            sort_by="abs_variance_pct",
        ),
        "tb_unreconciled": _top_records(
            tb.loc[~tb["Reconciled"].eq("Yes")],
            ["Entity", "Account Name", "Closing Balance (EUR)", "Variance %", "Reviewer Notes"],
            limit=10,
            sort_by="Closing Balance (EUR)",
        ),
        "fa_risks": _top_records(
            fa.loc[
                fa["Impairment (EUR)"].fillna(0).gt(0)
                | fa["Disposal Status"].isin(["Under Review", "Scheduled for Disposal"])
                | fa["Last Physical Verification"].astype(str).str.contains("Not Yet Verified", case=False, na=False)
            ],
            [
                "Entity",
                "Asset Description",
                "Impairment (EUR)",
                "Disposal Status",
                "Last Physical Verification",
            ],
            limit=10,
            sort_by="Impairment (EUR)",
        ),
        "checklist_bottlenecks": _top_records(
            checklist.loc[checklist["Status"].isin(["Blocked", "Waiting on Input", "Not Started"])],
            ["Category", "Task", "Status", "Priority", "Estimated Hours"],
            limit=10,
            sort_by="Estimated Hours",
        ),
    }

    summary = {
        "accounting_period": accounting_period["raw"],
        "accounting_period_label": accounting_period["label"],
        "pending_gl": pending_gl,
        "manual_jes": manual_jes,
        "posted_jes": posted_jes,
        "gl_straight_through_candidates": gl_agent["summary"]["straight_through_candidates"],
        "gl_manager_queue": gl_agent["summary"]["manager_queue"],
        "gl_controller_queue": gl_agent["summary"]["controller_queue"],
        "gl_cfo_queue": gl_agent["summary"]["cfo_queue"],
        "gl_approval_recoverable_hours": gl_agent["summary"]["recoverable_hours"],
        "bank_exceptions": bank_exceptions,
        "bank_total": bank_total,
        "bank_matched": bank_matched,
        "bank_stale": bank_stale,
        "bank_match_rate": bank_match_rate,
        "bank_auto_clear_candidates": bank_agent["summary"]["auto_clear_candidates"],
        "bank_journal_candidates": bank_agent["summary"]["journal_candidates"],
        "bank_escalations": bank_agent["summary"]["escalations"],
        "bank_manual_investigations": bank_agent["summary"]["manual_investigations"],
        "bank_statement_break_value": round(bank_agent["summary"]["statement_break_value"], 2),
        "ic_auto_matched": ic_agent["summary"]["auto_matched"],
        "ic_fx_mismatches": ic_agent["summary"]["fx_mismatches"],
        "ic_elimination_drafts": ic_agent["summary"]["elimination_drafts"],
        "ic_tp_flags": ic_agent["summary"]["tp_flags"],
        "journal_agent_ready_drafts": journal_agent["summary"]["erp_ready_drafts"],
        "journal_agent_review_needed": journal_agent["summary"]["review_needed"],
        "audit_drafts_reviewed": audit_agent["summary"]["drafts_reviewed"],
        "audit_ready_to_post": audit_agent["summary"]["ready_to_post"],
        "audit_conditional": audit_agent["summary"]["conditional_approval"],
        "audit_blocked": audit_agent["summary"]["blocked_for_review"],
        "audit_avg_control_score": audit_agent["summary"]["average_control_score"],
        "flux_mom_anomalies": flux_agent["summary"]["mom_anomalies"],
        "flux_yoy_anomalies": flux_agent["summary"]["yoy_anomalies"],
        "flux_unreconciled": flux_agent["summary"]["unreconciled"],
        "flux_top_variance_amount": flux_agent["summary"]["top_variance_amount"],
        "checklist_handoffs_at_risk": checklist_agent["summary"]["handoffs_at_risk"],
        "checklist_recoverable_hours": checklist_agent["summary"]["recoverable_hours"],
        "checklist_automation_candidates": checklist_agent["summary"]["automation_candidates"],
        "ic_exceptions": ic_exceptions,
        "ic_total_diff": round(ic_total_diff, 2),
        "accrual_risks": accrual_risks,
        "accrual_missing_docs": accrual_missing_docs,
        "high_risk_accruals": high_risk_accruals,
        "ap_3way_exceptions": ap_3way_exceptions,
        "ap_disputed": ap_disputed,
        "ar_overdue": ar_overdue,
        "ar_open_balance": round(ar_open_balance, 2),
        "blocked_tasks": blocked_tasks,
        "waiting_tasks": waiting_tasks,
        "not_started_tasks": not_started_tasks,
        "in_progress_tasks": in_progress_tasks,
        "completed_tasks": completed_tasks,
        "critical_open_tasks": critical_open_tasks,
        "unreconciled_tb": unreconciled_tb,
        "large_variances": large_variances,
        "max_tb_variance_pct": round(max_tb_variance_pct, 1),
        "fa_impairments": fa_impairments,
        "fa_review_flags": fa_review_flags,
        "fa_unverified": fa_unverified,
        "bank_exception_breakdown": bank_exception_breakdown,
        "checklist_status_counts": checklist_status_counts,
        "entity_count": len(entities),
        "riskiest_entity": riskiest_entity["entity"] if riskiest_entity else None,
        "riskiest_entity_blocker": riskiest_entity["driver_summary"] if riskiest_entity else None,
    }

    kpis = {
        "summary": summary,
        "areas": areas,
        "entities": entities,
        "details": details,
        "agents": {
            "gl": gl_agent,
            "bank": bank_agent,
            "ic": ic_agent,
            "journal": journal_agent,
            "audit": audit_agent,
            "checklist": checklist_agent,
            "flux": flux_agent,
        },
    }

    erp_posting = build_erp_posting_simulator(kpis)
    kpis["agents"]["erp_posting"] = erp_posting
    kpis["summary"].update(
        {
            "erp_auto_post": erp_posting["summary"]["auto_post"],
            "erp_ready_to_post": erp_posting["summary"]["ready_to_post"],
            "erp_manual_hold": erp_posting["summary"]["manual_hold"],
            "erp_auto_post_eligible": erp_posting["summary"]["eligible_for_auto_post"],
        }
    )

    return kpis


def _continuous_close_days_from_penalties(penalties):
    penalty_points = sum(item["points"] for item in penalties)
    days = 3.2 + (penalty_points / 60.0)
    return round(max(3.2, min(5.9, days)), 1)


def calculate_readiness_score(kpis):
    summary = kpis["summary"]
    score = 100
    penalties = []

    def add_penalty(reason, points, area):
        nonlocal score
        score -= points
        penalties.append({"reason": reason, "points": points, "area": area})

    if summary["pending_gl"] > 150:
        add_penalty("High GL approval backlog", 15, "GL & Journals")
    elif summary["pending_gl"] > 100:
        add_penalty("Elevated GL approval backlog", 10, "GL & Journals")

    if summary["manual_jes"] > 50:
        add_penalty("Heavy manual journal volume", 8, "GL & Journals")
    elif summary["manual_jes"] > 30:
        add_penalty("Moderate manual journal volume", 5, "GL & Journals")

    if summary["bank_exceptions"] >= 10:
        add_penalty("Bank reconciliation exceptions", 8, "Bank Reconciliation")
    elif summary["bank_exceptions"] >= 5:
        add_penalty("Some bank reconciliation exceptions", 5, "Bank Reconciliation")

    if summary["bank_stale"] > 0:
        add_penalty("Stale bank reconciling items", 4, "Bank Reconciliation")

    if summary["ic_exceptions"] >= 2:
        add_penalty("Intercompany elimination exceptions", 8, "Intercompany")
    elif summary["ic_exceptions"] == 1:
        add_penalty("One intercompany exception", 4, "Intercompany")

    if summary["accrual_risks"] > 5:
        add_penalty("Accrual backlog and approval risk", 6, "Accruals")

    if summary["accrual_missing_docs"] > 0:
        add_penalty("Accrual documentation gaps", 4, "Accruals")

    if summary["ap_3way_exceptions"] > 5:
        add_penalty("AP matching exceptions", 5, "AP Exceptions")

    if summary["blocked_tasks"] >= 3:
        add_penalty("Blocked checklist tasks", 8, "Close Checklist")
    elif summary["blocked_tasks"] > 0:
        add_penalty("Some blocked checklist tasks", 4, "Close Checklist")

    if summary["critical_open_tasks"] > 5:
        add_penalty("Too many critical tasks remain open", 6, "Close Checklist")

    if summary["large_variances"] > 25:
        add_penalty("Large trial balance variances need explanation", 4, "Trial Balance")

    if summary["fa_impairments"] > 0 or summary["fa_unverified"] > 0:
        add_penalty("Fixed asset control gaps need review", 2, "Fixed Assets")

    score = max(0, min(100, score))
    score = int(round(55 + (score * 0.45)))
    continuous_close_days = _continuous_close_days_from_penalties(penalties)

    if score >= 80:
        risk_level = "Low"
        predicted_days = 3.7
    elif score >= 65:
        risk_level = "Medium"
        predicted_days = 4.4
    else:
        risk_level = "High"
        predicted_days = 5.6

    gap_to_target_days = round(max(0, predicted_days - 4.0), 1)
    continuous_gap_to_target_days = round(max(0, continuous_close_days - 4.0), 1)

    return {
        "readiness_score": score,
        "risk_level": risk_level,
        "predicted_close_days": predicted_days,
        "continuous_close_days": continuous_close_days,
        "gap_to_target_days": gap_to_target_days,
        "continuous_gap_to_target_days": continuous_gap_to_target_days,
        "penalties": penalties,
    }


def _source_action_for_posting(source_agent):
    return {
        "Bank Agent": "bank_auto_clear",
        "IC Agent": "ic_post_drafts",
        "Journal Agent": "journal_post_ready",
    }.get(source_agent)


def _source_action_label(action_id):
    return {
        "gl_straight_through": "GL straight-through approval",
        "gl_fast_track_controller": "GL batched approvals",
        "bank_auto_clear": "Bank clear/post action",
        "ic_post_drafts": "IC elimination posting",
        "journal_post_ready": "Journal auto-post",
    }.get(action_id, "Manual release")


def build_erp_posting_simulator(kpis, selected_action_ids=None):
    selected_ids = set(selected_action_ids or [])
    gl_agent = kpis["agents"]["gl"]
    audit_agent = kpis["agents"]["audit"]
    rows = []

    gl_worklist = gl_agent["worklist"].copy()
    controller_slots = 20 if "gl_fast_track_controller" in selected_ids else 0

    if not gl_worklist.empty:
        for _, row in gl_worklist.iterrows():
            lane = row["Approval Lane"]
            outcome = "Manual Hold"
            eligible = "No"
            trigger_action = ""
            hold_reason = ""
            next_step = ""

            if lane == "Straight-through":
                eligible = "Yes"
                trigger_action = _source_action_label("gl_straight_through")
                if "gl_straight_through" in selected_ids:
                    outcome = "Auto-Post"
                    next_step = "Release the straight-through approval batch and post directly to ERP."
                else:
                    outcome = "Ready to Post"
                    next_step = "Approve the straight-through batch to release ERP posting."
            elif lane == "Manager queue":
                trigger_action = _source_action_label("gl_fast_track_controller")
                if "gl_fast_track_controller" in selected_ids:
                    outcome = "Ready to Post"
                    next_step = "Manager packet is batched and can be released to ERP after sign-off."
                else:
                    hold_reason = "Pending manager approval packet"
                    next_step = "Route into the manager approval packet."
            elif lane == "Controller queue":
                trigger_action = _source_action_label("gl_fast_track_controller")
                if controller_slots > 0:
                    outcome = "Ready to Post"
                    controller_slots -= 1
                    next_step = "Controller packet is fast-tracked and can be released after approval."
                elif "gl_fast_track_controller" in selected_ids:
                    hold_reason = "Controller queue remains above the fast-track batch limit"
                    next_step = "Keep in the next controller approval packet."
                else:
                    hold_reason = "Pending controller approval"
                    next_step = "Escalate through the controller approval queue."
            else:
                hold_reason = "Executive approval still required"
                next_step = "Hold for CFO sign-off before ERP posting."

            rows.append(
                {
                    "Posting Item ID": row["Journal Entry ID"],
                    "Source Agent": "GL Approval Agent",
                    "Entity": row["Entity"],
                    "Item Type": "GL Journal",
                    "Amount (EUR)": float(pd.to_numeric(row["Entry Amount (EUR)"], errors="coerce") or 0),
                    "Current Gate": lane,
                    "Posting Outcome": outcome,
                    "Auto-Post Eligible": eligible,
                    "Trigger Action": trigger_action or "Manual release",
                    "Required Approver": row["Required Approver"],
                    "Manual Hold Reason": hold_reason or "None",
                    "Next ERP Step": next_step,
                }
            )

    audit_pack = audit_agent["review_pack"].copy()
    if not audit_pack.empty:
        for _, row in audit_pack.iterrows():
            source_agent = row["Source Agent"]
            source_action = _source_action_for_posting(source_agent)
            recommendation = row["Posting Recommendation"]
            outcome = "Manual Hold"
            eligible = "No"
            hold_reason = ""
            next_step = ""

            if recommendation == "Ready to Post":
                eligible = "Yes"
                if source_action in selected_ids:
                    outcome = "Auto-Post"
                    next_step = "Control-cleared draft is released directly into ERP posting."
                else:
                    outcome = "Ready to Post"
                    next_step = "Draft is control-cleared and waiting for posting release."
            elif recommendation == "Conditional Approval":
                hold_reason = row["Primary Control Gap"]
                next_step = f"Obtain {row['Required Approver']} approval and clear the control gap."
            else:
                hold_reason = row["Primary Control Gap"]
                next_step = "Resolve blocking control issues before any ERP posting attempt."

            rows.append(
                {
                    "Posting Item ID": row["Journal ID"],
                    "Source Agent": source_agent,
                    "Entity": row["Entity"],
                    "Item Type": row["JE Type"],
                    "Amount (EUR)": float(pd.to_numeric(row["Amount (EUR)"], errors="coerce") or 0),
                    "Current Gate": recommendation,
                    "Posting Outcome": outcome,
                    "Auto-Post Eligible": eligible,
                    "Trigger Action": _source_action_label(source_action) if source_action else "Manual release",
                    "Required Approver": row["Required Approver"],
                    "Manual Hold Reason": hold_reason or "None",
                    "Next ERP Step": next_step,
                }
            )

    worklist = pd.DataFrame(rows)
    if worklist.empty:
        empty_status = pd.DataFrame(columns=["Posting Outcome", "Count"])
        empty_sources = pd.DataFrame(columns=["Source Agent", "Posting Outcome", "Count", "Amount (EUR)"])
        return {
            "summary": {
                "total_items": 0,
                "auto_post": 0,
                "ready_to_post": 0,
                "manual_hold": 0,
                "eligible_for_auto_post": 0,
                "auto_post_amount": 0.0,
                "ready_amount": 0.0,
                "hold_amount": 0.0,
            },
            "worklist": worklist,
            "auto_post_queue": worklist,
            "ready_queue": worklist,
            "manual_hold_queue": worklist,
            "status_breakdown": empty_status,
            "source_breakdown": empty_sources,
        }

    outcome_rank = {"Auto-Post": 0, "Ready to Post": 1, "Manual Hold": 2}
    worklist["Outcome Rank"] = worklist["Posting Outcome"].map(outcome_rank).fillna(9)
    worklist = worklist.sort_values(
        ["Outcome Rank", "Amount (EUR)"],
        ascending=[True, False],
    ).drop(columns=["Outcome Rank"])

    auto_post_queue = worklist.loc[worklist["Posting Outcome"].eq("Auto-Post")].copy()
    ready_queue = worklist.loc[worklist["Posting Outcome"].eq("Ready to Post")].copy()
    manual_hold_queue = worklist.loc[worklist["Posting Outcome"].eq("Manual Hold")].copy()

    status_breakdown = (
        worklist["Posting Outcome"]
        .value_counts()
        .reindex(["Auto-Post", "Ready to Post", "Manual Hold"], fill_value=0)
        .rename_axis("Posting Outcome")
        .reset_index(name="Count")
    )
    source_breakdown = (
        worklist.groupby(["Source Agent", "Posting Outcome"])
        .agg(
            Count=("Posting Item ID", "count"),
            **{"Amount (EUR)": ("Amount (EUR)", "sum")},
        )
        .reset_index()
    )
    source_breakdown["Amount (EUR)"] = source_breakdown["Amount (EUR)"].round(2)

    summary = {
        "total_items": int(worklist.shape[0]),
        "auto_post": int(auto_post_queue.shape[0]),
        "ready_to_post": int(ready_queue.shape[0]),
        "manual_hold": int(manual_hold_queue.shape[0]),
        "eligible_for_auto_post": int(worklist["Auto-Post Eligible"].eq("Yes").sum()),
        "auto_post_amount": round(float(auto_post_queue["Amount (EUR)"].sum()), 2),
        "ready_amount": round(float(ready_queue["Amount (EUR)"].sum()), 2),
        "hold_amount": round(float(manual_hold_queue["Amount (EUR)"].sum()), 2),
    }

    return {
        "summary": summary,
        "worklist": worklist,
        "auto_post_queue": auto_post_queue,
        "ready_queue": ready_queue,
        "manual_hold_queue": manual_hold_queue,
        "status_breakdown": status_breakdown,
        "source_breakdown": source_breakdown,
    }


def priority_engine(kpis):
    summary = kpis["summary"]
    areas = kpis["areas"]
    priorities = []

    def add(area_key, item, why, impact, hours_saved, priority_score, downstream_unlock):
        priorities.append(
            {
                "area_key": area_key,
                "area": AREA_LABELS[area_key],
                "priority_item": item,
                "why_it_matters": why,
                "impact": impact,
                "hours_saved_est": hours_saved,
                "priority_score": priority_score,
                "downstream_unlock": downstream_unlock,
            }
        )

    if summary["pending_gl"] > 0:
        add(
            "gl",
            "Clear pending GL reviews and approvals",
            f"{summary['pending_gl']} journals still require review or approval, including {summary['manual_jes']} manual entries.",
            "Unlocks account sign-off and cuts late close surprises.",
            min(14, 6 + summary["pending_gl"] // 30),
            10,
            "Final account review and downstream controller sign-off",
        )
    if summary["bank_exceptions"] > 0:
        add(
            "bank",
            "Resolve bank reconciliation exceptions",
            f"{summary['bank_exceptions']} cash exceptions remain open and {summary['bank_stale']} are stale.",
            "Completes cash reconciliation and removes manual investigation time.",
            min(8, 4 + summary["bank_exceptions"] // 3),
            9,
            "Cash certification and treasury close readiness",
        )
    if summary["ic_exceptions"] > 0:
        add(
            "intercompany",
            "Fix intercompany elimination gaps",
            f"{summary['ic_exceptions']} IC items still show {_format_currency(summary['ic_total_diff'])} unresolved.",
            "Prevents consolidation delay and elimination rework.",
            5,
            8,
            "Group consolidation and FX elimination posting",
        )
    if summary["blocked_tasks"] > 0 or summary["waiting_tasks"] > 0:
        add(
            "checklist",
            "Unblock critical close checklist tasks",
            f"{summary['blocked_tasks']} tasks are blocked, {summary['waiting_tasks']} are waiting on input, and {summary['critical_open_tasks']} critical items remain open.",
            "Removes dependency bottlenecks across reporting and reconciliation.",
            summary.get("checklist_recoverable_hours", 7),
            8,
            "Critical checklist handoffs across reconciliation, journals, and reporting",
        )
    if summary["accrual_risks"] > 0:
        add(
            "accruals",
            "Review high-risk accruals and missing support",
            f"{summary['accrual_risks']} accruals need action and {summary['accrual_missing_docs']} are missing support.",
            "Improves expense accuracy and lowers audit friction.",
            5,
            7,
            "Expense completeness and period-end journal quality",
        )
    if summary["ap_3way_exceptions"] > 0:
        add(
            "ap",
            "Clear AP 3-way match exceptions",
            f"{summary['ap_3way_exceptions']} AP documents are outside a clean 3-way match.",
            "Improves subledger confidence and reduces vendor reconciliation noise.",
            4,
            6,
            "AP subledger certification and invoice accuracy",
        )
    if summary["large_variances"] > 20:
        add(
            "tb",
            "Explain the largest trial balance variances",
            f"{summary['large_variances']} trial balance lines moved more than 10% and {summary['unreconciled_tb']} remain unreconciled.",
            "Sharpens controller commentary and surfaces hidden balance risks earlier.",
            3,
            5,
            "Management reporting narrative and controller review",
        )
    if summary["fa_impairments"] > 0 or summary["fa_unverified"] > 0:
        add(
            "fa",
            "Validate fixed-asset review items",
            f"{summary['fa_impairments']} impairments, {summary['fa_review_flags']} review/disposal flags, and {summary['fa_unverified']} unverified assets need checking.",
            "Reduces close control gaps in asset accounting.",
            2,
            4,
            "Asset subledger integrity and audit readiness",
        )

    priorities = sorted(
        priorities,
        key=lambda item: (
            item["priority_score"],
            item["hours_saved_est"],
            areas[item["area_key"]]["severity_score"],
        ),
        reverse=True,
    )
    return priorities


def build_scenario_actions(kpis):
    summary = kpis["summary"]
    agents = kpis["agents"]
    return [
        {
            "id": "gl_straight_through",
            "area": "GL Approval Agent",
            "title": f"Auto-approve {agents['gl']['summary']['straight_through_candidates']} straight-through journals",
            "description": "System-generated low-risk journals move from pending review into the approval batch automatically.",
            "hours_saved": agents["gl"]["summary"]["recoverable_hours"],
        },
        {
            "id": "gl_fast_track_controller",
            "area": "GL Approval Agent",
            "title": "Fast-track manager and controller approval packets",
            "description": "Route the non-CFO queue through batched approval packets and remove manual approval chasing.",
            "hours_saved": 8,
        },
        {
            "id": "bank_auto_clear",
            "area": "Bank Agent",
            "title": f"Clear {summary['bank_auto_clear_candidates']} timing items and post {summary['bank_journal_candidates']} bank drafts",
            "description": "Auto-clear recurring timing items and move bank journal candidates into ERP posting.",
            "hours_saved": 5,
        },
        {
            "id": "ic_post_drafts",
            "area": "IC Agent",
            "title": f"Post {summary['ic_elimination_drafts']} IC elimination drafts and clear matched pairs",
            "description": "Use the IC agent to close FX mismatches and remove elimination blockers from consolidation.",
            "hours_saved": 5,
        },
        {
            "id": "journal_post_ready",
            "area": "Journal Agent",
            "title": f"Post {summary['journal_agent_ready_drafts']} ERP-ready journal drafts",
            "description": "Move audit-cleared accrual and reclassification drafts into ERP posting and leave only review exceptions open.",
            "hours_saved": 5,
        },
        {
            "id": "checklist_unblock",
            "area": "Checklist Agent",
            "title": f"Auto-release blocked checklist handoffs ({summary['checklist_recoverable_hours']} hrs recoverable)",
            "description": "Auto-escalate blocked tasks, release completed handoffs, and notify owners on the critical path.",
            "hours_saved": summary["checklist_recoverable_hours"],
        },
        {
            "id": "ap_exception_sweep",
            "area": "AP Exceptions",
            "title": "Sweep low-risk AP matching exceptions",
            "description": "Resolve the easier 3-way exceptions in one batch to reduce subledger noise.",
            "hours_saved": 3,
        },
    ]


def simulate_close_scenario(kpis, selected_action_ids):
    scenario = deepcopy(kpis)
    summary = scenario["summary"]
    base_summary = kpis["summary"]
    selected_ids = set(selected_action_ids)

    if "gl_straight_through" in selected_ids:
        summary["pending_gl"] = max(0, summary["pending_gl"] - base_summary["gl_straight_through_candidates"])

    if "gl_fast_track_controller" in selected_ids:
        reduction = base_summary["gl_manager_queue"] + min(base_summary["gl_controller_queue"], 20)
        summary["pending_gl"] = max(0, summary["pending_gl"] - reduction)

    if "bank_auto_clear" in selected_ids:
        reduction = base_summary["bank_auto_clear_candidates"] + base_summary["bank_journal_candidates"]
        summary["bank_exceptions"] = max(0, summary["bank_exceptions"] - reduction)
        summary["bank_stale"] = max(0, summary["bank_stale"] - 2)

    if "ic_post_drafts" in selected_ids:
        summary["ic_exceptions"] = max(0, summary["ic_exceptions"] - base_summary["ic_elimination_drafts"])
        if summary["ic_exceptions"] == 0:
            summary["ic_total_diff"] = 0.0

    if "journal_post_ready" in selected_ids:
        remaining_review = max(1, base_summary["journal_agent_review_needed"])
        summary["accrual_risks"] = min(summary["accrual_risks"], remaining_review)
        summary["accrual_missing_docs"] = min(summary["accrual_missing_docs"], remaining_review)

    if "checklist_unblock" in selected_ids:
        summary["blocked_tasks"] = 0
        summary["waiting_tasks"] = 1
        summary["critical_open_tasks"] = max(3, summary["critical_open_tasks"] - 4)

    if "ap_exception_sweep" in selected_ids:
        summary["ap_3way_exceptions"] = max(0, summary["ap_3way_exceptions"] - 4)

    score = calculate_readiness_score(scenario)
    selected_actions = [action for action in build_scenario_actions(kpis) if action["id"] in selected_ids]
    total_hours_saved = sum(action["hours_saved"] for action in selected_actions)
    posting_simulator = build_erp_posting_simulator(kpis, selected_ids)

    return {
        "selected_actions": selected_actions,
        "score": score,
        "total_hours_saved": total_hours_saved,
        "posting_simulator": posting_simulator,
        "gets_below_four": score["continuous_close_days"] < 4.0,
    }


def generate_commentary(kpis, score):
    summary = kpis["summary"]
    entity = summary["riskiest_entity"]
    entity_blocker = summary["riskiest_entity_blocker"]

    return (
        f"NovaTech is currently at {score['readiness_score']}/100 close readiness, implying a "
        f"{score['risk_level'].lower()} risk of missing the 4-day target and an estimated {score['predicted_close_days']} day close. "
        f"The largest blockers are {summary['pending_gl']} pending GL approvals, {summary['bank_exceptions']} bank exceptions, "
        f"{summary['ic_exceptions']} intercompany gaps, and {summary['blocked_tasks']} blocked checklist tasks. "
        f"{entity} is the riskiest entity right now, driven by {entity_blocker}. "
        f"The fastest path back toward target is to clear GL backlog first, then close bank and intercompany exceptions, and unblock critical checklist steps."
    )


def _route_copilot_intent(prompt):
    text = (prompt or "").strip().lower()

    if not text:
        return "general"

    if any(keyword in text for keyword in ["gl agent", "gl approval", "journal approvals", "approval agent"]):
        return "gl_agent"
    if any(
        keyword in text
        for keyword in ["erp posting", "auto-post", "ready to post", "manual hold", "posting simulator"]
    ):
        return "erp_posting"
    if any(keyword in text for keyword in ["bank agent", "bank reconciliation agent", "bank rec agent", "cash agent"]):
        return "bank_agent"
    if any(
        keyword in text
        for keyword in [
            "ic agent",
            "intercompany agent",
            "intercompany reconciliation",
            "intercompany",
            "fx mismatch",
            "transfer pricing",
        ]
    ):
        return "ic_agent"
    if any(keyword in text for keyword in ["journal agent", "je agent", "journal entry agent", "what should the journal agent post"]):
        return "journal_agent"
    if any(
        keyword in text
        for keyword in [
            "checklist agent",
            "close checklist",
            "checklist tasks",
            "unblocked",
            "unblock",
            "unblock critical tasks",
            "waiting on input",
            "handoff",
        ]
    ):
        return "checklist_agent"
    if any(keyword in text for keyword in ["controller", "today", "fix first", "do today", "priority today"]):
        return "controller_actions"
    if any(
        keyword in text
        for keyword in [
            "audit",
            "compliance",
            "controls",
            "before posting",
            "what does audit need to review",
            "control completeness",
        ]
    ):
        return "audit_agent"
    if any(keyword in text for keyword in ["riskiest entity", "which entity", "entity risk", "entity is riskiest"]):
        return "riskiest_entity"
    if any(keyword in text for keyword in ["variance", "trial balance", "tb", "flux", "mom", "yoy"]):
        return "flux_agent"
    if any(keyword in text for keyword in ["automate", "automation", "auto", "automated"]):
        return "automation_opportunities"
    if any(keyword in text for keyword in ["below 4", "under 4", "4 days", "scenario", "simulator"]):
        return "below_four"
    if any(keyword in text for keyword in ["cfo", "executive", "summary", "board"]):
        return "cfo_summary"
    if any(keyword in text for keyword in ["delay", "blocker", "risk", "close"]):
        return "close_blockers"
    return "general"


def generate_copilot_response(prompt, kpis, score, priorities):
    summary = kpis["summary"]
    entities = kpis["entities"]
    details = kpis["details"]
    gl_agent = kpis["agents"]["gl"]
    bank_agent = kpis["agents"]["bank"]
    ic_agent = kpis["agents"]["ic"]
    journal_agent = kpis["agents"]["journal"]
    audit_agent = kpis["agents"]["audit"]
    checklist_agent = kpis["agents"]["checklist"]
    flux_agent = kpis["agents"]["flux"]
    erp_posting = kpis["agents"]["erp_posting"]
    scenario_actions = build_scenario_actions(kpis)
    top_priorities = priorities[:3]
    top_entity = entities[0] if entities else None
    intent = _route_copilot_intent(prompt)

    follow_ups = [
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

    if intent == "close_blockers":
        answer = (
            f"The close is most exposed by {summary['pending_gl']} pending GL approvals, "
            f"{summary['bank_exceptions']} unresolved bank exceptions, {summary['ic_exceptions']} intercompany gaps, "
            f"and {summary['blocked_tasks']} blocked checklist tasks. At the current pace NovaTech is tracking to "
            f"{score['predicted_close_days']} days, which is {score['gap_to_target_days']} days above the CFO target. "
            f"If the team clears GL backlog first and then resolves bank and intercompany exceptions, the close path improves fastest."
        )
        source_metrics = [
            f"{summary['pending_gl']} pending GL",
            f"{summary['bank_exceptions']} bank exceptions",
            f"{summary['ic_exceptions']} IC exceptions",
            f"{summary['blocked_tasks']} blocked tasks",
        ]
    elif intent == "controller_actions":
        action_lines = []
        for index, priority in enumerate(top_priorities, start=1):
            action_lines.append(
                f"{index}. {priority['priority_item']} because it unlocks {priority['downstream_unlock'].lower()}."
            )
        answer = (
            "The controller should focus on the three actions that unblock the most downstream work today:\n\n"
            + "\n".join(action_lines)
        )
        source_metrics = [
            f"{top_priorities[0]['hours_saved_est']} hrs recoverable",
            f"{summary['critical_open_tasks']} critical tasks open",
            f"{score['predicted_close_days']} day projected close",
        ]
    elif intent == "automation_opportunities":
        answer = (
            "The best automation targets are journal approval routing, bank matching for recurring exceptions, "
            "recurring accrual preparation, and variance commentary generation. Those four areas combine high manual effort "
            "with repeatable patterns already visible in the dataset."
        )
        source_metrics = [
            f"{summary['manual_jes']} manual JEs",
            f"{summary['bank_exceptions']} bank exceptions",
            f"{summary['accrual_risks']} accrual actions",
            f"{summary['large_variances']} TB variances",
        ]
    elif intent == "cfo_summary":
        answer = (
            f"NovaTech is currently at {score['readiness_score']}/100 close readiness and is projected to close in "
            f"{score['predicted_close_days']} days, missing the 4-day target by {score['gap_to_target_days']} days. "
            f"The main risk is operational, not transactional: GL approvals, cash exceptions, IC mismatches, and blocked checklist steps "
            f"are absorbing finance capacity that should be used to finish the close."
        )
        source_metrics = [
            f"{score['readiness_score']}/100 readiness",
            f"{score['predicted_close_days']} day close forecast",
            f"{summary['riskiest_entity']} highest entity risk",
        ]
    elif intent == "riskiest_entity" and top_entity:
        answer = (
            f"{top_entity['entity']} is the riskiest entity in the current close profile. Its primary blocker is "
            f"{top_entity['primary_blocker'].lower()}, driven by {top_entity['driver_summary'].lower()}. "
            f"It also carries {top_entity['pending_gl']} pending GL items, {top_entity['ap_exceptions']} AP exceptions, "
            f"and {top_entity['tb_large_variances']} large trial balance movements."
        )
        source_metrics = [
            top_entity["entity"],
            f"Risk {top_entity['risk_score']}/100",
            top_entity["driver_summary"],
        ]
    elif intent == "flux_agent":
        anomaly_rows = flux_agent["anomalies"].head(3)
        anomaly_lines = [
            f"{row['Entity']}: {row['Account Name']} at {row['Variance (EUR)']:.0f} EUR ({row['MoM %']:.1f}% MoM, {row['YoY %']:.1f}% YoY)"
            for _, row in anomaly_rows.iterrows()
        ]
        answer = (
            f"The flux agent found {flux_agent['summary']['mom_anomalies']} MoM and "
            f"{flux_agent['summary']['yoy_anomalies']} YoY anomalies. "
            f"Top movements are:\n\n"
            + "\n".join(f"- {line}" for line in anomaly_lines)
        )
        source_metrics = [
            f"{flux_agent['summary']['mom_anomalies']} MoM anomalies",
            f"{flux_agent['summary']['yoy_anomalies']} YoY anomalies",
            f"Top variance EUR {flux_agent['summary']['top_variance_amount']:.0f}",
        ]
    elif intent == "gl_agent":
        worklist = gl_agent["worklist"].head(3)
        action_lines = [
            f"- {row['Journal Entry ID']}: {row['Approval Lane']} | {row['Required Approver']} | EUR {row['Entry Amount (EUR)']:,.2f}"
            for _, row in worklist.iterrows()
        ]
        answer = (
            f"The GL approval agent is triaging {gl_agent['summary']['pending_items']} pending journals. "
            f"It sees {gl_agent['summary']['straight_through_candidates']} straight-through candidates, "
            f"{gl_agent['summary']['manager_queue']} manager approvals, "
            f"{gl_agent['summary']['controller_queue']} controller items, and "
            f"{gl_agent['summary']['cfo_queue']} CFO items. "
            f"The fastest win is auto-approving the straight-through queue and batching the manager/controller packets.\n\n"
            + "\n".join(action_lines)
        )
        source_metrics = [
            f"{gl_agent['summary']['straight_through_candidates']} straight-through",
            f"{gl_agent['summary']['manager_queue']} manager queue",
            f"{gl_agent['summary']['controller_queue']} controller queue",
            f"{gl_agent['summary']['cfo_queue']} CFO queue",
        ]
    elif intent == "erp_posting":
        holds = erp_posting["manual_hold_queue"].head(3)
        hold_lines = [
            f"- {row['Posting Item ID']}: {row['Source Agent']} | {row['Manual Hold Reason']}"
            for _, row in holds.iterrows()
        ]
        answer = (
            f"The ERP posting simulator currently sees {erp_posting['summary']['auto_post']} items on auto-post, "
            f"{erp_posting['summary']['ready_to_post']} ready to post, and "
            f"{erp_posting['summary']['manual_hold']} still on manual hold. "
            f"{erp_posting['summary']['eligible_for_auto_post']} items are structurally eligible for auto-post once the right agent actions are switched on.\n\n"
            + ("\n".join(hold_lines) if hold_lines else "No manual holds remain in the simulated posting queue.")
        )
        source_metrics = [
            f"{erp_posting['summary']['auto_post']} auto-post",
            f"{erp_posting['summary']['ready_to_post']} ready to post",
            f"{erp_posting['summary']['manual_hold']} manual hold",
            f"{erp_posting['summary']['eligible_for_auto_post']} eligible",
        ]
    elif intent == "bank_agent":
        worklist = bank_agent["worklist"].head(3)
        action_lines = [
            f"- {row['Entity']}: {row['Description']} -> {row['Recommended Action']}"
            for _, row in worklist.iterrows()
        ]
        answer = (
            f"The bank reconciliation agent is triaging {bank_agent['summary']['open_items']} open cash exceptions. "
            f"It sees {bank_agent['summary']['auto_clear_candidates']} auto-clear candidates, "
            f"{bank_agent['summary']['journal_candidates']} journal candidates, "
            f"{bank_agent['summary']['escalations']} escalations, and "
            f"{bank_agent['summary']['manual_investigations']} manual investigations.\n\n"
            + "\n".join(action_lines)
        )
        source_metrics = [
            f"{bank_agent['summary']['open_items']} bank exceptions",
            f"{bank_agent['summary']['journal_candidates']} journal candidates",
            f"{bank_agent['summary']['escalations']} escalations",
            f"{bank_agent['summary']['manual_investigations']} manual investigations",
        ]
    elif intent == "ic_agent":
        worklist = ic_agent["worklist"].head(3)
        action_lines = [
            f"- {row['IC Transaction ID']}: {row['Pair']} | {row['Issue Type']} | {row['Recommended Action']}"
            for _, row in worklist.iterrows()
        ]
        answer = (
            f"The intercompany agent has auto-matched {ic_agent['summary']['auto_matched']} of {ic_agent['summary']['total_pairs']} IC pairs. "
            f"It sees {ic_agent['summary']['open_exceptions']} open reconciliation exceptions, "
            f"{ic_agent['summary']['fx_mismatches']} FX-driven mismatches, "
            f"{ic_agent['summary']['elimination_drafts']} elimination entries ready to draft, and "
            f"{ic_agent['summary']['tp_flags']} transfer-pricing or agreement flags.\n\n"
            + "\n".join(action_lines)
        )
        source_metrics = [
            f"{ic_agent['summary']['auto_matched']} auto-matched",
            f"{ic_agent['summary']['open_exceptions']} open IC exceptions",
            f"{ic_agent['summary']['fx_mismatches']} FX mismatches",
            f"{ic_agent['summary']['tp_flags']} TP/agreement flags",
        ]
    elif intent == "journal_agent":
        drafts = journal_agent["erp_journal_drafts"].head(3)
        draft_lines = [
            f"- {row['Journal ID']}: {row['Entity']} | {row['JE Type']} | EUR {row['Amount (EUR)']:,.2f}"
            for _, row in drafts.iterrows()
        ]
        answer = (
            f"The journal entry agent has reviewed {journal_agent['summary']['candidate_items']} accrual candidates and prepared "
            f"{journal_agent['summary']['erp_ready_drafts']} draft JEs ready for ERP posting. "
            f"Those include {journal_agent['summary']['contract_backed_drafts']} contract-backed accruals and "
            f"{journal_agent['summary']['standard_reclasses']} standard reclassifications. Every candidate has an attached AI audit trail.\n\n"
            + "\n".join(draft_lines)
        )
        source_metrics = [
            f"{journal_agent['summary']['erp_ready_drafts']} ERP-ready drafts",
            f"{journal_agent['summary']['contract_backed_drafts']} contract-backed",
            f"{journal_agent['summary']['standard_reclasses']} reclasses",
            f"{journal_agent['summary']['audit_trails_attached']} audit trails",
        ]
    elif intent == "audit_agent":
        review_items = audit_agent["exception_queue"].head(3)
        review_lines = [
            f"- {row['Journal ID']}: {row['Posting Recommendation']} | {row['Primary Control Gap']} | {row['Required Approver']}"
            for _, row in review_items.iterrows()
        ]
        answer = (
            f"The audit and compliance agent reviewed {audit_agent['summary']['drafts_reviewed']} ERP-ready journal drafts. "
            f"{audit_agent['summary']['ready_to_post']} are ready to post, "
            f"{audit_agent['summary']['conditional_approval']} need conditional approval, and "
            f"{audit_agent['summary']['blocked_for_review']} are blocked for review. "
            f"The average control completeness score is {audit_agent['summary']['average_control_score']}/100.\n\n"
            + ("\n".join(review_lines) if review_lines else "No additional control exceptions are open.")
        )
        source_metrics = [
            f"{audit_agent['summary']['ready_to_post']} ready to post",
            f"{audit_agent['summary']['conditional_approval']} conditional",
            f"{audit_agent['summary']['blocked_for_review']} blocked",
            f"{audit_agent['summary']['average_control_score']}/100 avg score",
        ]
    elif intent == "checklist_agent":
        handoffs = checklist_agent["handoff_queue"].head(3)
        handoff_lines = [
            f"- Step {row['Step']}: {row['Task']} | {row['Status']} | {row['Action Lane']}"
            for _, row in handoffs.iterrows()
        ]
        answer = (
            f"The checklist agent is focused on {checklist_agent['summary']['blocked_tasks']} blocked tasks, "
            f"{checklist_agent['summary']['waiting_tasks']} waiting-on-input tasks, and "
            f"{checklist_agent['summary']['critical_open_tasks']} critical tasks still open. "
            f"It sees {checklist_agent['summary']['handoffs_at_risk']} dependency handoffs at risk and "
            f"{checklist_agent['summary']['recoverable_hours']} recoverable hours if the team clears the top bottlenecks first.\n\n"
            + ("\n".join(handoff_lines) if handoff_lines else "No checklist bottlenecks are currently queued.")
        )
        source_metrics = [
            f"{checklist_agent['summary']['blocked_tasks']} blocked",
            f"{checklist_agent['summary']['waiting_tasks']} waiting",
            f"{checklist_agent['summary']['critical_open_tasks']} critical open",
            f"{checklist_agent['summary']['recoverable_hours']} hrs recoverable",
        ]
    elif intent == "below_four":
        label_map = {
            "gl_straight_through": "GL auto-approve",
            "gl_fast_track_controller": "GL fast-track",
            "bank_auto_clear": "Bank clear/post",
            "ic_post_drafts": "IC elimination",
            "journal_post_ready": "Journal auto-post",
            "checklist_unblock": "Checklist unblock",
            "ap_exception_sweep": "AP sweep",
        }
        bundles = [
            ["gl_straight_through", "gl_fast_track_controller", "checklist_unblock"],
            ["gl_straight_through", "gl_fast_track_controller", "bank_auto_clear"],
            ["bank_auto_clear", "ic_post_drafts", "checklist_unblock"],
            ["gl_straight_through", "bank_auto_clear", "ic_post_drafts"],
        ]
        bundle_lines = []
        for bundle in bundles:
            scenario = simulate_close_scenario(kpis, bundle)
            if not scenario["gets_below_four"]:
                continue
            labels = [label_map[action["id"]] for action in scenario["selected_actions"]]
            bundle_lines.append(f"- {' + '.join(labels)} -> {scenario['score']['continuous_close_days']} day forecast")
        answer = (
            f"The continuous close forecast is currently {score['continuous_close_days']} days. "
            f"To get below 4.0, NovaClose needs a bundled move rather than a single fix. "
            f"The best bundles in the simulator right now are:\n\n"
            + ("\n".join(bundle_lines) if bundle_lines else "No tested bundle currently gets below 4.0 days.")
        )
        source_metrics = [
            f"{score['continuous_close_days']} day base forecast",
            f"{gl_agent['summary']['straight_through_candidates']} GL auto-approvals",
            f"{bank_agent['summary']['auto_clear_candidates']} bank auto-clears",
            f"{checklist_agent['summary']['recoverable_hours']} checklist hrs",
        ]
    else:
        answer = (
            "NovaClose Copilot is optimized for close blockers, today’s controller actions, automation opportunities, "
            "GL approval triage, bank reconciliation triage, intercompany matching, journal drafting, checklist unblocking, audit review, riskiest entity analysis, and trial balance variance summaries. Start with one of those questions and I’ll answer using the dataset already loaded into the app."
        )
        source_metrics = [
            f"{summary['pending_gl']} pending GL",
            f"{summary['bank_exceptions']} bank exceptions",
            f"{summary['ap_3way_exceptions']} AP issues",
        ]

    return {
        "intent": intent,
        "answer": answer,
        "source_metrics": source_metrics,
        "suggested_prompts": follow_ups,
    }


def run_analysis(path=DATA_FILE):
    data = load_data(path)
    kpis = calculate_kpis(data)
    score = calculate_readiness_score(kpis)
    priorities = priority_engine(kpis)
    commentary = generate_commentary(kpis, score)
    return {"kpis": kpis, "score": score, "priorities": priorities, "commentary": commentary}


if __name__ == "__main__":
    results = run_analysis()
    summary = results["kpis"]["summary"]
    print("\n=== NOVACLOSE AI ANALYSIS ===")
    print(f"Readiness Score: {results['score']['readiness_score']}/100")
    print(f"Risk Level: {results['score']['risk_level']}")
    print(f"Predicted Close Days: {results['score']['predicted_close_days']}")
    print(f"Gap To 4-Day Target: {results['score']['gap_to_target_days']}")
    print("\nHeadline KPIs:")
    print(f"- Pending GL approvals/reviews: {summary['pending_gl']}")
    print(f"- Manual journal entries: {summary['manual_jes']}")
    print(f"- Bank exceptions: {summary['bank_exceptions']}")
    print(f"- Intercompany exceptions: {summary['ic_exceptions']}")
    print(f"- AP 3-way exceptions: {summary['ap_3way_exceptions']}")
    print(f"- Blocked checklist tasks: {summary['blocked_tasks']}")
    print("\nTop Priorities:")
    for index, priority in enumerate(results["priorities"][:6], start=1):
        print(f"{index}. {priority['priority_item']} ({priority['area']}, score {priority['priority_score']})")
    print("\nCommentary:")
    print(results["commentary"])
