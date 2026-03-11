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

    bank_exception_mask = ~bank["Match Status"].eq("Matched")
    bank_exceptions = int(bank_exception_mask.sum())
    bank_matched = int(bank["Match Status"].eq("Matched").sum())
    bank_stale = int(bank["Days Outstanding"].fillna(0).gt(30).sum())
    bank_match_rate = float(safe_pct(bank_matched, len(bank)))
    bank_exception_breakdown = (
        bank.loc[bank_exception_mask, "Match Status"].value_counts().sort_values(ascending=False).to_dict()
    )
    ic_agent = _build_ic_agent(ic)
    bank_agent = _build_bank_agent(bank)
    journal_agent = _build_journal_agent(accruals)
    audit_agent = _build_audit_agent(bank_agent, ic_agent, journal_agent)
    checklist_agent = _build_checklist_agent(checklist)

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
        "bank_exceptions": bank_exceptions,
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

    return {
        "summary": summary,
        "areas": areas,
        "entities": entities,
        "details": details,
        "agents": {
            "bank": bank_agent,
            "ic": ic_agent,
            "journal": journal_agent,
            "audit": audit_agent,
            "checklist": checklist_agent,
        },
    }


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

    return {
        "readiness_score": score,
        "risk_level": risk_level,
        "predicted_close_days": predicted_days,
        "gap_to_target_days": gap_to_target_days,
        "penalties": penalties,
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
    if any(keyword in text for keyword in ["variance", "trial balance", "tb"]):
        return "tb_variances"
    if any(keyword in text for keyword in ["automate", "automation", "auto", "automated"]):
        return "automation_opportunities"
    if any(keyword in text for keyword in ["cfo", "executive", "summary", "board"]):
        return "cfo_summary"
    if any(keyword in text for keyword in ["delay", "blocker", "risk", "close"]):
        return "close_blockers"
    return "general"


def generate_copilot_response(prompt, kpis, score, priorities):
    summary = kpis["summary"]
    entities = kpis["entities"]
    details = kpis["details"]
    bank_agent = kpis["agents"]["bank"]
    ic_agent = kpis["agents"]["ic"]
    journal_agent = kpis["agents"]["journal"]
    audit_agent = kpis["agents"]["audit"]
    checklist_agent = kpis["agents"]["checklist"]
    top_priorities = priorities[:3]
    top_entity = entities[0] if entities else None
    top_variances = details["tb_variances"].head(3)
    intent = _route_copilot_intent(prompt)

    follow_ups = [
        "What issues will delay the close?",
        "What should the bank agent do next?",
        "What should the IC agent do next?",
        "What should the journal agent post?",
        "What checklist tasks need to be unblocked?",
        "What does audit need to review before posting?",
        "What should the controller do today?",
        "Which entity is riskiest?",
        "Explain the biggest variances.",
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
    elif intent == "tb_variances":
        variance_lines = [
            f"{row['Entity']}: {row['Account Name']} at {row['Variance %']}"
            for _, row in top_variances.iterrows()
        ]
        answer = (
            f"The trial balance shows {summary['large_variances']} lines moving more than 10% and "
            f"{summary['unreconciled_tb']} unreconciled accounts. The biggest movements right now are:\n\n"
            + "\n".join(f"- {line}" for line in variance_lines)
        )
        source_metrics = [
            f"{summary['large_variances']} large variances",
            f"{summary['unreconciled_tb']} unreconciled TB lines",
            f"Max variance {summary['max_tb_variance_pct']}%",
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
    else:
        answer = (
            "NovaClose Copilot is optimized for close blockers, today’s controller actions, automation opportunities, "
            "bank reconciliation triage, journal drafting, checklist unblocking, audit review, riskiest entity analysis, and trial balance variance summaries. Start with one of those questions and I’ll answer using the dataset already loaded into the app."
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
