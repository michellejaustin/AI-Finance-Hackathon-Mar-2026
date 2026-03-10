# NovaClose AI Demo Script

## 1. Opening line

"NovaTech does not have a transaction-processing problem. It has an exception-management problem. The close slows down because finance teams find issues too late, chase approvals manually, and lose time untangling dependencies."

## 2. Frame the problem

"NovaTech currently closes in 7 business days. The CFO wants to get below 4. We analyzed the month-end workbook and found that the biggest blockers are not SAP itself, but the operational friction around journals, reconciliations, accruals, and workflow."

## 3. Start in Command Center

Lead with the copilot and the readiness gauge.

Say:

- "This is NovaClose AI, our month-end close intelligence layer."
- "It gives finance a live readiness score, predicted close days, and a ranked view of what will delay the close."
- "The current demo projects a 4.4-day close, which is 0.4 days above target."

## 4. Ask the copilot first

Use:

- `What issues will delay the close?`

Then say:

"We designed the experience to feel like a finance copilot, but it is grounded in the workbook and returns finance-specific evidence behind every answer."

Call out the main blockers:

- 171 pending GL approvals or reviews
- 10 bank exceptions
- 2 intercompany gaps
- 8 AP 3-way match exceptions
- 3 blocked checklist tasks

## 5. Show the readiness gauge

Say:

- "The score is 67 out of 100."
- "That puts NovaTech at medium risk of missing the 4-day target."
- "This is not a static score. It is driven by actual penalty factors in the dataset."

Then point out:

- predicted close days
- gap to target
- riskiest entity

## 6. Move to Risk Atlas

Say:

"The Risk Atlas scans every sheet in the workbook. We keep the narrative centered on the biggest close blockers, but we also surface secondary signals from trial balance and fixed assets so judges can see full workbook coverage."

Walk the judges through:

- GL and journals
- bank reconciliation
- intercompany
- accruals
- AP exceptions
- checklist workflow
- trial balance movement
- fixed-asset control flags

## 7. Explain the AP logic correction

Say:

"We corrected one important finance logic issue during the build. The AP exception metric now counts only true AP 3-way match exceptions, so the app shows 8 real AP issues instead of inflating the count with AR rows."

This is a good credibility moment.

## 8. Move to Priority Engine

Say:

"This is where the tool becomes operational. We do not just describe issues. We rank what finance should fix first based on close impact and downstream unlock."

Walk through the top items:

1. Clear pending GL reviews and approvals
2. Resolve bank reconciliation exceptions
3. Unblock critical close checklist tasks
4. Fix intercompany elimination gaps

Then say:

"The goal is to shift finance effort from exception hunting to targeted resolution."

## 9. Move to Entity View

Say:

"Because NovaTech has five legal entities, we also roll the signals into entity-level risk. This makes the tool feel like a real close control tower rather than a single KPI dashboard."

Call out:

- which entity is riskiest
- what its primary blocker is
- how that helps management allocate attention

## 10. Impact statement

Use this line:

"NovaClose AI moves NovaTech from reactive close management to predictive close intelligence. Instead of finding problems on day 5 or 6, finance can detect blockers immediately, prioritize work, and shorten the path to target."

## 11. Quantified closeout

Suggested summary:

- 7-day current close
- 4.4-day projected outcome at current risk
- clear path to recover hours by fixing top priorities
- visibility across all 8 sheets and all 5 entities

## 12. Q&A backup answers

### Why is this AI and not just a dashboard?

"Because it does three intelligence tasks: it scans for exceptions across the close, prioritizes what matters operationally, and provides a copilot-style finance interface for explanation and action."

### Why is the copilot local and not tied to a live API?

"For the hackathon we optimized for demo reliability. The UI is designed to be provider-ready, but the current version keeps answers deterministic so it still works without internet access."

### What would you automate next?

"Journal approval routing, recurring bank matching, recurring accrual drafting, and variance commentary generation."

### Why does the entity view matter?

"Because close risk is never perfectly uniform across entities. The entity view helps the controller and CFO focus attention where the close is genuinely at risk."
