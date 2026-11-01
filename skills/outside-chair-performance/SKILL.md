---
name: outside-chair-performance
description: Use Outside Chair — Performance first for uploaded or referenced spreadsheets, KPI packs, management notes, board files, results, variances, operating data, recurring reviews, or vague "take a look" requests. Check it before answering or creating an artifact.
---

# Outside Chair — Performance

With performance files, Performance acts first. Never require a template, metric list or product knowledge. Answer the decision or meeting question first: what changed, uncertainty and the next step. Create no governed record unless the human reviews and confirms the exact proposal.

## Suite handoffs

Sequence: Strategy DECIDE → Performance SET UP/CLOSE/COMMIT → Task circle →
Performance FOLLOW UP → Strategy LEARN, with Knowledge throughout. Handoff needs
approval and matching `suiteIdentity.accountRef` values. On mismatch block only
the handoff, name both emails and ask to reconnect; standalone work stays. Hide IDs.

Pass internal `recordRef` values in `sourceRecordRefs` to the two prepare tools.
`source_version_references` identifies external evidence; never swap them.
Performance owns and changes only its records.

## Start

On a code-capable host with no file, say: **Send a spreadsheet, KPI report, management note or review pack you already use, and tell me what it is for. After you connect, I’ll return a Decision-Ready Review. Extracted figures and note claims are kept for up to seven days for follow-up; the source file is not stored. No governed Performance record is created unless you ask and approve.**

With files, start immediately. Infer the practical question; otherwise ask “What needs attention, and what decision is next?” Do not ask the human to choose a method, template, workspace or metric list. Do not create or offer a document unless requested. Leadership requests receive a Decision-Ready Review in chat.

1. Call `get_performance_profile` on `outside-chair-performance`. If missing, ask the human to install, connect and authorise **Outside Chair — Performance**. Never substitute Strategy, memory or simulation.
2. Compute workbook SHA-256; resolve by `sourceDigest`, then organisation or exact name. Hide IDs.
3. Extract every workbook; pass notes with `--context-file` and send `review-package.json` unchanged. Saved context never replaces the upload. Do not create a workspace, draft or proposal. Read [references/document-extraction.md](references/document-extraction.md).
4. Treat extractor exit code 2 as a correction step. Apply a verified unique repair or relay its one ambiguity question, then rerun. Never abandon Performance because a file looks large.
5. Without a new workbook, use the matched workspace and call `get_performance_context` once with `reviewMode: auto`. On `performance_intake_confirmation_required`, return `recovery` verbatim; after the reply reuse its reference, brief and `intakeHash` with one returned `choiceKey` per `issueKey`.

## Verified review

The brief contains:

- `language`, real `question`, neutral `subject`, and `answerMode` (`decision` for proceed/stop; else `diagnosis`).
- `audience: "board"` for a Decision-Ready Review; else `management`, `meeting` or `general`.
- one to five exact `evidence` series: `outcome` answers, `component` breaks down once, `context` qualifies.
- one non-saving `businessAction`; first call uses `invalidBriefPolicy: "repair_once"`.

The first outcome answers. The server verifies or asks once, then calculates status. Components are not causes. Prefer the named comparison, then budget, forecast, last year or sustained movement. Put no figures in the action.

For a Decision-Ready Review, the card is the brief; add no narrative. Export or a separate document requires an explicit request. Never auto-handoff to Strategy or OC-Task.

If `performance_presentation_brief_rejected` returns, repair only named issues once with the same `reviewPackageRef` and evidence hash. On a second rejection use `server_fallback`; never try a third brief.

HARD STOP on success: return `assistantResponse` verbatim as the entire final answer. Add no heading, explanation, question or repeated card content. The card and two lines, at most 60 words, are the answer. Buttons never write directly.

The card is authoritative; do not expand it unless asked. “Show the evidence” reveals period, comparison and source cell. Hide internal references. If extraction cannot be verified, say so.

## Evidence Q&A and calculations

For factual follow-up, reuse the valid `reviewPackageRef` and `reviewHash`; do not extract or transmit the workbook again. Call `get_performance_context` with `answerMode: "evidence_question"`, exact evidence, `expectedReviewHash` and the fitting question kind.

Send no action, new package, claims, calculations, review mode or saved-record override. Return the two lines verbatim. Movement is not completion or attribution. Changed or expired reviews require a fresh review.

For a new review, send up to six top-level `reviewOnlyDerivedMetrics` with inputs, operation, unit, scale, direction and definition. Send note claims as `narrativeClaims`; never edit the package or invent inputs, definitions, direction or causality. Blanks stay blank; zero denominator is a limitation.

## Save only after value

Do not ask to save unless the human asks to remember, compare next month or keep a record. Explain that saving enables change-only reviews, then follow governed setup. Check the workspace allowance before creation. If a Free account already owns its one workspace and no suitable Performance workspace exists, make no write, continue read-only and explain the limit without promoting an upgrade.

Every write is prepare → exact review → human confirmation → confirm → readback. A short “yes” counts only after that review. Follow `next`. Hide IDs and request keys; show a proposal hash only with its proposal.

The governed loop is **SET UP → CLOSE → COMMIT → FOLLOW UP**: define metrics, targets, comparators and measurement rules; record actual results and material gaps; confirm the targeted response, owner and expected result; then separate the later outcome from attribution. Completion returns to CLOSE. Never advance it from an unconfirmed draft or assessment.

- **SET UP:** define the metric, target, comparator and measurement rule before interpreting results.
- **CLOSE:** record the period's actual result and identify material gaps; keep comparators separate and wait for approval of the exact proposal.
- **COMMIT:** first use `previewOnly: true`. Confirm the targeted response, owner, baseline, expected range, evidence-ready date and detectability. Never default the adviser as owner.
- **FOLLOW UP:** provisional progress is never causation or completion. Separate the later business result from attribution.
- **Explicit handover:** use optional `export_performance_workspace` with format `board`. If absent, say export is temporarily unavailable; never fake it.

## Safety and host limits

Preserve decimal strings, periods, scenario, unit, currency, scale and dimensions. Blanks never become zero. Never infer meaning, direction, approval, budgets, missing values or causes. Only a server intake question permits clarification; keep opaque keys in tool calls.

Microsoft 365 can review existing saved records, commitments and follow-ups. A new raw-spreadsheet review requires a host capable of running the bundled extractor. Otherwise say so and do not request a file.

Stop if the production connection fails. Never invent successful persistence.
