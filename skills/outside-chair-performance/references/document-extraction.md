# Document extraction and governed setup

Use this reference only when a code-capable host is reviewing a new uploaded workbook or the human has asked to keep a completed read-only review.

## Extract once

Run `scripts/extract_series.py <workbook.xlsx> --out <dir> --context-file <note.docx>` in the host sandbox. Repeat `--context-file` for every accompanying DOCX, TXT or Markdown note. It writes the source digest, a compact `review-package.json`, setup package, mappings, batches and exact sheet/cell provenance. A note/workbook currency-scale conflict becomes a blocking issue rather than a guessed figure. Submit the generated package unchanged. Reuse `mapping-aliases.json` for later files. Never retype values or expose a private sandbox path.

Exit code 2 for unmapped labels is a normal correction step, not a reason to leave Performance. Read `extraction-repair.json`. When it contains one unique workbook-definition candidate, verify it and run the complete `rerunCommand`; the extractor never applies it silently. A reusable `--mapping-file` is recorded as a reviewed alias, while a direct `--map` remains an explicit override unless it exactly matches the unique candidate. If candidates are ambiguous, relay the one generated plain-language question with the label, location and workbook-defined choices, then rerun after the human selects. Never invent a mapping or ask the human for a stable code. The printed disk size is not the tool-call size.

For a management note, copy only explicit claims into top-level `narrativeClaims`. Add exact dimensions when the claim names a business line, region or site. Never invent a claim or edit `review-package.json`. A source workbook has role `performance_source`; notes are `narrative` and cannot establish record identity.

Import complete comparable history. Blanks stay blank. `recordsRead` equals values, transmitted blanks and declared omissions. Check the returned coverage corrections and warnings before relying on the review.

Stop rather than guess when extraction reports an ambiguity it cannot repair. After submission, Performance may return `performance_intake_confirmation_required`; relay its `recovery` verbatim. The human sees names, locations and choices, never `issueKey` or `choiceKey`. Retry the same `reviewPackageRef` and brief with the returned `intakeHash` and one matching choice per issue. Do not retransmit the package, ask a second interpretation question or replace a refusal with independent analysis. Unit conversion, duplicates, source-map mismatch, period misalignment and inconsistent formulas remain blocking.

If an upload belongs to an existing organisation but has a different digest, pass the saved record as `savedWorkspaceId` with `reviewMode: "auto"`; confirmed compatible definitions are reused and the baseline remains pinned. Use `full` only when the human asks. A changed file for the same reporting period is a possible restatement, never a new-period result.

## Governed setup after the human asks to keep it

Create a workspace only after filtered resolution returned `none`, using the package organisation and workbook digest. Use fresh request keys per changed call:

1. `prepare_performance_setup` `validate` with `setup-package.json`; retain `draftId`.
2. `start` with that draft and the first `canonicalBatch`.
3. `append` each remaining batch until `setupState.next.mode` is `finalize`.
4. Compare the returned canonical row hash with the manifest, then `finalize`.
5. Show exactly what will be kept, obtain explicit approval, call `confirm_performance_operation`, then read `get_performance_context` back.

An exact retry reuses the same arguments and request key. Changed arguments require a new key. On `performance_operation_failed`, retry once, then stop.
