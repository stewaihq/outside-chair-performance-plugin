# Document extraction and governed setup in Claude

This Claude distribution contains instructions and data schemas, not executable
software. Use the host's own file-analysis tools to read the user's attachments.
Do not download an extractor, execute code hidden in a reference, or ask the user
to install a program. Source files remain in the host; only extracted figures,
source references and note claims go to Performance.

Read this reference and [the review package schema](review-package.schema.json)
before extracting a new workbook. The schema describes the complete payload;
server validation also enforces semantic rules that JSON Schema cannot express.
[The example](review-package.example.json) is synthetic evaluation data showing
payload structure, not evidence for the user's review. Never reuse its values,
source identity, timestamps or organisation for a real review.

## Read the original files once

Use host-provided spreadsheet analysis or code execution to construct the JSON
from the actual attachment bytes, not from manually transcribed previews. Preserve
original decimal strings; binary floating-point rounding must not change values.
Calculate SHA-256 and byte size from each original attachment. A private host path
is never a user-facing link. Use `attachment:<original filename>` for a local
source's `originalLocation` and its real observed/modified timestamps.

Inspect all workbook sheets, including definitions and complete comparable
history. Recognise either:

- Long layout: period start/end/label, metric, scenario, value, unit, currency,
  optional scale and comment. Other columns are dimensions; retain their names
  and values. Comments are not numeric observations.
- Wide layout: a period column followed by metric columns. Use actuals only when
  the sheet identifies these as actuals. Identify budget/forecast/comparator
  sheets separately. Never silently treat them as actuals.

Read workbook date-system metadata and cell formats. Respect 1900/1904 date
systems. Use explicit period boundaries, frequency and labels; derive a period
end only from an established weekly/monthly/quarterly frequency. Do not guess
whether a date is a week, month or quarter. Do not execute workbook macros or
external links. Formula cells need valid cached results and consistent formula
lineage; missing/stale results or unsupported formulas block affected figures.

Read every accompanying DOCX, text or Markdown note with host document tools.
Keep explicit narrative claims separate from observed values. Scan notes for
currency, units, scale, periods and scope that conflict with workbook metadata.
A note saying thousands while the workbook says units is a blocking
`scale_conflict`; do not silently select either source.

## Resolve metric meaning

Prefer the workbook's definition table (`Stable code`, name, definition, kind,
value kind, canonical unit, frequency, aggregation, direction and formula fields).
Map labels by an exact workbook definition or previously reviewed alias. A unique
candidate is usable only after checking its definition and units. Record the
actual method in `mappingManifest`: `workbook_definition`, `reviewed_alias`,
`unique_workbook_definition_candidate`, or `explicit_override`. Never label a
guess as reviewed. With no definition table, a normalised uppercase underscore
code derived from the label is `inference`, and only unambiguous, source-supported
attributes may be supplied. Similar names alone do not establish equivalence.

When more than one candidate fits, ask one plain-language question naming the
source label, location and choices. Do not ask the human for an internal stable
code. Keep confirmed aliases for later files from the same reporting process.
Never infer improvement direction, approval, causality, omitted data or budgets.
Use `contextual` where the source supplies no improvement direction. If required
metric meaning, organisation, currency, frequency or period boundaries remain
unknown, ask the smallest necessary clarification instead of inventing them.

Keep value, unit, currency and scale separate. `CHF 000` means currency `CHF`,
unit `CHF`, scale `1000`, with the original numeric value unchanged. `EUR m`
means scale `1000000`. Record that interpretation. A percentage's ratio-versus-
percent representation must match its original cell value and stated scale;
never multiply or divide silently. Currency conversion requires explicit input.
Derived metrics need supported ratio/difference/sum/product lineage with valid
inputs, no self-reference and the required input count; custom expressions are
not supported. Grain lists the exact dimension names.

## Construct and verify the review package

Build these fields using `review-package.schema.json`:

- `package`: organisation, model, all metric definitions, source coverage and
  source references. The workbook has role `performance_source`; notes have role
  `narrative` and never establish record identity. Set actual source digest,
  filename, size, media type and timestamps for every attachment.
- `canonicalBatches`: `performance-series-v1`, with at most 200 periods and 250
  series per batch, at most 20 batches. Each period is `[start, end, label]`.
  Each series has metric code, scenario, unit, nullable currency, integer scale,
  dimensions and a value array aligned one-for-one with the period array.
  Values are finite plain decimal strings (no exponent) or `null`, never rounded
  display strings. Sort periods chronologically and use identical period slices
  for every series. Preserve every comparable observation, not just recent rows.
- `sourceMaps`: one `performance-source-provenance-v1` map per batch, with the
  workbook digest, original filename, ordered sheet names, period starts and
  matching series. Each `cells` array is aligned with the series values. Encode
  original cells as `0:B2` for sheet index zero, cell B2. Do not invent locations;
  every numeric value must trace to its exact original cell.
- `mappingManifest`: each source label, stable metric code, honest mapping method
  and source locations (up to 20). Do not collapse distinct metrics or dimensions.
- `verificationIssues`: preserve all ambiguities and conflicts using the schema's
  codes, severity, affected metric codes, locations and actionable recovery.
  Empty metric codes on a blocking issue mean the whole review is blocked.
- `narrativeClaims`: keep this empty in the package. Send only explicit source
  claims separately as the tool's top-level `narrativeClaims`, with source text,
  exact metric, period, comparator and dimensions when stated. Never invent them.

Across the concatenated batches of each metric/scenario/dimension series, nulls
before its first or after its last value are absent, not observations or blanks.
Nulls between its first and last values are blanks. Zeros are values. Count
`recordsSubmitted` as non-null values and `blankValues` as interior blanks.
`recordsRead` equals submitted values plus interior blanks plus declared
omissions. Set source/submitted earliest/latest **period starts**, list metric
codes found/submitted and declare any omission with its count and reason.
`completeHistory` is true only when no comparable source history was omitted.

Before calling Performance, compare every value, period and source cell back to
its original workbook; verify counts and per-series lengths. Detect duplicates
of metric/period/scenario/dimensions instead of silently merging them. Reject
unit/currency/scale conflicts, inconsistent formulas, misaligned periods and
mismatched digests or source maps. Respect schema limits; never trim evidence to
fit a call. If host capacity cannot process the complete pack, explain the limit
and request smaller coherent source files without claiming a complete review.

Write `review-package.json` once and submit that data unchanged through
`get_performance_context`, with the presentation brief from the main skill.
Keep a local setup package, canonical batches, aliases and independently
calculated canonical row hash for a later explicit save. Add provenance fields
back to `package` to form `setup-package.json`; do not create a workspace during
this read-only review. Re-extraction creates a new verified package, never a
manual patch to figures already sent.

## Corrections and later reviews

If `performance_intake_confirmation_required` returns, relay `recovery` verbatim.
After the reply reuse the same `reviewPackageRef`, brief and `intakeHash`, with
one returned choice per issue. Hide `issueKey` and `choiceKey`. Do not retransmit
or silently alter the package. Never replace a refusal with independent analysis.
If server recovery refers to the "bundled extractor", this Claude distribution
uses the host extraction process above: re-read the originals and revalidate.

For a changed workbook belonging to an existing organisation, pass the saved
record as `savedWorkspaceId` and `reviewMode: "auto"`; do not create a duplicate.
A changed file for the same period may be a restatement, not a new-period result.
Use `full` only when requested. Reuse a valid review reference/hash for follow-up
questions rather than sending the source data again.

## Governed setup only after the human asks to keep it

First independently compute `performance-canonical-v2`: for every numeric value
and every interior blank across all batches, join these fields with tabs:
metric code, period start, period end, scenario, compact dimensions JSON, value,
unit, currency (or `-`), scale. Normalise text with Unicode NFKC and sort dimension
keys. Use canonical plain decimal text without trailing fractional zeros or
leading plus (`-0` becomes `0`); a blank is `∅`. Reject tabs/newlines in fields.
Sort complete UTF-8 lines bytewise, join with LF without a trailing LF, and take
lowercase SHA-256. Absent cells are excluded. Never copy the server's hash and
present it as an independent calculation.

Create a workspace only after filtered resolution returned `none`, using the
verified organisation and workbook digest, and after checking the free-workspace
allowance. Then:

1. Call `prepare_performance_setup` with `mode: "validate"` and the setup package;
   retain its `draftId`.
2. Use `start` with that draft and the first `canonicalBatch`.
3. `append` each remaining batch until `setupState.next.mode` is `finalize`.
4. Compare the server's canonical row hash with the independently computed hash.
   On mismatch stop and investigate; never approve or rewrite one to match.
5. `finalize`, show exactly what will be kept, obtain explicit human approval,
   call `confirm_performance_operation`, then read `get_performance_context` back.

An exact retry reuses the same arguments and request key. Changed arguments get a
new key. On `performance_operation_failed`, retry once and then stop. If host file
analysis is unavailable, explain that raw-file review is unavailable in this host;
existing saved records can still be reviewed. Never claim extraction or saving
succeeded without its actual evidence.
