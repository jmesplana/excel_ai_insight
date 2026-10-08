# Configuring Jev classification and dataset reports

Import JSON on **Jev Classification → Configure Analysis → Import Configuration**. Start with [the complete example](examples/configurable-workflow.jev-config.json). Change `sourceColumns` and `report.evidenceColumns` to your spreadsheet's headers, and replace the example questions and labels with your own codebook. There are no dataset-specific column names, categories, framework tables or report topics in the workflow engine.

Existing version 1 configurations remain importable. Export writes version 2. Advanced settings remain visible in the question's **Question rules (JSON)** editor and the **Workflow and report settings (JSON)** editor. They survive import, UI edits, and export. Unsupported version 2 field names are rejected rather than silently treated as working features.

## Top-level fields

| Field | Meaning |
| --- | --- |
| `format` | Must be `aidstack-insights-jev-config` |
| `version` | `2` |
| `sourceColumns` | Source headers to evaluate; missing columns produce an import warning |
| `sheetName` | Optional informational label; the selected sheet controls the run |
| `model` | Optional model ID; otherwise API Settings/server defaults apply. After the first successful response the resolved version is saved for the run |
| `includeConfidence` | Show confidence and raw Score/Noul columns in the main results sheet; raw answers are retained regardless |
| `questions` | Ordered array of question definitions |
| `derived` | Optional lookup or composite outputs |
| `outputOrder` | Optional array of output names in results, export and report order. May mix question and derived outputs; unlisted outputs follow in their original order. Confidence/raw score columns follow their corresponding question. Execution still follows question dependencies. |
| `report` | Optional report definitions and evidence selection |
| `execution` | Optional throughput settings |

Credentials belong in API Settings or server environment variables, never configuration JSON. Importing a configuration does not call Jev. Output names, including generated audit columns, must not collide with source or other output columns.

## Questions

Every question has `outputColumnName`, `questionType`, `instructions`, and normally `options`.

- **choice:** 2–255 options. Each option is a string label or `{ "label": "Name", "description": ... }`. Descriptions and instructions can be strings, objects or arrays. Include an Other/Not stated option if the codebook is not exhaustive.
- **score:** 2–10 ordered rubric levels, lowest first, with the same label/description format. Main cells show the nearest level, with exact midpoints rounded upward; the audit and statistics retain the original fractional score.
- **noul:** a yes/no statement. Optional `options` contains exactly two display labels, yes first and no second. Optional `criteria` contains `true` and `false` descriptions. Neither an ambiguous answer nor an absent answer is automatically a No.

`review` sets thresholds per question:

```json
{"review": {"minConfidence": 0.7}}
```

Choice/Score answers below that threshold keep their assigned value but are marked for review. Default: 0.5. Confidence measures distribution concentration; it is not an accuracy guarantee.

Classification cells contain only configured labels (or configured lookup values). Missing, failed, blocked and unassigned answers leave blank cells; review and error messages are retained in the decision audit and review queue. “Needs review” appears as a classification value only when it is explicitly one of that column's configured labels.

```json
{"review": {"noMax": 0.15, "yesMin": 0.85}}
```

Noul answers at or below `noMax` are No, at or above `yesMin` are Yes, and between them need review. Defaults: 0.2 and 0.8. All probabilities remain available for auditing. Use separate Noul questions to capture multiple simultaneous themes; their percentages need not sum to 100%.

## Conditional questions

Use `dependsOn` to name earlier question outputs and `branches` to supply allowed options for each accepted combination:

```json
{
  "outputColumnName": "Detail",
  "questionType": "choice",
  "instructions": "Classify the detail in record, using decisions.Category as the parent.",
  "dependsOn": ["Category"],
  "branches": [
    {"when": {"Category": "Request"}, "options": ["Information", "Assistance", "Other"]},
    {"when": {"Category": "Appreciation"}, "options": ["Service", "Staff", "Other"]}
  ]
}
```

Each `when` must match every dependency exactly. Order parents before children. A missing branch, uncertain parent, failed parent or empty parent blocks the child. The workflow never guesses a branch. Questions at each available stage run together; genuine dependencies require another call. Later state contains `record` (the source context) and `decisions` (the accepted dependency values).

To classify a large taxonomy, define parent questions then provide each branch's child options. Do not place the entire taxonomy into one Choice if it exceeds 255 options. The engine is generic; taxonomy values live in JSON.

## Derived outputs

A lookup enforces allowed combinations without an AI call:

```json
{
  "outputColumnName": "Queue",
  "type": "lookup",
  "inputs": ["Category", "Detail"],
  "table": [
    {"when": {"Category": "Request", "Detail": "Information"}, "value": "Information team"}
  ]
}
```

Missing matches are flagged for review and leave the value blank. A lookup maps an assigned input label even when that input is uncertain, retaining the mapped value and propagating the review status. Failed, blocked or unassigned inputs block the lookup. Derived inputs may reference earlier outputs. Questions can depend on earlier questions; derived outputs run after the questions.

Composite scores use positive weights referencing Score questions:

```json
{"outputColumnName": "Priority", "type": "composite", "weights": {"Urgency": 2, "Impact": 1}}
```

Each raw score is normalized by its rubric's maximum index. The result is a weighted mean from 0 to 1, and is blocked if any input requires review. It is a configurable ranking index, not a probability or physical measurement.

## Reports

```json
{
  "report": {
    "title": "Community feedback findings",
    "instructions": "Describe leading topics, uncertainty and suggested follow-up with supporting counts.",
    "groupBy": ["Region"],
    "crossTabs": [["Category", "Urgency"]],
    "evidenceColumns": ["Feedback"],
    "maxExamples": 12,
    "maxGroups": 100,
    "maxCategories": 30
  }
}
```

Change every column name to one in your dataset or configured outputs. Leave grouping/cross-tab arrays empty when not needed. Blank grouping fields are excluded. Counts use every processed original record, including records beyond 5,000. Errors, empty inputs and blocked decisions do not become ordinary categories. Uncertain assigned Choice/Score values contribute to distributions and are explicitly flagged; uncertain Noul answers have no assigned yes/no value. Each distribution supplies its own denominator.

The report shows total source records, selected records, processed coverage, successful/review/failed/blocked/empty records and unprocessed records. A test run is partial relative to the full uploaded sheet. Scope is the selected sheet, not all sheets in a workbook.

`maxCategories` and `maxGroups` bound the evidence sent to the text model, not aggregation. Omitted counts are explicit. Downloads retain full tables. Examples are selected to illustrate distinct outcomes, capped by `maxExamples`, with 1,000 characters per evidence field. They are not a representative sample, and the written report must not claim to have read all raw text. Report packets over the server limit are rejected with instructions to reduce these settings, never silently chopped.

**Generate written report** uses the separately configured OpenAI/Azure provider. It explains exact aggregates and illustrative examples. The statistical report works with only Jev credentials. Follow-up chat on Jev results uses the same evidence packet. New themes outside the configured codebook require new questions and another classification run; the narrative is not a substitute for full-text theme discovery.

The Excel download includes results, summary distributions, raw score summaries, group breakdowns, cross-tabs, review queue, decision audit, and run metadata/configuration. A generated narrative is included when present. JSON downloads contain the summary or narrative plus evidence. Record IDs are 1-based imported data-record positions, not physical Excel row numbers. Cleaning/splitting results afterward preserves the report's original classified-record basis to avoid double-counting feedback.

## Execution and recovery

```json
{
  "execution": {
    "batchSize": 4,
    "workers": 4,
    "requestsPerMinute": 600,
    "tokensPerSecond": 100000,
    "maxAttempts": 3,
    "structuredState": true
  }
}
```

Defaults above favor reliability. Allowed ranges: batchSize 1–25; workers 1–8; requestsPerMinute 1–1,200; tokensPerSecond 1–250,000; maxAttempts 1–4. These application ceilings can be updated if the provider changes its limits; they are not dataset assumptions. Structured state preserves selected field names and values as JSON. Otherwise the legacy labeled-text format is used.

`tokensPerSecond` currently uses UTF-8 request bytes as a conservative token upper bound, including question instructions, criteria keys and descriptions. Large vocabularies can therefore be limited by application pacing even when Jev itself responds quickly. Raising worker count alone does not remove this limit. Actual input-token usage is recorded separately.

For a human-assigned source field, use `sourceBranches` on an independent Choice question to select its vocabulary **before** inference. This avoids asking the model to infer a known type or sending irrelevant vocabularies. Use `execution.structuredState: true` and select every field named in `when` as a source column:

```json
{
  "outputColumnName": "code",
  "questionType": "choice",
  "instructions": "Classify only the clause matching state.type_of_feedback.",
  "options": [],
  "sourceBranches": [
    {"when": {"type_of_feedback": "Questions"}, "options": ["Question A", "Question B"]},
    {"when": {"type_of_feedback": "Demandes"}, "options": ["Request A", "Request B"]}
  ]
}
```

Conditions match exact source values. All branches must use the same fields, with unique conditions and 2–255 options each. They can follow a dependency such as a source excerpt, but cannot be combined with decision `branches`. Unknown/missing source values produce a blank decision requiring review, without a model call for that question. Use lookups to derive category and subcode from the chosen codebook path; uncertainty carries through to those outputs.

Independent questions cannot read one another's answers. To keep operational assessments tied to the chosen code, give them `dependsOn: ["feedback_path"]` and reference `record.Feedback`, `record.type_of_feedback` and `decisions.feedback_path` in their instructions. This requires a second API stage. By default an uncertain dependency blocks its children. Setting `allowReviewDependencies: true` on a dependent question permits an assigned `review` answer to provide context; the child retains its answer and raw confidence but is also flagged for review. Failed, blocked or unassigned parents still block the child. This flag never makes an uncertain parent certain, including on retry.

For feedback containing multiple speech acts, an optional first Choice question can select an exact source excerpt:

```json
{
  "outputColumnName": "feedback_excerpt",
  "questionType": "choice",
  "instructions": "Choose the excerpt matching state.type_of_feedback, or no_match if ambiguous.",
  "options": [],
  "sourceExcerpt": {
    "column": "Feedback",
    "separators": [";", "?", "!", "\n", "."],
    "splitBefore": ["et vous", "et nous"]
  },
  "review": {"minConfidence": 0.7}
}
```

Use structured state and select the named source column. `separators` and `splitBefore` are literal strings, not regular expressions. Separators stay at the end of the preceding excerpt; splitBefore phrases start the next one. Decimal points between digits do not split. Candidates are exact source spans, with only surrounding whitespace excluded. One candidate is carried through without a selection API call or invented confidence; its semantic type match is left to the code question. Multiple candidates are sent as choices plus `no_match`. Missing, ambiguous or oversized selections remain unassigned for review; there is no full-text fallback or silent truncation. Every later question must depend on this first question, directly or indirectly.

Once selected, the source column is replaced with the chosen excerpt in **every downstream provider request**, including retries. Other selected source columns are preserved, so do not select another column containing a duplicate of the full feedback. Original spreadsheet cells remain unchanged. The excerpt output and audit retain the chosen text, selection confidence when applicable, method, and zero-based Unicode code-point start/end offsets (end exclusive). This isolates downstream evidence, but segmentation and model selection can still be wrong: inspect the excerpt when reviewing a decision. Other real safety concerns in the full row are not disproven or resolved by their exclusion from this type-specific excerpt.

`execution.reuseIdenticalRows: true` opts into reuse within a frozen run. Every selected source value participates in the key: select both feedback and its assigned type. IDs participate only if selected as source columns. All original rows and metadata remain in the output. Reused rows keep the original decisions and review flags, and do not double-count API usage. Errors, blocked decisions and unassigned review decisions are not cached across batches; retry mode preserves each row's own successful decisions. Recovery reconstructs the reuse index. The export identifies `reusedFromRow` and reports successful call count, reused row count and summed call durations (including pacing/retries, **not** wall-clock duration under concurrency).

The Ebola configuration builders and workload/regression scripts in `scripts/` use local files in `sample_jev/`. That directory and Excel workbooks are intentionally excluded from Git. Their sample-specific tests skip when the local configuration is absent; generic workflow tests still run. To rebuild a sample configuration, provide the source workbook and configuration at the paths documented by the script.

Retries handle transient network errors and HTTP 429/500/502/503/504/529, honor Retry-After, and stop within a batch time budget. Calls that exceed the remaining budget are marked failed so they can be retried. Lower the batch size for deep workflows. Rate pacing is shared within one backend process and estimates token pressure conservatively from request bytes. Multiple independent server instances sharing an API key still need an account-wide queue for strict quota enforcement.

**Continue unfinished run** processes rows not yet acknowledged. **Retry failed decisions** retries failed/blocked rows, reuses validated successful or uncertain answers, and recomputes derived outputs. An unresolved human-review parent stays blocked; retries do not override its uncertainty. Changing a rubric requires a new run, so old answers cannot silently be mixed with new definitions.

Choice and Score responses must include every option with a finite probability between 0 and 1. Validation allows a 1% total tolerance, including its numeric boundary. When all values use at most two decimal places, it also allows accumulated rounding only if their individual rounding intervals can contain a total of 1. All-zero and impossible totals remain errors. Raw probabilities, confidence and scores are preserved unchanged for auditing. If an older run contains `Jev probabilities do not sum to one`, update/restart the local server, restore the saved run if needed, choose **Retry failed decisions**, and download a fresh workbook. New total-validation errors include the observed total and option count for diagnosis.

Completed batches save the source records, frozen configuration and raw decisions in IndexedDB on this browser, excluding credentials. **Restore saved run** restores that snapshot without making API calls. Only the latest run is saved. **Remove saved checkpoint** deletes it. Download results before changing devices or clearing browser storage. Narrative text is part of the current result/export, not the batch checkpoint. Local storage failures pause further batches and leave completed results available in memory.

For very large datasets, browser memory/storage and the deployment request limits still apply. The table is paginated; it does not render the entire workbook into the DOM. This implementation is a resumable browser-driven workflow, not a durable background processing service.

## Laya and portable classifier selection

The same import/export workflow now supports Laya. New exports may contain
`classifier: {"provider": "laya", "model": "multilingual"}` (or provider `jev`).
Only provider and model belong in this object; credentials stay in Settings.
Existing configurations without it use the classifier currently selected in Settings.
The legacy root `model` field pins Jev only. See [local setup](LOCAL_INFERENCE.md).
