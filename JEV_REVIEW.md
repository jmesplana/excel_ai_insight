# JEV integration review and dataset report design

Reviewed 20 September 2026 against TypeSafe's official documentation and the local application. This records the pre-change audit. The configurable implementation that followed is documented in [JEV_CONFIGURATION.md](JEV_CONFIGURATION.md); findings below describe the original behavior.

## Conclusion

The row classification architecture is appropriate for Jev. The app supports all three primitives and sends the questions for each row together. It does not yet provide a reliable full-dataset findings report or use all the useful information Jev returns.

Keep Jev for semantic judgments. Calculate dataset statistics in code. For an original written report, use the existing OpenAI/Azure integration over verified aggregates and explicitly selected evidence. A Jev-only report can still contain exact tables, charts, and template-based sentences, without another model.

TypeSafe explicitly documents unreliable counting and the absence of free-text generation. Sending a whole workbook to Jev and asking it to summarize or count it is therefore the wrong implementation. [Official limitations](https://docs.typesafe.ai/model-jaggedness/jev-1.13)

## What already works

- Correct direct HTTP endpoint, bearer authentication, and `model` / `state` / `questions` request shape in `jev_provider.py`.
- Choice for a category, Score for an ordered rubric, Noul for a yes/no probability. Choice and Score option-count checks exist.
- One request carries all questions for one row. Eight workers process different rows concurrently; the browser sends batches of 25 rows.
- Selected source fields are labeled in the row context. Empty input skips the provider call.
- Optional confidence and raw Score/Noul values can be exported. The UI correctly explains that confidence is not the probability of correctness.
- Question configuration import/export, test runs, stop/resume within the current browser session, and Excel export already exist.

These core request patterns match the [quick start](https://docs.typesafe.ai/introduction/quickstart) and [primitive guidance](https://docs.typesafe.ai/primitives).

## Findings, in priority order

### 1. Existing chat cannot support a claim of whole-dataset coverage

`static/js/app.js:48–74` caps chat input at the first 5,000 rows, including after classification. `insights/chat.py:130–164` gives the text model the first five non-null examples per column, numeric summaries, and a keyword-selected analysis result, then truncates the assembled context to 8,000 characters. Later column details and computed results can be cut off.

Consequences: rows after 5,000 cannot inform chat; early examples are not representative thematic evidence; categorical distributions are not systematically included. A generic request for a summary does not fix these limits. Quick Insights in `static/js/results.js` contains row/column counts, missingness, unique values and three example values, rather than an overall classification report.

**Recommended change:** build an aggregate report from every result row before any context reduction. Always display selected sheet, total input rows, processed rows, successful rows, empty rows, failed rows and unprocessed rows. A test run or interrupted run must say it is partial. The current processing scope is one selected sheet, not automatically the entire workbook.

### 2. Transient errors become completed rows

`jev_provider.ask` makes one attempt. A 429, overload, or network failure becomes an error cell in `insights/jev.py`. The response still contains that row, so `BatchRun` advances its cursor. Resume retries unacknowledged batches, not failed rows inside acknowledged batches. A run can therefore finish with missing classifications.

Eight concurrent workers do not enforce a requests-per-minute limit. The configured per-call timeout is 60 seconds, while `vercel.json` gives the entire server invocation 60 seconds; multiple waves of requests can exceed that deadline. Checkpoints are in memory and disappear on page reload.

**Recommended change:** bounded retries with backoff and Retry-After handling, explicit per-row/per-question status, retry-failed-rows without rerunning successes, request/token pacing, and persisted checkpoints for large jobs. Keep backend work within the deployment deadline, or use a background job design for sustained workloads. The official SDK already provides retry behavior, but adopting it is optional; direct HTTP is supported. [API error and retry guidance](https://docs.typesafe.ai/api)

### 3. Useful probability information is discarded

`decode_answer` keeps only the winning Choice label. Score becomes a rounded level label, with its fractional value retained only when the optional confidence setting is enabled. Noul becomes Yes/No at 0.5, again retaining its probability only with that setting. Full distributions, actual returned model ID and token usage are discarded by the provider path.

**Recommended change:** always retain raw answers and provenance internally; let the checkbox control visible/exported columns only. Report category counts, continuous score summaries with their rubric, uncertainty counts and a review queue. Treat a rounded Score as an application policy, not the API's winning category. A fractional score is a rubric position, not a physical measurement.

Confidence describes the concentration of Choice/Score distributions. Noul has no separate confidence; values near 0.5 are ambiguous. Review thresholds should be configurable and evaluated on labeled data, not interpreted as guaranteed accuracy. [Confidence documentation](https://docs.typesafe.ai/confidence)

### 4. Framework 2 dependencies are not implemented

`examples/ebola-framework2.jev-config.json` asks Type and Topic independently. Topic instructions nevertheless require the topic to exist for “the chosen type.” Questions in one request cannot see each other's answers, so that wording cannot enforce the relationship. [Question independence and dependent requests](https://docs.typesafe.ai/primitives)

The repository contains `framework2_topic_grid.json` and a documented taxonomy of 444 unique leaf labels, but the running app does not load the grid, derive the sous-dimension, validate missing combinations, or classify leaf codes. The current configuration returns Type and Topic plus five operational judgments. Parts of `FRAMEWORK2_CODING.md` describe a lookup step that is not wired into the app.

**Recommended change:** first add deterministic lookup and flag invalid combinations. If detailed framework coding is required, add a second request with only the leaf options under the validated branch. If no valid branch exists, route to review or an explicit branch-resolution step. Dependent requests can be integrated into the existing per-row worker and checkpoint system; they do not inherently require abandoning batching. Compare independent Type/Topic classification with a single Choice over valid branches using labeled examples before selecting a permanent approach.

### 5. Malformed responses can silently produce plausible results

Local reproduction showed that an unknown Choice key is accepted, a missing Choice becomes `None`, and Noul `1.5` becomes Yes. Out-of-range Scores are clamped to the nearest endpoint. These are adapter validation gaps, not evidence that Jev has returned such responses in production.

**Recommended change:** validate response type, required answer fields, allowed keys, finite numeric ranges and distributions. Surface invalid responses as explicit failures rather than ordinary labels. Also reserve generated audit-column names: current output naming resolves main columns against source names but does not reserve all generated `__confidence` and `__score` names, allowing collisions.

## Capabilities worth exposing

| Capability | Current support | Useful next step |
| --- | --- | --- |
| Choice, Score, Noul | All available | Keep clear, narrow rubrics |
| Multiple questions per state | Implemented | Preserve this architecture |
| Probability distributions | Discarded | Retain for uncertainty and alternative labels |
| Noul decision thresholds | Fixed at 0.5 | Add yes / review / no policy per question |
| Separate labels and descriptions | Labels double as descriptions | Add definitions, exclusions and examples |
| Structured state and criteria | UI flattens to strings | Offer structured JSON configuration when needed |
| Multi-label themes | Possible through multiple Nouls | Ask one independent presence question per theme |
| Hierarchical classification | Not implemented | Narrow later questions using earlier answers |
| Composite scoring | Not implemented | Combine normalized scores with explicit weights in code |
| Model provenance and usage | Discarded | Save actual version and input-token usage |
| Review / escalation | Explanatory UI only | Add actionable review queue and corrections |

Structured instructions, option descriptions, rubric levels and Noul true/false criteria are supported by the API. They can express definitions more clearly than a bare label list. [Structured questions](https://docs.typesafe.ai/primitives/advanced)

A Choice selects one dominant label; it does not measure all topics present. If a row mentions both vaccination and access to treatment, independent Nouls can capture both. Such percentages may sum above 100%, and the report must explain that. Use a generative model or human review to propose new themes outside the codebook, then validate and run those questions over all rows before claiming their frequency.

Not every documented pattern needs a UI feature. Routing, reranking, verification and extraction are compositions of the same three primitives. Add them when a workflow needs them. [Patterns](https://docs.typesafe.ai/patterns)

## Concrete full-dataset report design

1. **Classify and preserve:** source row ID, question configuration/version, selected fields, raw answers, returned model, status and usage. Never confuse an error or empty input with a category.
2. **Aggregate all rows in code:** count each category; show percentages with explicit valid-response denominators; summarize raw scores; count uncertain decisions; build Type × Topic tables and optional geography/date breakdowns using source columns. Count failed answers separately for each question because a row can be partly successful.
3. **Display automatically at completion:** coverage banner, principal category distributions, score distributions, review counts and source-row links. A Jev-only user can use all of this without LLM credentials.
4. **Offer “Generate findings report”:** send the existing text provider a compact evidence packet containing exact aggregates, question definitions, coverage and selected examples. Display which provider generates the narrative. Any subset of examples must be labeled as examples; it must not determine dataset-wide counts.
5. **Ground the writing:** require supporting counts and source IDs, distinguish reported beliefs from verified events, identify missing context, and label recommendations as interpretation. Do not claim to have read every free-text row when only examples are supplied. For semantic coverage beyond the codebook, process every text row in bounded chunks and retain evidence references before synthesis.
6. **Export together:** analyzed rows, summary tables, review rows, report text and run metadata. Preserve original row IDs if cleaning splits a row into multiple records; otherwise totals may count duplicated source feedback.

Suggested report contents: coverage and data quality; leading types/topics; operational urgency and sensitivity; requested actions; differences by location/time where available; uncertain or inconsistent classifications; evidence examples; interpretation and suggested follow-up.

For probability-weighted reporting, an optional expected category count is the sum of that category's probabilities across valid rows. Label it as an estimate, separately from hard-label counts. Its quality depends on calibration on this dataset. It is not a confidence interval or a verified occurrence count.

## Model and scale considerations

As checked on the review date, TypeSafe lists `jev-1.13.0`, 64k total tokens per request and 32k for state plus the longest question, with published limits of 1,200 requests/minute and 250,000 tokens/second. Limits may change. Current pricing is $0.042 per million input tokens, with free output tokens. The app's 20-question limit is its own limit, not an established API question-count ceiling.

Pin a version for reproducible evaluations and log the actual response version. TypeSafe documents English as its strongest language; evaluate the French codebook against representative French feedback. Customer-specific fine-tuning is not offered; adaptation is through state, instructions and criteria. [Models and limits](https://docs.typesafe.ai/models)

## Validation and rollout

Run a human-reviewed sample covering frequent classes, rare classes, ambiguous/multi-topic feedback, empty records and relevant languages. Measure per-class precision/recall, confusion, review rate and error rate; tune thresholds on held-out examples. Keep source labels out of input if they are the ground truth being evaluated. Schema guarantees alone do not establish semantic accuracy.

Before large uploads, prioritize full-data aggregation, response validation, retained probabilities, failed-row recovery and rate/deadline handling. Then add the written report and any required hierarchical or multi-label coding.

Local verification: 27 frontend unit tests, 13 backend tests and all 3 Jev browser tests passed. Additional decoder probes confirmed the validation issues above. Provider responses are mocked in these tests; they do not establish live Jev accuracy, quota behavior or large-file throughput. No live classification of user data was performed for this review.
