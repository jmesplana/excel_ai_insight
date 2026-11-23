# Pipeline Studio

Open **Pipeline Studio** in the app navigation, or visit `/pipeline`. It is an ordered workflow editor for repeatable spreadsheet intake: load a source, map its columns, run configurable steps, review exceptions, and download an updated Excel master.

## Start with the feedback workflow

1. In **Choose workflow**, select **Code DRC community feedback**. This loads [`examples/ebola-feedback.pipeline.json`](examples/ebola-feedback.pipeline.json) in one click. It translates all source rows, then classifies the feedback types and aligns each with its original excerpt and English rendering. The **Pipeline classifications** worksheet preserves one row per classified source with corresponding fields joined by ` || `. After every eligible source has been classified or flagged, the workflow splits the validated pairs. Each message keeps its original ID and receives its own codebook path, category, subcode and code. The default provider is Azure; enter a deployment in Connections & models or override it on individual steps.
2. Click **Continue to files**. Upload the team's Excel/CSV file, select its worksheet and header row, and map `Feedback` to the actual source column. Select the original ID column, or choose no existing ID.
3. Optionally upload the current Excel master and select its main worksheet. Master headers must be on row 1. Without a master, the pipeline creates a new workbook.
4. Test a few source rows, inspect the results, then process the file. A test run creates a separate result; a full run replaces the browser's latest checkpoint, so download a test recovery file if you want to keep it.
5. Correct flagged records in the review queue, or discard them. Use **Choose download columns** and **Download completed sheet** for a single Excel worksheet containing accepted messages from this run. **Download audit workbook** includes the cumulative master, review queue, workflow and history. Use that workbook for later master imports.

[`examples/paired-feedback.pipeline.json`](examples/paired-feedback.pipeline.json) contains only generic paired extraction. Add translation, classification, validation or lookup steps for another context. **Append steps from configuration** can add a legacy analysis configuration or the supplied step 2 classification codebook after extraction.

The new module does not change the existing analysis and classification modes.

## Steps and model selection

The everyday screen has three stages: **Choose workflow → Add files → Review results**. Use **Edit steps & models** to open the workflow editor; only one step is expanded at a time. Advanced response settings and JSON rules are available inside a step. Use **Done editing** to return to the workflow summary.

Steps run in their displayed order. Use the arrow buttons to reorder them. Input columns are entered one per line; outputs from an earlier step can be selected as later inputs. The mapping panel maps original workflow input names to the new file's headers. Source columns are preserved; output names cannot overwrite them. A Split step is the explicit exception: it replaces selected list cells with aligned individual values.

| Step | Behavior | Providers |
| --- | --- | --- |
| Extract paired messages | Generates an array of objects, validates each excerpt against its source, and expands it to one row per object | Azure, OpenAI, Ollama |
| Analyze / translate | Generates configured text fields for each row; optional allowed values constrain each field | Azure, OpenAI, Ollama |
| Classify | Selects one allowed label; supports imported source/decision branches | Azure, OpenAI, Ollama, Jev, Laya |
| Split aligned lists | Splits all selected columns using the exact separator; unequal counts or empty entries go to review | No AI |
| Validate fields | Requires populated inputs; optional `rules` restrict allowed values | No AI |
| Look up a derived value | Maps exact combinations of input values to a configured output | No AI |

Each AI step has its own provider, model/deployment, instructions, output fields and token limit. Empty model/deployment fields use that provider's connection settings or server environment. Configure each provider separately; a mixed Azure/Jev workflow sends only the relevant credentials with each request. API keys are not included in exported configurations, checkpoints or recovery files.

Laya uses the existing browser-to-local connector, with its existing origin restrictions. Run `npm run laya -- --origin <app origin>` and set the connector address in Connections & models. Laya is limited to 100 choices per question, so some large Ebola branches require Jev or a text model. Ollama uses the existing server-side local provider.

Cloud text steps use Chat Completions structured output. Choose a compatible Azure deployment/OpenAI model. Advanced `responseMode: "json"` supports models without strict schemas; application validation still applies. Refusals, incomplete output, invalid labels and non-verbatim excerpts cannot enter the accepted master. Output-validation and non-rate-limit transient failures get at most two attempts per text step. For HTTP 429 responses, the browser waits for the provider’s Retry-After delay (including Azure millisecond headers), or uses increasing delays when none is supplied. It pauses after three rate-limit responses; waits longer than a minute also pause for manual continuation. Explicit quota/billing and credential failures pause immediately. Completed steps and cooldowns are saved, so retrying resumes at the failed step. Jev uses its existing bounded retry policy.

Jev/Laya confidence thresholds use their actual classifier outputs. Azure/OpenAI/Ollama do not fabricate comparable confidence scores: they use explicit review outcomes and application validation. `reviewValues` can list labels (such as `Needs review`) that must route to review. The supplied Ebola workflow sets this for its status and no-code fallback. In the DRC translation-first workflow, a translation flagged for uncertain meaning still goes through message extraction and codebook path selection. The preview and review sheet show tentative English, feedback types, paths, and the categories those paths map to. Each split message keeps the translation flag. Retrying a completed translation review reopens its source row so its messages can be classified and split together again. Other review outcomes stop the record at the affected step.

The legacy codebook importer currently supports Choice questions and lookup tables. Score/Noul, sourceExcerpt selection and composite scores remain available in the original classification module. Separate legacy excerpt/type analyses are imported as separate analyses; use a paired-extraction step to fix their alignment problem.

## Pairing and provenance

For `Merci pour votre aide.Où est le centre ?Quand ouvre-t-il ?`, the extraction schema asks for three objects, each containing an excerpt and type. It does not independently generate two delimiter lists. Repeated types are preserved. Application checks validate required fields, allowed types and verbatim substrings; semantic completeness and classification accuracy still need evaluation on reviewed examples.

Each result preserves the source columns and these additional fields:

| Field | Meaning |
| --- | --- |
| `_source_dataset_id` | SHA-256 fingerprint of the uploaded/downloaded file bytes |
| `_source_filename`, `_source_sheet`, `_source_row` | Original file, worksheet and physical Excel row, including header/blank-row offsets |
| `_source_record_id` | Dataset + worksheet + physical row; shared by extracted messages from that source |
| `_original_id` | Copy of the selected original ID; also retained in its original column |
| `_message_number` | Message ordinal within the source; nested expansions use dotted ordinals |
| `_feedback_id` | Unique positive integer allocated for each accepted/review result within this master lineage |
| `_record_key` | Stable expansion key used to avoid duplicate appends |
| `_original_record` | Original source row as JSON, retained when current inputs are corrected |
| `_status`, `_failed_step`, `_reason` | Processing outcome and review information |

Example citation: `Feedback 104582, Beni.xlsx, Sheet1, original Excel row 83, message 2`.

Corrections/retries preserve assigned IDs. If a failed extraction is corrected and then produces multiple messages, the first child retains the original review ID and additional children receive new IDs. Preserve the same latest master when processing successive imports: independent copies of a master cannot coordinate a globally unique numeric sequence.

## Excel output, duplicate prevention and recovery

The downloaded workbook contains the original master worksheet with accepted rows appended, plus **Pipeline review**, **Pipeline imports**, **Pipeline audit**, and **Pipeline workflow**. Other worksheets and existing worksheet cell/formula objects are retained; newly added columns are placed after existing headers. This is a data-processing export using the app's XLSX library, not a lossless editor for every advanced Excel feature. Use a copy of complex workbooks containing macros, charts or external links.

Review and discarded records remain outside accepted data. Pending work is excluded from partial Excel output and retained in the checkpoint/recovery file. An unfinished import is marked in the import log. Reimporting its dataset into that partial master is blocked: restore the recovery file to finish it, rather than risk skipping unfinished sibling messages.

A completed source record already in this master or its review sheet is skipped on reimport. This prevents duplicate imports of the same file bytes; identical text from another source row is preserved. An edited file has a new fingerprint and is treated as a new dataset. Resolve existing review records through the saved run/recovery file rather than expecting a new import to retry them.

Routine autosaves write the changed source and its messages to IndexedDB after each step. The results panel shows workflow progress, active processing time, an estimate for the current pass and screen wake protection status. After an interruption, reopen Pipeline Studio in the same browser, choose **Open saved run**, then **Continue processing**. An interrupted request can be retried from its last saved step. Every 1,000 processed records, the panel offers a downloadable recovery backup. **Saved work → Resume previous run** loads the latest checkpoint; recovery JSON lets you move it to another browser. Recovery files contain source records, master data, configuration and audit history, so store them with the corresponding data. Credentials must be re-entered on another browser.

To change a model, prompt, label list or other settings during review, edit the workflow and click **Use these changes for unfinished records**. The step order, types, inputs and output names must stay the same. Accepted records are retained; retry the affected review records explicitly. Old configurations are preserved in the run's revision history and exported workflow sheet.

This version runs while its browser tab remains open. It supports uploaded files and direct public file URLs with browser CORS access, not scheduled background jobs or authenticated SharePoint connectors. It is an ordered workflow builder, not an arbitrary branching graph or a full n8n replacement. Codebook branches are supported inside classification steps.

## Configuration example

```json
{
  "format": "aidstack-insights-pipeline",
  "version": 1,
  "name": "Translate extracted feedback",
  "idColumn": "ID",
  "mappings": {"feedback_excerpt": "Comment"},
  "nodes": [{
    "id": "translate",
    "name": "English translation",
    "type": "analyze",
    "inputs": ["feedback_excerpt"],
    "provider": "azure",
    "model": "",
    "instructions": "Translate this excerpt faithfully into English. Flag unclear text for review.",
    "fields": [{"name": "translation_en", "options": []}],
    "maxOutputTokens": 4096,
    "responseMode": "schema"
  }]
}
```

Structured output implementation reference: [OpenAI structured outputs](https://developers.openai.com/api/docs/guides/structured-outputs?api-mode=chat).

Excel cells have a 32,767-character limit. If an audit or record value exceeds it, the export stores the complete value in numbered chunks in **Pipeline overflow** and puts a reference in the original cell. Concatenate the chunks for the referenced run and entry ID in part order to recover the exact text; recovery JSON also retains the original value.

Temporary rate limits and quota errors are displayed separately when the provider identifies the cause. A generic 429 is treated as a rate limit, but persistent failures may still require an account or deployment quota check. Review records show the stopped step and the original extraction input; empty output columns do not mean the source record was deleted.


Configured result columns are shown even if every source row stops at extraction. The compact preview names how many result columns are displayed; **Show all data columns** includes every configured result and the source columns. The compact preview also includes columns from steps currently awaiting review. In the preview and the **Pipeline review** worksheet, each stopped output shows `[Needs Review] | …` with the tentative result, or a labeled translation or source excerpt when no tentative result is available. Later steps show **Not run** in the preview and stay blank in Excel. **Retry review records** retries all currently flagged records from their stopped steps, including records outside the 100-row preview. Accepted records keep their results and IDs.

French, Congolese Swahili, Lingala and mixtures are expected in the feedback presets. Language mixing alone is not grounds for review. The model uses the supplied context and local speech patterns to infer each requested output. Extraction can recognize a question even if its referent is uncertain. When the model initially returns a review decision at extraction, the provider gets one more attempt within the two-attempt limit to determine whether the ambiguity actually changes a required output. Unresolved decisions stay in review. Validated excerpt/type pairs are still required for acceptance. **Recheck extraction reviews** retries these records in the saved run, preserving IDs and completed work, without automatically retrying reviewed classification decisions.
