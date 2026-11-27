# Pipeline Studio

Open **Pipeline Studio** in the app navigation, or visit `/pipeline`. It is an ordered workflow editor for repeatable spreadsheet intake: load a source, map its columns, run configurable steps, review exceptions, and download an updated Excel master.

## Start with the feedback workflow

1. In **Choose workflow**, pick **DRC Ebola framework** from the preset list. This loads [`presets/drc-ebola-framework.pipeline.json`](presets/drc-ebola-framework.pipeline.json). It translates all source rows, then classifies the feedback types and aligns each with its original excerpt and English rendering. The **Pipeline classifications** worksheet preserves one row per classified source with corresponding fields joined by ` || `. After every eligible source has been classified or flagged, the workflow splits the validated pairs. Each message keeps its original ID and receives its own codebook path, category, subcode and code. The default provider is Azure; enter a deployment in Connections & models or override it on individual steps.
2. Click **Continue to files**. Upload the team's Excel/CSV file, select its worksheet and header row, and map `Feedback` to the actual source column. Select the original ID column, or choose no existing ID.
3. Optionally upload the current Excel master and select its main worksheet. Master headers must be on row 1. Without a master, the pipeline creates a new workbook.
4. Test a few source rows, inspect the results, then process the file. A test run creates a separate result; a full run replaces the browser's latest checkpoint, so download a test recovery file if you want to keep it.
5. Review flagged records (see [Reviewing flagged records](#reviewing-flagged-records)), or discard them. Use **Choose download columns** and **Download completed sheet** for a single Excel worksheet containing accepted messages from this run. Under **Feedback records**, choose **All records** or a status filter and click **Download filtered records** to save every match as Excel or CSV. The final **Pipeline status** column includes each record’s status at download time. **Download audit workbook** includes the cumulative master, review queue, workflow and history. Use that workbook for later master imports.

The **DRC feedback messages** preset ([`presets/drc-feedback-messages.pipeline.json`](presets/drc-feedback-messages.pipeline.json)) only separates messages and pairs them with feedback types. [`examples/paired-feedback.pipeline.json`](examples/paired-feedback.pipeline.json) is an older-format example of the same extraction. Add translation, classification, validation or lookup steps for another context. **Append steps from configuration** can add a legacy analysis configuration or the supplied step 2 classification codebook after extraction.

The new module does not change the existing analysis and classification modes.

## Sign-in and server credentials

Set `AUTH_ALLOWED_EMAILS` and/or `AUTH_ALLOWED_DOMAINS` (comma-separated; domains match exactly) to require sign-in. Opening Pipeline Studio then asks for an email address. An approved address receives a sign-in link that works for 15 minutes; any other address is told it is not on the approved list. A session lasts 7 days, and removing an address from the list signs it out on its next request.

While sign-in is on, the server's AI credentials (`AZURE_OPENAI_*`, `OPENAI_API_KEY`, `TYPESAFE_API_KEY`) are available only to signed-in users, anywhere in the app. A key entered in **AI connections** takes precedence. A provider left blank there uses the server's connection; a browser-supplied endpoint is only ever used with a browser-supplied key. Sign-in also needs `SECRET_KEY`, `APP_BASE_URL` and the SMTP settings `EMAIL_SERVER_HOST`, `EMAIL_SERVER_PORT`, `EMAIL_SERVER_USER`, `EMAIL_SERVER_PASSWORD` and `EMAIL_FROM` (see `.env.example`). Links are stateless: one can be reused until it expires. The rest of the app is unaffected.

## Presets and workflow files

Every `*.pipeline.json` file in the `presets` folder appears in the preset list, using its `name` and `description`. To add a framework, save a workflow with **Saved work → Save workflow as a file**, add a `description`, and put the file in `presets`. No code change is needed.

A workflow file contains everything the workflow does. The server applies no hidden rules based on language, place names or column names. The context, step guidance and review behaviour of the DRC presets are written in their files and can be edited in **Edit steps & models**:

| Setting | Where | Effect |
| --- | --- | --- |
| `context` | Workflow (**Context & language notes**) | Setting, languages and local terms, added to the system instructions of each AI step that uses context |
| `useContext` | Step (checkbox) | `false` leaves the workflow context out of this step. Default `true` |
| `systemNotes` | Step (**Model guidance**) | Step-specific guidance added after the context |
| `recheck` | Step (**Model guidance**) | When the model asks for review, it gets one more attempt with this instruction. Extraction steps have a built-in recheck unless this is set to `""` |
| `recheckWithReason` | Step (advanced JSON) | The recheck also receives the first review reason |
| `interpretFirst` | Step (advanced JSON) | The model writes a short plain-meaning sentence (`plainMeaning`) before deciding |
| `valuesBeforeStatus` | Step (advanced JSON) | The model gives its answer before the ok/review decision |
| `fallbackOption` | Classify step (advanced JSON) | A “no code fits” label: choosing it is a review outcome, the recheck excludes it, and it is never written as a result |
| `allowSkip` | Step (checkbox) | The model may answer “skip” when the text is not language at all (random typing). The record becomes **Skipped** and can be processed anyway. Off by default, including in the presets; turning it on changes that step's prompt, so check it on a test batch |
| `screening` | Workflow (JSON) | Rules applied before any AI step: `requireLetters`, `minLetters`, `randomTyping`, optional `columns` (default: the workflow's input columns). See [Skipped and discarded records](#skipped-and-discarded-records) |
| `tentativeOnReview` | Step (advanced JSON) | When this step is flagged, the record continues through the next steps tentatively and stays in review |
| `roles` | Workflow (JSON) | Names the columns that tentative review, ID preservation on retry and the **Pipeline classifications** sheet use: `sourceTranslation`, `messageTranslation`, `excerpt`, `messageType`, `code`, `category`, `tentativeColumns`. Without roles these features are off |

Workflow files are `version: 2`. Version 1 files and saved runs still load. Before version 2, the server recognised DRC wording (“Congolese Swahili”, “in the DRC”) and DRC column names. A version 1 workflow containing such steps is upgraded on import or restore: those steps receive the settings of the equivalent step in the DRC Ebola framework preset, so they send the same requests as before.

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

## Reviewing flagged records

Click **Review** on a flagged record. The dialog offers two ways forward:

- **Ask the AI again.** Write a **Note for the model**, for example “‘maso’ here means ‘words’, not ‘eyes’”, and click **Save & retry with AI**. The note goes to the model as trusted reviewer guidance, separately from the feedback text, so it is never translated, coded or quoted as an excerpt. It stays with the record and the messages split from it for every later step, and is exported in `_reviewer_note`. **Correct the input text** is only for real errors in the text, such as a transcription mistake: extracted excerpts must match the text used, although the original row is always kept in `_original_record`. Jev and Laya steps do not receive notes.
- **Enter the result yourself.** For translation/analysis, classification and lookup steps, type or choose the answer and click **Accept my answer**. Choices are limited to the step's allowed values, including the codebook branch for this record; a `fallbackOption` is not offered. The record continues with the next steps automatically. The step is recorded as a reviewer decision: `_reviewer_decisions` lists the steps a person answered, and the audit history shows `reviewer` as the provider. For a flagged DRC translation, entering the English reopens the source, and its messages are extracted and coded again; messages found again with the same excerpt and type keep their IDs. Extraction and split steps cannot be answered this way: add a note or correct the input.

## Skipped and discarded records

**Skipped** records never reached a decision, because they cannot be feedback. The DRC presets screen every row before any AI request, without cost:

- no letters at all: empty cells, numbers, symbols (`12345`, `---`);
- fewer than `minLetters` letters (default 2);
- every word looks like random typing: no vowels in 5+ letters (`sdfgh`), one repeated letter (`jjjj`), a repeated 2–3 letter chunk (`sdfsdf`, `hahaha`), or 5 keys in a row on a QWERTY or AZERTY keyboard (`qwert`, `azert`).

The rules are deliberately narrow, so real Lingala, Swahili and French text passes (including `polepole`, `malamu malamu`, abbreviations such as `MSF`). Short random text such as `kdjfa` still goes to the workflow; switch on `allowSkip` for a step to let the model skip it, or discard it in review. Skipped records show their reason in `_skip_reason`, are counted separately, and are listed on the review sheet. **Process anyway** sends a skipped record through the workflow.

**Discarded** records were rejected by a person. Use **Discard** on a flagged record, or tick several records and use **Discard selected**. Choose a reason (test or random typing, no meaningful content, cannot be translated, duplicate, or your own). When a source row was split into several messages, you can discard all of its flagged messages at once. The reason is kept in `_discard_reason` and the audit history. **Restore** returns a discarded record to review.

Skipped and discarded records stay out of the accepted master. Because they are on the review sheet, importing the same file into that master again does not bring them back.

## Speed

Pipeline Studio sends several requests at once: 4 by default, adjustable from 1 to 8 in **Run options** (remembered in this browser). Laya and Ollama always take one request at a time. Records still move through the workflow in the same phases, and each record goes through the same steps with the same inputs as when running one at a time. Only the order in which new feedback IDs are assigned can differ. After a rate-limit, quota or credential error, no new requests are sent; requests already sent finish and are saved. If a provider keeps asking you to slow down, lower **Requests at a time**.

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
