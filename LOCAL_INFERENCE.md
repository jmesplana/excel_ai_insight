# Local classification and text generation

Jev, OpenAI and Azure Foundry remain available. Laya adds local classification;
Ollama adds local analysis, written reports and chat. Laya connects directly from
your browser to a local connector, including when the website is hosted on Vercel.
Ollama still uses the Flask backend, so Ollama requires the app backend to run on
the same computer. There is no automatic cloud fallback.

## Ollama

Start Ollama, then open **API Settings → Text generation → Ollama · local**.
Click **Find installed models**, choose a downloaded model, and test the connection.
The default address is `http://127.0.0.1:11434` (do not append `/v1`).
Use an instruction/chat model that supports your task; structured analysis requires
JSON output. The app uses Ollama's OpenAI-compatible completions and streaming API.

## Laya

Python 3.10+ is required. The optional service uses its own `.venv-laya` environment,
so its PyTorch/Transformers dependencies do not change the Flask environment.

```sh
npm run setup
npm run setup:laya
npm run laya -- --origin https://YOUR-APP.vercel.app
```

The first service start downloads and loads the multilingual checkpoint. Wait for
startup to complete while online. Then open **API Settings → Classification →
Laya · local**, select **Multilingual**, and test the connection. The default
browser connector address is **`http://127.0.0.1:8001`**, prefilled in Settings.
The underlying model service uses port 8000; do not enter that raw model address
in browser Settings. No key is needed unless you set `LAYA_API_KEY`
when starting the service; that optional key has its own settings field.

For subsequent starts with cached model files and no network:

```sh
npm run laya:offline
```

Keep that terminal open and use the app on Vercel. You do not need to run the
website locally. Settings shows the exact startup command for the current website;
copy it instead of the example URL above. The allowed website is remembered in
`.laya-origins.json` (not committed). To use a different domain, restart with its
`--origin`. Exact origins are required; arbitrary websites and wildcard domains
are not permitted. `LAYA_ALLOWED_ORIGINS` can specify a comma-separated list.
Local app origins on localhost/127.0.0.1 port 8080 are also allowed.

Allow local-network access if your browser prompts. Browser policies vary; use an
up-to-date browser that permits HTTPS pages to access loopback services. Do not
disable browser security. Connection failures show setup/permission guidance.
The connector handles CORS and private-network preflights and only exposes Laya
classification routes; it cannot proxy requests to cloud providers. Both services
bind only to loopback.

For fully offline use, also start the website locally with `npm run dev` in another
terminal. The hosted website itself is not cached as a PWA.

Set `LAYA_PYTHON` during setup to choose a
Python interpreter, `LAYA_PORT` to change the model port, or `LAYA_BRIDGE_PORT`
to change the browser connector port (then update Settings). To preload other checkpoints,
set `LAYA_MODELS=english,multilingual,typed-decisions` before the initial online
start; download every checkpoint you intend to use before going offline.

Laya is a separate model, with different accuracy and confidence behavior. Compare
it with manually reviewed rows before reusing Jev review thresholds. The integration
uses one row at a time and at most 100 choices per question (including branches).
Multilingual requests use an 8192-token budget; shorter checkpoint budgets and
model tokenization can still truncate long input/instructions. Large codebooks
should use staged questions. Classifier failures remain retryable errors, with no
request to Jev. The existing Jev probability decoder validates both providers.

## Configuration import and export

Use the same **Import Configuration** and **Export Configuration** controls in
**Classification (Jev / Laya)**. Existing Jev v1/v2 configurations still work.
Questions, source columns, review rules, dependencies, branches, derived columns,
report settings and execution settings round-trip as before.

New exports also include a `classifier` object, for example:

```json
{"classifier": {"provider": "laya", "model": "multilingual"}}
```

Import restores that selection. Change it in Settings to reuse the same codebook
with another provider. Legacy `model` remains a Jev model pin and is ignored for
Laya. API keys and local service addresses are not exported. Checkpoints freeze the
run's classifier and local service address; retries/resumes do not switch providers
when Settings changes. Keys are read from the matching current profile.

## Saving settings and offline scope

OpenAI, Azure, Jev, optional Laya authentication, Ollama, and WHO each have separate
fields. **Save all settings** saves every profile, including hidden tabs, when
**Remember all provider settings** is checked. Clearing a field removes its saved
value. Unchecking Remember removes saved profiles on Save while leaving the current
page's values usable until reload. Server environment defaults still apply to empty
fields. Old shared keys are moved into the provider selected when they were saved;
an already-overwritten key cannot be recovered.

All frontend scripts, styles, icons and fonts are bundled locally. Once dependencies
and models have been downloaded, spreadsheet import, local inference, statistical
reports and exports work without internet. Cloud providers and WHO ICD lookup
still require internet.

For the backend / local connector, server defaults are: `LLM_PROVIDER=ollama`, `OLLAMA_BASE_URL`, `OLLAMA_MODEL`,
`CLASSIFICATION_PROVIDER=laya`, `LAYA_BASE_URL`, `LAYA_MODEL`, `LAYA_API_KEY`.
Browser selections take precedence. Service addresses must be loopback HTTP(S)
URLs without paths. Imported codebooks do not override addresses or credentials.

References: [Laya](https://github.com/NandhaKishorM/laya),
[Ollama API compatibility](https://docs.ollama.com/api/openai-compatibility).

## Verification

The integration was checked with 40 backend tests, 31 frontend unit tests, and 23
browser tests, including independent saved credentials, legacy-key migration,
configuration round trips, and a browser session that blocks external requests.
Live checks used Ollama `gemma3:1b` for completion and streamed reporting, and Laya
0.3.21 multilingual for choice, score, Noul, structured instructions and dependency
stages on synthetic French text. Laya passed again after restarting with
`HF_HUB_OFFLINE=1` and `TRANSFORMERS_OFFLINE=1`. These checks verify integration;
they are not an accuracy evaluation of a production codebook.

The browser-to-local connector is also tested from a simulated HTTPS website,
using a real local connector and a deterministic model server. The test grants
the browser local-network permission, exercises real CORS, verifies failed-row
retry and XLSX export, and asserts no classification POST reaches the hosted site.
An actual Vercel deployment has not been made by this change.
