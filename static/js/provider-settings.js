import { classifierFetch } from './classifier-transport.js';
/** Independent provider profiles. Only the selected profile travels with a request. */
export const API_KEY_STORAGE_KEY = 'excel_ai_insight_openai_api_key';
export const PROVIDER_STORAGE_KEY = 'excel_ai_insight_provider';
export const AZURE_API_KEY_STORAGE_KEY = 'excel_ai_insight_azure_api_key';
export const AZURE_ENDPOINT_STORAGE_KEY = 'excel_ai_insight_azure_endpoint';
export const AZURE_DEPLOYMENT_STORAGE_KEY = 'excel_ai_insight_azure_deployment';
export const AZURE_API_VERSION_STORAGE_KEY = 'excel_ai_insight_azure_api_version';
export const OPENAI_MODEL_STORAGE_KEY = 'excel_ai_insight_openai_model';
export const JEV_API_KEY_STORAGE_KEY = 'excel_ai_insight_jev_api_key';
export const JEV_MODEL_STORAGE_KEY = 'excel_ai_insight_jev_model';
const CLASSIFIER = 'excel_ai_insight_classifier';
const REMEMBER = 'excel_ai_insight_remember_settings';
const fields = {
    'modal-api-key': API_KEY_STORAGE_KEY,
    'modal-openai-model': OPENAI_MODEL_STORAGE_KEY,
    'modal-azure-api-key': AZURE_API_KEY_STORAGE_KEY,
    'modal-azure-endpoint': AZURE_ENDPOINT_STORAGE_KEY,
    'modal-azure-deployment': AZURE_DEPLOYMENT_STORAGE_KEY,
    'modal-azure-api-version': AZURE_API_VERSION_STORAGE_KEY,
    'modal-jev-api-key': JEV_API_KEY_STORAGE_KEY,
    'modal-jev-model': JEV_MODEL_STORAGE_KEY,
    'modal-ollama-endpoint': 'excel_ai_insight_ollama_endpoint',
    'modal-ollama-model': 'excel_ai_insight_ollama_model',
    'modal-laya-endpoint': 'excel_ai_insight_laya_connector_endpoint',
    'modal-laya-model': 'excel_ai_insight_laya_model',
    'modal-laya-api-key': 'excel_ai_insight_laya_api_key',
    'modal-icd-client-id': 'excel_ai_insight_icd_client_id',
    'modal-icd-client-secret': 'excel_ai_insight_icd_client_secret'
};
const field = id => {
    const element = document.getElementById(id);
    return (element ? element.value : localStorage.getItem(fields[id]))?.trim() || null;
};
const selected = (name, key, fallback) => document.querySelector(`input[name="${name}"]:checked`)?.value
    || localStorage.getItem(key) || fallback;

export function getLLMConfig() {
    const provider = selected('llm-provider', PROVIDER_STORAGE_KEY, 'openai');
    if (provider === 'ollama') return {provider, ollamaEndpoint: field('modal-ollama-endpoint'), ollamaModel: field('modal-ollama-model')};
    if (provider === 'azure') return {provider, apiKey: field('modal-azure-api-key'),
        azureEndpoint: field('modal-azure-endpoint'), azureDeployment: field('modal-azure-deployment'),
        azureApiVersion: field('modal-azure-api-version')};
    return {provider, apiKey: field('modal-api-key'), model: field('modal-openai-model')};
}
export function hasLocalCredentials() {
    const c = getLLMConfig();
    return c.provider === 'ollama' ? !!c.ollamaModel : !!(c.apiKey && (c.provider !== 'azure' || (c.azureEndpoint && c.azureDeployment)));
}
export function getJevConfig() {
    const classificationProvider = selected('classification-provider', CLASSIFIER, 'jev');
    if (classificationProvider === 'laya') return {classificationProvider,
        layaEndpoint: field('modal-laya-endpoint'), layaModel: field('modal-laya-model') || 'multilingual',
        layaApiKey: field('modal-laya-api-key')};
    return {classificationProvider, jevApiKey: field('modal-jev-api-key'), jevModel: field('modal-jev-model')};
}
export function hasJevCredentials() {
    const c = getJevConfig();
    return c.classificationProvider === 'laya' || !!c.jevApiKey;
}
export function syncProviderUI() {
    const llm = getLLMConfig().provider, classifier = getJevConfig().classificationProvider;
    document.querySelectorAll('[data-llm-panel]').forEach(el => el.hidden = el.dataset.llmPanel !== llm);
    document.querySelectorAll('[data-classifier-panel]').forEach(el => el.hidden = el.dataset.classifierPanel !== classifier);
    document.querySelectorAll('[data-classifier-label]').forEach(el => el.textContent = classifier === 'laya' ? 'Laya · local' : 'Jev · cloud');
}
export function restoreSettings() {
    // The previous shared key belongs to the provider selected when it was saved.
    const legacy = localStorage.getItem('excel_ai_insight_api_key');
    if (legacy) {
        const key = localStorage.getItem(PROVIDER_STORAGE_KEY) === 'azure' ? AZURE_API_KEY_STORAGE_KEY : API_KEY_STORAGE_KEY;
        if (!localStorage.getItem(key)) localStorage.setItem(key, legacy);
        localStorage.removeItem('excel_ai_insight_api_key');
    }
    for (const [id, key] of Object.entries(fields)) {
        const el = document.getElementById(id);
        if (el && localStorage.getItem(key) !== null) el.value = localStorage.getItem(key);
    }
    const provider = localStorage.getItem(PROVIDER_STORAGE_KEY) || 'openai';
    const classifier = localStorage.getItem(CLASSIFIER) || 'jev';
    const llmRadio = document.getElementById(`provider-${provider}`);
    if (llmRadio) llmRadio.checked = true;
    const classRadio = document.getElementById(`classifier-${classifier}`);
    if (classRadio) classRadio.checked = true;
    document.getElementById('modal-save-api-key').checked = localStorage.getItem(REMEMBER) === 'true'
        || Object.values(fields).some(key => !!localStorage.getItem(key));
    syncProviderUI();
}
export function saveSettings() {
    const remember = document.getElementById('modal-save-api-key').checked;
    for (const [id, key] of Object.entries(fields)) {
        const value = field(id);
        if (remember && value) localStorage.setItem(key, value);
        else localStorage.removeItem(key);
    }
    for (const [key, value] of [[PROVIDER_STORAGE_KEY, getLLMConfig().provider], [CLASSIFIER, getJevConfig().classificationProvider]]) {
        if (remember) localStorage.setItem(key, value);
        else localStorage.removeItem(key);
    }
    localStorage.setItem(REMEMBER, String(remember));
    syncProviderUI();
}
export function initProviderSettings() {
    restoreSettings();
    document.getElementById('laya-start-command').textContent = `npm run laya -- --origin ${location.origin}`;
    const modal = new bootstrap.Modal(document.getElementById('api-settings-modal'));
    document.getElementById('api-settings-btn').onclick = () => modal.show();
    document.querySelectorAll('input[name="llm-provider"], input[name="classification-provider"]').forEach(el => el.onchange = syncProviderUI);
    document.getElementById('save-api-settings').onclick = () => {
        try { saveSettings(); modal.hide(); }
        catch { document.getElementById('settings-status').textContent = 'Could not save browser settings. Current fields remain available for this session.'; }
    };
    document.querySelectorAll('[data-reveal]').forEach(button => button.onclick = () => {
        const input = document.getElementById(button.dataset.reveal);
        input.type = input.type === 'password' ? 'text' : 'password';
        button.textContent = input.type === 'password' ? 'Show' : 'Hide';
    });
    for (const [buttonId, resultId, endpoint, config] of [
        ['test-connection-btn', 'test-connection-result', '/test_connection', getLLMConfig],
        ['test-jev-connection-btn', 'test-jev-connection-result', '/test_jev_connection', getJevConfig]
    ]) {
        const button = document.getElementById(buttonId), out = document.getElementById(resultId);
        button.onclick = async () => {
            button.disabled = true; out.textContent = 'Testing connection…';
            try {
                const payload = config();
                const response = endpoint === '/test_jev_connection'
                    ? await classifierFetch(endpoint, payload)
                    : await fetch(endpoint, {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(payload)});
                const result = await response.json();
                out.textContent = response.ok && result.ok ? `Connected to ${result.provider} (${result.model}).` : result.error || 'Connection failed.';
            } catch (error) { out.textContent = error.message || 'Could not connect. Check that the app and selected service are running.'; }
            finally { button.disabled = false; }
        };
    }
    document.getElementById('ollama-models-btn').onclick = async event => {
        const button = event.currentTarget, out = document.getElementById('ollama-models-status');
        button.disabled = true; out.textContent = 'Looking for installed models…';
        try {
            const response = await fetch('/ollama_models', {method:'POST', headers:{'Content-Type':'application/json'},
                body:JSON.stringify({ollamaEndpoint:field('modal-ollama-endpoint')})});
            const result = await response.json();
            if (!response.ok) throw new Error(result.error);
            const list = document.getElementById('ollama-model-options'); list.replaceChildren();
            result.models.forEach(name => { const option = document.createElement('option'); option.value = name; list.append(option); });
            if (!field('modal-ollama-model') && result.models.length) document.getElementById('modal-ollama-model').value = result.models[0];
            out.textContent = result.models.length ? `${result.models.length} installed model(s). Choose one in the Model field.` : 'No models installed. Download a model with Ollama first.';
        } catch (e) { out.textContent = e.message; }
        finally { button.disabled = false; }
    };
}
