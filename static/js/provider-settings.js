/** Provider configuration uses browser values first, then server defaults. */
export const API_KEY_STORAGE_KEY = 'excel_ai_insight_api_key';
export const PROVIDER_STORAGE_KEY = 'excel_ai_insight_provider';
export const AZURE_ENDPOINT_STORAGE_KEY = 'excel_ai_insight_azure_endpoint';
export const AZURE_DEPLOYMENT_STORAGE_KEY = 'excel_ai_insight_azure_deployment';
export const AZURE_API_VERSION_STORAGE_KEY = 'excel_ai_insight_azure_api_version';
export const OPENAI_MODEL_STORAGE_KEY = 'excel_ai_insight_openai_model';

export function getLLMConfig() {
    const value = (id, key) => document.getElementById(id)?.value?.trim()
        || localStorage.getItem(key)?.trim() || null;
    const provider = document.querySelector('input[name="llm-provider"]:checked')?.value
        || localStorage.getItem(PROVIDER_STORAGE_KEY) || 'openai';
    const config = { provider, apiKey: value('modal-api-key', API_KEY_STORAGE_KEY) };
    if (provider === 'azure') {
        config.azureEndpoint = value('modal-azure-endpoint', AZURE_ENDPOINT_STORAGE_KEY);
        config.azureDeployment = value('modal-azure-deployment', AZURE_DEPLOYMENT_STORAGE_KEY);
        config.azureApiVersion = value('modal-azure-api-version', AZURE_API_VERSION_STORAGE_KEY);
    } else {
        config.model = value('modal-openai-model', OPENAI_MODEL_STORAGE_KEY);
    }
    return config;
}

export function hasLocalCredentials() {
    const config = getLLMConfig();
    return !!(config.apiKey && (config.provider !== 'azure'
        || (config.azureEndpoint && config.azureDeployment)));
}

export function syncProviderUI() {
    const isAzure = document.querySelector('input[name="llm-provider"]:checked')?.value === 'azure';
    document.getElementById('azure-settings')?.classList.toggle('d-none', !isAzure);
    document.getElementById('openai-model-settings')?.classList.toggle('d-none', isAzure);
    document.querySelectorAll('.provider-azure-only').forEach(el => el.classList.toggle('d-none', !isAzure));
    document.querySelectorAll('.provider-openai-only').forEach(el => el.classList.toggle('d-none', isAzure));
}
