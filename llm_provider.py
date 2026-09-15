"""
LLM provider abstraction.

The app supports two backends that both speak the OpenAI chat-completions API:

  * "openai" -- api.openai.com, keyed by a plain API key. The model name is
    sent as-is (e.g. "gpt-4o-mini").
  * "azure"  -- an Azure AI Foundry / Azure OpenAI resource, keyed by an
    endpoint + key + API version. On Azure the "model" argument is really the
    *deployment name* your organization chose, which often differs from the
    underlying model name.

Configuration is resolved per request with this precedence:

    1. Values the browser sent in the request body (the Settings modal).
    2. Environment variables (.env), for a shared org-wide deployment.

That ordering lets an individual override a server default without one being
required, and lets the server supply credentials so users never type a key.
"""

import os

from openai import OpenAI, AzureOpenAI

# Default model used when neither the request nor the environment names one.
# On Azure this is a deployment name, so it is far more likely to be overridden.
DEFAULT_OPENAI_MODEL = "gpt-4o-mini"

# Azure requires an explicit API version. This one supports the chat
# completions and streaming features the app relies on.
DEFAULT_AZURE_API_VERSION = "2024-10-21"


class LLMConfigError(ValueError):
    """Raised when the resolved provider configuration is unusable."""


def _clean(value):
    """Normalize a request/env value to a non-empty string or None."""
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def resolve_config(data=None):
    """
    Resolve the effective LLM configuration for one request.

    Args:
        data: the parsed JSON request body (or None). Recognized keys are
            provider, apiKey, azureEndpoint, azureDeployment, azureApiVersion
            and model.

    Returns:
        A dict with keys: provider, api_key, model, endpoint, api_version.

    Raises:
        LLMConfigError: if required values are missing for the provider.
    """
    data = data or {}

    provider = (
        _clean(data.get('provider'))
        or _clean(os.environ.get('LLM_PROVIDER'))
        or 'openai'
    ).lower()

    if provider not in ('openai', 'azure'):
        raise LLMConfigError(
            f"Unknown provider '{provider}'. Expected 'openai' or 'azure'."
        )

    if provider == 'azure':
        api_key = (
            _clean(data.get('apiKey'))
            or _clean(os.environ.get('AZURE_OPENAI_API_KEY'))
        )
        endpoint = (
            _clean(data.get('azureEndpoint'))
            or _clean(os.environ.get('AZURE_OPENAI_ENDPOINT'))
        )
        model = (
            _clean(data.get('azureDeployment'))
            or _clean(data.get('model'))
            or _clean(os.environ.get('AZURE_OPENAI_DEPLOYMENT'))
        )
        api_version = (
            _clean(data.get('azureApiVersion'))
            or _clean(os.environ.get('AZURE_OPENAI_API_VERSION'))
            or DEFAULT_AZURE_API_VERSION
        )

        missing = []
        if not api_key:
            missing.append("API key")
        if not endpoint:
            missing.append("endpoint")
        if not model:
            missing.append("deployment name")
        if missing:
            raise LLMConfigError(
                "Azure AI Foundry is selected but the "
                + ", ".join(missing)
                + " is missing. Add it in Settings or set the matching "
                  "AZURE_OPENAI_* environment variable."
            )

        # The SDK builds request URLs by appending to the endpoint, so a
        # trailing slash would produce a double slash on some gateways.
        endpoint = endpoint.rstrip('/')

        return {
            "provider": "azure",
            "api_key": api_key,
            "model": model,
            "endpoint": endpoint,
            "api_version": api_version,
        }

    api_key = (
        _clean(data.get('apiKey'))
        or _clean(os.environ.get('OPENAI_API_KEY'))
    )
    if not api_key:
        raise LLMConfigError(
            "Missing OpenAI API key. Add it in Settings or set OPENAI_API_KEY."
        )

    model = (
        _clean(data.get('model'))
        or _clean(os.environ.get('OPENAI_MODEL'))
        or DEFAULT_OPENAI_MODEL
    )

    return {
        "provider": "openai",
        "api_key": api_key,
        "model": model,
        "endpoint": None,
        "api_version": None,
    }


def build_client(config):
    """Construct the SDK client for a resolved config."""
    if config["provider"] == "azure":
        return AzureOpenAI(
            api_key=config["api_key"],
            azure_endpoint=config["endpoint"],
            api_version=config["api_version"],
        )
    return OpenAI(api_key=config["api_key"])


def get_client_and_model(data=None):
    """
    Resolve config and build a client in one step.

    Returns:
        (client, model) where model is the model name on OpenAI or the
        deployment name on Azure -- in both cases the value to pass as the
        `model` argument of chat.completions.create().
    """
    config = resolve_config(data)
    return build_client(config), config["model"]


def provider_label(config_or_data=None):
    """Human-readable provider name, for error messages shown to the user."""
    provider = (config_or_data or {}).get('provider') or 'openai'
    return "Azure AI Foundry" if provider == 'azure' else "OpenAI"
