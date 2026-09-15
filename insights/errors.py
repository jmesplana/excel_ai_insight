"""Public error messages never include provider responses or credentials."""
from llm_provider import LLMConfigError
from openai import AuthenticationError, RateLimitError, APIConnectionError


class AnalysisOutputError(ValueError):
    """Safe, application-generated validation message."""


def public_error(error):
    if isinstance(error, AnalysisOutputError):
        return str(error)
    if isinstance(error, LLMConfigError):
        return 'AI provider configuration is incomplete. Check API Settings.'
    if isinstance(error, AuthenticationError):
        return 'Authentication failed. Check API Settings.'
    if isinstance(error, RateLimitError):
        return 'The provider is busy or its usage limit was reached. Retry later.'
    if isinstance(error, APIConnectionError):
        return 'Could not reach the AI provider. Check your connection and retry.'
    return 'The request could not be completed. Check your settings and retry.'
