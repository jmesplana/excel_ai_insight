"""Public error messages never include provider responses or credentials."""
from llm_provider import LLMConfigError
from jev_provider import JevConfigError, JevError
from openai import AuthenticationError, RateLimitError, APIConnectionError


class AnalysisOutputError(ValueError):
    """Safe, application-generated validation message."""


def public_error(error):
    if isinstance(error, AnalysisOutputError):
        return str(error)
    # Jev config errors describe the user's own question setup ("a Choice needs
    # at least 2 options"), so they are safe -- and necessary -- to show as-is.
    if isinstance(error, JevConfigError):
        return str(error)
    # JevError messages are built by jev_provider from status codes and the
    # API's own `error` field; they never carry the key or the request body.
    if isinstance(error, JevError):
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
