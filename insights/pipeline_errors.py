"""Safe provider diagnostics for resumable pipeline work; never echo error bodies."""
import math
import time
from datetime import timezone
from email.utils import parsedate_to_datetime
from openai import RateLimitError, AuthenticationError
from .errors import public_error


def retry_delay(headers, now=None):
    now = time.time() if now is None else now
    delays = []
    for name, scale in [('retry-after-ms', .001), ('x-ms-retry-after-ms', .001), ('retry-after', 1)]:
        value = headers.get(name)
        if value is None:
            continue
        try:
            delay = float(value) * scale
        except (TypeError, ValueError):
            if name != 'retry-after':
                continue
            try:
                date = parsedate_to_datetime(value)
                if date.tzinfo is None:
                    date = date.replace(tzinfo=timezone.utc)
                delay = date.timestamp() - now
            except (TypeError, ValueError, OverflowError):
                continue
        if math.isfinite(delay) and delay >= 0:
            delays.append(delay)
    return max(delays) if delays else None


def provider_failure(exc):
    if isinstance(exc, RateLimitError):
        body = exc.body if isinstance(exc.body, dict) else {}
        error = body.get('error', body)
        if not isinstance(error, dict):
            error = {}
        codes = {str(error.get(k, '')).lower() for k in ('code', 'type')}
        if codes & {'insufficient_quota', 'quota_exceeded', 'billing_hard_limit_reached', 'billing_not_active', 'usage_limit_reached', 'insufficient_balance'}:
            return {'status': 'error', 'errorCode': 'quota_exhausted', 'retryable': False,
                    'reason': 'The provider reports an exhausted quota or billing limit. Check the account or deployment quota, or choose another configured model before retrying.'}
        headers = getattr(getattr(exc, 'response', None), 'headers', {}) or {}
        return {'status': 'error', 'errorCode': 'rate_limit', 'retryable': True,
                'retryAfterSeconds': retry_delay(headers),
                'reason': 'The provider returned HTTP 429 (too many requests or tokens). The pipeline will wait before retrying this step. If it persists, check the deployment limits and account quota.'}
    if isinstance(exc, AuthenticationError):
        return {'status': 'error', 'errorCode': 'authentication', 'retryable': False,
                'reason': 'The provider rejected these credentials. Check AI connections before retrying.'}
    return {'status': 'error', 'errorCode': 'provider_error', 'reason': public_error(exc)}
