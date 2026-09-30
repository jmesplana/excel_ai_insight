"""Optional local Laya HTTP integration. No cloud fallback or model downloads here."""
import requests
from jev_provider import JevError, JevConfigError


def validate_configs(configs):
    for question in configs or []:
        options = [question.get('options', [])] + [b.get('options', []) for b in question.get('branches', [])]
        if any(len(values) > 100 for values in options):
            raise JevConfigError('Laya supports up to 100 choices per question. Split this question into stages or use Jev.')


def ask(state, questions, *, config, **kwargs):
    # Send questions separately to stay within Laya's combined 512-option limit.
    answers, usage = {}, {'input_tokens': 0, 'output_tokens': 0}
    with requests.Session() as session:
        session.trust_env = False
        for name, question in questions.items():
            payload = {'model': config['model'], 'state': state, 'questions': {name: question}}
            if config['model'] == 'multilingual':
                payload['max_len'] = 8192
            try:
                response = session.post(config['endpoint'] + '/v1/systemone', json=payload,
                    headers={'Authorization': 'Bearer ' + config['api_key']} if config['api_key'] else {},
                    timeout=180, allow_redirects=False)
                if response.status_code in (401, 403):
                    raise JevError('Laya authentication failed. Check its optional API key.')
                if response.status_code != 200:
                    raise JevError(f'Laya rejected the request (HTTP {response.status_code}). Check the local service.')
                body = response.json()
                if not isinstance(body, dict) or not isinstance(body.get('answers'), dict) or set(body['answers']) != {name}:
                    raise JevError('Laya returned missing or unexpected answers.')
                answers.update(body['answers'])
                for key in usage:
                    value = (body.get('usage') or {}).get(key, 0)
                    if isinstance(value, (int, float)):
                        usage[key] += value
            except (requests.RequestException, ValueError):
                raise JevError('Could not read a response from local Laya. Start laya-serve, check its address and retry unfinished rows.') from None
    return {'model': 'laya:' + config['model'], 'answers': answers, 'usage': usage}
