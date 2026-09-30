"""Browser-to-local Laya connector. Run alongside laya-serve, never on Vercel."""
import os
from urllib.parse import urlsplit
from flask import Flask, request, jsonify
from local_services import local_url

DEFAULT_ORIGINS = ('http://127.0.0.1:8080', 'http://localhost:8080')


def allowed_origins(value):
    result = set(DEFAULT_ORIGINS)
    for origin in value.split(','):
        origin = origin.strip().rstrip('/')
        if not origin:
            continue
        url = urlsplit(origin)
        if (url.scheme not in ('https', 'http') or not url.hostname or url.path
                or url.query or url.fragment or url.username or url.password
                or '*' in origin or (url.scheme == 'http' and url.hostname not in ('localhost', '127.0.0.1', '::1'))):
            raise ValueError('Allowed sites must be exact HTTPS origins, or HTTP localhost origins.')
        result.add(origin)
    return result


def create_bridge(origins=None, laya_endpoint=None):
    from insights.jev import bp
    app = Flask(__name__)
    app.config['MAX_CONTENT_LENGTH'] = 8 * 1024 * 1024
    sites = allowed_origins(origins if origins is not None else os.environ.get('LAYA_ALLOWED_ORIGINS', ''))
    endpoint = local_url(laya_endpoint or os.environ.get('LAYA_BASE_URL', 'http://127.0.0.1:8000'))
    app.register_blueprint(bp)
    paths = {'/test_jev_connection', '/validate_jev_config', '/analyze_batch_jev'}

    @app.before_request
    def authorize():
        if request.host.split(':')[0] not in ('localhost', '127.0.0.1'):
            return jsonify(error='Invalid local connector host.'), 403
        origin = request.headers.get('Origin')
        if request.path == '/health' and request.method == 'GET' and not origin:
            return None
        if origin not in sites:
            return jsonify(error='This website is not allowed. Restart the connector with --origin and your exact website URL.'), 403
        if request.path == '/health' and request.method == 'GET':
            return None
        if request.path not in paths:
            return jsonify(error='Unknown local connector route.'), 404
        if request.method == 'OPTIONS':
            if request.headers.get('Access-Control-Request-Method') != 'POST':
                return jsonify(error='Only JSON POST requests are supported.'), 405
            return '', 204
        if request.method != 'POST' or not request.is_json:
            return jsonify(error='Use a JSON POST request.'), 415
        data = request.get_json(silent=True)
        if not isinstance(data, dict):
            return jsonify(error='Expected a JSON object.'), 400
        # The connector only runs Laya against the operator's local service.
        # Never let a browser select Jev or redirect data to a different host.
        data['classificationProvider'] = 'laya'
        data['layaEndpoint'] = endpoint
        if request.path == '/analyze_batch_jev' and (not isinstance(data.get('rows'), list) or len(data['rows']) != 1):
            return jsonify(error='The local connector accepts one row per batch.'), 400

    @app.after_request
    def cors(response):
        origin = request.headers.get('Origin')
        if origin in sites:
            response.headers['Access-Control-Allow-Origin'] = origin
            response.headers['Access-Control-Allow-Methods'] = 'POST, GET, OPTIONS'
            response.headers['Access-Control-Allow-Headers'] = 'Content-Type'
            response.headers['Access-Control-Allow-Private-Network'] = 'true'
            response.headers['Access-Control-Max-Age'] = '600'
        response.headers['Vary'] = 'Origin'
        response.headers['Cache-Control'] = 'no-store'
        return response

    @app.route('/health')
    def health():
        return jsonify(ok=True, service='aidstack-laya-connector', version=1)

    @app.errorhandler(413)
    def oversized(error):
        return jsonify(error='Request too large for the local connector.'), 413

    return app


if __name__ == '__main__':
    port = int(os.environ.get('LAYA_BRIDGE_PORT', '8001'))
    print(f'Browser connector: http://127.0.0.1:{port}', flush=True)
    print('Allowed websites: ' + ', '.join(sorted(allowed_origins(os.environ.get('LAYA_ALLOWED_ORIGINS', '')))), flush=True)
    create_bridge().run(host='127.0.0.1', port=port, debug=False, use_reloader=False)
