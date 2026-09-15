"""Application entry point; feature routes live in insights/."""
import os
from flask import Flask, render_template, jsonify
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass


def create_app():
    app = Flask(__name__)
    app.config['MAX_CONTENT_LENGTH'] = 8 * 1024 * 1024
    from insights.analysis import bp as analysis
    from insights.chat import bp as chat
    from insights.icd import bp as icd
    for blueprint in (analysis, chat, icd):
        app.register_blueprint(blueprint)

    @app.errorhandler(413)
    def too_large(error):
        return jsonify(error='Request too large. Use fewer rows or smaller cells.'), 413

    @app.route('/')
    def index():
        return render_template('index.html')

    @app.route('/about')
    def about():
        return render_template('about.html')

    @app.route('/how-it-works')
    def how_it_works():
        return render_template('how-it-works.html')

    return app


app = create_app()

if __name__ == '__main__':
    # On WSL the working copy lives on the Windows mount (/mnt/c), where inotify
    # does not fire — the default reloader silently misses edits. Force the
    # polling ("stat") reloader so code changes are always picked up.
    port = int(os.environ.get("PORT", 8080))
    host = os.environ.get("HOST", "127.0.0.1")
    debug = os.environ.get("FLASK_DEBUG", "1") not in ("0", "false", "False")
    app.run(host=host, port=port, debug=debug, reloader_type="stat")
