"""Email magic-link sign-in for Pipeline Studio.

Sign-in is on only when an approved list is configured (AUTH_ALLOWED_EMAILS and/or
AUTH_ALLOWED_DOMAINS). Then Pipeline Studio and its API require a signed-in, approved
email, and the server's AI credentials (environment variables) are available only to
signed-in users, anywhere in the app. Keys a person enters in AI connections still take
precedence. Without a list the app behaves as before.

Links are stateless: a signed token that expires after LINK_MINUTES. Nothing is
stored on the server, so a link can be reused until it expires (this also keeps
mail scanners that prefetch links from using it up).
"""
import os
import re
import smtplib
from datetime import timedelta
from email.message import EmailMessage
from urllib.parse import urlsplit

from flask import (Blueprint, current_app, has_request_context, jsonify, redirect,
                   render_template, request, session, url_for)
from itsdangerous import BadSignature, SignatureExpired, URLSafeTimedSerializer

bp = Blueprint('auth', __name__)
LINK_MINUTES = 15
SESSION_DAYS = 7
EMAIL_PATTERN = re.compile(r'^[^@\s]{1,64}@[^@\s]{1,255}\.[^@\s]{2,}$')
NOT_APPROVED = 'This email address is not on the approved list.'


def _items(name):
    return {part.strip().lower() for part in os.environ.get(name, '').split(',') if part.strip()}


def auth_enabled():
    return bool(_items('AUTH_ALLOWED_EMAILS') or _items('AUTH_ALLOWED_DOMAINS'))


def normalize(email):
    return str(email or '').strip().lower()


def is_allowed(email):
    email = normalize(email)
    if not EMAIL_PATTERN.match(email):
        return False
    # Exact domain match: sub.ifrc.org or ifrc.org.example.com are not ifrc.org.
    return email in _items('AUTH_ALLOWED_EMAILS') or email.rsplit('@', 1)[1] in _items('AUTH_ALLOWED_DOMAINS')


def current_user():
    """The signed-in email, if it is still approved. Removing someone from the list signs them out."""
    email = session.get('email')
    return email if email and is_allowed(email) else None


def server_credentials_allowed():
    """Whether this request may use the server's AI credentials."""
    if not auth_enabled() or not has_request_context():
        return True
    return current_user() is not None


def protected(path):
    return path == '/pipeline' or path.startswith('/api/pipeline/')


def safe_next(target):
    target = str(target or '')
    parts = urlsplit(target)
    if not target.startswith('/') or target.startswith('//') or '\\' in target or parts.scheme or parts.netloc:
        return '/pipeline'
    return target


def _serializer():
    return URLSafeTimedSerializer(current_app.secret_key, salt='pipeline-magic-link')


def _base_url():
    # Set APP_BASE_URL in production so links never depend on the request's Host header.
    return (os.environ.get('APP_BASE_URL') or request.url_root).rstrip('/')


def send_link(email, url):
    host = os.environ.get('EMAIL_SERVER_HOST')
    sender = os.environ.get('EMAIL_FROM')
    if not host or not sender:
        raise RuntimeError('Email sending is not configured (EMAIL_SERVER_HOST and EMAIL_FROM).')
    port = int(os.environ.get('EMAIL_SERVER_PORT') or 587)
    message = EmailMessage()
    message['Subject'] = 'Sign in to Aidstack Insights Pipeline Studio'
    message['From'] = sender
    message['To'] = email
    message.set_content(render_template('email/magic-link.txt', url=url, minutes=LINK_MINUTES))
    message.add_alternative(render_template('email/magic-link.html', url=url, minutes=LINK_MINUTES), subtype='html')
    smtp = smtplib.SMTP_SSL(host, port, timeout=15) if port == 465 else smtplib.SMTP(host, port, timeout=15)
    with smtp:
        if port != 465:
            smtp.starttls()
        user, password = os.environ.get('EMAIL_SERVER_USER'), os.environ.get('EMAIL_SERVER_PASSWORD')
        if user and password:
            smtp.login(user, password)
        smtp.send_message(message)


def init_app(app):
    app.secret_key = os.environ.get('SECRET_KEY') or app.secret_key
    app.config.update(
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SAMESITE='Lax',
        SESSION_COOKIE_SECURE=bool(os.environ.get('VERCEL')) or os.environ.get('SESSION_COOKIE_SECURE') == '1',
        PERMANENT_SESSION_LIFETIME=timedelta(days=SESSION_DAYS),
    )
    app.register_blueprint(bp)

    @app.before_request
    def require_sign_in():
        if not auth_enabled() or not protected(request.path):
            return None
        if not current_app.secret_key:
            return jsonify(error='Sign-in is enabled but SECRET_KEY is not set.'), 503
        if current_user():
            return None
        if request.path.startswith('/api/'):
            return jsonify(error='Your sign-in has expired. Sign in again to continue.', errorCode='signin'), 401
        return redirect(url_for('auth.login', next=request.full_path.rstrip('?')))

    @app.context_processor
    def auth_context():
        return {'auth_enabled': auth_enabled(), 'signed_in_email': current_user() if auth_enabled() else None}


@bp.route('/login', methods=['GET', 'POST'])
def login():
    next_path = safe_next(request.values.get('next'))
    if not auth_enabled():
        return redirect(next_path)
    if request.method == 'GET':
        if current_user():
            return redirect(next_path)
        error = {'expired': 'That sign-in link has expired. Request a new one.',
                 'invalid': 'That sign-in link is not valid. Request a new one.'}.get(request.args.get('error'))
        return render_template('login.html', next_path=next_path, error=error, sent=False, email='')
    email = normalize(request.form.get('email'))
    if not EMAIL_PATTERN.match(email):
        return render_template('login.html', next_path=next_path, error='Enter a valid email address.', sent=False, email=email), 400
    if not is_allowed(email):
        return render_template('login.html', next_path=next_path, error=NOT_APPROVED, sent=False, email=email), 403
    if not current_app.secret_key:
        return render_template('login.html', next_path=next_path, error='Sign-in is not configured yet (SECRET_KEY).', sent=False, email=email), 503
    token = _serializer().dumps({'email': email, 'next': next_path})
    try:
        send_link(email, f"{_base_url()}{url_for('auth.verify')}?token={token}")
    except Exception:
        current_app.logger.exception('Could not send the sign-in email')
        return render_template('login.html', next_path=next_path, error='We could not send the email. Try again in a few minutes.', sent=False, email=email), 502
    return render_template('login.html', next_path=next_path, error=None, sent=True, email=email)


@bp.route('/auth/verify')
def verify():
    if not auth_enabled() or not current_app.secret_key:
        return redirect('/')
    try:
        data = _serializer().loads(request.args.get('token', ''), max_age=LINK_MINUTES * 60)
    except SignatureExpired:
        return redirect(url_for('auth.login', error='expired'))
    except BadSignature:
        return redirect(url_for('auth.login', error='invalid'))
    email = normalize(data.get('email') if isinstance(data, dict) else '')
    if not is_allowed(email):
        return render_template('login.html', next_path='/pipeline', error=NOT_APPROVED, sent=False, email=email), 403
    session.clear()
    session.permanent = True
    session['email'] = email
    return redirect(safe_next(data.get('next')))


@bp.route('/logout', methods=['POST'])
def logout():
    session.clear()
    return redirect('/')

