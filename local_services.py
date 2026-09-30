"""Local inference endpoints; browser requests may only target loopback services."""
from urllib.parse import urlsplit


def local_url(value, error_type=ValueError):
    try:
        url = urlsplit(str(value).strip())
        if (url.scheme not in ('http', 'https') or url.hostname not in
                ('127.0.0.1', 'localhost', '::1') or url.username or url.password
                or url.query or url.fragment or url.path not in ('', '/')
                or url.port == 0):
            raise ValueError()
        _ = url.port
    except ValueError:
        raise error_type('Use a local service address such as http://127.0.0.1:11434 (no path).')
    return url.geturl().rstrip('/')
