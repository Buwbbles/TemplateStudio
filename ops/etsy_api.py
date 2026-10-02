"""Small Etsy Open API v3 helper for Template Studio (shop Starnetup 68377150).
Reads the key from ETSY_API_KEY (keystring:secret) and the OAuth token from .etsy/token.json.
Refreshes the token when it is older than 50 minutes. Never prints secrets."""
import json, os, subprocess, time, urllib.parse, urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TOKEN = os.path.join(ROOT, '.etsy', 'token.json')
SHOP = 68377150
BASE = 'https://openapi.etsy.com/v3/application'
KEY = os.environ['ETSY_API_KEY']


def _tok():
    t = json.load(open(TOKEN))
    if time.time() - t.get('obtained_at', 0) > 3000:
        data = urllib.parse.urlencode({'grant_type': 'refresh_token', 'client_id': KEY.split(':')[0],
                                       'refresh_token': t['refresh_token']}).encode()
        r = json.load(urllib.request.urlopen(urllib.request.Request(
            'https://api.etsy.com/v3/public/oauth/token', data=data)))
        t.update(r)
        t['obtained_at'] = int(time.time())
        with open(TOKEN, 'w') as f:
            json.dump(t, f)
        os.chmod(TOKEN, 0o600)
    return t['access_token']


def call(method, path, form=None, files=None):
    """form: list of (k, v) pairs (form-encoded); files: list of (field, path) for multipart."""
    cmd = ['curl', '-s', '-X', method, BASE + path, '-H', 'x-api-key: ' + KEY,
           '-H', 'Authorization: Bearer ' + _tok(), '-w', '\n%{http_code}']
    if files:
        for k, v in (form or []):
            cmd += ['-F', f'{k}={v}']
        for k, p in files:
            cmd += ['-F', f'{k}=@{p}']
    elif form is not None:
        for k, v in form:
            cmd += ['--data-urlencode', f'{k}={v}']
    out = subprocess.run(cmd, capture_output=True, text=True).stdout
    body, _, code = out.rpartition('\n')
    try:
        body = json.loads(body)
    except Exception:
        pass
    return int(code or 0), body
