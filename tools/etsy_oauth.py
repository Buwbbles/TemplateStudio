#!/usr/bin/env python3
"""One-time Etsy OAuth (PKCE) sign-in for Template Studio.

Run in the background. The Commander opens http://localhost:3003/start in a
browser, approves on Etsy, and the token is saved to ../.etsy/token.json (0600).
No key or token is ever printed.

Prerequisite: the Etsy app "starnet" must list this exact Callback URL:
    http://localhost:3003/oauth/redirect
"""
import base64, hashlib, http.server, json, os, secrets, time, urllib.parse, urllib.request

PORT = 3003
REDIRECT = f"http://localhost:{PORT}/oauth/redirect"
SCOPES = "listings_r listings_w listings_d shops_r shops_w transactions_r profile_r email_r"
CLIENT_ID = os.environ["ETSY_API_KEY"].split(":", 1)[0]
OUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".etsy")
TOKEN_PATH = os.path.join(OUT_DIR, "token.json")

verifier = base64.urlsafe_b64encode(secrets.token_bytes(48)).rstrip(b"=").decode()
challenge = base64.urlsafe_b64encode(hashlib.sha256(verifier.encode()).digest()).rstrip(b"=").decode()
state = secrets.token_urlsafe(16)
AUTH_URL = "https://www.etsy.com/oauth/connect?" + urllib.parse.urlencode({
    "response_type": "code", "redirect_uri": REDIRECT, "scope": SCOPES,
    "client_id": CLIENT_ID, "state": state,
    "code_challenge": challenge, "code_challenge_method": "S256"})
DONE = {"ok": False}


def exchange(code):
    body = urllib.parse.urlencode({"grant_type": "authorization_code", "client_id": CLIENT_ID,
                                   "redirect_uri": REDIRECT, "code": code,
                                   "code_verifier": verifier}).encode()
    req = urllib.request.Request("https://api.etsy.com/v3/public/oauth/token", data=body,
                                 headers={"Content-Type": "application/x-www-form-urlencoded"})
    with urllib.request.urlopen(req, timeout=30) as r:
        tok = json.load(r)
    tok["obtained_at"] = int(time.time())
    os.makedirs(OUT_DIR, exist_ok=True)
    with open(TOKEN_PATH, "w") as f:
        json.dump(tok, f)
    os.chmod(TOKEN_PATH, 0o600)


class H(http.server.BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _send(self, code, text, loc=None):
        self.send_response(code)
        if loc:
            self.send_header("Location", loc)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write(text.encode())

    def do_GET(self):
        u = urllib.parse.urlparse(self.path)
        if u.path == "/start":
            return self._send(302, "", AUTH_URL)
        if u.path == "/oauth/redirect":
            q = urllib.parse.parse_qs(u.query)
            if q.get("state", [""])[0] != state:
                return self._send(400, "<h2>State mismatch. Open /start again.</h2>")
            if "error" in q:
                print("ETSY_OAUTH_ERROR", q.get("error"), q.get("error_description"), flush=True)
                return self._send(400, f"<h2>Etsy said: {q.get('error_description', q['error'])[0]}</h2>")
            try:
                exchange(q["code"][0])
            except Exception as e:  # report type only, never secrets
                print("ETSY_OAUTH_EXCHANGE_FAILED", type(e).__name__, getattr(e, "code", ""), flush=True)
                return self._send(500, "<h2>Token exchange failed. Tell Lucille.</h2>")
            DONE["ok"] = True
            print("ETSY_OAUTH_OK token saved", flush=True)
            return self._send(200, "<h2>Done. Template Studio is connected to your Etsy shop. You can close this tab.</h2>")
        self._send(404, "<h2>Open /start</h2>")


if __name__ == "__main__":
    srv = http.server.HTTPServer(("127.0.0.1", PORT), H)
    srv.timeout = 5
    print(f"listening on http://localhost:{PORT}/start", flush=True)
    end = time.time() + 6 * 3600
    while not DONE["ok"] and time.time() < end:
        srv.handle_request()
    print("exiting", "ok" if DONE["ok"] else "timeout", flush=True)
