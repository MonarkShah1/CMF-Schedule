"""
Shared-password gate.

The site is on the public internet, so it cannot be wide open. This is a single
team password plus a signed session cookie — proportionate for one shop floor,
and a real barrier rather than security theatre:

  * the cookie is HMAC-signed with SECRET_KEY, so it cannot be forged
  * the password is compared with compare_digest, so it cannot be timed
  * sessions expire, and the cookie is HttpOnly + SameSite

It is deliberately NOT per-user accounts. Names are still recorded against every
action for the audit trail; this only controls who gets in the door. Per-user
logins with roles are worth adding when someone outside the shop needs access.
"""

import base64
import hashlib
import hmac
import json
import os
import time

SESSION_COOKIE = "cmf_session"
SESSION_DAYS = 30

# Set in the hosting dashboard. Empty disables the gate, which is what local
# development wants and what production must never do.
APP_PASSWORD = os.environ.get("APP_PASSWORD", "")
SECRET_KEY = os.environ.get("SECRET_KEY", "")

PUBLIC_PATHS = {"/login", "/health", "/static", "/favicon.ico"}


def enabled() -> bool:
    return bool(APP_PASSWORD)


def _secret() -> bytes:
    # Falling back to the password keeps a misconfigured deploy locked rather
    # than silently unsigned.
    return (SECRET_KEY or APP_PASSWORD or "insecure-dev-only").encode()


def _sign(payload: bytes) -> str:
    sig = hmac.new(_secret(), payload, hashlib.sha256).digest()
    return (base64.urlsafe_b64encode(payload).decode().rstrip("=") + "."
            + base64.urlsafe_b64encode(sig).decode().rstrip("="))


def _unpad(s: str) -> bytes:
    return base64.urlsafe_b64decode(s + "=" * (-len(s) % 4))


def make_session(name: str = "") -> str:
    payload = json.dumps({
        "n": name,
        "exp": int(time.time()) + SESSION_DAYS * 86400,
    }, separators=(",", ":")).encode()
    return _sign(payload)


def read_session(token: str):
    """Returns the session dict, or None if forged, malformed or expired."""
    if not token or "." not in token:
        return None
    body, sig = token.rsplit(".", 1)
    try:
        payload = _unpad(body)
        expected = hmac.new(_secret(), payload, hashlib.sha256).digest()
        if not hmac.compare_digest(_unpad(sig), expected):
            return None
        data = json.loads(payload)
    except (ValueError, TypeError, json.JSONDecodeError):
        return None
    if data.get("exp", 0) < time.time():
        return None
    return data


def password_ok(candidate: str) -> bool:
    return hmac.compare_digest((candidate or "").encode(), APP_PASSWORD.encode())


def is_public(path: str) -> bool:
    return any(path == p or path.startswith(p + "/") for p in PUBLIC_PATHS)
