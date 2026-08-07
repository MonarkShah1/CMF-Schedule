"""Auth regression — the gate that stands between your shop data and the internet."""

import importlib
import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(
    os.path.abspath(__file__)))))


@pytest.fixture
def secured(monkeypatch):
    """Reload the app with a password configured."""
    monkeypatch.setenv("APP_PASSWORD", "shop-floor-secret")
    monkeypatch.setenv("SECRET_KEY", "test-signing-key")
    from v2.app import auth as auth_mod
    importlib.reload(auth_mod)
    from v2.app import main as main_mod
    importlib.reload(main_mod)
    from fastapi.testclient import TestClient
    return TestClient(main_mod.app), auth_mod


@pytest.fixture(autouse=True)
def _restore():
    yield
    os.environ.pop("APP_PASSWORD", None)
    os.environ.pop("SECRET_KEY", None)
    from v2.app import auth as auth_mod
    importlib.reload(auth_mod)
    from v2.app import main as main_mod
    importlib.reload(main_mod)


def test_pages_redirect_to_login_when_signed_out(secured):
    client, _ = secured
    for url in ["/", "/board", "/purchasing", "/orders/new"]:
        r = client.get(url, follow_redirects=False)
        assert r.status_code == 303, url
        assert "/login" in r.headers["location"], url


def test_health_stays_public(secured):
    """The host needs an unauthenticated health check to know the app is up."""
    client, _ = secured
    assert client.get("/health").status_code == 200


def test_correct_password_signs_in(secured):
    client, _ = secured
    r = client.post("/login", data={"password": "shop-floor-secret", "name": "Sam"},
                    follow_redirects=False)
    assert r.status_code == 303
    assert client.get("/board").status_code == 200


def test_wrong_password_is_refused(secured):
    client, _ = secured
    r = client.post("/login", data={"password": "guess"}, follow_redirects=False)
    assert "error=1" in r.headers["location"]
    assert client.get("/", follow_redirects=False).status_code == 303


def test_forged_cookie_is_rejected(secured):
    client, auth = secured
    client.cookies.set(auth.SESSION_COOKIE, "eyJuIjoiaGFja2VyIn0.notarealsignature")
    assert client.get("/", follow_redirects=False).status_code == 303


def test_tampered_payload_is_rejected(secured):
    """Flipping the payload must invalidate the signature."""
    client, auth = secured
    good = auth.make_session("Sam")
    body, sig = good.rsplit(".", 1)
    client.cookies.set(auth.SESSION_COOKIE, body[:-2] + "XY." + sig)
    assert client.get("/", follow_redirects=False).status_code == 303


def test_expired_session_is_rejected(secured):
    client, auth = secured
    import base64, json, time
    payload = json.dumps({"n": "Sam", "exp": int(time.time()) - 10},
                         separators=(",", ":")).encode()
    client.cookies.set(auth.SESSION_COOKIE, auth._sign(payload))
    assert client.get("/", follow_redirects=False).status_code == 303


def test_open_redirect_is_blocked(secured):
    """?next=//evil.com must not bounce the user off-site after login."""
    client, _ = secured
    r = client.post("/login", data={"password": "shop-floor-secret",
                                    "next": "//evil.example.com"},
                    follow_redirects=False)
    assert r.headers["location"] == "/"


def test_logout_clears_the_session(secured):
    client, _ = secured
    client.post("/login", data={"password": "shop-floor-secret"},
                follow_redirects=False)
    assert client.get("/board").status_code == 200
    client.post("/logout", follow_redirects=False)
    assert client.get("/", follow_redirects=False).status_code == 303


def test_no_password_configured_leaves_the_app_open():
    """Local development runs without a password; production must set one."""
    os.environ.pop("APP_PASSWORD", None)
    from v2.app import auth as auth_mod
    importlib.reload(auth_mod)
    assert auth_mod.enabled() is False
