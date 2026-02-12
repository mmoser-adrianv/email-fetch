import msal
from flask import session
import app_config


def build_msal_app(cache=None):
    return msal.ConfidentialClientApplication(
        app_config.CLIENT_ID,
        authority=app_config.AUTHORITY,
        client_credential=app_config.CLIENT_SECRET,
        token_cache=cache,
    )


def load_cache():
    cache = msal.SerializableTokenCache()
    if session.get("token_cache"):
        cache.deserialize(session["token_cache"])
    return cache


def save_cache(cache):
    if cache.has_state_changed:
        session["token_cache"] = cache.serialize()


def build_auth_code_flow(scopes=None, redirect_uri=None):
    cache = load_cache()
    app = build_msal_app(cache=cache)
    flow = app.initiate_auth_code_flow(
        scopes or app_config.SCOPE,
        redirect_uri=redirect_uri,
        domain_hint="mmoser.com",
    )
    save_cache(cache)
    return flow


def get_token_from_cache(scopes=None):
    cache = load_cache()
    app = build_msal_app(cache=cache)
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(
            scopes or app_config.SCOPE,
            account=accounts[0],
        )
        save_cache(cache)
        return result
    return None
