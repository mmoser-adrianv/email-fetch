from flask import (
    Flask, redirect, render_template, request,
    session, url_for, jsonify, send_file,
)
from flask_session import Session
from werkzeug.middleware.proxy_fix import ProxyFix

import app_config
import auth_helper
import graph_helper

app = Flask(__name__)
app.config.from_object(app_config)
Session(app)
app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1)


@app.route("/")
def index():
    if not session.get("user"):
        return redirect(url_for("login"))
    return render_template("index.html", user=session["user"])


@app.route("/login")
def login():
    flow = auth_helper.build_auth_code_flow(
        redirect_uri=url_for("authorized", _external=True)
    )
    session["auth_flow"] = flow
    return render_template("login.html", auth_uri=flow["auth_uri"])


@app.route("/getAToken")
def authorized():
    try:
        cache = auth_helper.load_cache()
        msal_app = auth_helper.build_msal_app(cache=cache)
        result = msal_app.acquire_token_by_auth_code_flow(
            session.get("auth_flow", {}),
            request.args,
        )
        if "error" in result:
            return render_template("auth_error.html", result=result)
        session["user"] = result.get("id_token_claims")
        auth_helper.save_cache(cache)
    except ValueError:
        pass
    return redirect(url_for("index"))


@app.route("/logout")
def logout():
    session.clear()
    return redirect(
        f"{app_config.AUTHORITY}/oauth2/v2.0/logout"
        f"?post_logout_redirect_uri={url_for('index', _external=True)}"
    )


@app.route("/api/people/search")
def search_people():
    if not session.get("user"):
        return jsonify({"error": "login_required"}), 401
    query = request.args.get("q", "")
    if not query or len(query) < 2:
        return jsonify([])
    token = auth_helper.get_token_from_cache()
    if not token:
        return jsonify({"error": "login_required"}), 401
    results = graph_helper.search_people(token["access_token"], query)
    return jsonify(results)


@app.route("/api/messages")
def get_messages():
    if not session.get("user"):
        return jsonify({"error": "login_required"}), 401
    email = request.args.get("email", "")
    if not email:
        return jsonify({"error": "email parameter required"}), 400
    token = auth_helper.get_token_from_cache()
    if not token:
        return jsonify({"error": "login_required"}), 401
    result = graph_helper.get_user_messages(token["access_token"], email)

    # If there was an error, return it directly
    if "error" in result:
        return jsonify(result)

    # Store full result info in session so the download route can use it
    session["last_results"] = result
    # Return just the messages list to the frontend (backward-compatible)
    return jsonify(result["messages"])


@app.route("/api/messages/download")
def download_messages():
    if not session.get("user"):
        return jsonify({"error": "login_required"}), 401
    result_info = session.get("last_results")
    if not result_info:
        return jsonify({"error": "No messages to download. Fetch emails first."}), 400
    token = auth_helper.get_token_from_cache()
    if not token:
        return jsonify({"error": "login_required"}), 401

    try:
        zip_buffer = graph_helper.download_emails_zip(
            token["access_token"], result_info
        )
        return send_file(
            zip_buffer,
            mimetype="application/zip",
            as_attachment=True,
            download_name="emails.zip",
        )
    except Exception as e:
        return jsonify({"error": f"Download failed: {str(e)}"}), 500


if __name__ == "__main__":
    app.run(host="localhost", port=5000, debug=True)
