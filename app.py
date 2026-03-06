import os

from flask import (
    Flask, redirect, render_template, request,
    session, url_for, jsonify, send_file, Response,
)
from flask_session import Session
from werkzeug.middleware.proxy_fix import ProxyFix

import app_config
import auth_helper
import graph_helper
from flask_migrate import Migrate
from models import db, save_email_to_db

app = Flask(__name__)
app.config.from_object(app_config)
Session(app)
app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1)

db.init_app(app)
Migrate(app, db)


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


@app.route("/api/messages/download/progress")
def download_progress():
    """SSE endpoint that streams progress while building the ZIP."""
    if not session.get("user"):
        return jsonify({"error": "login_required"}), 401
    result_info = session.get("last_results")
    if not result_info:
        return jsonify({"error": "No messages to download. Fetch emails first."}), 400
    token = auth_helper.get_token_from_cache()
    if not token:
        return jsonify({"error": "login_required"}), 401

    return Response(
        graph_helper.download_emails_zip_progress(
            token["access_token"], result_info
        ),
        mimetype="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
        },
    )


@app.route("/api/messages/download/file")
def download_file():
    """Serve a completed temp ZIP file and clean it up."""
    if not session.get("user"):
        return jsonify({"error": "login_required"}), 401
    path = request.args.get("path", "")
    # Security: only allow files from the temp directory with our prefix
    if not os.path.basename(path).startswith("emails_") or \
       not path.endswith(".zip"):
        return jsonify({"error": "Invalid file"}), 400
    if not os.path.isfile(path):
        return jsonify({"error": "File not found"}), 404

    try:
        return send_file(
            path,
            mimetype="application/zip",
            as_attachment=True,
            download_name="emails.zip",
        )
    finally:
        # Clean up temp file after sending
        try:
            os.unlink(path)
        except OSError:
            pass


@app.route("/ingest")
def ingest():
    if not session.get("user"):
        return redirect(url_for("login"))
    return render_template("ingest.html", user=session["user"])


@app.route("/api/ingest/page")
def ingest_page():
    if not session.get("user"):
        return jsonify({"error": "login_required"}), 401
    email = request.args.get("email", "")
    if not email:
        return jsonify({"error": "email parameter required"}), 400
    next_link = request.args.get("nextLink", "") or None
    token = auth_helper.get_token_from_cache()
    if not token:
        return jsonify({"error": "login_required"}), 401

    messages, next_link_out = graph_helper.get_messages_page(
        token["access_token"], email, next_link=next_link
    )

    if messages is None:
        if next_link_out == "group_not_supported":
            return jsonify({"error": "Group mailboxes are not supported on the ingest page."}), 400
        return jsonify({"error": "Failed to fetch messages from Microsoft Graph."}), 502

    return jsonify({"messages": messages, "nextLink": next_link_out})


@app.route("/api/ingest/run", methods=["POST"])
def ingest_run():
    if not session.get("user"):
        return jsonify({"error": "login_required"}), 401
    data = request.get_json(silent=True) or {}
    message_id = data.get("messageId", "")
    if not message_id:
        return jsonify({"error": "messageId required"}), 400
    token = auth_helper.get_token_from_cache()
    if not token:
        return jsonify({"error": "login_required"}), 401

    access_token = token["access_token"]

    detail = graph_helper.get_message_detail(access_token, message_id)
    if not detail:
        return jsonify({"error": "Failed to fetch message detail"}), 502

    attachments = []
    if detail.get("hasAttachments"):
        attachments = graph_helper.get_message_attachments(access_token, message_id)

    saved, skipped = save_email_to_db(detail, attachments)

    return jsonify({
        "saved": saved,
        "skipped": skipped,
        "subject": detail.get("subject", "(no subject)"),
        "attachmentCount": len(attachments),
    })


@app.route("/api/auth/token-refresh", methods=["POST"])
def token_refresh():
    """Called periodically by the frontend to silently refresh the access token."""
    if not session.get("user"):
        return jsonify({"error": "login_required"}), 401
    token = auth_helper.get_token_from_cache()
    if not token:
        return jsonify({"error": "login_required"}), 401
    return jsonify({"ok": True})


if __name__ == "__main__":
    app.run(host="localhost", port=5000, debug=True)
