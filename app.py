import json
import os

from flask import (
    Flask, redirect, render_template, request,
    session, url_for, jsonify, send_file, Response, stream_with_context,
)
from flask_session import Session
from werkzeug.middleware.proxy_fix import ProxyFix

import app_config
import auth_helper
import graph_helper
from flask_migrate import Migrate
from models import (
    db, save_email_to_db, ChatSession, ChatMessage, Email, Project,
    TeamsChat, TeamsMessage,
    upsert_teams_chat, save_teams_messages_to_db, clean_body_for_export,
)

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


@app.route("/api/emails/export")
def export_emails_json():
    if not session.get("user"):
        return jsonify({"error": "login_required"}), 401
    project_id = request.args.get("project_id", type=int)
    query = Email.query.options(
        db.joinedload(Email.attachments),
        db.joinedload(Email.project),
    )
    if project_id:
        query = query.filter(Email.project_id == project_id)
    emails = query.all()
    data = []
    for email in emails:
        # Use new separate fields; fall back to legacy merged recipients for old rows
        to = json.loads(email.to_recipients) if email.to_recipients else (
            json.loads(email.recipients) if email.recipients else []
        )
        cc = json.loads(email.cc_recipients) if email.cc_recipients else []
        bcc = json.loads(email.bcc_recipients) if email.bcc_recipients else []
        data.append({
            "id": email.id,
            "message_id": email.message_id,
            "thread_id": email.conversation_id or None,
            "project_name": email.project.title if email.project else None,
            "project_number": email.project.project_number if email.project else None,
            "subject": email.subject,
            "date": email.date_received.isoformat() if email.date_received else None,
            "from": email.sender_email,
            "from_name": email.sender_name,
            "to": to,
            "cc": cc,
            "bcc": bcc,
            "body": clean_body_for_export(email.body_text) if email.body_text else None,
            "attachments": [a.filename for a in email.attachments],
        })
    from datetime import datetime
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    json_bytes = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
    return Response(
        json_bytes,
        mimetype="application/json",
        headers={"Content-Disposition": f'attachment; filename="emails_{timestamp}.json"'},
    )


@app.route("/api/projects", methods=["GET"])
def list_projects():
    if not session.get("user"):
        return jsonify({"error": "login_required"}), 401
    projects = Project.query.order_by(Project.created_at.desc()).all()
    return jsonify([{
        "id": p.id,
        "title": p.title,
        "project_number": p.project_number,
        "created_at": p.created_at.isoformat() if p.created_at else None,
    } for p in projects])


@app.route("/api/projects", methods=["POST"])
def create_project():
    if not session.get("user"):
        return jsonify({"error": "login_required"}), 401
    data = request.get_json(silent=True) or {}
    title = (data.get("title") or "").strip()
    if not title:
        return jsonify({"error": "title required"}), 400
    project = Project(
        title=title,
        project_number=(data.get("project_number") or "").strip() or None,
    )
    db.session.add(project)
    db.session.commit()
    return jsonify({
        "id": project.id,
        "title": project.title,
        "project_number": project.project_number,
    }), 201


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
    searched_email = data.get("searchedEmail", "") or None
    project_id = data.get("projectId") or None
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

    saved, skipped = save_email_to_db(detail, attachments, searched_email=searched_email, project_id=project_id)

    if saved:
        try:
            import chroma_helper
            from models import Email
            email_row = Email.query.filter_by(message_id=message_id).first()
            if email_row:
                chroma_helper.embed_and_upsert(email_row)
        except Exception as e:
            app.logger.warning(f"ChromaDB embed failed: {e}")

    return jsonify({
        "saved": saved,
        "skipped": skipped,
        "subject": detail.get("subject", "(no subject)"),
        "attachmentCount": len(attachments),
    })


@app.route("/emails-by-search")
def emails_by_search():
    if not session.get("user"):
        return redirect(url_for("login"))
    return render_template("emails_by_search.html", user=session["user"])


@app.route("/api/emails/by-search")
def api_emails_by_search_list():
    if not session.get("user"):
        return jsonify({"error": "login_required"}), 401
    from sqlalchemy import func
    rows = (
        db.session.query(Email.searched_email, func.count(Email.id).label("count"))
        .filter(Email.searched_email.isnot(None))
        .group_by(Email.searched_email)
        .order_by(func.count(Email.id).desc())
        .all()
    )
    return jsonify([{"searched_email": r.searched_email, "count": r.count} for r in rows])


@app.route("/api/emails/by-search/<path:email>")
def api_emails_by_search_detail(email):
    if not session.get("user"):
        return jsonify({"error": "login_required"}), 401
    emails = (
        Email.query
        .filter_by(searched_email=email)
        .order_by(Email.date_received.desc())
        .all()
    )
    return jsonify([{
        "id": e.id,
        "subject": e.subject,
        "sender_email": e.sender_email,
        "sender_name": e.sender_name,
        "date_received": e.date_received.isoformat() if e.date_received else None,
        "attachment_count": len(e.attachments),
    } for e in emails])


@app.route("/search")
def search():
    if not session.get("user"):
        return redirect(url_for("login"))
    return render_template("search.html", user=session["user"])


@app.route("/api/search")
def api_search():
    if not session.get("user"):
        return jsonify({"error": "login_required"}), 401
    query = request.args.get("q", "").strip()
    if not query:
        return jsonify({"results": []})
    try:
        import chroma_helper
        results = chroma_helper.search_emails(query, n_results=10)
        return jsonify({"results": results})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/chat")
def chat():
    if not session.get("user"):
        return redirect(url_for("login"))
    user_oid = session["user"].get("oid", "")
    sessions = (
        ChatSession.query
        .filter_by(user_oid=user_oid)
        .order_by(ChatSession.updated_at.desc())
        .all()
    )
    sessions_data = [
        {"id": s.id, "title": s.title or "Untitled", "updated_at": s.updated_at.isoformat() if s.updated_at else ""}
        for s in sessions
    ]
    return render_template("chat.html", user=session["user"], chat_sessions=sessions_data)


def _parse_query_for_dates(message):
    """
    Use GPT to extract a cleaned semantic query and optional date range from the user's message.
    Returns: {"semantic_query": str, "date_start": str|None, "date_end": str|None}
    Falls back to the original message with no dates on any failure.
    """
    import json
    from datetime import datetime, timezone

    today_str = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    system_prompt = (
        "You are a query parser for an email search system. "
        f"Today's date (UTC) is {today_str}. "
        "Given a user's natural language question about emails, extract:\n"
        "1. 'semantic_query': the question with ALL time-based words removed (e.g. 'last week', "
        "'yesterday', 'recently', 'this month', 'last week only'). Keep names, topics, and non-temporal context.\n"
        "2. 'date_start': start of date range as ISO 8601 string (YYYY-MM-DDTHH:MM:SS+00:00), or null.\n"
        "3. 'date_end': end of date range as ISO 8601 string (YYYY-MM-DDTHH:MM:SS+00:00), or null.\n\n"
        "Date rules:\n"
        "- 'last week' = Monday-to-Sunday week before the current week\n"
        "- 'this week' = Monday of current week to today\n"
        "- 'yesterday' = previous calendar day, full day\n"
        "- 'today' = current calendar day\n"
        "- 'last month' = full previous calendar month\n"
        "- 'last N days' = today minus N days to today\n"
        "- 'recently' = last 7 days\n"
        "If no temporal qualifier, return null for both dates.\n"
        'Respond ONLY with JSON: {"semantic_query": "...", "date_start": "...", "date_end": "..."}'
    )
    fallback = {"semantic_query": message, "date_start": None, "date_end": None}
    try:
        from openai import OpenAI
        client = OpenAI(api_key=app_config.OPENAI_API_KEY)
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": message},
            ],
            temperature=0,
            max_tokens=200,
        )
        parsed = json.loads(response.choices[0].message.content.strip())
        return {
            "semantic_query": parsed.get("semantic_query", message) or message,
            "date_start": parsed.get("date_start") or None,
            "date_end": parsed.get("date_end") or None,
        }
    except Exception:
        return fallback


_CHAT_SYSTEM_PROMPT = (
    "You are an email assistant with access to a database of ingested emails. "
    "When email context is provided, synthesize across all provided emails to give a "
    "comprehensive answer — do not just describe one email if multiple are relevant. "
    "Use bullet points for lists of tasks or items. Include the email subject and date "
    "when referencing a specific email. "
    "If the context seems incomplete, say so and suggest the user try a more specific query. "
    "If no relevant context is found, respond: 'I couldn't find relevant emails for that query. "
    "Try rephrasing or being more specific about the project name, person, or time period.' "
    "Never invent email content not shown to you."
)


def _generate_sub_queries(semantic_query):
    """Return 2-3 varied search query rephrases for better recall. Falls back to [semantic_query] on failure."""
    import json
    from openai import OpenAI
    system = (
        "You are a search query generator for an email database. "
        "Given a question or topic, return 2-3 short, varied search queries that would find "
        "relevant emails via semantic similarity. Each query should approach the topic differently. "
        "Output ONLY a JSON array of strings, e.g. [\"query 1\", \"query 2\"]"
    )
    try:
        client = OpenAI(api_key=app_config.OPENAI_API_KEY)
        resp = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": system},
                {"role": "user", "content": semantic_query},
            ],
            temperature=0.3,
            max_tokens=120,
        )
        queries = json.loads(resp.choices[0].message.content.strip())
        if isinstance(queries, list) and all(isinstance(q, str) for q in queries):
            return queries[:3]
    except Exception:
        pass
    return [semantic_query]


@app.route("/api/chat", methods=["POST"])
def api_chat():
    if not session.get("user"):
        return jsonify({"error": "login_required"}), 401

    data = request.get_json(silent=True) or {}
    user_message = data.get("message", "").strip()
    reset = data.get("reset", False)
    load_session_id = data.get("session_id")  # load an existing session

    if not user_message:
        return jsonify({"error": "message required"}), 400

    if reset:
        session.pop("chat_response_id", None)
        session.pop("chat_session_id", None)

    # If loading a specific session, restore its response_id
    if load_session_id:
        user_oid = session["user"].get("oid", "")
        db_session = ChatSession.query.filter_by(id=load_session_id, user_oid=user_oid).first()
        if db_session:
            session["chat_session_id"] = db_session.id
            session["chat_response_id"] = db_session.openai_response_id

    user_oid = session["user"].get("oid", "")
    previous_response_id = session.get("chat_response_id")

    # Get or create a ChatSession for this conversation
    chat_session_id = session.get("chat_session_id")
    if not chat_session_id:
        title = user_message[:60]
        new_db_session = ChatSession(user_oid=user_oid, title=title)
        db.session.add(new_db_session)
        db.session.commit()
        chat_session_id = new_db_session.id
        session["chat_session_id"] = chat_session_id

    # Save the user message
    db.session.add(ChatMessage(session_id=chat_session_id, role="user", content=user_message))
    db.session.commit()

    def generate():
        import json
        import chroma_helper
        from openai import OpenAI

        client = OpenAI(api_key=app_config.OPENAI_API_KEY)

        # Extract temporal intent so we can apply a date filter in ChromaDB
        parsed = _parse_query_for_dates(user_message)
        semantic_query = parsed["semantic_query"]
        date_start = parsed["date_start"]
        date_end = parsed["date_end"]

        where_clause = None
        if date_start and date_end:
            where_clause = {"$and": [
                {"date_received": {"$gte": date_start}},
                {"date_received": {"$lte": date_end}},
            ]}
        elif date_start:
            where_clause = {"date_received": {"$gte": date_start}}
        elif date_end:
            where_clause = {"date_received": {"$lte": date_end}}

        # Generate varied sub-queries for better recall
        sub_queries = _generate_sub_queries(semantic_query)

        # Multi-query retrieval with deduplication
        raw_results = chroma_helper.search_emails_multi_query(
            sub_queries, n_results_each=10, where_clause=where_clause
        )
        context = chroma_helper.format_results_as_context(raw_results)

        # Fallback 1: date filter produced no results — retry without it
        if not context and where_clause is not None:
            raw_results = chroma_helper.search_emails_multi_query(sub_queries, n_results_each=10)
            context = chroma_helper.format_results_as_context(raw_results)
            if context:
                context = "(Note: No emails found in the specified time period. Showing related emails from other dates:)\n\n" + context

        # Fallback 2: nothing found — loosen distance threshold as last resort
        if not context:
            raw_results = chroma_helper.search_emails_multi_query(sub_queries, n_results_each=5)
            context = chroma_helper.format_results_as_context(raw_results, distance_threshold=1.6)
            if context:
                context = "(Note: These emails may be loosely related to your query:)\n\n" + context

        if context:
            full_input = (
                f"Relevant emails from the database:\n\n{context}\n\n"
                f"---\n\nUser question: {user_message}"
            )
        else:
            full_input = (
                f"(No relevant emails found in the database for this query.)\n\n"
                f"User question: {user_message}"
            )

        try:
            kwargs = {
                "model": "gpt-4o-mini",
                "instructions": _CHAT_SYSTEM_PROMPT,
                "input": full_input,
                "stream": True,
            }
            if previous_response_id:
                kwargs["previous_response_id"] = previous_response_id

            stream = client.responses.create(**kwargs)

            response_id = None
            assistant_text = ""
            for event in stream:
                event_type = getattr(event, "type", None)

                if event_type == "response.created":
                    response_id = event.response.id

                elif event_type == "response.output_text.delta":
                    chunk = event.delta
                    if chunk:
                        assistant_text += chunk
                        yield f"data: {json.dumps({'delta': chunk})}\n\n"

                elif event_type == "response.completed":
                    if hasattr(event, "response") and event.response.id:
                        response_id = event.response.id

            if response_id:
                session["chat_response_id"] = response_id
                session.modified = True

            # Persist assistant message and update session response_id
            if assistant_text and chat_session_id:
                db.session.add(ChatMessage(session_id=chat_session_id, role="assistant", content=assistant_text))
                db_sess = ChatSession.query.get(chat_session_id)
                if db_sess:
                    db_sess.openai_response_id = response_id
                    db_sess.updated_at = db.func.now()
                db.session.commit()

            yield f"data: {json.dumps({'done': True, 'session_id': chat_session_id})}\n\n"

        except Exception as e:
            yield f"data: {json.dumps({'error': str(e)})}\n\n"

    return Response(
        stream_with_context(generate()),
        mimetype="text/event-stream",
        headers={
            "Cache-Control": "no-cache",
            "X-Accel-Buffering": "no",
        },
    )


@app.route("/api/chat/reset", methods=["POST"])
def api_chat_reset():
    if not session.get("user"):
        return jsonify({"error": "login_required"}), 401
    session.pop("chat_response_id", None)
    session.pop("chat_session_id", None)
    return jsonify({"ok": True})


@app.route("/api/chat/sessions")
def api_chat_sessions():
    if not session.get("user"):
        return jsonify({"error": "login_required"}), 401
    user_oid = session["user"].get("oid", "")
    sessions = (
        ChatSession.query
        .filter_by(user_oid=user_oid)
        .order_by(ChatSession.updated_at.desc())
        .all()
    )
    return jsonify([
        {"id": s.id, "title": s.title or "Untitled", "updated_at": s.updated_at.isoformat() if s.updated_at else ""}
        for s in sessions
    ])


@app.route("/api/chat/sessions/<int:session_id>")
def api_chat_session(session_id):
    if not session.get("user"):
        return jsonify({"error": "login_required"}), 401
    user_oid = session["user"].get("oid", "")
    db_session = ChatSession.query.filter_by(id=session_id, user_oid=user_oid).first()
    if not db_session:
        return jsonify({"error": "not found"}), 404
    return jsonify({
        "id": db_session.id,
        "title": db_session.title or "Untitled",
        "openai_response_id": db_session.openai_response_id,
        "messages": [
            {"role": m.role, "content": m.content}
            for m in db_session.messages
        ],
    })


@app.route("/api/chat/sessions/<int:session_id>", methods=["DELETE"])
def api_chat_session_delete(session_id):
    if not session.get("user"):
        return jsonify({"error": "login_required"}), 401
    user_oid = session["user"].get("oid", "")
    db_session = ChatSession.query.filter_by(id=session_id, user_oid=user_oid).first()
    if not db_session:
        return jsonify({"error": "not found"}), 404
    db.session.delete(db_session)
    db.session.commit()
    # Clear Flask session if this was the active session
    if session.get("chat_session_id") == session_id:
        session.pop("chat_session_id", None)
        session.pop("chat_response_id", None)
    return jsonify({"ok": True})


# ── Teams Chat Scraper ────────────────────────────────────────────────────────

@app.route("/teams")
def teams():
    if not session.get("user"):
        return redirect(url_for("login"))

    chats = (
        TeamsChat.query
        .order_by(TeamsChat.scraped_at.desc().nullslast())
        .all()
    )
    scraped_chats = []
    for c in chats:
        scraped_chats.append({
            "chat_id": c.chat_id,
            "display_name": c.display_name,
            "member_count": c.member_count,
            "scraped_at": c.scraped_at,
            "message_count": len(c.messages),
        })
    return render_template("teams.html", user=session["user"], scraped_chats=scraped_chats)


@app.route("/api/teams/chats/search")
def api_teams_chats_search():
    if not session.get("user"):
        return jsonify({"error": "login_required"}), 401
    token = auth_helper.get_token_from_cache()
    if not token or "access_token" not in token:
        return jsonify({"error": "login_required"}), 401

    q = request.args.get("q", "").strip()
    if len(q) < 2:
        return jsonify([])

    result = graph_helper.search_teams_chats(token["access_token"], q)
    if isinstance(result, dict) and "error" in result:
        err_code = (result.get("error") or {}).get("error", {}).get("code", "")
        if err_code in ("InvalidAuthenticationToken", "AuthenticationError", "Unauthorized"):
            return jsonify({"error": "login_required"}), 401
        return jsonify(result), 502
    return jsonify(result)


@app.route("/api/teams/messages/page")
def api_teams_messages_page():
    if not session.get("user"):
        return jsonify({"error": "login_required"}), 401
    token = auth_helper.get_token_from_cache()
    if not token or "access_token" not in token:
        return jsonify({"error": "login_required"}), 401

    chat_id = request.args.get("chatId", "").strip()
    if not chat_id:
        return jsonify({"error": "chatId required"}), 400

    next_link = request.args.get("nextLink") or None
    messages, next_link_out = graph_helper.get_teams_messages_page(
        token["access_token"], chat_id, next_link=next_link
    )
    if messages is None:
        err_detail = next_link_out or {}
        return jsonify({"error": "Failed to fetch messages from Graph API", "detail": err_detail}), 502

    return jsonify({"messages": messages, "nextLink": next_link_out})


@app.route("/api/teams/messages/save", methods=["POST"])
def api_teams_messages_save():
    if not session.get("user"):
        return jsonify({"error": "login_required"}), 401

    data = request.get_json(force=True) or {}
    chat_id = data.get("chatId", "").strip()
    chat_topic = data.get("chatTopic", "")
    messages = data.get("messages", [])

    if not chat_id:
        return jsonify({"error": "chatId required"}), 400

    upsert_teams_chat({"id": chat_id, "topic": chat_topic})
    result = save_teams_messages_to_db(chat_id, messages)
    return jsonify(result)


@app.route("/api/teams/chats/complete", methods=["POST"])
def api_teams_chats_complete():
    if not session.get("user"):
        return jsonify({"error": "login_required"}), 401
    token = auth_helper.get_token_from_cache()
    if not token or "access_token" not in token:
        return jsonify({"error": "login_required"}), 401

    chat_id = request.args.get("chatId", "").strip()
    if not chat_id:
        return jsonify({"error": "chatId required"}), 400

    from datetime import datetime
    chat_row = TeamsChat.query.filter_by(chat_id=chat_id).first()
    if chat_row:
        chat_row.scraped_at = datetime.utcnow()
        member_count = graph_helper.get_teams_chat_members_count(token["access_token"], chat_id)
        if member_count is not None:
            chat_row.member_count = member_count
        from models import db as _db
        _db.session.commit()

    return jsonify({"ok": True})


@app.route("/api/teams/chats")
def api_teams_chats_list():
    if not session.get("user"):
        return jsonify({"error": "login_required"}), 401

    chats = TeamsChat.query.order_by(TeamsChat.scraped_at.desc().nullslast()).all()
    return jsonify([
        {
            "id": c.id,
            "chat_id": c.chat_id,
            "display_name": c.display_name,
            "member_count": c.member_count,
            "scraped_at": c.scraped_at.isoformat() if c.scraped_at else None,
            "message_count": len(c.messages),
        }
        for c in chats
    ])


@app.route("/api/teams/messages/export")
def api_teams_messages_export():
    if not session.get("user"):
        return jsonify({"error": "login_required"}), 401

    chat_id = request.args.get("chatId", "").strip()
    if not chat_id:
        return jsonify({"error": "chatId required"}), 400

    chat = TeamsChat.query.filter_by(chat_id=chat_id).first()
    if not chat:
        return jsonify({"error": "Chat not found"}), 404

    messages = (
        TeamsMessage.query
        .filter_by(teams_chat_id=chat.id)
        .order_by(TeamsMessage.created_date_time.asc())
        .all()
    )

    payload = {
        "chat_id": chat.chat_id,
        "display_name": chat.display_name,
        "member_count": chat.member_count,
        "scraped_at": chat.scraped_at.isoformat() if chat.scraped_at else None,
        "messages": [
            {
                "message_id": m.message_id,
                "sender_name": m.sender_name,
                "sender_email": m.sender_email,
                "created_date_time": m.created_date_time.isoformat() if m.created_date_time else None,
                "content": m.content_text,
            }
            for m in messages
        ],
    }

    safe_name = (chat.display_name or chat.chat_id)[:40].replace(" ", "_")
    filename = f"teams_{safe_name}.json"

    return Response(
        json.dumps(payload, indent=2, ensure_ascii=False),
        mimetype="application/json",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )



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
    app.run(host="localhost", port=5000, debug=True, use_reloader=False)
