import base64
import html
import os
import re
from html.parser import HTMLParser
from io import BytesIO

from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()


_SKIP_TAGS = {"style", "script", "head"}
# Block-level tags that act as line separators when stripped.
_BLOCK_TAGS = {
    "p", "div", "br", "li", "tr", "td", "th",
    "h1", "h2", "h3", "h4", "h5", "h6",
    "blockquote", "hr", "pre",
}


class _HTMLStripper(HTMLParser):
    def __init__(self):
        super().__init__()
        self._parts = []
        self._skip = False

    def handle_starttag(self, tag, attrs):
        t = tag.lower()
        if t in _SKIP_TAGS:
            self._skip = True
        elif t in _BLOCK_TAGS:
            self._parts.append('\n')

    def handle_endtag(self, tag):
        t = tag.lower()
        if t in _SKIP_TAGS:
            self._skip = False
        elif t in _BLOCK_TAGS and t != 'br':
            self._parts.append('\n')

    def handle_data(self, data):
        if not self._skip:
            self._parts.append(data)

    def handle_comment(self, data):
        pass  # drop HTML comments


def strip_html_tags(raw: str) -> str:
    """Strip HTML tags and decode entities, returning plain text."""
    stripper = _HTMLStripper()
    stripper.feed(raw or "")
    return html.unescape("".join(stripper._parts))


# Patterns that always start a new paragraph when encountered inline.
_PARA_BREAK_BEFORE = re.compile(
    r'(?<!\n)'
    r'('
    r'From:\s|Sent:\s|To:\s|Cc:\s|Subject:\s'
    r'|Thanks\s*&\s*Regards|Warm\s+Regards|Best\s+Regards|Kind\s+Regards|Regards[,.]'
    r'|-----\s*Original'
    r')',
    re.IGNORECASE,
)

# Bullet variants to normalise to "- ".
_BULLET_RE = re.compile(r'^[\u2022\u2023\u25e6\u2043\u2219*]\s+', re.MULTILINE)


def clean_body_for_export(text: str) -> str:
    """Clean extracted email body text for JSONB storage / export.

    Fixes structural issues (run-on lines, broken line breaks, noisy metadata)
    without rewriting or summarising any content.
    """
    if not text:
        return text

    # 1. Normalise line endings.
    text = text.replace('\r\n', '\n').replace('\r', '\n')

    # 2. Non-breaking spaces act as paragraph separators in stripped HTML —
    #    convert each run to a single newline.
    text = re.sub(r'\xa0+', '\n', text)

    # 3. Convert literal escaped newline sequences to real newlines.
    text = text.replace('\\r\\n', '\n').replace('\\r', '\n').replace('\\n', '\n')

    # 4. Replace corrupted/replacement characters (e.g. mangled em-dashes).
    text = text.replace('\ufffd', '\u2013')  # U+FFFD → en-dash
    # Drop other non-printable control characters (except newline/tab).
    text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', text)

    # 5. Insert paragraph breaks before structural email markers.
    text = _PARA_BREAK_BEFORE.sub(r'\n\n\1', text)

    # 6. Rejoin lines broken mid-phrase (multiple passes).
    #    a) No terminal punctuation, continuation starts with lowercase or digit.
    text = re.sub(r'(?<![.!?:,\-])\n(?=[a-z0-9])', ' ', text)
    #    b) Line ends with a dangling conjunction, preposition, or & (clearly
    #       mid-phrase even when the next word is capitalised).
    text = re.sub(
        r'(\b(?:of|and|or|the|an?|to|for|with|by|as|in|on|at|its|our|their|a)\b|&)'
        r'\s*\n\s*(\S)',
        r'\1 \2',
        text,
        flags=re.IGNORECASE,
    )
    #    c) A single capitalised word on its own line that follows a line
    #       without terminal punctuation is very likely a split heading/label.
    text = re.sub(r'(?<![.!?:\n])\n([A-Z][a-z]+\n)', r' \1', text)

    # 7. Collapse runs of 3+ newlines to a single blank line.
    text = re.sub(r'\n{3,}', '\n\n', text)

    # 8. Normalise horizontal whitespace within each line.
    text = re.sub(r'[^\S\n]{2,}', ' ', text)

    # 9. Strip leading/trailing whitespace per line.
    lines = [line.strip() for line in text.split('\n')]
    text = '\n'.join(lines)

    # 10. Standardise bullet points to "- ".
    text = _BULLET_RE.sub('- ', text)

    # 11. Remove duplicate consecutive blank lines introduced by earlier steps.
    text = re.sub(r'\n{3,}', '\n\n', text)

    return text.strip()


class Project(db.Model):
    __tablename__ = "projects"

    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(255), nullable=False)
    project_number = db.Column(db.String(64), nullable=True)
    created_at = db.Column(db.DateTime, server_default=db.func.now())

    emails = db.relationship("Email", backref="project", lazy=True)


class Email(db.Model):
    __tablename__ = "emails"

    id = db.Column(db.Integer, primary_key=True)
    message_id = db.Column(db.String(512), unique=True, nullable=False, index=True)
    sender_email = db.Column(db.String(255))
    sender_name = db.Column(db.String(255))
    recipients = db.Column(db.Text)  # JSON-encoded list of TO+CC addresses (legacy)
    to_recipients = db.Column(db.Text)   # JSON-encoded list of TO addresses
    cc_recipients = db.Column(db.Text)   # JSON-encoded list of CC addresses
    bcc_recipients = db.Column(db.Text)  # JSON-encoded list of BCC addresses
    conversation_id = db.Column(db.String(512), index=True)  # MS Graph conversationId
    date_received = db.Column(db.DateTime(timezone=True))
    subject = db.Column(db.Text)
    body = db.Column(db.Text)
    body_text = db.Column(db.Text)  # HTML-stripped plain text
    searched_email = db.Column(db.String(255), index=True)
    project_id = db.Column(db.Integer, db.ForeignKey("projects.id"), nullable=True)
    created_at = db.Column(db.DateTime, server_default=db.func.now())

    attachments = db.relationship(
        "EmailAttachment", backref="email", lazy=True, cascade="all, delete-orphan"
    )


class EmailAttachment(db.Model):
    __tablename__ = "email_attachments"

    id = db.Column(db.Integer, primary_key=True)
    email_id = db.Column(db.Integer, db.ForeignKey("emails.id"), nullable=False)
    filename = db.Column(db.String(512))
    content_type = db.Column(db.String(255))
    extracted_text = db.Column(db.Text)


class ChatSession(db.Model):
    __tablename__ = "chat_sessions"

    id = db.Column(db.Integer, primary_key=True)
    user_oid = db.Column(db.String(64), nullable=False, index=True)
    title = db.Column(db.String(255))
    openai_response_id = db.Column(db.String(255))
    created_at = db.Column(db.DateTime, server_default=db.func.now())
    updated_at = db.Column(db.DateTime, server_default=db.func.now(), onupdate=db.func.now())

    messages = db.relationship(
        "ChatMessage", backref="session", lazy=True, cascade="all, delete-orphan",
        order_by="ChatMessage.id"
    )


class ChatMessage(db.Model):
    __tablename__ = "chat_messages"

    id = db.Column(db.Integer, primary_key=True)
    session_id = db.Column(db.Integer, db.ForeignKey("chat_sessions.id"), nullable=False)
    role = db.Column(db.String(16))  # "user" or "assistant"
    content = db.Column(db.Text)
    created_at = db.Column(db.DateTime, server_default=db.func.now())


def extract_text_from_attachment(filename, content_type, content_bytes):
    """Return scraped text from an attachment, or None if not extractable."""
    ext = os.path.splitext(filename or "")[1].lower()

    # Plain-text formats
    if ext in (".txt", ".csv", ".log", ".md", ".json", ".xml", ".html", ".htm"):
        try:
            return content_bytes.decode("utf-8", errors="replace")
        except Exception:
            return None

    # PDF
    if ext == ".pdf" or (content_type or "").lower() == "application/pdf":
        try:
            from pypdf import PdfReader

            reader = PdfReader(BytesIO(content_bytes))
            pages = [page.extract_text() or "" for page in reader.pages]
            return "\n".join(pages).strip() or None
        except Exception:
            return None

    # Word (docx)
    if ext == ".docx" or content_type in (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ):
        try:
            from docx import Document

            doc = Document(BytesIO(content_bytes))
            return "\n".join(p.text for p in doc.paragraphs).strip() or None
        except Exception:
            return None

    return None


class TeamsChat(db.Model):
    __tablename__ = "teams_chats"

    id = db.Column(db.Integer, primary_key=True)
    chat_id = db.Column(db.String(512), unique=True, nullable=False, index=True)
    display_name = db.Column(db.String(512), nullable=True)
    chat_type = db.Column(db.String(64), nullable=True)
    last_updated_date_time = db.Column(db.DateTime(timezone=True), nullable=True)
    member_count = db.Column(db.Integer, nullable=True)
    scraped_at = db.Column(db.DateTime, nullable=True)
    created_at = db.Column(db.DateTime, server_default=db.func.now())

    messages = db.relationship(
        "TeamsMessage", backref="chat", lazy=True, cascade="all, delete-orphan",
        order_by="TeamsMessage.id"
    )


class TeamsMessage(db.Model):
    __tablename__ = "teams_messages"

    id = db.Column(db.Integer, primary_key=True)
    teams_chat_id = db.Column(db.Integer, db.ForeignKey("teams_chats.id"), nullable=False)
    message_id = db.Column(db.String(512), unique=True, nullable=False, index=True)
    sender_name = db.Column(db.String(255), nullable=True)
    sender_email = db.Column(db.String(255), nullable=True)
    content_html = db.Column(db.Text, nullable=True)
    content_text = db.Column(db.Text, nullable=True)
    created_date_time = db.Column(db.DateTime(timezone=True), nullable=True)
    message_type = db.Column(db.String(64), nullable=True)
    created_at = db.Column(db.DateTime, server_default=db.func.now())



def upsert_teams_chat(chat_data):
    """Create or update a TeamsChat row from a Graph API chat object.

    chat_data keys: id, topic (optional), chatType (optional),
    lastUpdatedDateTime (optional).
    Returns the TeamsChat row.
    """
    from datetime import datetime, timezone

    chat_id = chat_data.get("id", "")
    if not chat_id:
        return None

    row = TeamsChat.query.filter_by(chat_id=chat_id).first()
    if not row:
        row = TeamsChat(chat_id=chat_id)
        db.session.add(row)

    row.display_name = chat_data.get("topic") or chat_data.get("display_name")
    row.chat_type = chat_data.get("chatType")

    lud = chat_data.get("lastUpdatedDateTime")
    if lud:
        try:
            row.last_updated_date_time = datetime.fromisoformat(
                lud.replace("Z", "+00:00")
            )
        except ValueError:
            pass

    db.session.commit()
    return row


def save_teams_messages_to_db(chat_id_str, messages_batch):
    """Insert new Teams messages, skipping duplicates and system events.

    Returns {"saved": N, "skipped": M}.
    """
    from datetime import datetime, timezone

    chat_row = TeamsChat.query.filter_by(chat_id=chat_id_str).first()
    if not chat_row:
        return {"saved": 0, "skipped": len(messages_batch)}

    saved = 0
    skipped = 0

    for msg in messages_batch:
        msg_id = msg.get("id", "")
        if not msg_id:
            skipped += 1
            continue

        # Skip system event messages (joins, renames, etc.)
        if msg.get("messageType", "message") != "message":
            skipped += 1
            continue

        if TeamsMessage.query.filter_by(message_id=msg_id).first():
            skipped += 1
            continue

        content_html = msg.get("contentHtml") or ""
        content_text = strip_html_tags(content_html) if content_html else ""

        created_str = msg.get("createdDateTime", "")
        created_dt = None
        if created_str:
            try:
                created_dt = datetime.fromisoformat(
                    created_str.replace("Z", "+00:00")
                )
            except ValueError:
                pass

        msg_row = TeamsMessage(
            teams_chat_id=chat_row.id,
            message_id=msg_id,
            sender_name=msg.get("senderName") or "",
            sender_email=msg.get("senderEmail") or "",
            content_html=content_html,
            content_text=content_text,
            created_date_time=created_dt,
            message_type=msg.get("messageType", "message"),
        )
        db.session.add(msg_row)
        saved += 1

    db.session.commit()
    return {"saved": saved, "skipped": skipped}


# ── Teams Teams / Channels / Posts ────────────────────────────────────────────

class TeamsTeam(db.Model):
    __tablename__ = "teams_teams"

    id = db.Column(db.Integer, primary_key=True)
    team_id = db.Column(db.String(512), unique=True, nullable=False, index=True)
    display_name = db.Column(db.String(512), nullable=True)
    description = db.Column(db.Text, nullable=True)
    scraped_at = db.Column(db.DateTime, nullable=True)
    created_at = db.Column(db.DateTime, server_default=db.func.now())

    channels = db.relationship(
        "TeamsChannel", backref="team", lazy=True, cascade="all, delete-orphan"
    )


class TeamsChannel(db.Model):
    __tablename__ = "teams_channels"

    id = db.Column(db.Integer, primary_key=True)
    teams_team_id = db.Column(db.Integer, db.ForeignKey("teams_teams.id"), nullable=False)
    channel_id = db.Column(db.String(512), unique=True, nullable=False, index=True)
    display_name = db.Column(db.String(512), nullable=True)
    description = db.Column(db.Text, nullable=True)
    scraped_at = db.Column(db.DateTime, nullable=True)
    created_at = db.Column(db.DateTime, server_default=db.func.now())

    posts = db.relationship(
        "TeamsChannelPost", backref="channel", lazy=True, cascade="all, delete-orphan"
    )


class TeamsChannelPost(db.Model):
    __tablename__ = "teams_channel_posts"

    id = db.Column(db.Integer, primary_key=True)
    teams_channel_id = db.Column(db.Integer, db.ForeignKey("teams_channels.id"), nullable=False)
    message_id = db.Column(db.String(512), unique=True, nullable=False, index=True)
    sender_name = db.Column(db.String(255), nullable=True)
    sender_email = db.Column(db.String(255), nullable=True)
    subject = db.Column(db.String(512), nullable=True)
    content_html = db.Column(db.Text, nullable=True)
    content_text = db.Column(db.Text, nullable=True)
    created_date_time = db.Column(db.DateTime(timezone=True), nullable=True)
    importance = db.Column(db.String(32), nullable=True)
    web_url = db.Column(db.Text, nullable=True)
    message_type = db.Column(db.String(64), nullable=True)
    created_at = db.Column(db.DateTime, server_default=db.func.now())


def upsert_teams_team(team_data):
    """Create or update a TeamsTeam row from a Graph API joinedTeams item.

    team_data keys: id, displayName (optional), description (optional).
    Returns the TeamsTeam row.
    """
    team_id = team_data.get("id", "")
    if not team_id:
        return None

    row = TeamsTeam.query.filter_by(team_id=team_id).first()
    if not row:
        row = TeamsTeam(team_id=team_id)
        db.session.add(row)

    row.display_name = team_data.get("displayName") or team_data.get("display_name")
    row.description = team_data.get("description")

    db.session.commit()
    return row


def upsert_teams_channel(team_db_id, channel_data):
    """Create or update a TeamsChannel row.

    channel_data keys: id, displayName (optional), description (optional).
    Returns the TeamsChannel row.
    """
    channel_id = channel_data.get("id", "")
    if not channel_id:
        return None

    row = TeamsChannel.query.filter_by(channel_id=channel_id).first()
    if not row:
        row = TeamsChannel(channel_id=channel_id, teams_team_id=team_db_id)
        db.session.add(row)

    row.display_name = channel_data.get("displayName") or channel_data.get("display_name")
    row.description = channel_data.get("description")

    db.session.commit()
    return row


def save_teams_channel_posts_to_db(channel_db_id, posts_batch):
    """Insert new Teams channel posts, skipping duplicates and system events.

    Returns {"saved": N, "skipped": M, "rows": [TeamsChannelPost, ...]}.
    rows contains only the newly saved rows (for immediate ChromaDB embedding).
    """
    from datetime import datetime

    saved = 0
    skipped = 0
    new_rows = []

    for post in posts_batch:
        msg_id = post.get("id", "")
        if not msg_id:
            skipped += 1
            continue

        if TeamsChannelPost.query.filter_by(message_id=msg_id).first():
            skipped += 1
            continue

        content_html = post.get("contentHtml") or ""
        content_text = strip_html_tags(content_html) if content_html else ""

        created_str = post.get("createdDateTime", "")
        created_dt = None
        if created_str:
            try:
                created_dt = datetime.fromisoformat(created_str.replace("Z", "+00:00"))
            except ValueError:
                pass

        post_row = TeamsChannelPost(
            teams_channel_id=channel_db_id,
            message_id=msg_id,
            sender_name=post.get("senderName") or "",
            sender_email=post.get("senderEmail") or "",
            subject=post.get("subject") or "",
            content_html=content_html,
            content_text=content_text,
            created_date_time=created_dt,
            importance=post.get("importance") or "normal",
            web_url=post.get("webUrl") or "",
            message_type=post.get("messageType", "message"),
        )
        db.session.add(post_row)
        new_rows.append(post_row)
        saved += 1

    db.session.commit()
    return {"saved": saved, "skipped": skipped, "rows": new_rows}


class CalendarEvent(db.Model):
    __tablename__ = "calendar_events"

    id = db.Column(db.Integer, primary_key=True)
    event_id = db.Column(db.String(512), unique=True, nullable=False, index=True)
    subject = db.Column(db.Text)
    organizer_email = db.Column(db.String(255))
    organizer_name = db.Column(db.String(255))
    start_datetime = db.Column(db.DateTime(timezone=True))
    end_datetime = db.Column(db.DateTime(timezone=True))
    timezone = db.Column(db.String(64))
    location = db.Column(db.Text)
    body_html = db.Column(db.Text)
    body_text = db.Column(db.Text)
    is_online_meeting = db.Column(db.Boolean, default=False)
    online_meeting_url = db.Column(db.Text)
    join_url = db.Column(db.Text)
    web_link = db.Column(db.Text)
    searched_email = db.Column(db.String(255), index=True)
    created_at = db.Column(db.DateTime, server_default=db.func.now())

    attendees = db.relationship(
        "CalendarEventAttendee", backref="event", lazy=True, cascade="all, delete-orphan"
    )


class CalendarEventAttendee(db.Model):
    __tablename__ = "calendar_event_attendees"

    id = db.Column(db.Integer, primary_key=True)
    calendar_event_id = db.Column(db.Integer, db.ForeignKey("calendar_events.id"), nullable=False)
    email = db.Column(db.String(255))
    name = db.Column(db.String(255))
    attendee_type = db.Column(db.String(32))    # "required", "optional", "resource"
    response_status = db.Column(db.String(32))  # "accepted", "declined", "tentativelyAccepted", "none"


def save_calendar_event_to_db(event_data, searched_email=None):
    """Persist one calendar event and its attendees. Returns (saved: bool, skipped: bool).

    Uses the Graph API event id as the deduplication key.
    """
    from datetime import datetime

    graph_id = event_data.get("id", "")
    if not graph_id:
        return False, False

    existing = CalendarEvent.query.filter_by(event_id=graph_id).first()
    if existing:
        if existing.searched_email is None and searched_email:
            existing.searched_email = searched_email
            db.session.commit()
        return False, True

    # Parse start/end datetimes
    start_info = event_data.get("start") or {}
    end_info = event_data.get("end") or {}
    tz = start_info.get("timeZone", "UTC")

    def _parse_dt(dt_str):
        if not dt_str:
            return None
        try:
            return datetime.fromisoformat(dt_str.replace("Z", "+00:00"))
        except ValueError:
            return None

    start_dt = _parse_dt(start_info.get("dateTime", ""))
    end_dt = _parse_dt(end_info.get("dateTime", ""))

    # Location
    location_info = event_data.get("location") or {}
    location = location_info.get("displayName", "") or ""

    # Body
    body_info = event_data.get("body") or {}
    body_html = body_info.get("content", "")
    body_content_type = body_info.get("contentType", "html")
    if body_content_type.lower() == "html":
        body_text = clean_body_for_export(strip_html_tags(body_html))
    else:
        body_text = clean_body_for_export(body_html or "")

    # Organizer
    org = (event_data.get("organizer") or {}).get("emailAddress") or {}
    organizer_email = org.get("address", "")
    organizer_name = org.get("name", "")

    # Online meeting URLs
    online_meeting_url = event_data.get("onlineMeetingUrl") or ""
    join_url = (event_data.get("onlineMeeting") or {}).get("joinUrl") or ""

    event_row = CalendarEvent(
        event_id=graph_id,
        subject=event_data.get("subject", "(no subject)"),
        organizer_email=organizer_email,
        organizer_name=organizer_name,
        start_datetime=start_dt,
        end_datetime=end_dt,
        timezone=tz,
        location=location,
        body_html=body_html,
        body_text=body_text,
        is_online_meeting=bool(event_data.get("isOnlineMeeting", False)),
        online_meeting_url=online_meeting_url,
        join_url=join_url,
        web_link=event_data.get("webLink", ""),
        searched_email=searched_email,
    )
    db.session.add(event_row)
    db.session.flush()

    for att in event_data.get("attendees") or []:
        email_info = (att.get("emailAddress") or {})
        status_info = (att.get("status") or {})
        db.session.add(CalendarEventAttendee(
            calendar_event_id=event_row.id,
            email=email_info.get("address", ""),
            name=email_info.get("name", ""),
            attendee_type=att.get("type", "required"),
            response_status=status_info.get("response", "none"),
        ))

    db.session.commit()
    return True, False


def save_email_to_db(message_detail, attachment_list, searched_email=None, project_id=None):
    """Persist one email + its attachments. Returns (saved: bool, skipped: bool).

    Uses the Graph API message id as the deduplication key.
    """
    import json
    from datetime import datetime, timezone

    graph_id = message_detail.get("id", "")
    if not graph_id:
        return False, False

    # Duplicate check — backfill any missing fields on existing records
    existing = Email.query.filter_by(message_id=graph_id).first()
    if existing:
        updated = False
        if existing.body_text is None and existing.body:
            existing.body_text = clean_body_for_export(strip_html_tags(existing.body))
            updated = True
        if existing.searched_email is None and searched_email:
            existing.searched_email = searched_email
            updated = True
        if existing.project_id is None and project_id is not None:
            existing.project_id = project_id
            updated = True
        if updated:
            db.session.commit()
        return False, True

    # Parse recipients
    to_recipients = [
        r.get("emailAddress", {}).get("address", "")
        for r in message_detail.get("toRecipients", [])
    ]
    cc_recipients = [
        r.get("emailAddress", {}).get("address", "")
        for r in message_detail.get("ccRecipients", [])
    ]
    bcc_recipients = [
        r.get("emailAddress", {}).get("address", "")
        for r in message_detail.get("bccRecipients", [])
    ]
    recipients_json = json.dumps(list(dict.fromkeys(to_recipients + cc_recipients)))

    # Parse date
    date_str = message_detail.get("receivedDateTime", "")
    date_received = None
    if date_str:
        try:
            date_received = datetime.fromisoformat(date_str.replace("Z", "+00:00"))
        except ValueError:
            pass

    body_info = message_detail.get("body", {})
    body_content = body_info.get("content", "")
    body_content_type = body_info.get("contentType", "html")
    if body_content_type.lower() == "html":
        body_text_content = clean_body_for_export(strip_html_tags(body_content))
    else:
        body_text_content = clean_body_for_export(body_content or "")

    email_row = Email(
        message_id=graph_id,
        sender_email=message_detail.get("from", {})
            .get("emailAddress", {}).get("address", ""),
        sender_name=message_detail.get("from", {})
            .get("emailAddress", {}).get("name", ""),
        recipients=recipients_json,
        to_recipients=json.dumps(to_recipients),
        cc_recipients=json.dumps(cc_recipients),
        bcc_recipients=json.dumps(bcc_recipients),
        conversation_id=message_detail.get("conversationId", ""),
        date_received=date_received,
        subject=message_detail.get("subject", "(no subject)"),
        body=body_content,
        body_text=body_text_content,
        searched_email=searched_email,
        project_id=project_id,
    )
    db.session.add(email_row)
    db.session.flush()  # get email_row.id before committing

    for att in attachment_list:
        att_name = att.get("name", "attachment")
        att_content_type = att.get("contentType", "")
        att_b64 = att.get("contentBytes", "")

        extracted = None
        if att_b64:
            try:
                att_bytes = base64.b64decode(att_b64)
                extracted = extract_text_from_attachment(
                    att_name, att_content_type, att_bytes
                )
            except Exception:
                pass

        db.session.add(EmailAttachment(
            email_id=email_row.id,
            filename=att_name,
            content_type=att_content_type,
            extracted_text=extracted,
        ))

    db.session.commit()
    return True, False
