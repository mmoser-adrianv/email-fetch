import base64
import html
import os
import re
from html.parser import HTMLParser
from io import BytesIO

from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()


_SKIP_TAGS = {"style", "script", "head"}


class _HTMLStripper(HTMLParser):
    def __init__(self):
        super().__init__()
        self._parts = []
        self._skip = False

    def handle_starttag(self, tag, attrs):
        if tag.lower() in _SKIP_TAGS:
            self._skip = True

    def handle_endtag(self, tag):
        if tag.lower() in _SKIP_TAGS:
            self._skip = False

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
    r'|Thanks\s*&\s*Regards|Warm\s+Regards|Best\s+Regards|Kind\s+Regards'
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

    # 3. Remove literal backslash-n sequences (not real newlines).
    text = text.replace('\\n', ' ')

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


class Email(db.Model):
    __tablename__ = "emails"

    id = db.Column(db.Integer, primary_key=True)
    message_id = db.Column(db.String(512), unique=True, nullable=False, index=True)
    sender_email = db.Column(db.String(255))
    sender_name = db.Column(db.String(255))
    recipients = db.Column(db.Text)  # JSON-encoded list of addresses
    date_received = db.Column(db.DateTime(timezone=True))
    subject = db.Column(db.Text)
    body = db.Column(db.Text)
    body_text = db.Column(db.Text)  # HTML-stripped plain text
    searched_email = db.Column(db.String(255), index=True)
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


def save_email_to_db(message_detail, attachment_list, searched_email=None):
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
            existing.body_text = strip_html_tags(existing.body)
            updated = True
        if existing.searched_email is None and searched_email:
            existing.searched_email = searched_email
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
    body_text_content = (
        strip_html_tags(body_content)
        if body_content_type.lower() == "html"
        else body_content
    )

    email_row = Email(
        message_id=graph_id,
        sender_email=message_detail.get("from", {})
            .get("emailAddress", {}).get("address", ""),
        sender_name=message_detail.get("from", {})
            .get("emailAddress", {}).get("name", ""),
        recipients=recipients_json,
        date_received=date_received,
        subject=message_detail.get("subject", "(no subject)"),
        body=body_content,
        body_text=body_text_content,
        searched_email=searched_email,
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
