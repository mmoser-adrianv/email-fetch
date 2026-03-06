import base64
import html
import os
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


def save_email_to_db(message_detail, attachment_list):
    """Persist one email + its attachments. Returns (saved: bool, skipped: bool).

    Uses the Graph API message id as the deduplication key.
    """
    import json
    from datetime import datetime, timezone

    graph_id = message_detail.get("id", "")
    if not graph_id:
        return False, False

    # Duplicate check — backfill body_text if missing
    existing = Email.query.filter_by(message_id=graph_id).first()
    if existing:
        if existing.body_text is None and existing.body:
            existing.body_text = strip_html_tags(existing.body)
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
