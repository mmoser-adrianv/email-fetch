import re
import zipfile
from io import BytesIO
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.utils import formatdate
import base64

import requests
import app_config

GRAPH_URL = app_config.GRAPH_ENDPOINT


def search_people(access_token, query):
    headers = {"Authorization": f"Bearer {access_token}"}
    url = (
        f"{GRAPH_URL}/me/people/"
        f"?$search={query}"
        f"&$select=displayName,scoredEmailAddresses"
        f"&$top=10"
    )
    response = requests.get(url, headers=headers, timeout=10)
    if response.status_code != 200:
        return {"error": response.json()}

    results = []
    for person in response.json().get("value", []):
        emails = person.get("scoredEmailAddresses", [])
        email = emails[0]["address"] if emails else None
        if email:
            results.append({
                "displayName": person.get("displayName", ""),
                "email": email,
            })
    return results


def get_user_messages(access_token, user_email, top=10):
    """Fetch messages from a user's mailbox.

    Tries /users/{email}/messages first. If the email belongs to a
    Microsoft 365 Group (ErrorGroupIsUsedInNonGroupURI), it falls back
    to finding the group by email and fetching its conversations via
    /groups/{id}/conversations.

    Returns a dict with:
      - "messages": list of message dicts
      - "source": "user" or "group"
      - "email": the queried email
      - "groupId": (only if source is "group")
    """
    headers = {"Authorization": f"Bearer {access_token}"}

    # --- Attempt 1: regular user mailbox ---
    url = (
        f"{GRAPH_URL}/users/{user_email}/messages"
        f"?$select=id,subject,from,receivedDateTime,bodyPreview"
        f"&$top={top}"
        f"&$orderby=receivedDateTime desc"
    )
    response = requests.get(url, headers=headers, timeout=30)

    if response.status_code == 200:
        return {
            "messages": _parse_messages(response.json()),
            "source": "user",
            "email": user_email,
        }

    # Check if the error is the Group-shard error
    error_body = response.json()
    error_code = (
        error_body.get("error", {}).get("code", "")
    )

    if error_code == "ErrorGroupIsUsedInNonGroupURI":
        return _get_group_threads(headers, user_email, top)

    return {"error": error_body}


def _get_group_threads(headers, group_email, top=10):
    """Look up a Microsoft 365 Group by its mail address and fetch its
    conversation threads."""

    # Find the group by its email (proxyAddresses or mail)
    filter_query = f"mail eq '{group_email}'"
    groups_url = (
        f"{GRAPH_URL}/groups"
        f"?$filter={filter_query}"
        f"&$select=id,displayName,mail"
    )
    resp = requests.get(groups_url, headers=headers, timeout=10)
    if resp.status_code != 200:
        return {"error": resp.json()}

    groups = resp.json().get("value", [])
    if not groups:
        return {"error": f"No group found with email {group_email}"}

    group_id = groups[0]["id"]

    # Fetch the group's conversation threads
    threads_url = (
        f"{GRAPH_URL}/groups/{group_id}/threads"
        f"?$select=id,topic,lastDeliveredDateTime,preview"
        f"&$top={top}"
        f"&$orderby=lastDeliveredDateTime desc"
    )
    resp = requests.get(threads_url, headers=headers, timeout=30)
    if resp.status_code != 200:
        return {"error": resp.json()}

    messages = []
    for thread in resp.json().get("value", []):
        messages.append({
            "id": thread.get("id", ""),
            "subject": thread.get("topic", "(no subject)"),
            "from": groups[0].get("mail", ""),
            "fromName": groups[0].get("displayName", ""),
            "received": thread.get("lastDeliveredDateTime", ""),
            "preview": thread.get("preview", ""),
        })
    # Enforce the limit (Graph may ignore $top on threads)
    messages = messages[:top]
    return {
        "messages": messages,
        "source": "group",
        "email": group_email,
        "groupId": group_id,
    }


def _parse_messages(data):
    """Parse the standard /messages response into a flat list."""
    messages = []
    for msg in data.get("value", []):
        messages.append({
            "id": msg.get("id", ""),
            "subject": msg.get("subject", "(no subject)"),
            "from": msg.get("from", {}).get("emailAddress", {}).get("address", ""),
            "fromName": msg.get("from", {}).get("emailAddress", {}).get("name", ""),
            "received": msg.get("receivedDateTime", ""),
            "preview": msg.get("bodyPreview", ""),
        })
    return messages


# --------------- EML / ZIP download helpers ---------------

def _sanitize_filename(name, max_len=80):
    """Remove characters unsafe for filenames."""
    name = re.sub(r'[\\/*?:"<>|]', "", name)
    name = name.strip(". ")
    return name[:max_len] if name else "email"


def get_message_mime(access_token, email, message_id):
    """Download the raw MIME (.eml) content of a single message.

    GET /users/{email}/messages/{id}/$value
    Returns bytes (the full RFC‑822 message including attachments).
    """
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"{GRAPH_URL}/users/{email}/messages/{message_id}/$value"
    resp = requests.get(url, headers=headers, timeout=60)
    if resp.status_code != 200:
        return None
    return resp.content


def get_group_thread_mime(access_token, group_id, thread_id):
    """Build an EML file from the first post in a group thread.

    The /posts/{id}/$value endpoint is not reliably supported for group
    posts, so instead we:
      1. GET the first post with full body and sender info
      2. GET any attachments on that post
      3. Construct a proper MIME email (EML) from those parts

    Returns bytes or None on failure.
    """
    headers = {"Authorization": f"Bearer {access_token}"}

    # 1. Get the first post with body content
    posts_url = (
        f"{GRAPH_URL}/groups/{group_id}/threads/{thread_id}/posts"
        f"?$select=id,body,from,receivedDateTime,hasAttachments"
        f"&$top=1"
    )
    resp = requests.get(posts_url, headers=headers, timeout=60)
    if resp.status_code != 200:
        return None

    posts = resp.json().get("value", [])
    if not posts:
        return None

    post = posts[0]
    post_id = post["id"]
    body = post.get("body", {})
    body_content = body.get("content", "")
    body_type = body.get("contentType", "text")
    sender = post.get("from", {}).get("emailAddress", {})
    sender_email = sender.get("address", "unknown@group")
    sender_name = sender.get("name", "")
    received = post.get("receivedDateTime", "")

    # 2. Get the thread topic for the Subject header
    thread_url = (
        f"{GRAPH_URL}/groups/{group_id}/threads/{thread_id}"
        f"?$select=topic"
    )
    tresp = requests.get(thread_url, headers=headers, timeout=10)
    subject = ""
    if tresp.status_code == 200:
        subject = tresp.json().get("topic", "(no subject)")

    # 3. Build the MIME message
    msg = MIMEMultipart("mixed")
    msg["Subject"] = subject
    msg["From"] = f"{sender_name} <{sender_email}>" if sender_name else sender_email
    msg["Date"] = received or formatdate(localtime=True)

    # Body part
    if body_type == "html":
        body_part = MIMEText(body_content, "html", "utf-8")
    else:
        body_part = MIMEText(body_content, "plain", "utf-8")
    msg.attach(body_part)

    # 4. Fetch and attach any attachments
    if post.get("hasAttachments"):
        att_url = (
            f"{GRAPH_URL}/groups/{group_id}/threads/{thread_id}"
            f"/posts/{post_id}/attachments"
        )
        att_resp = requests.get(att_url, headers=headers, timeout=60)
        if att_resp.status_code == 200:
            for att in att_resp.json().get("value", []):
                att_name = att.get("name", "attachment")
                att_content_type = att.get("contentType", "application/octet-stream")
                att_bytes_b64 = att.get("contentBytes", "")

                if att_bytes_b64:
                    maintype, _, subtype = att_content_type.partition("/")
                    subtype = subtype or "octet-stream"
                    part = MIMEBase(maintype, subtype)
                    part.set_payload(base64.b64decode(att_bytes_b64))
                    encoders.encode_base64(part)
                    part.add_header(
                        "Content-Disposition", "attachment",
                        filename=att_name,
                    )
                    msg.attach(part)

    return msg.as_bytes()


def download_emails_zip(access_token, result_info):
    """Build an in-memory ZIP of .eml files for the given messages.

    Parameters:
        access_token: Graph API bearer token
        result_info: dict with keys "messages", "source", "email",
                     and optionally "groupId"

    Returns a BytesIO containing the ZIP archive, or None on failure.
    """
    messages = result_info.get("messages", [])
    source = result_info.get("source", "user")
    email = result_info.get("email", "")
    group_id = result_info.get("groupId", "")

    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for idx, msg in enumerate(messages, start=1):
            subject = _sanitize_filename(msg.get("subject", "email"))
            filename = f"{idx:02d}_{subject}.eml"

            if source == "user":
                mime_bytes = get_message_mime(
                    access_token, email, msg["id"]
                )
            else:
                mime_bytes = get_group_thread_mime(
                    access_token, group_id, msg["id"]
                )

            if mime_bytes:
                zf.writestr(filename, mime_bytes)
            else:
                # Write a placeholder so the user knows it failed
                zf.writestr(
                    filename.replace(".eml", "_ERROR.txt"),
                    f"Failed to download: {msg.get('subject', 'Unknown')}\n",
                )

    zip_buffer.seek(0)
    return zip_buffer
