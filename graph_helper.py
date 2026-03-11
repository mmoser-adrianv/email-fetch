import json
import re
import tempfile
import zipfile
from urllib.parse import quote
from io import BytesIO
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.utils import formatdate, make_msgid, parsedate_to_datetime
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
        f"&$top=20"
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


def get_user_messages(access_token, user_email, top=20):
    """Fetch messages from the signed-in user's mailbox where the searched
    person appears as a recipient or CC.

    Uses /me/messages with a $filter on toRecipients and ccRecipients.
    If the email belongs to a Microsoft 365 Group
    (ErrorGroupIsUsedInNonGroupURI), it falls back to fetching group
    conversations.

    Returns a dict with:
      - "messages": list of message dicts
      - "source": "user" or "group"
      - "email": the queried email
      - "groupId": (only if source is "group")
    """
    headers = {"Authorization": f"Bearer {access_token}"}

    # Proactively check if this is a Microsoft 365 Group email so we read
    # from the group's conversation threads rather than the user's inbox.
    group_id, _name = resolve_group_id(access_token, user_email)
    if group_id:
        return _get_group_messages(headers, user_email, top)

    # --- Search the signed-in user's mailbox for emails to/cc the person ---
    # Note: toRecipients/ccRecipients are NOT filterable in Graph API.
    # Use $search with "to:" and "cc:" instead (KQL syntax).
    search_query = f'"to:{user_email} OR cc:{user_email}"'
    url = (
        f"{GRAPH_URL}/me/messages"
        f"?$search={search_query}"
        f"&$select=id,subject,from,receivedDateTime,bodyPreview"
        f"&$top={top}"
    )
    response = requests.get(url, headers=headers, timeout=30)

    if response.status_code == 200:
        messages = _parse_messages(response.json())
        # $search returns results by relevance; re-sort by date
        messages.sort(key=lambda m: m.get("received", ""), reverse=True)
        return {
            "messages": messages,
            "source": "user",
            "email": user_email,
        }

    # Fallback: if Graph still signals this is a group, try the group path.
    error_body = response.json()
    error_code = error_body.get("error", {}).get("code", "")
    if error_code == "ErrorGroupIsUsedInNonGroupURI":
        return _get_group_messages(headers, user_email, top)

    return {"error": error_body}


def _get_group_all_posts(headers, group_id, max_posts=200):
    """Expand all conversation threads of a Microsoft 365 Group into their
    individual posts (one per email delivery).

    Microsoft 365 Groups store email as conversation threads; each reply or
    forwarded message is a separate post within a thread.  This function
    walks every thread and returns one dict per post so callers see individual
    emails rather than thread summaries.

    Returns a list sorted by receivedDateTime descending.
    Each dict has keys: id, subject, from, fromName, received, preview,
    hasAttachments.  The id is a composite "{thread_id}||{post_id}" string.
    """
    all_posts = []

    threads_url = (
        f"{GRAPH_URL}/groups/{group_id}/threads"
        f"?$select=id,topic,lastDeliveredDateTime"
        f"&$top=50"
        f"&$orderby=lastDeliveredDateTime desc"
    )

    while threads_url and len(all_posts) < max_posts:
        resp = requests.get(threads_url, headers=headers, timeout=30)
        if resp.status_code != 200:
            break
        data = resp.json()

        for thread in data.get("value", []):
            if len(all_posts) >= max_posts:
                break
            thread_id = thread.get("id", "")
            topic = thread.get("topic", "(no subject)")

            posts_url = (
                f"{GRAPH_URL}/groups/{group_id}/threads/{quote(thread_id, safe='')}/posts"
                f"?$select=id,from,receivedDateTime,hasAttachments,body"
                f"&$top=50"
            )
            while posts_url and len(all_posts) < max_posts:
                posts_resp = requests.get(posts_url, headers=headers, timeout=30)
                if posts_resp.status_code != 200:
                    break
                posts_data = posts_resp.json()
                for post in posts_data.get("value", []):
                    sender = (post.get("from") or {}).get("emailAddress") or {}
                    body = post.get("body") or {}
                    body_content = body.get("content", "")
                    if body.get("contentType") == "html":
                        preview = re.sub(r"<[^>]+>", " ", body_content)
                        preview = " ".join(preview.split())[:200]
                    else:
                        preview = body_content[:200]
                    all_posts.append({
                        "id": f"{thread_id}||{post.get('id', '')}",
                        "subject": topic,
                        "from": sender.get("address", ""),
                        "fromName": sender.get("name", ""),
                        "received": post.get("receivedDateTime", ""),
                        "preview": preview,
                        "hasAttachments": post.get("hasAttachments", False),
                    })
                posts_url = posts_data.get("@odata.nextLink")

        threads_url = data.get("@odata.nextLink")

    all_posts.sort(key=lambda m: m.get("received", ""), reverse=True)
    return all_posts


def _get_group_messages(headers, group_email, top=20):
    """Look up a Microsoft 365 Group by its mail address and return individual
    posts (emails) by expanding all conversation threads."""

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

    posts = _get_group_all_posts(headers, group_id, max_posts=max(top * 5, 100))
    return {
        "messages": posts[:top],
        "source": "group",
        "groupMessageType": "post",
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

    GET /me/messages/{id}/$value
    Returns bytes (the full RFC‑822 message including attachments).
    """
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"{GRAPH_URL}/me/messages/{message_id}/$value"
    resp = requests.get(url, headers=headers, timeout=60)
    if resp.status_code != 200:
        return None
    return resp.content


def _build_post_eml(headers, group_id, thread_id, post, subject,
                     group_email):
    """Build a MIME message (EML) from a single group post.

    Returns a MIMEMultipart object.
    """
    post_id = post["id"]
    body = post.get("body", {})
    body_content = body.get("content", "")
    body_type = body.get("contentType", "text")
    sender = post.get("from", {}).get("emailAddress", {})
    sender_email = sender.get("address", "unknown@group")
    sender_name = sender.get("name", "")
    received = post.get("receivedDateTime", "")

    # Build the MIME message with all standard headers
    msg = MIMEMultipart("mixed")
    msg["Subject"] = subject
    msg["From"] = (
        f"{sender_name} <{sender_email}>" if sender_name else sender_email
    )
    msg["To"] = group_email
    msg["Message-ID"] = make_msgid(
        idstring=post_id[:16], domain="graph.microsoft.com"
    )

    # Format the Date header properly from ISO 8601
    if received:
        try:
            dt = parsedate_to_datetime(
                received.replace("T", " ").replace("Z", " +0000")
            )
            msg["Date"] = formatdate(dt.timestamp(), localtime=False)
        except (ValueError, TypeError):
            msg["Date"] = formatdate(localtime=True)
    else:
        msg["Date"] = formatdate(localtime=True)

    msg["MIME-Version"] = "1.0"

    # Body part
    if body_type == "html":
        body_part = MIMEText(body_content, "html", "utf-8")
    else:
        body_part = MIMEText(body_content, "plain", "utf-8")
    msg.attach(body_part)

    # Fetch and attach any attachments
    if post.get("hasAttachments"):
        att_url = (
            f"{GRAPH_URL}/groups/{group_id}/threads/{thread_id}"
            f"/posts/{post_id}/attachments"
        )
        att_resp = requests.get(att_url, headers=headers, timeout=60)
        if att_resp.status_code == 200:
            for att in att_resp.json().get("value", []):
                att_name = att.get("name", "attachment")
                att_content_type = att.get(
                    "contentType", "application/octet-stream"
                )
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

    return msg


def get_group_post_detail(access_token, group_id, thread_id, post_id):
    """Fetch full detail of a single post within a group conversation thread.

    Returns a message-like dict compatible with save_email_to_db, or None.
    """
    headers = {"Authorization": f"Bearer {access_token}"}
    url = (
        f"{GRAPH_URL}/groups/{group_id}/threads/{quote(thread_id, safe='')}"
        f"/posts/{quote(post_id, safe='')}"
        f"?$select=id,from,receivedDateTime,body,hasAttachments"
    )
    resp = requests.get(url, headers=headers, timeout=30)
    if resp.status_code != 200:
        return None
    post = resp.json()

    # Fetch thread topic for subject
    tresp = requests.get(
        f"{GRAPH_URL}/groups/{group_id}/threads/{quote(thread_id, safe='')}?$select=topic",
        headers=headers, timeout=10,
    )
    subject = tresp.json().get("topic", "(no subject)") if tresp.status_code == 200 else "(no subject)"

    return {
        "id": f"{thread_id}||{post_id}",
        "internetMessageId": post.get("id", f"{thread_id}||{post_id}"),
        "subject": subject,
        "from": post.get("from") or {},
        "toRecipients": post.get("toRecipients") or [],
        "ccRecipients": post.get("ccRecipients") or [],
        "receivedDateTime": post.get("receivedDateTime", ""),
        "body": post.get("body") or {},
        "hasAttachments": post.get("hasAttachments", False),
    }


def get_group_post_attachments(access_token, group_id, thread_id, post_id):
    """Fetch attachments for a single post within a group conversation thread."""
    headers = {"Authorization": f"Bearer {access_token}"}
    attachments = []
    url = (
        f"{GRAPH_URL}/groups/{group_id}/threads/{quote(thread_id, safe='')}"
        f"/posts/{quote(post_id, safe='')}/attachments"
        f"?$select=id,name,contentType,size,isInline&$top=1"
    )
    while url:
        resp = requests.get(url, headers=headers, timeout=60)
        if resp.status_code != 200:
            break
        data = resp.json()
        for att in data.get("value", []):
            att_id = att.get("id")
            if att_id:
                value_resp = requests.get(
                    f"{GRAPH_URL}/groups/{group_id}/threads/{quote(thread_id, safe='')}"
                    f"/posts/{quote(post_id, safe='')}/attachments/{att_id}/$value",
                    headers=headers,
                    timeout=60,
                )
                if value_resp.status_code == 200:
                    att["contentBytes"] = base64.b64encode(value_resp.content).decode()
            attachments.append(att)
        url = data.get("@odata.nextLink")
    return attachments


def get_group_post_mime(access_token, group_id, thread_id, post_id):
    """Build an EML from a specific post within a group conversation thread."""
    headers = {"Authorization": f"Bearer {access_token}"}
    url = (
        f"{GRAPH_URL}/groups/{group_id}/threads/{quote(thread_id, safe='')}"
        f"/posts/{quote(post_id, safe='')}"
        f"?$select=id,body,from,receivedDateTime,hasAttachments"
    )
    resp = requests.get(url, headers=headers, timeout=60)
    if resp.status_code != 200:
        return None
    post = resp.json()

    tresp = requests.get(
        f"{GRAPH_URL}/groups/{group_id}/threads/{quote(thread_id, safe='')}?$select=topic",
        headers=headers, timeout=10,
    )
    subject = tresp.json().get("topic", "(no subject)") if tresp.status_code == 200 else "(no subject)"

    gresp = requests.get(
        f"{GRAPH_URL}/groups/{group_id}?$select=mail",
        headers=headers, timeout=10,
    )
    group_email = gresp.json().get("mail", "") if gresp.status_code == 200 else ""

    msg = _build_post_eml(headers, group_id, thread_id, post, subject, group_email)
    return msg.as_bytes()


# Kept for backward compatibility (was used for mailbox-type group messages)
def get_group_message_mime(access_token, group_id, message_id):
    """Download the raw MIME content of an individual message from a group's
    shared mailbox via GET /users/{group_id}/messages/{message_id}/$value.

    Returns bytes or None on failure.
    """
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"{GRAPH_URL}/users/{group_id}/messages/{message_id}/$value"
    resp = requests.get(url, headers=headers, timeout=60)
    if resp.status_code != 200:
        return None
    return resp.content


def get_group_thread_mime(access_token, group_id, thread_id):
    """Build an EML file from the first post in a group thread.

    The /posts/{id}/$value endpoint is not reliably supported for group
    posts, so instead we:
      1. GET the first post with full body and sender info
      2. GET thread metadata for subject and group email
      3. GET any attachments on the post
      4. Construct a proper MIME email with full headers (To, Message-ID,
         MIME-Version, Date, etc.)

    Returns bytes or None on failure.
    """
    headers = {"Authorization": f"Bearer {access_token}"}

    # 1. Get the first post in the thread
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

    # 2. Get thread metadata (topic + group email for To header)
    thread_url = (
        f"{GRAPH_URL}/groups/{group_id}/threads/{thread_id}"
        f"?$select=topic"
    )
    tresp = requests.get(thread_url, headers=headers, timeout=10)
    subject = "(no subject)"
    if tresp.status_code == 200:
        subject = tresp.json().get("topic", "(no subject)")

    # Get group email for the To header
    group_url = f"{GRAPH_URL}/groups/{group_id}?$select=mail"
    gresp = requests.get(group_url, headers=headers, timeout=10)
    group_email = ""
    if gresp.status_code == 200:
        group_email = gresp.json().get("mail", "")

    # 3. Build and return the EML
    msg = _build_post_eml(
        headers, group_id, thread_id, posts[0],
        subject, group_email,
    )
    return msg.as_bytes()


def get_messages_page(access_token, user_email, next_link=None, top=50, group_id=None):
    """Fetch one page of message summaries for archiving.

    Returns (messages_list, next_link_or_None, group_id_or_None).
    messages_list contains dicts with keys: id, subject, hasAttachments.

    If group_id is provided, fetches directly from that group's mailbox
    without an extra resolution round-trip.
    """
    headers = {"Authorization": f"Bearer {access_token}"}

    if next_link:
        # Continuation URL — only occurs for regular mailbox pagination.
        response = requests.get(next_link, headers=headers, timeout=30)
        if response.status_code != 200:
            return None, None, None
        data = response.json()
        messages = [
            {
                "id": m.get("id", ""),
                "subject": m.get("subject", "(no subject)"),
                "hasAttachments": m.get("hasAttachments", False),
            }
            for m in data.get("value", [])
        ]
        return messages, data.get("@odata.nextLink"), None

    # If group_id is supplied directly, skip resolution and fetch group posts.
    if not group_id:
        group_id, _name = resolve_group_id(access_token, user_email)

    if group_id:
        posts = _get_group_all_posts(headers, group_id, max_posts=500)
        summaries = [
            {
                "id": p["id"],
                "subject": p["subject"],
                "hasAttachments": p["hasAttachments"],
            }
            for p in posts
        ]
        return summaries, None, group_id

    return None, None, None


def get_message_detail(access_token, message_id, group_id=None):
    """Fetch full details of a single message (body, recipients, etc.).

    When group_id is provided, reads from the group's shared mailbox via
    GET /users/{group_id}/messages/{message_id} instead of /me/messages.
    """
    headers = {"Authorization": f"Bearer {access_token}"}
    select = (
        "id,internetMessageId,subject,from,toRecipients,ccRecipients,"
        "receivedDateTime,body,hasAttachments"
    )
    mailbox = f"users/{group_id}" if group_id else "me"
    url = f"{GRAPH_URL}/{mailbox}/messages/{message_id}?$select={select}"
    response = requests.get(url, headers=headers, timeout=30)
    if response.status_code != 200:
        return None
    return response.json()


def get_message_attachments(access_token, message_id, group_id=None):
    """Fetch attachment metadata and content for a single message.

    Paginates one attachment at a time ($top=1) to keep each JSON response
    small, since Graph API may ignore $select and include contentBytes.
    Downloads binary content via the /$value endpoint to avoid JSON parsing
    problems entirely — binary responses have no parse limit.

    When group_id is provided, reads from the group's shared mailbox via
    /users/{group_id}/messages/{message_id}/attachments.
    """
    headers = {"Authorization": f"Bearer {access_token}"}
    attachments = []

    mailbox = f"users/{group_id}" if group_id else "me"
    url = (
        f"{GRAPH_URL}/{mailbox}/messages/{message_id}/attachments"
        "?$select=id,name,contentType,size,isInline&$top=1"
    )

    while url:
        response = requests.get(url, headers=headers, timeout=60)
        if response.status_code != 200:
            break
        try:
            data = response.json()
        except Exception:
            import logging
            logging.getLogger(__name__).warning(
                "Attachment list JSON parse failed for message %s (status=%s, body_start=%r)",
                message_id, response.status_code, response.text[:200],
            )
            break

        for att in data.get("value", []):
            att_id = att.get("id")
            if att_id:
                # Fetch binary content via /$value — returns raw bytes, not
                # base64 JSON, so there is no JSON size/truncation concern.
                value_resp = requests.get(
                    f"{GRAPH_URL}/{mailbox}/messages/{message_id}/attachments/{att_id}/$value",
                    headers=headers,
                    timeout=60,
                )
                if value_resp.status_code == 200:
                    att["contentBytes"] = base64.b64encode(value_resp.content).decode()
            attachments.append(att)

        url = data.get("@odata.nextLink")

    return attachments


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

    group_message_type = result_info.get("groupMessageType", "thread")
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for idx, msg in enumerate(messages, start=1):
            subject = _sanitize_filename(msg.get("subject", "email"))
            filename = f"{idx:02d}_{subject}.eml"

            if source == "user":
                mime_bytes = get_message_mime(
                    access_token, email, msg["id"]
                )
            elif group_message_type == "post" and "||" in msg["id"]:
                tid, pid = msg["id"].split("||", 1)
                mime_bytes = get_group_post_mime(access_token, group_id, tid, pid)
            elif group_message_type == "mailbox":
                mime_bytes = get_group_message_mime(
                    access_token, group_id, msg["id"]
                )
            else:
                mime_bytes = get_group_thread_mime(
                    access_token, group_id, msg["id"]
                )

            if mime_bytes:
                zf.writestr(filename, mime_bytes)
            else:
                zf.writestr(
                    filename.replace(".eml", "_ERROR.txt"),
                    f"Failed to download: {msg.get('subject', 'Unknown')}\n",
                )

    zip_buffer.seek(0)
    return zip_buffer


def search_teams_chats(access_token, query):
    """Search the signed-in user's group chats by topic name.

    Returns a list of chat dicts with keys: id, topic, chatType,
    lastUpdatedDateTime. Falls back to fetching all group chats and
    filtering client-side if the tenant does not support $search on chats.
    Returns {"error": ...} on unrecoverable failure.
    """
    headers = {"Authorization": f"Bearer {access_token}"}
    url = (
        f"{GRAPH_URL}/me/chats"
        f'?$search="{query}"'
        f"&$filter=chatType eq 'group'"
        f"&$select=id,topic,chatType,lastUpdatedDateTime"
        f"&$top=20"
    )
    response = requests.get(url, headers=headers, timeout=15)

    if response.status_code == 200:
        return response.json().get("value", [])

    # Some tenants do not support $search on /me/chats — fall back to a
    # plain fetch with client-side filtering.
    if response.status_code in (400, 501):
        fb = requests.get(f"{GRAPH_URL}/me/chats", headers=headers, timeout=15)
        if fb.status_code == 200:
            q_lower = query.lower()
            return [
                c for c in fb.json().get("value", [])
                if c.get("chatType") == "group"
                and q_lower in (c.get("topic") or "").lower()
            ]
        return {"error": {"fallback_status": fb.status_code, "detail": fb.json()}}

    return {"error": response.json()}


def get_teams_chat_members_count(access_token, chat_id):
    """Return the number of members in a Teams chat, or None on failure."""
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"{GRAPH_URL}/chats/{chat_id}/members?$select=id"
    response = requests.get(url, headers=headers, timeout=10)
    if response.status_code != 200:
        return None
    return len(response.json().get("value", []))


def get_teams_messages_page(access_token, chat_id, next_link=None, top=50):
    """Fetch one page of messages from a Teams chat.

    Returns (messages_list, next_link_or_None).
    Each message dict has: id, messageType, createdDateTime,
    senderName, senderEmail, contentHtml.

    Note: Graph API returns the sender's displayName but not their email
    address in chatMessage resources. senderEmail is stored as empty string.
    """
    headers = {"Authorization": f"Bearer {access_token}"}

    if next_link:
        url = next_link
    else:
        url = f"{GRAPH_URL}/chats/{chat_id}/messages?$top={top}"

    response = requests.get(url, headers=headers, timeout=30)
    if response.status_code != 200:
        return None, {"graph_status": response.status_code, "detail": response.json()}

    data = response.json()
    messages = []
    for m in data.get("value", []):
        from_user = (m.get("from") or {}).get("user") or {}

        messages.append({
            "id": m.get("id", ""),
            "messageType": m.get("messageType", "message"),
            "createdDateTime": m.get("createdDateTime", ""),
            "senderName": from_user.get("displayName", ""),
            "senderEmail": "",  # not available in chatMessage resource
            "contentHtml": (m.get("body") or {}).get("content", ""),
        })

    next_link_out = data.get("@odata.nextLink")
    return messages, next_link_out


def get_joined_teams(access_token):
    """Return all Teams the signed-in user is a member of.

    Returns a list of dicts with keys: id, displayName, description.
    Returns {"error": ...} on failure.
    """
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"{GRAPH_URL}/me/joinedTeams?$select=id,displayName,description"
    response = requests.get(url, headers=headers, timeout=15)
    if response.status_code == 200:
        return response.json().get("value", [])
    return {"error": response.json()}


def get_team_channels(access_token, team_id):
    """Return all channels in a Team.

    Returns a list of dicts with keys: id, displayName, description.
    Returns {"error": ...} on failure.
    """
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"{GRAPH_URL}/teams/{team_id}/channels?$select=id,displayName,description"
    response = requests.get(url, headers=headers, timeout=15)
    if response.status_code == 200:
        return response.json().get("value", [])
    return {"error": response.json()}


def get_channel_messages_page(access_token, team_id, channel_id, next_link=None, top=50):
    """Fetch one page of top-level posts from a Teams channel.

    Returns (messages_list, next_link_or_None).
    Each message dict has: id, messageType, createdDateTime,
    senderName, senderEmail, contentHtml, subject, importance, webUrl.

    Note: senderEmail is empty string — not available in channelMessage resource.
    """
    headers = {"Authorization": f"Bearer {access_token}"}

    if next_link:
        url = next_link
    else:
        url = f"{GRAPH_URL}/teams/{team_id}/channels/{channel_id}/messages?$top={top}"

    response = requests.get(url, headers=headers, timeout=30)
    if response.status_code != 200:
        return None, {"graph_status": response.status_code, "detail": response.json()}

    data = response.json()
    messages = []
    for m in data.get("value", []):
        from_user = (m.get("from") or {}).get("user") or {}

        messages.append({
            "id": m.get("id", ""),
            "messageType": m.get("messageType", "message"),
            "createdDateTime": m.get("createdDateTime", ""),
            "senderName": from_user.get("displayName", ""),
            "senderEmail": "",  # not available in channelMessage resource
            "contentHtml": (m.get("body") or {}).get("content", ""),
            "subject": m.get("subject") or "",
            "importance": m.get("importance") or "normal",
            "webUrl": m.get("webUrl") or "",
        })

    next_link_out = data.get("@odata.nextLink")
    return messages, next_link_out


def search_teams_by_name(access_token, query):
    """Return joined Teams whose displayName contains the query string (case-insensitive).

    Returns a filtered list of {id, displayName, description} dicts.
    Returns {"error": ...} on failure.
    """
    teams = get_joined_teams(access_token)
    if isinstance(teams, dict) and "error" in teams:
        return teams
    q_lower = query.lower()
    return [t for t in teams if q_lower in (t.get("displayName") or "").lower()]


def resolve_group_id(access_token, group_email):
    """Look up a Microsoft 365 Group by its email address.

    Returns (group_id, display_name) or (None, None) if not found.
    """
    headers = {"Authorization": f"Bearer {access_token}"}
    url = (
        f"{GRAPH_URL}/groups"
        f"?$filter=mail eq '{group_email}'"
        f"&$select=id,displayName,mail"
    )
    resp = requests.get(url, headers=headers, timeout=10)
    if resp.status_code != 200:
        return None, None
    groups = resp.json().get("value", [])
    if not groups:
        return None, None
    return groups[0]["id"], groups[0].get("displayName", "")


def get_user_groups(access_token):
    """Return all Microsoft 365 Groups (Unified groups with mailboxes) the user is a member of.

    Returns a list of {id, displayName, mail} dicts, sorted by displayName.
    """
    headers = {"Authorization": f"Bearer {access_token}"}
    url = (
        f"{GRAPH_URL}/me/memberOf/microsoft.graph.group"
        f"?$filter=groupTypes/any(c:c eq 'Unified')"
        f"&$select=id,displayName,mail"
        f"&$top=100"
    )
    groups = []
    while url:
        resp = requests.get(url, headers=headers, timeout=15)
        if resp.status_code != 200:
            break
        data = resp.json()
        for g in data.get("value", []):
            if g.get("mail"):
                groups.append({"id": g["id"], "displayName": g.get("displayName", ""), "mail": g["mail"]})
        url = data.get("@odata.nextLink")
    groups.sort(key=lambda g: g["displayName"].lower())
    return groups


def get_group_calendar_events_page(access_token, group_id, next_link=None, top=50):
    """Fetch one page of calendar events from a Microsoft 365 Group.

    Returns (events_list, next_link_or_None).
    Each event dict has: id, subject, start, end, organizer, isOnlineMeeting, bodyPreview, location.
    """
    headers = {"Authorization": f"Bearer {access_token}"}

    if next_link:
        url = next_link
    else:
        select = "id,subject,start,end,organizer,isOnlineMeeting,bodyPreview,location"
        url = (
            f"{GRAPH_URL}/groups/{group_id}/events"
            f"?$select={select}"
            f"&$orderby=start/dateTime desc"
            f"&$top={top}"
        )

    resp = requests.get(url, headers=headers, timeout=30)
    if resp.status_code != 200:
        return None, {"graph_status": resp.status_code, "detail": resp.json()}

    data = resp.json()
    events = []
    for e in data.get("value", []):
        org = (e.get("organizer") or {}).get("emailAddress") or {}
        loc = (e.get("location") or {}).get("displayName", "")
        events.append({
            "id": e.get("id", ""),
            "subject": e.get("subject", "(no subject)"),
            "start": e.get("start", {}),
            "end": e.get("end", {}),
            "organizer": org.get("name", "") or org.get("address", ""),
            "isOnlineMeeting": e.get("isOnlineMeeting", False),
            "bodyPreview": e.get("bodyPreview", ""),
            "location": loc,
        })

    next_link_out = data.get("@odata.nextLink")
    return events, next_link_out


def get_calendar_event_detail(access_token, group_id, event_id):
    """Fetch full details of a single group calendar event.

    Returns the full event dict or None on failure.
    """
    headers = {"Authorization": f"Bearer {access_token}"}
    select = (
        "id,subject,start,end,location,body,organizer,attendees,"
        "isOnlineMeeting,onlineMeetingUrl,onlineMeeting,webLink"
    )
    url = f"{GRAPH_URL}/groups/{group_id}/events/{event_id}?$select={select}"
    resp = requests.get(url, headers=headers, timeout=30)
    if resp.status_code != 200:
        return None
    return resp.json()


def download_emails_zip_progress(access_token, result_info):
    """Generator that yields SSE events while building the ZIP.

    Each yield is a Server-Sent Event string:
      - progress events:  data: {"current": 3, "total": 20, "subject": "..."}
      - done event:       data: {"done": true, "file": "<temp_path>"}
      - error event:      data: {"error": "..."}

    The final ZIP is written to a temp file whose path is sent in the
    done event so the client can fetch it via a separate download route.
    """
    messages = result_info.get("messages", [])
    source = result_info.get("source", "user")
    email = result_info.get("email", "")
    group_id = result_info.get("groupId", "")
    group_message_type = result_info.get("groupMessageType", "thread")
    total = len(messages)

    try:
        tmp = tempfile.NamedTemporaryFile(
            delete=False, suffix=".zip", prefix="emails_"
        )
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
            for idx, msg in enumerate(messages, start=1):
                subject = msg.get("subject", "email")

                # Send progress event
                evt = json.dumps({
                    "current": idx,
                    "total": total,
                    "subject": subject,
                })
                yield f"data: {evt}\n\n"

                safe_subject = _sanitize_filename(subject)
                filename = f"{idx:02d}_{safe_subject}.eml"

                if source == "user":
                    mime_bytes = get_message_mime(
                        access_token, email, msg["id"]
                    )
                elif group_message_type == "post" and "||" in msg["id"]:
                    tid, pid = msg["id"].split("||", 1)
                    mime_bytes = get_group_post_mime(access_token, group_id, tid, pid)
                elif group_message_type == "mailbox":
                    mime_bytes = get_group_message_mime(
                        access_token, group_id, msg["id"]
                    )
                else:
                    mime_bytes = get_group_thread_mime(
                        access_token, group_id, msg["id"]
                    )

                if mime_bytes:
                    zf.writestr(filename, mime_bytes)
                else:
                    zf.writestr(
                        filename.replace(".eml", "_ERROR.txt"),
                        f"Failed to download: {subject}\n",
                    )

        tmp.close()

        done_evt = json.dumps({"done": True, "file": tmp.name})
        yield f"data: {done_evt}\n\n"

    except Exception as e:
        err_evt = json.dumps({"error": str(e)})
        yield f"data: {err_evt}\n\n"


