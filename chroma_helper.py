from urllib.parse import urlparse

import app_config

_chroma_client = None
_openai_client = None
_collection = None


def _get_collection():
    global _chroma_client, _collection
    if _collection is None:
        import chromadb
        parsed = urlparse(app_config.CHROMA_URL)
        host = parsed.hostname or "127.0.0.1"
        port = parsed.port or 8001
        ssl = parsed.scheme == "https"
        _chroma_client = chromadb.HttpClient(host=host, port=port, ssl=ssl)
        _collection = _chroma_client.get_or_create_collection("emails")
    return _collection


def _get_openai():
    global _openai_client
    if _openai_client is None:
        from openai import OpenAI
        _openai_client = OpenAI(api_key=app_config.OPENAI_API_KEY)
    return _openai_client


def _build_embedding_text(email):
    parts = [
        f"Subject: {email.subject or ''}",
        f"From: {email.sender_name or ''} <{email.sender_email or ''}>",
        f"To: {email.recipients or ''}",
        f"Date: {email.date_received.isoformat() if email.date_received else ''}",
        "",
        email.body_text or "",
    ]
    for att in email.attachments:
        if att.extracted_text:
            parts.append(f"\n--- Attachment: {att.filename} ---")
            parts.append(att.extracted_text)
    return "\n".join(parts)


def embed_and_upsert(email):
    """Generate an OpenAI embedding for an email and upsert it into ChromaDB."""
    text = _build_embedding_text(email)
    text = text[:24000]  # ~6,000 tokens, safely under the 8,192 token limit

    response = _get_openai().embeddings.create(
        input=text,
        model="text-embedding-3-small",
    )
    embedding = response.data[0].embedding

    _get_collection().upsert(
        ids=[str(email.id)],
        embeddings=[embedding],
        documents=[text[:2000]],
        metadatas=[{
            "email_id": email.id,
            "subject": email.subject or "",
            "sender_email": email.sender_email or "",
            "sender_name": email.sender_name or "",
            "date_received": email.date_received.isoformat() if email.date_received else "",
            "recipients": email.recipients or "",
        }],
    )


def search_emails(query, n_results=10):
    """Embed a query and return the nearest emails from ChromaDB."""
    response = _get_openai().embeddings.create(
        input=query,
        model="text-embedding-3-small",
    )
    embedding = response.data[0].embedding

    results = _get_collection().query(
        query_embeddings=[embedding],
        n_results=n_results,
        include=["metadatas", "distances", "documents"],
    )

    output = []
    for i, meta in enumerate(results["metadatas"][0]):
        output.append({
            **meta,
            "distance": results["distances"][0][i],
            "snippet": results["documents"][0][i][:300] if results["documents"][0] else "",
        })
    return output
