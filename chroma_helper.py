from urllib.parse import urlparse

import app_config

DISTANCE_THRESHOLD = 1.2  # L2 distance; ChromaDB default metric is L2

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
            "snippet": results["documents"][0][i][:800] if results["documents"][0] else "",
        })
    return output


def search_emails_filtered(query, n_results=20, where_clause=None):
    """Embed a query and return nearest emails from ChromaDB, with optional metadata filtering."""
    response = _get_openai().embeddings.create(
        input=query,
        model="text-embedding-3-small",
    )
    embedding = response.data[0].embedding

    kwargs = {
        "query_embeddings": [embedding],
        "n_results": n_results,
        "include": ["metadatas", "distances", "documents"],
    }
    if where_clause:
        kwargs["where"] = where_clause

    results = _get_collection().query(**kwargs)

    output = []
    for i, meta in enumerate(results["metadatas"][0]):
        output.append({
            **meta,
            "distance": results["distances"][0][i],
            "snippet": results["documents"][0][i][:800] if results["documents"][0] else "",
        })
    return output


def format_results_as_context(results, distance_threshold=DISTANCE_THRESHOLD):
    """Format a list of search result dicts into a context string for the LLM. Returns '' if empty."""
    filtered = [r for r in results if r.get("distance", 999) <= distance_threshold]
    if not filtered:
        return ""
    parts = []
    for i, r in enumerate(filtered, 1):
        lines = [
            f"[Email {i}]",
            f"Subject: {r.get('subject', '(no subject)')}",
            f"From: {r.get('sender_name', '')} <{r.get('sender_email', '')}>",
            f"Date: {r.get('date_received', '')}",
            f"Excerpt: {r.get('snippet', '')[:1000]}",
        ]
        parts.append("\n".join(lines))
    return "\n\n".join(parts)


def search_emails_multi_query(queries, n_results_each=10, where_clause=None):
    """Run multiple queries and return deduplicated results ranked by best (lowest) distance."""
    seen = {}
    for q in queries:
        try:
            results = search_emails_filtered(q, n_results=n_results_each, where_clause=where_clause)
        except Exception:
            continue
        for r in results:
            eid = r.get("email_id")
            if eid not in seen or r["distance"] < seen[eid]["distance"]:
                seen[eid] = r
    return sorted(seen.values(), key=lambda x: x["distance"])


def get_context_for_chat(query, n_results=10, where_clause=None, distance_threshold=DISTANCE_THRESHOLD):
    """Query ChromaDB and return a formatted email context string for RAG. Returns '' on any error."""
    try:
        if where_clause is not None:
            results = search_emails_filtered(query, n_results=n_results, where_clause=where_clause)
        else:
            results = search_emails(query, n_results=n_results)
    except Exception:
        return ""
    return format_results_as_context(results, distance_threshold=distance_threshold)
