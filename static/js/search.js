function esc(str) {
    return String(str ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
}

function formatDate(iso) {
    if (!iso) return "";
    try {
        return new Date(iso).toLocaleString();
    } catch {
        return iso;
    }
}

function formatScore(distance) {
    // Convert cosine distance to a 0-100% similarity score
    const similarity = Math.round((1 - distance) * 100);
    return `${similarity}% match`;
}

function renderResults(results) {
    const container = document.getElementById("results-container");
    const noResults = document.getElementById("no-results");

    container.innerHTML = "";

    if (!results.length) {
        noResults.style.display = "";
        return;
    }
    noResults.style.display = "none";

    results.forEach(r => {
        const card = document.createElement("div");
        card.className = "result-card";
        card.innerHTML = `
            <div class="result-subject">
                <span class="result-score">${esc(formatScore(r.distance))}</span>
                ${esc(r.subject || "(no subject)")}
            </div>
            <div class="result-meta">
                From: ${esc(r.sender_name || r.sender_email || "")}
                ${r.sender_name && r.sender_email ? `&lt;${esc(r.sender_email)}&gt;` : ""}
                &nbsp;&middot;&nbsp;
                ${esc(formatDate(r.date_received))}
            </div>
            ${r.snippet ? `<div class="result-snippet">${esc(r.snippet)}</div>` : ""}
        `;
        container.appendChild(card);
    });
}

document.getElementById("search-form").addEventListener("submit", async (e) => {
    e.preventDefault();

    const query = document.getElementById("query-input").value.trim();
    if (!query) return;

    const errorEl = document.getElementById("search-error");
    const spinner = document.getElementById("loading-spinner");
    const noResults = document.getElementById("no-results");
    const container = document.getElementById("results-container");
    const btn = document.getElementById("search-btn");

    errorEl.style.display = "none";
    noResults.style.display = "none";
    container.innerHTML = "";
    spinner.style.display = "";
    btn.disabled = true;

    try {
        const resp = await fetch(`/api/search?q=${encodeURIComponent(query)}`);
        const data = await resp.json();

        if (!resp.ok || data.error) {
            errorEl.textContent = data.error || "Search failed.";
            errorEl.style.display = "";
            return;
        }

        renderResults(data.results || []);
    } catch (err) {
        errorEl.textContent = "Network error. Please try again.";
        errorEl.style.display = "";
    } finally {
        spinner.style.display = "none";
        btn.disabled = false;
    }
});
