const chatWindow = document.getElementById("chat-window");
const thinkingEl = document.getElementById("thinking-indicator");
const chatInput  = document.getElementById("chat-input");
const sendBtn    = document.getElementById("send-btn");
const newChatBtn = document.getElementById("new-chat-btn");
const errorEl    = document.getElementById("chat-error");
const sessionList = document.getElementById("session-list");

let activeSessionId = null;
let sessions = [...INITIAL_SESSIONS];

// ── Sidebar rendering ──────────────────────────────────────────

function renderSidebar() {
    sessionList.innerHTML = "";
    sessions.forEach(s => {
        const item = document.createElement("div");
        item.className = "session-item" + (s.id === activeSessionId ? " active" : "");
        item.dataset.id = s.id;

        const title = document.createElement("span");
        title.className = "session-item-title";
        title.textContent = s.title || "Untitled";
        title.title = s.title || "Untitled";

        const del = document.createElement("button");
        del.className = "session-item-delete";
        del.textContent = "×";
        del.title = "Delete";
        del.addEventListener("click", (e) => {
            e.stopPropagation();
            deleteSession(s.id);
        });

        item.appendChild(title);
        item.appendChild(del);
        item.addEventListener("click", () => loadSession(s.id));
        sessionList.appendChild(item);
    });
}

async function refreshSidebar() {
    try {
        const resp = await fetch("/api/chat/sessions");
        if (resp.ok) {
            sessions = await resp.json();
            renderSidebar();
        }
    } catch (_) {}
}

// ── Session loading ────────────────────────────────────────────

async function loadSession(id) {
    try {
        const resp = await fetch(`/api/chat/sessions/${id}`);
        if (!resp.ok) return;
        const data = await resp.json();

        activeSessionId = id;
        clearChat();
        data.messages.forEach(m => appendMessage(m.role, m.content));
        renderSidebar();
        chatWindow.scrollTop = chatWindow.scrollHeight;
    } catch (_) {}
}

async function deleteSession(id) {
    try {
        await fetch(`/api/chat/sessions/${id}`, { method: "DELETE" });
        sessions = sessions.filter(s => s.id !== id);
        if (activeSessionId === id) {
            activeSessionId = null;
            clearChat();
        }
        renderSidebar();
    } catch (_) {}
}

// ── Chat UI helpers ────────────────────────────────────────────

function clearChat() {
    Array.from(chatWindow.children).forEach(el => {
        if (el !== thinkingEl) el.remove();
    });
    const emptyState = document.createElement("div");
    emptyState.id = "empty-state";
    emptyState.innerHTML = `
        <svg width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5">
            <path d="M21 15a2 2 0 0 1-2 2H7l-4 4V5a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2z"/>
        </svg>
        <p>Email Chat</p>
        <small>Ask anything about your emails</small>`;
    chatWindow.insertBefore(emptyState, thinkingEl);
    errorEl.style.display = "none";
    chatInput.value = "";
    sendBtn.disabled = true;
}

function appendMessage(role, text) {
    const emptyState = document.getElementById("empty-state");
    if (emptyState) emptyState.remove();

    const div = document.createElement("div");
    if (role === "user") {
        div.className = "msg-user";
        const bubble = document.createElement("div");
        bubble.className = "msg-user-bubble";
        bubble.textContent = text;
        div.appendChild(bubble);
    } else {
        div.className = "msg-assistant";
        const content = document.createElement("div");
        content.className = "msg-assistant-content";
        content.innerHTML = marked.parse(text);
        div.appendChild(content);
    }
    chatWindow.insertBefore(div, thinkingEl);
    chatWindow.scrollTop = chatWindow.scrollHeight;
    return div;
}

// ── Input handling ─────────────────────────────────────────────

chatInput.addEventListener("input", () => {
    sendBtn.disabled = chatInput.value.trim() === "";
});

chatInput.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !sendBtn.disabled) {
        e.preventDefault();
        sendMessage();
    }
});

sendBtn.addEventListener("click", sendMessage);

newChatBtn.addEventListener("click", async () => {
    await fetch("/api/chat/reset", { method: "POST" }).catch(() => {});
    activeSessionId = null;
    clearChat();
    renderSidebar();
});

// ── Send message ───────────────────────────────────────────────

async function sendMessage() {
    const message = chatInput.value.trim();
    if (!message) return;

    errorEl.style.display = "none";
    chatInput.value = "";
    sendBtn.disabled = true;
    chatInput.disabled = true;

    appendMessage("user", message);

    thinkingEl.style.display = "block";
    chatWindow.scrollTop = chatWindow.scrollHeight;

    const emptyState = document.getElementById("empty-state");
    if (emptyState) emptyState.remove();

    const assistantBubble = document.createElement("div");
    assistantBubble.className = "msg-assistant";
    const assistantContent = document.createElement("div");
    assistantContent.className = "msg-assistant-content";
    assistantBubble.appendChild(assistantContent);
    assistantBubble._rawText = "";
    assistantBubble.style.display = "none";
    chatWindow.insertBefore(assistantBubble, thinkingEl);

    try {
        const resp = await fetch("/api/chat", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ message }),
        });

        if (!resp.ok) {
            const errData = await resp.json().catch(() => ({}));
            throw new Error(errData.error || `HTTP ${resp.status}`);
        }

        const reader = resp.body.getReader();
        const decoder = new TextDecoder();
        let buffer = "";

        while (true) {
            const { done, value } = await reader.read();
            if (done) break;

            buffer += decoder.decode(value, { stream: true });
            const parts = buffer.split("\n\n");
            buffer = parts.pop();

            for (const part of parts) {
                const line = part.trim();
                if (!line.startsWith("data:")) continue;
                const jsonStr = line.slice("data:".length).trim();
                let evt;
                try { evt = JSON.parse(jsonStr); } catch { continue; }

                if (evt.error) {
                    throw new Error(evt.error);
                }
                if (evt.done) {
                    if (evt.session_id && evt.session_id !== activeSessionId) {
                        activeSessionId = evt.session_id;
                    }
                    await refreshSidebar();
                    break;
                }
                if (evt.delta) {
                    thinkingEl.style.display = "none";
                    assistantBubble.style.display = "";
                    assistantBubble._rawText += evt.delta;
                    assistantContent.innerHTML = marked.parse(assistantBubble._rawText);
                    chatWindow.scrollTop = chatWindow.scrollHeight;
                }
            }
        }
    } catch (err) {
        if (!assistantBubble.textContent) assistantBubble.remove();
        errorEl.textContent = "Error: " + err.message;
        errorEl.style.display = "";
    } finally {
        thinkingEl.style.display = "none";
        chatInput.disabled = false;
        sendBtn.disabled = chatInput.value.trim() === "";
        chatInput.focus();
    }
}

// ── Init ───────────────────────────────────────────────────────
renderSidebar();
