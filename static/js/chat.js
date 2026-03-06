const chatWindow = document.getElementById("chat-window");
const thinkingEl = document.getElementById("thinking-indicator");
const chatInput  = document.getElementById("chat-input");
const sendBtn    = document.getElementById("send-btn");
const newChatBtn = document.getElementById("new-chat-btn");
const errorEl    = document.getElementById("chat-error");

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

newChatBtn.addEventListener("click", () => {
    fetch("/api/chat/reset", { method: "POST" }).catch(() => {});
    Array.from(chatWindow.children).forEach(el => {
        if (el !== thinkingEl) el.remove();
    });
    errorEl.style.display = "none";
    chatInput.value = "";
    sendBtn.disabled = true;
});

function appendMessage(role, text) {
    const div = document.createElement("div");
    div.className = role === "user" ? "msg-user" : "msg-assistant";
    if (role === "user") {
        div.textContent = text;
    } else {
        div.innerHTML = marked.parse(text);
    }
    chatWindow.insertBefore(div, thinkingEl);
    chatWindow.scrollTop = chatWindow.scrollHeight;
    return div;
}

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

    const assistantBubble = document.createElement("div");
    assistantBubble.className = "msg-assistant";
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
                    break;
                }
                if (evt.delta) {
                    thinkingEl.style.display = "none";
                    assistantBubble.style.display = "";
                    assistantBubble._rawText += evt.delta;
                    assistantBubble.innerHTML = marked.parse(assistantBubble._rawText);
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
