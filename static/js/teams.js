(function () {
    var searchInput = document.getElementById("chat-search");
    var dropdown = document.getElementById("suggestions-dropdown");
    var scrapeBtn = document.getElementById("scrape-btn");
    var stopBtn = document.getElementById("stop-btn");
    var progressPanel = document.getElementById("progress-panel");
    var logWrap = document.getElementById("scrape-log-wrap");
    var logBody = document.getElementById("scrape-log-body");
    var countProcessed = document.getElementById("count-processed");
    var countSaved = document.getElementById("count-saved");
    var countSkipped = document.getElementById("count-skipped");
    var searchSpinner = document.getElementById("search-spinner");
    var statusText = document.getElementById("status-text");
    var progressBar = document.getElementById("progress-bar");

    var selectedChat = null;  // {id, topic}
    var debounceTimer = null;
    var stopped = false;
    var isRunning = false;
    var batchIndex = 0;

    // ── Wake Lock ─────────────────────────────────────────────────────────────

    var wakeLock = null;

    function acquireWakeLock() {
        if ("wakeLock" in navigator) {
            navigator.wakeLock.request("screen").then(function (lock) {
                wakeLock = lock;
            }).catch(function () {});
        }
    }

    function releaseWakeLock() {
        if (wakeLock) { wakeLock.release(); wakeLock = null; }
    }

    document.addEventListener("visibilitychange", function () {
        if (document.visibilityState === "visible" && isRunning && wakeLock === null) {
            acquireWakeLock();
        }
    });

    // ── Token refresh ─────────────────────────────────────────────────────────

    var tokenRefreshInterval = null;

    function startTokenRefresh() {
        tokenRefreshInterval = setInterval(function () {
            fetch("/api/auth/token-refresh", { method: "POST" })
                .then(function (res) {
                    if (res.status === 401) { window.location.href = "/login"; }
                });
        }, 4 * 60 * 1000); // every 4 minutes
    }

    function stopTokenRefresh() {
        if (tokenRefreshInterval) { clearInterval(tokenRefreshInterval); tokenRefreshInterval = null; }
    }

    // ── Status helper ─────────────────────────────────────────────────────────

    function setStatus(text, spinning) {
        if (statusText) statusText.textContent = text;
        if (progressBar) progressBar.style.display = spinning ? "" : "none";
        if (progressPanel) progressPanel.classList.toggle("is-running", spinning);
    }

    // ── Autocomplete ──────────────────────────────────────────────────────────

    searchInput.addEventListener("input", function () {
        var query = this.value.trim();
        clearTimeout(debounceTimer);
        selectedChat = null;
        scrapeBtn.disabled = true;

        if (query.length < 2) {
            dropdown.style.display = "none";
            if (searchSpinner) searchSpinner.style.display = "none";
            return;
        }

        if (searchSpinner) searchSpinner.style.display = "block";

        debounceTimer = setTimeout(function () {
            fetch("/api/teams/chats/search?q=" + encodeURIComponent(query))
                .then(function (res) {
                    if (res.status === 401) { window.location.href = "/login"; return null; }
                    return res.json();
                })
                .then(function (data) {
                    if (searchSpinner) searchSpinner.style.display = "none";
                    if (data) renderDropdown(data);
                })
                .catch(function () { if (searchSpinner) searchSpinner.style.display = "none"; });
        }, 300);
    });

    function renderDropdown(chats) {
        dropdown.innerHTML = "";
        if (!chats || chats.length === 0 || chats.error) {
            dropdown.style.display = "none";
            return;
        }
        chats.forEach(function (chat) {
            var item = document.createElement("div");
            item.className = "suggestion-item";
            item.textContent = chat.topic || "(unnamed chat)";
            item.addEventListener("click", function () {
                selectedChat = { id: chat.id, topic: chat.topic || "" };
                searchInput.value = chat.topic || "(unnamed chat)";
                dropdown.style.display = "none";
                scrapeBtn.disabled = false;
            });
            dropdown.appendChild(item);
        });
        dropdown.style.display = "block";
    }

    document.addEventListener("click", function (e) {
        if (!searchInput.contains(e.target) && !dropdown.contains(e.target)) {
            dropdown.style.display = "none";
        }
    });

    // ── Scrape loop ───────────────────────────────────────────────────────────

    scrapeBtn.addEventListener("click", function () {
        if (!selectedChat) return;

        stopped = false;
        isRunning = true;
        batchIndex = 0;
        logBody.innerHTML = "";
        countProcessed.textContent = "0";
        countSaved.textContent = "0";
        countSkipped.textContent = "0";
        setStatus("Starting…", true);

        scrapeBtn.style.display = "none";
        stopBtn.style.display = "";
        progressPanel.style.display = "";
        logWrap.style.display = "";

        acquireWakeLock();
        startTokenRefresh();
        fetchPage(selectedChat.id, null);
    });

    stopBtn.addEventListener("click", function () {
        stopped = true;
        setStatus("Stopping after current batch…", true);
    });

    function fetchPage(chatId, nextLink) {
        if (stopped) {
            finish("Stopped by user.");
            return;
        }

        var url = "/api/teams/messages/page?chatId=" + encodeURIComponent(chatId);
        if (nextLink) url += "&nextLink=" + encodeURIComponent(nextLink);

        setStatus("Fetching next batch…", true);

        fetch(url)
            .then(function (res) {
                if (res.status === 401) { window.location.href = "/login"; return null; }
                return res.json();
            })
            .then(function (data) {
                if (!data) return;
                if (data.error) {
                    finish("Error: " + escapeHtml(data.error));
                    return;
                }
                var messages = data.messages || [];
                var next = data.nextLink || null;

                if (messages.length === 0) {
                    markComplete(chatId);
                    return;
                }

                saveBatch(chatId, selectedChat.topic, messages, next);
            })
            .catch(function (err) {
                finish("Network error: " + err.message);
            });
    }

    function saveBatch(chatId, chatTopic, messages, nextLink) {
        if (stopped) {
            markComplete(chatId);
            return;
        }

        batchIndex++;
        var currentBatch = batchIndex;

        // Build date range label from the batch
        var dates = messages
            .map(function (m) { return m.createdDateTime || ""; })
            .filter(Boolean)
            .sort();
        var dateLabel = dates.length > 0
            ? dates[0].slice(0, 10) + (dates.length > 1 ? " – " + dates[dates.length - 1].slice(0, 10) : "")
            : "—";

        setStatus("Saving batch " + currentBatch + "…", true);

        fetch("/api/teams/messages/save", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ chatId: chatId, chatTopic: chatTopic, messages: messages }),
        })
            .then(function (res) {
                if (res.status === 401) { window.location.href = "/login"; return null; }
                return res.json();
            })
            .then(function (result) {
                if (!result) return;

                var batchSaved = result.saved || 0;
                var batchSkipped = result.skipped || 0;
                var batchTotal = batchSaved + batchSkipped;

                countProcessed.textContent = parseInt(countProcessed.textContent, 10) + batchTotal;
                countSaved.textContent = parseInt(countSaved.textContent, 10) + batchSaved;
                countSkipped.textContent = parseInt(countSkipped.textContent, 10) + batchSkipped;

                var labelClass = batchSaved > 0 ? "badge-saved" : "badge-skip";
                var label = batchSaved + " saved, " + batchSkipped + " skipped";

                var tr = document.createElement("tr");
                tr.innerHTML = "<td>" + currentBatch + "</td>"
                    + "<td>" + escapeHtml(dateLabel) + "</td>"
                    + "<td>" + batchTotal + "</td>"
                    + "<td><span class='badge " + labelClass + "'>" + escapeHtml(label) + "</span></td>";
                logBody.insertBefore(tr, logBody.firstChild);

                if (nextLink && !stopped) {
                    fetchPage(chatId, nextLink);
                } else {
                    markComplete(chatId);
                }
            })
            .catch(function (err) {
                finish("Save error: " + err.message);
            });
    }

    function markComplete(chatId) {
        fetch("/api/teams/chats/complete?chatId=" + encodeURIComponent(chatId), { method: "POST" })
            .then(function () {
                finish("Done — all messages processed.");
            })
            .catch(function () {
                finish("Done — all messages processed.");
            });
    }

    function finish(message) {
        isRunning = false;
        releaseWakeLock();
        stopTokenRefresh();
        setStatus(message, false);
        stopBtn.style.display = "none";
        scrapeBtn.style.display = "";
        scrapeBtn.disabled = false;
    }

    function escapeHtml(text) {
        var div = document.createElement("div");
        div.textContent = text || "";
        return div.innerHTML;
    }

})();
