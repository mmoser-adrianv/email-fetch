(function () {
    var searchInput = document.getElementById("people-search");
    var dropdown = document.getElementById("suggestions-dropdown");
    var goBtn = document.getElementById("go-btn");
    var stopBtn = document.getElementById("stop-btn");
    var progressPanel = document.getElementById("progress-panel");
    var logWrap = document.getElementById("ingest-log-wrap");
    var logBody = document.getElementById("ingest-log-body");
    var countProcessed = document.getElementById("count-processed");
    var countSaved = document.getElementById("count-saved");
    var countSkipped = document.getElementById("count-skipped");
    var searchSpinner = document.getElementById("search-spinner");
    var statusText = document.getElementById("status-text");
    var progressBar = document.getElementById("progress-bar");

    var selectedEmail = null;
    var debounceTimer = null;
    var stopped = false;
    var isRunning = false;
    var rowIndex = 0;

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
        selectedEmail = null;
        goBtn.disabled = true;

        if (query.length < 2) {
            dropdown.style.display = "none";
            if (searchSpinner) searchSpinner.style.display = "none";
            return;
        }

        if (searchSpinner) searchSpinner.style.display = "block";

        debounceTimer = setTimeout(function () {
            fetch("/api/people/search?q=" + encodeURIComponent(query))
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

    function renderDropdown(people) {
        dropdown.innerHTML = "";
        if (!people || people.length === 0 || people.error) {
            dropdown.style.display = "none";
            return;
        }
        people.forEach(function (person) {
            var item = document.createElement("div");
            item.className = "suggestion-item";
            item.textContent = person.displayName + " (" + person.email + ")";
            item.addEventListener("click", function () {
                selectedEmail = person.email;
                searchInput.value = person.displayName + " (" + person.email + ")";
                dropdown.style.display = "none";
                goBtn.disabled = false;
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

    // ── Ingest loop ──────────────────────────────────────────────────────────

    goBtn.addEventListener("click", function () {
        if (!selectedEmail) return;

        // Reset state
        stopped = false;
        isRunning = true;
        rowIndex = 0;
        logBody.innerHTML = "";
        countProcessed.textContent = "0";
        countSaved.textContent = "0";
        countSkipped.textContent = "0";
        setStatus("Starting…", true);

        goBtn.style.display = "none";
        stopBtn.style.display = "";
        progressPanel.style.display = "";
        logWrap.style.display = "";

        acquireWakeLock();
        startTokenRefresh();
        fetchPage(selectedEmail, null);
    });

    stopBtn.addEventListener("click", function () {
        stopped = true;
        setStatus("Stopping after current email…", true);
    });

    function fetchPage(email, nextLink) {
        if (stopped) {
            finish("Stopped by user.");
            return;
        }

        var url = "/api/ingest/page?email=" + encodeURIComponent(email);
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
                    finish("Done — no more emails.");
                    return;
                }

                processMessages(email, messages, 0, next);
            })
            .catch(function (err) {
                finish("Network error: " + err.message);
            });
    }

    function processMessages(email, messages, idx, nextLink) {
        if (stopped) {
            finish("Stopped by user.");
            return;
        }
        if (idx >= messages.length) {
            // Batch done — fetch next page
            if (nextLink) {
                fetchPage(email, nextLink);
            } else {
                finish("Done — all emails processed.");
            }
            return;
        }

        var msg = messages[idx];
        rowIndex++;
        var currentRow = rowIndex;
        setStatus("Processing: " + escapeHtml(msg.subject || "(no subject)"), true);

        fetch("/api/ingest/run", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ messageId: msg.id }),
        })
            .then(function (res) {
                if (res.status === 401) { window.location.href = "/login"; return null; }
                return res.json();
            })
            .then(function (result) {
                if (!result) return;

                var processed = parseInt(countProcessed.textContent, 10) + 1;
                countProcessed.textContent = processed;

                var label, labelClass;
                if (result.error) {
                    label = "Error";
                    labelClass = "badge-error";
                } else if (result.skipped) {
                    label = "Skipped";
                    labelClass = "badge-skip";
                    countSkipped.textContent = parseInt(countSkipped.textContent, 10) + 1;
                } else {
                    label = "Saved";
                    labelClass = "badge-saved";
                    countSaved.textContent = parseInt(countSaved.textContent, 10) + 1;
                }

                var attText = result.attachmentCount > 0
                    ? result.attachmentCount + " attachment(s)"
                    : "—";

                var tr = document.createElement("tr");
                tr.innerHTML = "<td>" + currentRow + "</td>"
                    + "<td>" + escapeHtml(result.subject || msg.subject || "") + "</td>"
                    + "<td>" + escapeHtml(attText) + "</td>"
                    + "<td><span class='badge " + labelClass + "'>" + label + "</span></td>";
                logBody.insertBefore(tr, logBody.firstChild);

                processMessages(email, messages, idx + 1, nextLink);
            })
            .catch(function (err) {
                finish("Network error: " + err.message);
            });
    }

    function finish(message) {
        isRunning = false;
        releaseWakeLock();
        stopTokenRefresh();
        setStatus(message, false);
        stopBtn.style.display = "none";
        goBtn.style.display = "";
        goBtn.disabled = false;
    }

    function escapeHtml(text) {
        var div = document.createElement("div");
        div.textContent = text || "";
        return div.innerHTML;
    }
})();
