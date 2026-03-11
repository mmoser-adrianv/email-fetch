(function () {
    var teamSearch = document.getElementById("team-search");
    var teamDropdown = document.getElementById("team-suggestions-dropdown");
    var channelSelect = document.getElementById("channel-select");
    var scrapeBtn = document.getElementById("channel-scrape-btn");
    var stopBtn = document.getElementById("channel-stop-btn");
    var progressPanel = document.getElementById("channel-progress-panel");
    var logWrap = document.getElementById("channel-scrape-log-wrap");
    var logBody = document.getElementById("channel-scrape-log-body");
    var countProcessed = document.getElementById("ch-count-processed");
    var countSaved = document.getElementById("ch-count-saved");
    var countSkipped = document.getElementById("ch-count-skipped");
    var teamSearchSpinner = document.getElementById("team-search-spinner");
    var statusText = document.getElementById("channel-status-text");
    var progressBar = document.getElementById("channel-progress-bar");

    var selectedTeam = null;   // {id, displayName}
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
        }, 4 * 60 * 1000);
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

    // ── Team autocomplete ─────────────────────────────────────────────────────

    teamSearch.addEventListener("input", function () {
        var query = this.value.trim();
        clearTimeout(debounceTimer);
        selectedTeam = null;
        channelSelect.innerHTML = "<option value=''>— select a channel —</option>";
        channelSelect.disabled = true;
        scrapeBtn.disabled = true;

        if (query.length < 2) {
            teamDropdown.style.display = "none";
            if (teamSearchSpinner) teamSearchSpinner.style.display = "none";
            return;
        }

        if (teamSearchSpinner) teamSearchSpinner.style.display = "block";

        debounceTimer = setTimeout(function () {
            fetch("/api/teams/teams/search?q=" + encodeURIComponent(query))
                .then(function (res) {
                    if (res.status === 401) { window.location.href = "/login"; return null; }
                    return res.json();
                })
                .then(function (data) {
                    if (teamSearchSpinner) teamSearchSpinner.style.display = "none";
                    if (data) renderTeamDropdown(data);
                })
                .catch(function () { if (teamSearchSpinner) teamSearchSpinner.style.display = "none"; });
        }, 300);
    });

    function renderTeamDropdown(teams) {
        teamDropdown.innerHTML = "";
        if (!teams || teams.length === 0 || teams.error) {
            teamDropdown.style.display = "none";
            return;
        }
        teams.forEach(function (team) {
            var item = document.createElement("div");
            item.className = "suggestion-item";
            item.textContent = team.displayName || "(unnamed team)";
            item.addEventListener("click", function () {
                selectedTeam = { id: team.id, displayName: team.displayName || "" };
                teamSearch.value = team.displayName || "(unnamed team)";
                teamDropdown.style.display = "none";
                loadChannels(team.id);
            });
            teamDropdown.appendChild(item);
        });
        teamDropdown.style.display = "block";
    }

    document.addEventListener("click", function (e) {
        if (!teamSearch.contains(e.target) && !teamDropdown.contains(e.target)) {
            teamDropdown.style.display = "none";
        }
    });

    // ── Channel loading ────────────────────────────────────────────────────────

    function loadChannels(teamId) {
        channelSelect.innerHTML = "<option value=''>Loading channels…</option>";
        channelSelect.disabled = true;
        scrapeBtn.disabled = true;

        fetch("/api/teams/teams/channels?teamId=" + encodeURIComponent(teamId))
            .then(function (res) {
                if (res.status === 401) { window.location.href = "/login"; return null; }
                return res.json();
            })
            .then(function (data) {
                channelSelect.innerHTML = "<option value=''>— select a channel —</option>";
                if (!data || data.error || data.length === 0) {
                    channelSelect.innerHTML = "<option value=''>No channels found</option>";
                    return;
                }
                data.forEach(function (ch) {
                    var opt = document.createElement("option");
                    opt.value = ch.id;
                    opt.dataset.name = ch.displayName || "";
                    opt.textContent = ch.displayName || "(unnamed)";
                    channelSelect.appendChild(opt);
                });
                channelSelect.disabled = false;
            })
            .catch(function () {
                channelSelect.innerHTML = "<option value=''>Error loading channels</option>";
            });
    }

    channelSelect.addEventListener("change", function () {
        scrapeBtn.disabled = !this.value;
    });

    // ── Scrape loop ───────────────────────────────────────────────────────────

    scrapeBtn.addEventListener("click", function () {
        if (!selectedTeam || !channelSelect.value) return;

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

        var channelId = channelSelect.value;
        var channelName = channelSelect.options[channelSelect.selectedIndex].dataset.name || "";
        fetchPage(selectedTeam.id, channelId, channelName, null);
    });

    stopBtn.addEventListener("click", function () {
        stopped = true;
        setStatus("Stopping after current batch…", true);
    });

    function fetchPage(teamId, channelId, channelName, nextLink) {
        if (stopped) {
            markComplete(channelId);
            return;
        }

        var url = "/api/teams/channel/messages/page"
            + "?teamId=" + encodeURIComponent(teamId)
            + "&channelId=" + encodeURIComponent(channelId);
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
                    markComplete(channelId);
                    return;
                }

                saveBatch(teamId, channelId, channelName, messages, next);
            })
            .catch(function (err) {
                finish("Network error: " + err.message);
            });
    }

    function saveBatch(teamId, channelId, channelName, messages, nextLink) {
        if (stopped) {
            markComplete(channelId);
            return;
        }

        batchIndex++;
        var currentBatch = batchIndex;

        var dates = messages
            .map(function (m) { return m.createdDateTime || ""; })
            .filter(Boolean)
            .sort();
        var dateLabel = dates.length > 0
            ? dates[0].slice(0, 10) + (dates.length > 1 ? " – " + dates[dates.length - 1].slice(0, 10) : "")
            : "—";

        setStatus("Saving batch " + currentBatch + "…", true);

        fetch("/api/teams/channel/posts/save", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                teamId: teamId,
                channelId: channelId,
                teamName: selectedTeam.displayName,
                channelName: channelName,
                posts: messages,
            }),
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
                    fetchPage(teamId, channelId, channelName, nextLink);
                } else {
                    markComplete(channelId);
                }
            })
            .catch(function (err) {
                finish("Save error: " + err.message);
            });
    }

    function markComplete(channelId) {
        fetch("/api/teams/channel/complete?channelId=" + encodeURIComponent(channelId), { method: "POST" })
            .then(function () {
                finish("Done — all posts processed.");
            })
            .catch(function () {
                finish("Done — all posts processed.");
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
