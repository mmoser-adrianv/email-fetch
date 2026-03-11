(function () {
    var groupSelect = document.getElementById("group-select");
    var goBtn = document.getElementById("go-btn");
    var stopBtn = document.getElementById("stop-btn");
    var progressPanel = document.getElementById("progress-panel");
    var logWrap = document.getElementById("ingest-log-wrap");
    var logBody = document.getElementById("ingest-log-body");
    var countProcessed = document.getElementById("count-processed");
    var countSaved = document.getElementById("count-saved");
    var countSkipped = document.getElementById("count-skipped");
    var statusText = document.getElementById("status-text");
    var progressBar = document.getElementById("progress-bar");
    var projectSelect = document.getElementById("project-select");
    var projectTitleInput = document.getElementById("project-title");
    var projectNumberInput = document.getElementById("project-number");
    var downloadLink = document.getElementById("download-link");

    function updateDownloadLink() {
        var pid = projectSelect.value;
        if (pid) {
            downloadLink.href = "/api/emails/export?project_id=" + encodeURIComponent(pid);
            downloadLink.textContent = "Download Project Emails (JSON)";
        } else {
            downloadLink.href = "/api/emails/export";
            downloadLink.textContent = "Download All Emails (JSON)";
        }
    }

    var selectedEmail = null;
    var stopped = false;
    var isRunning = false;
    var rowIndex = 0;
    var currentProjectId = null;
    var currentGroupId = null;

    // ── Projects ──────────────────────────────────────────────────────────────

    function loadProjects() {
        fetch("/api/projects")
            .then(function (res) { return res.json(); })
            .then(function (projects) {
                projects.forEach(function (p) {
                    var opt = document.createElement("option");
                    opt.value = p.id;
                    opt.textContent = p.title + (p.project_number ? " (" + p.project_number + ")" : "");
                    opt.dataset.number = p.project_number || "";
                    projectSelect.appendChild(opt);
                });
            })
            .catch(function () {});
    }

    projectSelect.addEventListener("change", function () {
        var selected = this.value;
        if (selected) {
            var opt = this.options[this.selectedIndex];
            projectTitleInput.value = opt.textContent.replace(/ \(.*\)$/, "");
            projectNumberInput.value = opt.dataset.number || "";
            projectTitleInput.disabled = true;
            projectNumberInput.disabled = true;
        } else {
            projectTitleInput.value = "";
            projectNumberInput.value = "";
            projectTitleInput.disabled = false;
            projectNumberInput.disabled = false;
        }
        updateDownloadLink();
    });

    loadProjects();

    // ── Groups dropdown ───────────────────────────────────────────────────────

    function loadGroups() {
        fetch("/api/groups")
            .then(function (res) {
                if (res.status === 401) { window.location.href = "/login"; return null; }
                return res.json();
            })
            .then(function (groups) {
                if (!groups) return;
                groupSelect.innerHTML = "";
                if (groups.length === 0) {
                    groupSelect.innerHTML = "<option value=''>No group mailboxes found</option>";
                    return;
                }
                var placeholder = document.createElement("option");
                placeholder.value = "";
                placeholder.textContent = "Select a group mailbox…";
                groupSelect.appendChild(placeholder);
                groups.forEach(function (group) {
                    var opt = document.createElement("option");
                    opt.value = group.id;
                    opt.textContent = group.displayName + " (" + group.mail + ")";
                    opt.dataset.mail = group.mail;
                    groupSelect.appendChild(opt);
                });
                groupSelect.disabled = false;
            })
            .catch(function () {
                groupSelect.innerHTML = "<option value=''>Failed to load groups</option>";
            });
    }

    groupSelect.addEventListener("change", function () {
        var opt = this.options[this.selectedIndex];
        if (this.value) {
            selectedEmail = opt.dataset.mail;
            currentGroupId = this.value;
            goBtn.disabled = false;
        } else {
            selectedEmail = null;
            currentGroupId = null;
            goBtn.disabled = true;
        }
    });

    loadGroups();

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

    // ── Ingest loop ───────────────────────────────────────────────────────────

    goBtn.addEventListener("click", function () {
        if (!selectedEmail) return;

        stopped = false;
        isRunning = true;
        rowIndex = 0;
        currentProjectId = null;
        logBody.innerHTML = "";
        countProcessed.textContent = "0";
        countSaved.textContent = "0";
        countSkipped.textContent = "0";
        setStatus("Starting…", true);

        goBtn.style.display = "none";
        stopBtn.style.display = "";
        progressPanel.style.display = "";
        logWrap.style.display = "";
        groupSelect.disabled = true;

        acquireWakeLock();
        startTokenRefresh();

        var existingProjectId = projectSelect.value;
        if (existingProjectId) {
            currentProjectId = parseInt(existingProjectId, 10);
            fetchPage(selectedEmail, null);
        } else {
            var title = projectTitleInput.value.trim();
            if (title) {
                fetch("/api/projects", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ title: title, project_number: projectNumberInput.value.trim() }),
                })
                    .then(function (res) { return res.json(); })
                    .then(function (data) {
                        if (data.id) {
                            currentProjectId = data.id;
                            var opt = document.createElement("option");
                            opt.value = data.id;
                            opt.textContent = data.title + (data.project_number ? " (" + data.project_number + ")" : "");
                            opt.dataset.number = data.project_number || "";
                            projectSelect.appendChild(opt);
                            projectSelect.value = data.id;
                            projectTitleInput.disabled = true;
                            projectNumberInput.disabled = true;
                            updateDownloadLink();
                        }
                        fetchPage(selectedEmail, null);
                    })
                    .catch(function () { fetchPage(selectedEmail, null); });
            } else {
                fetchPage(selectedEmail, null);
            }
        }
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

        var url = "/api/ingest/page?email=" + encodeURIComponent(email)
            + "&groupId=" + encodeURIComponent(currentGroupId);
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
            body: JSON.stringify({ messageId: msg.id, searchedEmail: email, projectId: currentProjectId, groupId: currentGroupId }),
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
                var processed = parseInt(countProcessed.textContent, 10) + 1;
                countProcessed.textContent = processed;

                var tr = document.createElement("tr");
                tr.innerHTML = "<td>" + currentRow + "</td>"
                    + "<td>" + escapeHtml(msg.subject || "") + "</td>"
                    + "<td>—</td>"
                    + "<td><span class='badge badge-error'>Error</span></td>";
                logBody.insertBefore(tr, logBody.firstChild);

                processMessages(email, messages, idx + 1, nextLink);
            });
    }

    function finish(message) {
        isRunning = false;
        releaseWakeLock();
        stopTokenRefresh();
        setStatus(message, false);
        stopBtn.style.display = "none";
        goBtn.style.display = "";
        goBtn.disabled = !selectedEmail;
        groupSelect.disabled = false;
    }

    function escapeHtml(text) {
        var div = document.createElement("div");
        div.textContent = text || "";
        return div.innerHTML;
    }
})();
