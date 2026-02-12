(function () {
    var searchInput = document.getElementById("people-search");
    var dropdown = document.getElementById("suggestions-dropdown");
    var fetchBtn = document.getElementById("fetch-emails-btn");
    var emailResults = document.getElementById("email-results");
    var selectedEmail = null;
    var debounceTimer = null;

    searchInput.addEventListener("input", function () {
        var query = this.value.trim();
        clearTimeout(debounceTimer);
        selectedEmail = null;
        fetchBtn.disabled = true;

        if (query.length < 2) {
            dropdown.style.display = "none";
            return;
        }

        debounceTimer = setTimeout(function () {
            fetch("/api/people/search?q=" + encodeURIComponent(query))
                .then(function (res) {
                    if (res.status === 401) {
                        window.location.href = "/login";
                        return [];
                    }
                    return res.json();
                })
                .then(function (data) {
                    if (data) renderDropdown(data);
                });
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
                fetchBtn.disabled = false;
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

    fetchBtn.addEventListener("click", function () {
        if (!selectedEmail) return;
        fetchBtn.disabled = true;
        fetchBtn.textContent = "Loading...";
        emailResults.innerHTML = "<p>Fetching emails...</p>";

        fetch("/api/messages?email=" + encodeURIComponent(selectedEmail))
            .then(function (res) {
                if (res.status === 401) {
                    window.location.href = "/login";
                    return null;
                }
                return res.json();
            })
            .then(function (data) {
                if (data) renderEmails(data);
                fetchBtn.disabled = false;
                fetchBtn.textContent = "Fetch Emails";
            });
    });

    function renderEmails(messages) {
        if (messages.error) {
            emailResults.innerHTML = "<p class='error'>Error: " + escapeHtml(JSON.stringify(messages.error)) + "</p>";
            return;
        }
        if (!Array.isArray(messages) || messages.length === 0) {
            emailResults.innerHTML = "<p>No messages found.</p>";
            return;
        }
        var html = "<table><thead><tr>"
            + "<th>From</th><th>Subject</th>"
            + "<th>Received</th><th>Preview</th>"
            + "</tr></thead><tbody>";
        messages.forEach(function (msg) {
            var date = new Date(msg.received).toLocaleString();
            html += "<tr>"
                + "<td>" + escapeHtml(msg.fromName || msg.from) + "</td>"
                + "<td>" + escapeHtml(msg.subject) + "</td>"
                + "<td>" + escapeHtml(date) + "</td>"
                + "<td>" + escapeHtml(msg.preview) + "</td>"
                + "</tr>";
        });
        html += "</tbody></table>";
        html += '<button id="download-zip-btn" class="btn btn-download">Download All as ZIP</button>';
        emailResults.innerHTML = html;

        // Attach download handler
        document.getElementById("download-zip-btn").addEventListener("click", function () {
            var dlBtn = this;
            dlBtn.disabled = true;
            dlBtn.textContent = "Preparing download...";

            fetch("/api/messages/download")
                .then(function (res) {
                    if (res.status === 401) {
                        window.location.href = "/login";
                        return null;
                    }
                    if (!res.ok) {
                        return res.json().then(function (err) {
                            throw new Error(err.error || "Download failed");
                        });
                    }
                    return res.blob();
                })
                .then(function (blob) {
                    if (!blob) return;
                    var url = URL.createObjectURL(blob);
                    var a = document.createElement("a");
                    a.href = url;
                    a.download = "emails.zip";
                    document.body.appendChild(a);
                    a.click();
                    document.body.removeChild(a);
                    URL.revokeObjectURL(url);
                    dlBtn.disabled = false;
                    dlBtn.textContent = "Download All as ZIP";
                })
                .catch(function (err) {
                    dlBtn.disabled = false;
                    dlBtn.textContent = "Download All as ZIP";
                    alert("Download failed: " + err.message);
                });
        });
    }

    function escapeHtml(text) {
        var div = document.createElement("div");
        div.textContent = text || "";
        return div.innerHTML;
    }
})();
