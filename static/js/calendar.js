'use strict';

// ── State ─────────────────────────────────────────────────────────────────────
let isRunning = false;
let stopped   = false;
let processed = 0;
let saved     = 0;
let skipped   = 0;
let wakeLock  = null;

// ── DOM refs ──────────────────────────────────────────────────────────────────
const emailInput    = document.getElementById('email-input');
const scrapeBtn     = document.getElementById('scrape-btn');
const stopBtn       = document.getElementById('stop-btn');
const errorMsg      = document.getElementById('error-msg');
const progressPanel = document.getElementById('progress-panel');
const progressBar   = document.getElementById('progress-bar');
const statusText    = document.getElementById('status-text');
const logWrap       = document.getElementById('scrape-log-wrap');
const logBody       = document.getElementById('scrape-log-body');
const countProcessed = document.getElementById('count-processed');
const countSaved     = document.getElementById('count-saved');
const countSkipped   = document.getElementById('count-skipped');

// ── Helpers ───────────────────────────────────────────────────────────────────
function escapeHtml(str) {
    const d = document.createElement('div');
    d.textContent = str || '';
    return d.innerHTML;
}

function formatDate(dtStr) {
    if (!dtStr) return '—';
    try {
        return new Date(dtStr).toLocaleString(undefined, {
            year: 'numeric', month: 'short', day: 'numeric',
            hour: '2-digit', minute: '2-digit',
        });
    } catch { return dtStr; }
}

function setStatus(text) {
    statusText.textContent = text;
}

function showError(msg) {
    errorMsg.textContent = msg;
    errorMsg.style.display = msg ? 'block' : 'none';
}

function updateCounters() {
    countProcessed.textContent = processed;
    countSaved.textContent     = saved;
    countSkipped.textContent   = skipped;
}

function addLogRow(subject, start, organizer, isOnline, result) {
    const badge = result === 'saved'   ? '<span class="badge-saved">saved</span>'
                : result === 'skipped' ? '<span class="badge-skip">skip</span>'
                :                        '<span class="badge-error">error</span>';

    const onlineMark = isOnline ? '✓' : '—';

    const tr = document.createElement('tr');
    tr.innerHTML = `
        <td>${escapeHtml(subject)}</td>
        <td>${escapeHtml(formatDate(start))}</td>
        <td>${escapeHtml(organizer || '—')}</td>
        <td style="text-align:center;">${onlineMark}</td>
        <td>${badge}</td>
    `;
    logBody.insertBefore(tr, logBody.firstChild);
}

// ── Wake lock ─────────────────────────────────────────────────────────────────
async function acquireWakeLock() {
    if ('wakeLock' in navigator) {
        try { wakeLock = await navigator.wakeLock.request('screen'); } catch { /* ignore */ }
    }
}

function releaseWakeLock() {
    if (wakeLock) { wakeLock.release(); wakeLock = null; }
}

document.addEventListener('visibilitychange', () => {
    if (isRunning && document.visibilityState === 'visible') acquireWakeLock();
});

// ── Token refresh ─────────────────────────────────────────────────────────────
setInterval(async () => {
    if (!isRunning) return;
    try {
        const r = await fetch('/api/auth/token-refresh', { method: 'POST' });
        if (r.status === 401) window.location.href = '/login';
    } catch { /* ignore */ }
}, 4 * 60 * 1000);

// ── UI state helpers ──────────────────────────────────────────────────────────
function startUI() {
    isRunning = true;
    stopped   = false;
    processed = saved = skipped = 0;
    updateCounters();

    emailInput.disabled = true;
    scrapeBtn.style.display = 'none';
    stopBtn.style.display   = '';
    showError('');

    progressPanel.style.display = '';
    progressPanel.classList.add('is-running');
    progressBar.style.display   = '';
    logWrap.style.display       = '';
    logBody.innerHTML           = '';
}

function finishUI(msg) {
    isRunning = false;
    releaseWakeLock();

    emailInput.disabled = false;
    scrapeBtn.style.display = '';
    stopBtn.style.display   = 'none';

    progressPanel.classList.remove('is-running');
    progressBar.style.display = 'none';
    setStatus(msg);
}

// ── Main scrape loop ──────────────────────────────────────────────────────────
scrapeBtn.addEventListener('click', async () => {
    const email = emailInput.value.trim();
    if (!email) { showError('Please enter a group email address.'); return; }

    startUI();
    await acquireWakeLock();

    // 1. Resolve group email → group ID
    setStatus('Resolving group email…');
    let groupId;
    try {
        const r = await fetch(`/api/calendar/resolve?email=${encodeURIComponent(email)}`);
        if (r.status === 401) { window.location.href = '/login'; return; }
        const data = await r.json();
        if (data.error) {
            finishUI('');
            showError(data.error);
            return;
        }
        groupId = data.groupId;
        setStatus(`Found group: ${data.displayName || email}. Fetching events…`);
    } catch (e) {
        finishUI('');
        showError('Network error while resolving group email.');
        return;
    }

    // 2. Paginate + process events
    await processPage(groupId, email, null);
});

async function processPage(groupId, searchedEmail, nextLink) {
    if (stopped) {
        finishUI(`Stopped. Processed ${processed}, saved ${saved}, skipped ${skipped}.`);
        return;
    }

    let pageData;
    try {
        const url = `/api/calendar/events/page?groupId=${encodeURIComponent(groupId)}`
            + (nextLink ? `&nextLink=${encodeURIComponent(nextLink)}` : '');
        const r = await fetch(url);
        if (r.status === 401) { window.location.href = '/login'; return; }
        pageData = await r.json();
        if (pageData.error) {
            finishUI(`Error fetching events: ${pageData.error}`);
            return;
        }
    } catch (e) {
        finishUI('Network error while fetching events page.');
        return;
    }

    const events = pageData.events || [];
    if (events.length === 0 && !pageData.nextLink) {
        finishUI(`Done. Processed ${processed}, saved ${saved}, skipped ${skipped}.`);
        return;
    }

    // Process each event in the page sequentially
    await processEvents(events, groupId, searchedEmail, 0, pageData.nextLink);
}

async function processEvents(events, groupId, searchedEmail, idx, nextLink) {
    if (stopped) {
        finishUI(`Stopped. Processed ${processed}, saved ${saved}, skipped ${skipped}.`);
        return;
    }

    if (idx >= events.length) {
        // Page done — fetch next page or finish
        if (nextLink) {
            setStatus(`Fetching next page… (${processed} processed so far)`);
            await processPage(groupId, searchedEmail, nextLink);
        } else {
            finishUI(`Done. Processed ${processed}, saved ${saved}, skipped ${skipped}.`);
        }
        return;
    }

    const ev = events[idx];
    setStatus(`Processing: ${ev.subject || '(no subject)'}`);

    let result = 'error';
    let startDt = (ev.start || {}).dateTime || '';
    let organizer = ev.organizer || '';

    try {
        const r = await fetch('/api/calendar/events/run', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                groupId: groupId,
                eventId: ev.id,
                searchedEmail: searchedEmail,
            }),
        });
        if (r.status === 401) { window.location.href = '/login'; return; }
        const data = await r.json();
        if (data.error) {
            result = 'error';
        } else if (data.saved) {
            result = 'saved';
            saved++;
        } else {
            result = 'skipped';
            skipped++;
        }
        startDt = data.start || startDt;
    } catch (e) {
        result = 'error';
    }

    processed++;
    updateCounters();
    addLogRow(ev.subject || '(no subject)', startDt, organizer, ev.isOnlineMeeting, result);

    // Recurse to next event
    await processEvents(events, groupId, searchedEmail, idx + 1, nextLink);
}

// ── Stop button ───────────────────────────────────────────────────────────────
stopBtn.addEventListener('click', () => {
    stopped = true;
    stopBtn.disabled = true;
    setStatus('Stopping after current event…');
});
