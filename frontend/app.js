/* ─────────────────────────────────────────────────────────────────────────
   Interview Assistant — Frontend Application Logic
   ───────────────────────────────────────────────────────────────────────── */

const API_BASE = 'http://127.0.0.1:3001/api';

// ── Browser guard — Chrome / Edge only ────────────────────────────────────
const isChrome = /Chrome\//.test(navigator.userAgent) && /Google Inc/.test(navigator.vendor);
const isEdge   = /Edg\//.test(navigator.userAgent);
const SUPPORTED_BROWSER = isChrome || isEdge;
const SpeechRecognitionAPI = window.SpeechRecognition || window.webkitSpeechRecognition || null;

// ── State ──────────────────────────────────────────────────────────────────
const state = {
  role: '',
  sessionId: null,
  mode: 'ai',
  cacheSize: 0,
  interviewCount: 0,
  sessionHistory: [],
  isListening: false,
  recognition: null,
  continuousActive: false   // true while the "Ask Question" mode is ON
};

// ── DOM References ─────────────────────────────────────────────────────────
const $ = id => document.getElementById(id);

const roleSelect       = $('roleSelect');
const btnStart         = $('btnStart');
const setupPanel       = $('setupPanel');
const appMain          = $('appMain');
const roleBadge        = $('roleBadge');
const sessionInfo      = $('sessionInfo');
const btnMic           = $('btnMic');
const btnStop          = $('btnStop');
const micLabel         = $('micLabel');
const micStatus        = $('micStatus');
const btnEnd           = $('btnEnd');
const questionText     = $('questionText');
const pointsBody       = $('pointsBody');
const illustratedBody  = $('illustratedBody');
const sourceRow        = $('sourceRow');
const sourceTag        = $('sourceTag');
const loadingOverlay   = $('loadingOverlay');
const setupStats       = $('setupStats');
const headerMeta       = $('headerMeta');
const manualInput      = $('manualInput');
const btnSubmit        = $('btnSubmit');
const btnHistoryToggle = $('btnHistoryToggle');
const historyDrawer    = $('historyDrawer');
const historyList      = $('historyList');
const btnCloseHistory  = $('btnCloseHistory');

// ── Setup ──────────────────────────────────────────────────────────────────
window.addEventListener('DOMContentLoaded', () => {
  // Block unsupported browsers immediately
  if (!SUPPORTED_BROWSER) {
    document.body.innerHTML = `
      <div class="browser-block">
        <div class="browser-block-card">
          <div class="browser-block-icon">🚫</div>
          <h2>Unsupported Browser</h2>
          <p>Interview Assistant requires <strong>Google Chrome</strong> or <strong>Microsoft Edge</strong> for full voice support.</p>
          <div class="browser-links">
            <a href="https://www.google.com/chrome/" target="_blank">Download Chrome</a>
            <a href="https://www.microsoft.com/edge" target="_blank">Download Edge</a>
          </div>
        </div>
      </div>`;
    return;
  }

  loadStats();

  roleSelect.addEventListener('change', () => {
    btnStart.disabled = !roleSelect.value;
  });

  btnStart.addEventListener('click', startSession);
  btnEnd.addEventListener('click', endSession);
  btnSubmit.addEventListener('click', () => submitQuestion(manualInput.value.trim()));
  manualInput.addEventListener('keydown', e => {
    if (e.key === 'Enter') submitQuestion(manualInput.value.trim());
  });
  btnMic.addEventListener('click', activateContinuousListening);
  btnStop.addEventListener('click', deactivateContinuousListening);
  btnHistoryToggle.addEventListener('click', () => historyDrawer.classList.toggle('hidden'));
  btnCloseHistory.addEventListener('click', () => historyDrawer.classList.add('hidden'));
});

// ── Load Stats ─────────────────────────────────────────────────────────────
async function loadStats() {
  try {
    const res = await fetch(`${API_BASE}/stats`);
    if (!res.ok) return;
    const { stats } = await res.json();
    const chips = Object.entries(stats)
      .map(([role, count]) => `<span class="stat-chip">${role}: ${count} sessions</span>`)
      .join('');
    setupStats.innerHTML = chips
      ? `<div style="margin-top:8px">Previous sessions:</div>${chips}`
      : '';
  } catch {
    // Backend might not be running yet — silently ignore
  }
}

// ── Start Session ──────────────────────────────────────────────────────────
async function startSession() {
  const role = roleSelect.value;
  if (!role) return;

  showLoading(true);
  try {
    const res = await fetch(`${API_BASE}/session/start`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ role })
    });

    if (!res.ok) throw new Error('Failed to start session');
    const data = await res.json();

    state.role = role;
    state.sessionId = data.sessionId;
    state.mode = data.mode;
    state.cacheSize = data.cacheSize || 0;
    state.interviewCount = data.interviewCount;
    state.sessionHistory = [];

    // Switch view
    setupPanel.classList.add('hidden');
    appMain.classList.remove('hidden');

    roleBadge.textContent = role;
    updateSessionInfo();
    headerMeta.textContent = `Session #${data.interviewCount} • ${role}`;

  } catch (err) {
    alert('Could not connect to the backend. Make sure the server is running.\n\n' + err.message);
  } finally {
    showLoading(false);
  }
}

function updateSessionInfo() {
  let modeLabel;
  if (state.mode === 'cache+ai') {
    modeLabel = `📊 Excel first (${state.cacheSize} saved), AI fallback`;
  } else {
    modeLabel = `🤖 AI only (no Excel data yet for this role)`;
  }
  sessionInfo.textContent = `Session: ${state.sessionId?.substring(0, 8)}… • ${modeLabel}`;
}

// ── End Session ────────────────────────────────────────────────────────────
function endSession() {
  deactivateContinuousListening();
  setupPanel.classList.remove('hidden');
  appMain.classList.add('hidden');
  historyDrawer.classList.add('hidden');
  roleSelect.value = '';
  btnStart.disabled = true;
  headerMeta.textContent = '';
  loadStats();
}

// ── Submit Question ────────────────────────────────────────────────────────
async function submitQuestion(question) {
  if (!question) return;

  manualInput.value = '';
  questionText.textContent = question;
  sourceRow.classList.add('hidden');
  pointsBody.innerHTML = '<div class="placeholder-msg">Generating answer…</div>';
  illustratedBody.innerHTML = '<div class="placeholder-msg">Generating answer…</div>';
  showLoading(true);

  try {
    const res = await fetch(`${API_BASE}/answer`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        role: state.role,
        question,
        sessionId: state.sessionId
      })
    });

    if (!res.ok) {
      const err = await res.json();
      throw new Error(err.error || 'Server error');
    }

    const { points, illustrated, source, matched_role } = await res.json();

    renderPoints(points);
    renderIllustrated(illustrated);
    renderSource(source, matched_role);

    // Save to session history
    state.sessionHistory.push({ question, points, illustrated, source, matched_role });
    updateHistoryDrawer();

  } catch (err) {
    pointsBody.innerHTML = `<div class="placeholder-msg" style="color:#b91c1c">Error: ${err.message}</div>`;
    illustratedBody.innerHTML = `<div class="placeholder-msg" style="color:#b91c1c">Please try again.</div>`;
  } finally {
    showLoading(false);
    // Restart mic automatically if continuous mode is still active
    if (state.continuousActive && !state.isListening) {
      micStatus.textContent = '🔴 Live — ready for next question…';
      setTimeout(_startOneSpeechCycle, 600);
    }
  }
}

// ── Render Functions ───────────────────────────────────────────────────────
function renderPoints(points) {
  if (!points || points.length === 0) {
    pointsBody.innerHTML = '<div class="placeholder-msg">No key points available.</div>';
    return;
  }
  const items = points.map((p, i) =>
    `<li>
      <span class="point-num">${i + 1}</span>
      <span class="point-text">${escapeHtml(p)}</span>
    </li>`
  ).join('');
  pointsBody.innerHTML = `<ul class="points-list">${items}</ul>`;
}

function renderIllustrated(text) {
  if (!text) {
    illustratedBody.innerHTML = '<div class="placeholder-msg">No illustrated answer available.</div>';
    return;
  }
  illustratedBody.innerHTML = `<div class="illustrated-text">${escapeHtml(text)}</div>`;
}

function renderSource(source, matchedRole) {
  sourceRow.classList.remove('hidden');

  let label, cls;
  if (source === 'cache') {
    label = `Excel Cache — ${escapeHtml(matchedRole || '')}`;
    cls   = 'source-chip source-cache';
  } else if (source === 'cache_cross') {
    label = `Excel Cache (cross-role) — found in: ${escapeHtml(matchedRole || '')}`;
    cls   = 'source-chip source-cross';
  } else {
    label = `AI Generated — ${escapeHtml(matchedRole || '')}`;
    cls   = 'source-chip source-ai';
  }

  sourceTag.innerHTML = `<span class="${cls}">${label}</span>`;
}

// ── History Drawer ─────────────────────────────────────────────────────────
function updateHistoryDrawer() {
  if (state.sessionHistory.length === 0) {
    historyList.innerHTML = '<div style="color:var(--muted);font-size:13px;text-align:center;margin-top:20px">No questions yet this session.</div>';
    return;
  }
  historyList.innerHTML = state.sessionHistory.slice().reverse().map(item =>
    `<div class="history-item">
      <div class="hi-q">${escapeHtml(item.question)}</div>
      <div class="hi-pts">${(item.points || []).slice(0, 2).map(p => `• ${p}`).join('<br>')}</div>
    </div>`
  ).join('');
}

// ── Continuous Speech Recognition ─────────────────────────────────────────

/**
 * Activate continuous listening mode.
 * The mic stays on after each answer and restarts automatically.
 * Stopped only by clicking "Stop" or ending the session.
 */
function activateContinuousListening() {
  const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
  if (!SpeechRecognition) {
    alert('Speech recognition is not supported in this browser.\nPlease use Google Chrome or Microsoft Edge.\n\nYou can still type your question manually.');
    return;
  }
  if (state.continuousActive) return;

  state.continuousActive = true;

  // Show Stop button, update Ask Question button appearance
  btnStop.classList.remove('hidden');
  btnMic.classList.add('listening');
  micLabel.textContent = 'Listening…';
  micStatus.textContent = '🔴 Live — mic stays on';

  _startOneSpeechCycle();
}

/** Stop continuous listening mode entirely. */
function deactivateContinuousListening() {
  state.continuousActive = false;
  state.isListening = false;
  if (state.recognition) {
    try { state.recognition.stop(); } catch {}
    state.recognition = null;
  }
  btnStop.classList.add('hidden');
  btnMic.classList.remove('listening');
  micLabel.textContent = 'Ask Question';
  micStatus.textContent = '';
}

/**
 * Start a single speech recognition cycle.
 * When a final result arrives, submit it; when the cycle ends,
 * restart automatically if continuousActive is still true.
 */
function _startOneSpeechCycle() {
  if (!state.continuousActive) return;

  const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
  const recognition = new SpeechRecognition();
  recognition.lang = 'en-US';
  recognition.interimResults = true;
  recognition.maxAlternatives = 1;
  recognition.continuous = false;   // one utterance per cycle; we loop manually

  state.isListening = true;
  state.recognition = recognition;

  recognition.onstart = () => {
    btnMic.classList.add('listening');
    micLabel.textContent = 'Listening…';
    micStatus.textContent = '🔴 Live — speak your question';
  };

  recognition.onresult = (event) => {
    let interim = '';
    let final = '';
    for (let i = event.resultIndex; i < event.results.length; i++) {
      if (event.results[i].isFinal) {
        final += event.results[i][0].transcript;
      } else {
        interim += event.results[i][0].transcript;
      }
    }
    questionText.textContent = final || interim || 'Listening…';
    if (final) {
      // Stop this cycle; answer will restart the mic when done
      state.isListening = false;
      try { recognition.stop(); } catch {}
      submitQuestion(final.trim());
    }
  };

  recognition.onerror = (event) => {
    state.isListening = false;
    if (event.error === 'no-speech') {
      // No speech heard — silently restart
      if (state.continuousActive) {
        micStatus.textContent = '🔴 Live — no speech detected, retrying…';
        setTimeout(_startOneSpeechCycle, 300);
      }
    } else if (event.error === 'aborted') {
      // Deliberately stopped — do nothing
    } else {
      micStatus.textContent = `Mic error: ${event.error}`;
      if (state.continuousActive) setTimeout(_startOneSpeechCycle, 800);
    }
  };

  recognition.onend = () => {
    state.isListening = false;
    state.recognition = null;
    // If no question was submitted (silence / interruption), restart the cycle
    if (state.continuousActive) {
      setTimeout(_startOneSpeechCycle, 300);
    }
  };

  recognition.start();
}

// ── Utilities ──────────────────────────────────────────────────────────────
function showLoading(show) {
  loadingOverlay.classList.toggle('hidden', !show);
}

function escapeHtml(str) {
  if (typeof str !== 'string') return '';
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
