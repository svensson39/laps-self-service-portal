/**
 * app.js - LAPS Self-Service Portal frontend.
 *
 * Extracted from index.html so the Static Web App can serve a Content-Security-Policy
 * without 'unsafe-inline' in script-src. Loaded with defer, so the DOM is ready and the
 * MSAL bundle (a classic script earlier in the document) has already executed.
 */

  'use strict';

  // ── Theme management ──────────────────────────────────────────────────────
  function getPreferredTheme() {
    const stored = localStorage.getItem('laps-theme');
    if (stored) return stored;
    return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  }

  function applyTheme(theme) {
    document.documentElement.setAttribute('data-theme', theme);
    localStorage.setItem('laps-theme', theme);
    const moonPath = 'M12 3a9 9 0 1 0 9 9c0-.46-.04-.92-.1-1.36a5.389 5.389 0 0 1-4.4 2.26 5.403 5.403 0 0 1-3.14-9.8c-.44-.06-.9-.1-1.36-.1z';
    const sunPath = 'M12 7c-2.76 0-5 2.24-5 5s2.24 5 5 5 5-2.24 5-5-2.24-5-5-5zM2 13h2c.55 0 1-.45 1-1s-.45-1-1-1H2c-.55 0-1 .45-1 1s.45 1 1 1zm18 0h2c.55 0 1-.45 1-1s-.45-1-1-1h-2c-.55 0-1 .45-1 1s.45 1 1 1zM11 2v2c0 .55.45 1 1 1s1-.45 1-1V2c0-.55-.45-1-1-1s-1 .45-1 1zm0 18v2c0 .55.45 1 1 1s1-.45 1-1v-2c0-.55-.45-1-1-1s-1 .45-1 1zM5.99 4.58a.996.996 0 0 0-1.41 0 .996.996 0 0 0 0 1.41l1.06 1.06c.39.39 1.03.39 1.41 0s.39-1.03 0-1.41L5.99 4.58zm12.37 12.37a.996.996 0 0 0-1.41 0 .996.996 0 0 0 0 1.41l1.06 1.06c.39.39 1.03.39 1.41 0a.996.996 0 0 0 0-1.41l-1.06-1.06zm1.06-10.96a.996.996 0 0 0 0-1.41.996.996 0 0 0-1.41 0l-1.06 1.06c-.39.39-.39 1.03 0 1.41s1.03.39 1.41 0l1.06-1.06zM7.05 18.36a.996.996 0 0 0 0-1.41.996.996 0 0 0-1.41 0l-1.06 1.06c-.39.39-.39 1.03 0 1.41s1.03.39 1.41 0l1.06-1.06z';
    const iconPath = theme === 'dark' ? sunPath : moonPath;
    document.querySelectorAll('#themeIconLogin path, #themeIconTopbar path').forEach(p => {
      p.setAttribute('d', iconPath);
    });
  }

  function toggleTheme() {
    const current = document.documentElement.getAttribute('data-theme') || 'light';
    applyTheme(current === 'dark' ? 'light' : 'dark');
  }

  applyTheme(getPreferredTheme());

  // ── Configuration ──────────────────────────────────────────────────────────
  // Values are read from window.LAPS_CONFIG (set by config.js) with fallbacks.
  // Replace the fallback strings before deploying, or create config.js.
  const C = Object.assign({
    msalClientId:          '',
    tenantId:              '',
    apiBaseUrl:            '',
    apiScope:              '',
    passwordTimeout:       60,
    justificationMinLength: 10,
    // Session limits (server enforces the absolute cap via auth_time as well)
    sessionMaxAgeMinutes:  480,   // 0 = unlimited absolute session length
    sessionIdleMinutes:    30,    // 0 = no idle logout
  }, window.LAPS_CONFIG ?? {});

  const MSAL_CONFIG = {
    auth: {
      clientId:    C.msalClientId,
      authority:   `https://login.microsoftonline.com/${C.tenantId}`,
      redirectUri: window.location.origin,
    },
    cache: {
      cacheLocation:          'sessionStorage',
      storeAuthStateInCookie: false,
    },
  };

  // ── MSAL setup ─────────────────────────────────────────────────────────────
  const msal = new window.msal.PublicClientApplication(MSAL_CONFIG);

  // ── App state ──────────────────────────────────────────────────────────────
  let activeDevice   = null;   // { id, name } currently selected
  let countdownTimer = null;
  let countdownTotal = C.passwordTimeout;
  let passwordCopied = false;  // did we put the password on the system clipboard?

  // ── View management ────────────────────────────────────────────────────────
  const views = {
    login:   document.getElementById('view-login'),
    loading: document.getElementById('view-loading'),
    devices: document.getElementById('view-devices'),
    error:   document.getElementById('view-error'),
  };
  const appTopbar  = document.getElementById('app-topbar');
  const appFooter   = document.getElementById('app-footer');

  function showView(name) {
    Object.values(views).forEach(v => v.classList.remove('active'));
    views[name].classList.add('active');
    // Show topbar + footer on all views except the login screen
    // (login has its own inline "powered by" credit)
    const authenticated = name !== 'login';
    appTopbar.classList.toggle('hidden', !authenticated);
    appFooter.classList.toggle('hidden', !authenticated);
  }

  function setLoading(msg = 'Loading…') {
    document.getElementById('loading-msg').textContent = msg;
    showView('loading');
  }

  function showGlobalError(title, msg, retryFn) {
    document.getElementById('error-title').textContent = title;
    document.getElementById('error-msg').textContent   = msg;
    document.getElementById('btn-error-retry').onclick = retryFn ?? (() => boot());
    showView('error');
  }

  // ── Modal management ───────────────────────────────────────────────────────
  const overlay = document.getElementById('modal-overlay');
  const phases  = {
    justification: document.getElementById('phase-justification'),
    loading:       document.getElementById('phase-loading'),
    password:      document.getElementById('phase-password'),
    error:         document.getElementById('phase-error'),
  };

  function openModal(phase) {
    overlay.classList.remove('hidden');
    Object.values(phases).forEach(p => p.classList.remove('active'));
    phases[phase].classList.add('active');
    document.body.style.overflow = 'hidden';
  }

  function closeModal() {
    overlay.classList.add('hidden');
    Object.values(phases).forEach(p => p.classList.remove('active'));
    document.body.style.overflow = '';
    stopCountdown();
    clearPasswordField();
    resetJustificationForm();
    activeDevice = null;
  }

  function switchModalPhase(phase) {
    Object.values(phases).forEach(p => p.classList.remove('active'));
    phases[phase].classList.add('active');
  }

  // ── Authentication ─────────────────────────────────────────────────────────
  async function login() {
    await msal.loginRedirect({
      scopes: ['openid', 'profile', 'User.Read'],
      prompt: 'select_account',
    });
  }

  async function getToken() {
    const account = msal.getActiveAccount();
    if (!account) throw new Error('No authenticated account.');
    try {
      const res = await msal.acquireTokenSilent({ account, scopes: [C.apiScope] });
      return res.accessToken;
    } catch (err) {
      if (err instanceof window.msal.InteractionRequiredAuthError) {
        // Redirect back to /api/laps – page will reload and pick up result
        await msal.acquireTokenRedirect({ account, scopes: [C.apiScope] });
        return ''; // unreachable – redirect happens
      }
      throw err;
    }
  }

  // ── API helpers ────────────────────────────────────────────────────────────
  async function apiFetch(path, opts = {}) {
    const token = await getToken();
    const res   = await fetch(C.apiBaseUrl + path, {
      ...opts,
      headers: {
        'Content-Type':  'application/json',
        'Authorization': `Bearer ${token}`,
        ...(opts.headers ?? {}),
      },
    });
    const body = await res.json().catch(() => ({}));
    if (!res.ok) {
      const err = new Error(body.message ?? `HTTP ${res.status}`);
      err.status = res.status;
      err.code   = body.error;
      throw err;
    }
    return body;
  }

  // ── Device list ────────────────────────────────────────────────────────────
  async function loadDevices() {
    setLoading('Loading your devices…');
    let data;
    try {
      data = await apiFetch('/api/my-devices');
    } catch (err) {
      // Session expired server-side (auth_time limit or revoked token): the
      // cached MSAL account is stale, so sign out cleanly instead of retrying.
      if (err.status === 401) {
        doLogout();
        return;
      }
      showGlobalError(
        'Failed to load devices',
        err.message ?? 'Could not connect to the backend. Please try again.',
        loadDevices
      );
      return;
    }
    renderDevices(data.devices ?? []);
  }

  function renderDevices(devices) {
    const account = msal.getActiveAccount();
    document.getElementById('topbar-upn').textContent =
      account?.username ?? account?.name ?? '';

    const grid = document.getElementById('device-grid');
    grid.innerHTML = '';

    if (devices.length === 0) {
      grid.innerHTML = `
        <div class="empty" role="status">
          <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M20 4H4c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm0 14H4V6h16v12z"/></svg>
          <p>No managed devices found for your account.</p>
        </div>`;
      showView('devices');
      return;
    }

    devices.forEach(device => {
      const isMgd    = device.isManaged;
      const lastSeen = device.lastSignIn ? fmtDate(device.lastSignIn) : 'Unknown';
      const badge    = isMgd
        ? '<span class="badge badge-success">Managed</span>'
        : '<span class="badge badge-neutral">Unmanaged</span>';

      const card = document.createElement('div');
      card.className = 'device-card';
      card.setAttribute('role', 'listitem');
      card.innerHTML = `
        <div class="device-icon" aria-hidden="true">
          <svg viewBox="0 0 24 24"><path d="M20 4H4c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm0 14H4V6h16v12zm-8-2l-4-4 1.41-1.41L12 13.17l6.59-6.58L20 8l-8 8z"/></svg>
        </div>
        <div class="device-info">
          <span class="device-name">${esc(device.name)}</span>
          <span class="device-meta">${esc(device.operatingSystem || 'Windows')} &nbsp;·&nbsp; Last seen: ${esc(lastSeen)}</span>
        </div>
        <div style="display:flex;align-items:center;gap:10px;flex-shrink:0">
          ${badge}
          <button class="btn btn-primary btn-sm" type="button"
                  data-id="${esc(device.id)}" data-name="${esc(device.name)}">
            Get password
          </button>
        </div>`;
      grid.appendChild(card);
    });

    // Bind "Get password" buttons
    grid.querySelectorAll('[data-id]').forEach(btn => {
      btn.addEventListener('click', () => {
        activeDevice = { id: btn.dataset.id, name: btn.dataset.name };
        openJustificationModal(activeDevice.name);
      });
    });

    showView('devices');
  }

  // ── Justification modal ────────────────────────────────────────────────────
  function openJustificationModal(deviceName) {
    document.getElementById('modal-device-name').textContent = `Device: ${deviceName}`;
    document.getElementById('min-len').textContent = C.justificationMinLength;
    resetJustificationForm();
    openModal('justification');
    document.getElementById('justification-input').focus();
  }

  function resetJustificationForm() {
    const ta = document.getElementById('justification-input');
    if (ta) {
      ta.value = '';
      document.getElementById('char-count').textContent = '0';
      document.getElementById('btn-modal-submit').disabled = true;
    }
  }

  // ── LAPS password retrieval ────────────────────────────────────────────────
  async function requestPassword() {
    const justification = document.getElementById('justification-input').value.trim();
    if (!justification || !activeDevice) return;

    switchModalPhase('loading');

    let data;
    try {
      data = await apiFetch('/api/laps-password', {
        method: 'POST',
        body:   JSON.stringify({ deviceId: activeDevice.id, justification }),
      });
    } catch (err) {
      // Server-side session expiry mid-flow: sign out cleanly.
      if (err.status === 401) {
        doLogout();
        return;
      }
      const msg = err.status === 403 ? 'This device is not registered to your account.'
                : err.status === 404 ? 'No LAPS password is stored for this device in Intune.'
                : err.message ?? 'An unexpected error occurred.';
      document.getElementById('modal-error-box').textContent = msg;
      document.getElementById('btn-modal-retry').onclick = () => {
        openJustificationModal(activeDevice.name);
      };
      switchModalPhase('error');
      return;
    }

    displayPassword(data);
  }

  // ── Password display ───────────────────────────────────────────────────────
  function displayPassword({ deviceName, accountName, password, passwordCreated, nextRotation }) {
    document.getElementById('pw-device-name').textContent  = `Device: ${deviceName}`;
    document.getElementById('pw-account-name').textContent = accountName ? `Account: ${accountName}` : '';
    document.getElementById('pw-input').value             = password;
    document.getElementById('pw-input').type              = 'password';
    document.getElementById('pw-copy-confirm').classList.add('hidden');
    passwordCopied = false;
    document.getElementById('pw-created').textContent =
      `Password created: ${passwordCreated ? fmtDate(passwordCreated) : 'Unknown'}`;
    document.getElementById('pw-next-rotation').textContent =
      `Next rotation: ${nextRotation ? fmtDate(nextRotation) : 'Unknown'}`;

    switchModalPhase('password');
    startCountdown(C.passwordTimeout);
  }

  function clearPasswordField() {
    clearClipboardIfCopied();
    const inp = document.getElementById('pw-input');
    if (inp) { inp.value = ''; inp.type = 'password'; }
    // Restore pw field in case it was hidden due to expiry
    document.getElementById('pw-field-wrap')?.classList.remove('hidden');
    document.getElementById('pw-expired-msg')?.classList.add('hidden');
    const fill = document.getElementById('countdown-fill');
    if (fill) { fill.classList.remove('expired'); }
    stopCountdown();
  }

  // ── Countdown ──────────────────────────────────────────────────────────────
  function startCountdown(total) {
    stopCountdown();
    countdownTotal = total;
    let remaining  = total;
    updateCountdown(remaining, total);

    countdownTimer = setInterval(() => {
      remaining -= 1;
      updateCountdown(remaining, total);
      if (remaining <= 0) {
        stopCountdown();
        clearPasswordField();
        // Replace password field with expiry notice
        document.getElementById('pw-field-wrap').classList.add('hidden');
        document.getElementById('pw-copy-confirm').classList.add('hidden');
        document.getElementById('pw-expired-msg').classList.remove('hidden');
        // Update countdown bar to show fully expired
        const fill = document.getElementById('countdown-fill');
        fill.style.width = '0%';
        fill.classList.remove('urgent');
        fill.classList.add('expired');
        document.getElementById('countdown-text').textContent = 'Expired';
      }
    }, 1000);
  }

  function updateCountdown(remaining, total) {
    const pct    = Math.max(0, remaining / total * 100);
    const fill   = document.getElementById('countdown-fill');
    const label  = document.getElementById('countdown-text');
    fill.style.width = `${pct}%`;
    fill.classList.toggle('urgent', remaining <= 10);
    label.textContent = `${remaining}s`;
  }

  function stopCountdown() {
    if (countdownTimer) { clearInterval(countdownTimer); countdownTimer = null; }
  }

  // ── Utilities ──────────────────────────────────────────────────────────────
  function esc(str) {
    return String(str ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function fmtDate(iso) {
    try {
      return new Intl.DateTimeFormat(undefined, { dateStyle: 'medium', timeStyle: 'short' })
        .format(new Date(iso));
    } catch { return iso; }
  }

  async function copyToClipboard(text) {
    try {
      await navigator.clipboard.writeText(text);
      passwordCopied = true;
      const confirm = document.getElementById('pw-copy-confirm');
      confirm.classList.remove('hidden');
      setTimeout(() => confirm.classList.add('hidden'), 2500);
    } catch { /* clipboard unavailable */ }
  }

  /**
   * Best-effort clipboard wipe once the password expires or the modal closes.
   * Only runs if we were the ones who wrote it, so unrelated clipboard content
   * is left alone. The write can fail when the document is not focused - that
   * is tolerable, but it does mean the clipboard is not a guaranteed-clean
   * channel for a credential.
   */
  async function clearClipboardIfCopied() {
    if (!passwordCopied) return;
    passwordCopied = false;
    try {
      await navigator.clipboard.writeText("");
    } catch { /* clipboard unavailable or document not focused */ }
  }

  // ── Event listeners ────────────────────────────────────────────────────────

  // Login
  document.getElementById('btn-login').addEventListener('click', login);

  // Logout
  document.getElementById('btn-logout').addEventListener('click', doLogout);

  // Theme toggles
  document.querySelectorAll('.theme-toggle').forEach(btn => {
    btn.addEventListener('click', toggleTheme);
  });

  // Error retry
  // (onclick set dynamically in showGlobalError)

  // Justification textarea
  document.getElementById('justification-input').addEventListener('input', () => {
    const len = document.getElementById('justification-input').value.trim().length;
    document.getElementById('char-count').textContent = len;
    document.getElementById('btn-modal-submit').disabled = (len < C.justificationMinLength);
  });

  // Submit justification
  document.getElementById('btn-modal-submit').addEventListener('click', requestPassword);

  // Enter key in textarea
  document.getElementById('justification-input').addEventListener('keydown', e => {
    if (e.key === 'Enter' && e.ctrlKey) requestPassword();
  });

  // Cancel / close buttons
  document.getElementById('btn-modal-cancel').addEventListener('click',    closeModal);
  document.getElementById('btn-modal-close-j').addEventListener('click',   closeModal);
  document.getElementById('btn-modal-close-pw').addEventListener('click',  closeModal);
  document.getElementById('btn-modal-close-e').addEventListener('click',   closeModal);
  document.getElementById('btn-modal-close-e2').addEventListener('click',  closeModal);

  // Close on overlay click
  overlay.addEventListener('click', e => { if (e.target === overlay) closeModal(); });

  // ESC closes modal
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape' && !overlay.classList.contains('hidden')) closeModal();
  });

  // Toggle password visibility
  document.getElementById('btn-toggle-pw').addEventListener('click', () => {
    const inp = document.getElementById('pw-input');
    const eye = document.getElementById('icon-eye');
    if (inp.type === 'password') {
      inp.type = 'text';
      eye.innerHTML = '<path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24M1 1l22 22"/>';
    } else {
      inp.type = 'password';
      eye.innerHTML = '<path d="M12 4.5C7 4.5 2.73 7.61 1 12c1.73 4.39 6 7.5 11 7.5s9.27-3.11 11-7.5c-1.73-4.39-6-7.5-11-7.5zm0 12.5a5 5 0 1 1 0-10 5 5 0 0 1 0 10zm0-8a3 3 0 1 0 0 6 3 3 0 0 0 0-6z"/>';
    }
  });

  // Copy password
  document.getElementById('btn-copy-pw').addEventListener('click', () => {
    copyToClipboard(document.getElementById('pw-input').value);
  });

  // ── Boot ───────────────────────────────────────────────────────────────────
  // Session timer state (set after a successful sign-in)
  let sessionLogoutTimer  = null;   // absolute session cap (from sign-in time)
  let idleLogoutTimer     = null;   // idle timeout (reset on user activity)
  let sessionStartTime    = null;   // epoch ms of first successful boot

  /**
   * Sign the user out via MSAL redirect. Clears the cached MSAL session, so a
   * reload lands on the login view instead of silently reacquiring tokens.
   */
  function doLogout() {
    stopSessionTimers();
    const account = msal.getActiveAccount();
    msal.logoutRedirect({ account, postLogoutRedirectUri: window.location.origin })
      .catch(() => { window.location.reload(); });
  }

  function stopSessionTimers() {
    if (sessionLogoutTimer) { clearTimeout(sessionLogoutTimer); sessionLogoutTimer = null; }
    if (idleLogoutTimer)    { clearTimeout(idleLogoutTimer);    idleLogoutTimer = null; }
  }

  function scheduleIdleLogout() {
    if (idleLogoutTimer) clearTimeout(idleLogoutTimer);
    if (!C.sessionIdleMinutes || C.sessionIdleMinutes <= 0) return;
    idleLogoutTimer = setTimeout(() => {
      stopSessionTimers();
      showGlobalError('Session expired', 'You were signed out after a period of inactivity.');
    }, C.sessionIdleMinutes * 60 * 1000);
  }

  /** Arm the absolute session cap once per sign-in (not reset by activity). */
  function scheduleAbsoluteLogout() {
    if (sessionLogoutTimer) return;
    if (!C.sessionMaxAgeMinutes || C.sessionMaxAgeMinutes <= 0) return;
    const elapsed = sessionStartTime ? (Date.now() - sessionStartTime) : 0;
    const remaining = Math.max(0, C.sessionMaxAgeMinutes * 60 * 1000 - elapsed);
    sessionLogoutTimer = setTimeout(() => {
      stopSessionTimers();
      showGlobalError('Session expired', 'Your session has reached its maximum duration. Please sign in again.');
    }, remaining);
  }

  // User activity resets the idle timer. Passive listeners – no scrolling cost.
  const ACTIVITY_EVENTS = ['click', 'keydown', 'pointerdown', 'scroll', 'touchstart'];
  ACTIVITY_EVENTS.forEach(evt =>
    document.addEventListener(evt, () => {
      if (sessionStartTime === null) return;   // not signed in yet
      scheduleIdleLogout();
    }, { passive: true }));

  // ── Boot ───────────────────────────────────────────────────────────────────
  function boot() {
    setLoading('Signing in…');

    msal.handleRedirectPromise()
      .then(response => {
        if (response?.account) msal.setActiveAccount(response.account);

        const accounts = msal.getAllAccounts();
        if (accounts.length === 0) {
          showView('login');
          return;
        }

        if (!msal.getActiveAccount()) msal.setActiveAccount(accounts[0]);

        // ── Session limits: arm on first authenticated boot ────────────────
        if (sessionStartTime === null) sessionStartTime = Date.now();
        scheduleAbsoluteLogout();
        scheduleIdleLogout();

        loadDevices();
      })
      .catch(err => {
        console.error('[MSAL]', err);
        showGlobalError(
          'Authentication error',
          err.message ?? 'Sign-in failed. Please reload the page and try again.'
        );
      });
  }

  boot();
