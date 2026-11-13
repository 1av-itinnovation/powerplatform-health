/* 1AV Project Health Monitor
 * Live mode: signs in with Microsoft (MSAL), calls the Power Automate "API" flow
 * with the user's token, and renders the JSON it returns.
 * Demo mode: no apiUrl in config.js, or ?demo=1 in the address -> data/sample.json.
 * The page renders whatever projects the flow returns, so a new project needs no page change.
 */
(function () {
  const C = window.MONITOR_CONFIG || {};
  const demo = !C.apiUrl || new URLSearchParams(location.search).has("demo");
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const TZ = "Asia/Manila";
  const STATUS_LABEL = { healthy: "Healthy", warning: "Warning", failing: "Failing", stale: "No data" };
  const SVC_LABEL = { ok: "Working", advisory: "Advisory", degraded: "Degraded", down: "Outage" };

  let msalApp = null, account = null, data = null, selected = null, nextAt = 0, timer = null;

  /* ---------- time helpers ---------- */
  const fmtTime = (iso) => iso ? new Date(iso).toLocaleString("en-PH", { timeZone: TZ, month: "short", day: "numeric", hour: "numeric", minute: "2-digit" }) : "—";
  function ago(iso) {
    if (!iso) return "—";
    const m = Math.round((Date.now() - new Date(iso)) / 60000);
    if (m < 1) return "just now";
    if (m < 60) return m + " min ago";
    const h = Math.floor(m / 60);
    return h < 48 ? h + " h ago" : Math.floor(h / 24) + " d ago";
  }

  /* ---------- auth ---------- */
  async function initAuth() {
    msalApp = new msal.PublicClientApplication({
      auth: {
        clientId: C.clientId,
        authority: "https://login.microsoftonline.com/" + C.tenantId,
        redirectUri: location.origin + location.pathname
      },
      cache: { cacheLocation: "sessionStorage" }
    });
    await msalApp.initialize();
    const res = await msalApp.handleRedirectPromise();
    account = res?.account || msalApp.getAllAccounts()[0] || null;
    if (account) msalApp.setActiveAccount(account);
  }

  async function getToken() {
    const req = { scopes: C.scopes, account };
    try {
      return (await msalApp.acquireTokenSilent(req)).accessToken;
    } catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError) { await msalApp.acquireTokenRedirect(req); return null; }
      throw e;
    }
  }

  function signIn() { msalApp.loginRedirect({ scopes: C.scopes }); }
  function signOut() { msalApp.logoutRedirect({ account }); }

  /* ---------- data ---------- */
  async function load() {
    $("refreshBtn").disabled = true;
    try {
      if (demo) {
        data = await (await fetch("data/sample.json", { cache: "no-store" })).json();
        // keep sample timestamps relative to now so demo mode never looks stale
        const shift = Date.now() - 6 * 60000 - new Date(data.generatedAt);
        data = JSON.parse(JSON.stringify(data), (k, v) => (typeof v === "string" && /^\d{4}-\d\d-\d\dT/.test(v)) ? new Date(new Date(v).getTime() + shift).toISOString() : v);
      } else {
        const token = await getToken();
        if (!token) return;
        const r = await fetch(C.apiUrl, {
          method: "POST",
          headers: { "Authorization": "Bearer " + token, "Content-Type": "application/json" },
          body: JSON.stringify({ action: "status" })
        });
        if (!r.ok) throw new Error("The monitor flow replied " + r.status + " " + r.statusText);
        data = await r.json();
      }
      render();
      banner();
    } catch (e) {
      console.error(e);
      showBanner("fail", "Could not load data: " + esc(e.message) + ". Showing the last data received.");
    } finally {
      $("refreshBtn").disabled = false;
      schedule();
    }
  }

  function schedule() {
    clearInterval(timer);
    nextAt = Date.now() + (C.refreshMinutes || 15) * 60000;
    timer = setInterval(() => {
      const s = Math.max(0, Math.round((nextAt - Date.now()) / 1000));
      $("countdown").textContent = "· next refresh " + Math.floor(s / 60) + ":" + String(s % 60).padStart(2, "0");
      if (data) $("updated").textContent = "Updated " + ago(data.generatedAt);
      if (s === 0) { clearInterval(timer); load(); }
    }, 1000);
  }

  /* ---------- banners ---------- */
  function showBanner(kind, html) { const b = $("banner"); b.className = "banner " + kind; b.innerHTML = html; }
  function banner() {
    const age = (Date.now() - new Date(data.generatedAt)) / 60000;
    if (age > (C.staleMinutes || 35)) showBanner("warn", "The collector has not updated for " + ago(data.generatedAt).replace(" ago", "") + ". Check the collector flow's run history.");
    else if (demo || data.demo) showBanner("info", "Demo mode: showing sample data. Add your client ID and flow URL in <b>config.js</b> to go live.");
    else $("banner").className = "banner hidden";
  }

  /* ---------- render ---------- */
  function render() {
    $("signin").classList.add("hidden");
    $("content").classList.remove("hidden");
    $("updated").textContent = "Updated " + ago(data.generatedAt);

    $("services").innerHTML = (data.services || []).map((s) =>
      `<div class="svc ${esc(s.status)}"><b>${esc(s.name)}</b><span>${esc(SVC_LABEL[s.status] || s.status)}</span>${s.note ? `<small>${esc(s.note)}</small>` : ""}</div>`).join("");

    const P = data.projects || [];
    const runs = P.reduce((a, p) => a + (p.runs || 0), 0);
    const failed = P.reduce((a, p) => a + (p.failed || 0), 0);
    const flowCount = P.reduce((a, p) => a + (p.flows || []).length, 0);
    const healthy = P.filter((p) => p.status === "healthy").length;
    const uptime = runs ? (100 * (runs - failed) / runs) : 100;
    $("kpis").innerHTML = [
      [uptime.toFixed(1) + "%", "Overall automation success (24 h)"],
      [runs.toLocaleString(), "Flow runs across " + flowCount + " flows"],
      [failed.toLocaleString(), "Failed runs (24 h)"],
      [healthy + " / " + P.length, "Projects healthy"]
    ].map(([v, l]) => `<div class="card kpi"><div class="v">${v}</div><div class="l">${l}</div></div>`).join("");

    $("projects").innerHTML = P.map(card).join("");
    document.querySelectorAll(".proj").forEach((el) => el.addEventListener("click", () => openDetail(el.dataset.key)));

    const I = (data.incidents || []).slice().sort((a, b) => (a.severity === "failing" ? 0 : 1) - (b.severity === "failing" ? 0 : 1) || new Date(b.time) - new Date(a.time));
    $("incidents").innerHTML = I.length ? I.map((i) =>
      `<div class="inc"><div class="t"><span class="dot ${esc(i.severity)}"></span>${esc(ago(i.time))}</div><div class="p">${esc(i.project)}</div><div><b>${esc(i.source)}</b> — ${esc(i.message)}</div></div>`).join("")
      : `<div class="empty">No open issues in the last 24 hours.</div>`;

    if (selected) openDetail(selected, true);
  }

  function bars(hourly) {
    const h = hourly || [];
    const max = Math.max(1, ...h.map(([o, f]) => o + f));
    return `<div class="bars">${h.map(([o, f]) =>
      `<div class="bar" title="${o} ok · ${f} failed"><i style="height:${(o / max) * 38}px"></i><i class="f" style="height:${f ? Math.max(3, (f / max) * 38) : 0}px"></i></div>`).join("")}</div>
      <div class="axis"><span>24 h ago</span><span>now</span></div>`;
  }

  function card(p) {
    const st = p.status || "stale";
    const lastUse = (p.apps || []).map((a) => a.lastUsed).filter(Boolean).sort().pop();
    return `<div class="card proj ${st} ${selected === p.key ? "sel" : ""}" data-key="${esc(p.key)}">
      <div class="ph"><div><h3>${esc(p.name)}</h3><div class="pre">${esc(p.prefix)} · ${(p.flows || []).length} flow${(p.flows || []).length === 1 ? "" : "s"}</div></div>
        <span class="pill ${st}">${STATUS_LABEL[st] || st}</span></div>
      <div class="nums">
        <div><b>${p.uptimePct != null ? Number(p.uptimePct).toFixed(1) + "%" : "—"}</b><span>uptime</span></div>
        <div><b>${(p.runs || 0).toLocaleString()}</b><span>runs</span></div>
        <div><b style="color:${p.failed ? "var(--fail)" : "inherit"}">${p.failed || 0}</b><span>failed</span></div>
      </div>
      ${p.hourly ? bars(p.hourly) : ""}
      ${p.summary ? `<p class="summary">${esc(p.summary)}</p>` : ""}
      <div class="apps">${(p.apps || []).length ? "App last used " + ago(lastUse) : "No Power App · device checks"}</div>
    </div>`;
  }

  function openDetail(key, keepScroll) {
    const p = (data.projects || []).find((x) => x.key === key);
    if (!p) return;
    selected = key;
    document.querySelectorAll(".proj").forEach((el) => el.classList.toggle("sel", el.dataset.key === key));
    const flows = (p.flows || []).slice().sort((a, b) => (b.failed || 0) - (a.failed || 0) || a.name.localeCompare(b.name));
    const d = $("detail");
    d.className = "card detail";
    d.innerHTML = `<button class="x" title="Close" onclick="Monitor.closeDetail()">×</button>
      <h3>${esc(p.name)}</h3><div class="muted">${esc(p.summary || "")}</div>
      <div class="dgrid">
        <div><table><thead><tr><th>Flow</th><th>Runs</th><th>Failed</th><th>Last run</th><th>Last error</th></tr></thead><tbody>
          ${flows.map((f) => `<tr><td>${esc(f.name)}</td><td>${f.runs}</td><td style="color:${f.failed ? "var(--fail)" : "inherit"};font-weight:${f.failed ? 600 : 400}">${f.failed}</td>
            <td>${esc(fmtTime(f.lastRun))}</td><td class="err">${esc(f.lastError || "")}</td></tr>`).join("") || `<tr><td colspan="5" class="muted">No flows found with prefix ${esc(p.prefix)}</td></tr>`}
        </tbody></table></div>
        <div>
          ${(p.apps || []).length ? `<h4 class="sec" style="margin-top:0">Apps</h4>` + p.apps.map((a) =>
            `<div class="check"><span>${esc(a.name)}</span><span class="muted">${esc(ago(a.lastUsed))}${a.activeHours != null ? " · active " + a.activeHours + "/24 h" : ""}</span></div>`).join("") : ""}
          ${(p.checks || []).length ? `<h4 class="sec">Checks</h4>` + p.checks.map((c) =>
            `<div class="check"><span><span class="dot ${esc(c.status)}"></span>${esc(c.label)}</span><span class="muted">${esc(c.value)}</span></div>`).join("") : ""}
        </div>
      </div>`;
    if (!keepScroll) d.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function closeDetail() { selected = null; $("detail").className = "detail hidden"; document.querySelectorAll(".proj").forEach((el) => el.classList.remove("sel")); }

  /* ---------- start ---------- */
  async function start() {
    $("refreshBtn").addEventListener("click", load);
    if (demo) {
      $("authBtn").classList.add("hidden");
      return load();
    }
    try { await initAuth(); }
    catch (e) { showBanner("fail", "Sign-in setup failed: " + esc(e.message)); return; }
    if (!account) {
      $("signin").classList.remove("hidden");
      $("authBtn").textContent = "Sign in";
      $("authBtn").onclick = signIn;
      return;
    }
    $("authBtn").textContent = "Sign out";
    $("authBtn").title = account.username;
    $("authBtn").onclick = signOut;
    load();
  }

  window.Monitor = { signIn, closeDetail };
  start();
})();
