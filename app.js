/* Vehicle Check – LBH Vehicles
   Offline-first phone app. Checks are saved on the phone (IndexedDB) and sent
   to the Google Apps Script endpoint in config.js whenever there is signal. */
(function () {
  "use strict";
  const CFG = window.APP_CONFIG || {};
  const ITEMS = window.CHECK_ITEMS, PARTS = window.SERVICE_PARTS;
  const $ = (s, el = document) => el.querySelector(s);
  const $$ = (s, el = document) => Array.from(el.querySelectorAll(s));
  const view = $("#view");

  // ---------- small helpers ----------
  const ls = {
    get(k, d) { try { const v = localStorage.getItem("vc." + k); return v == null ? d : JSON.parse(v); } catch (e) { return d; } },
    set(k, v) { try { localStorage.setItem("vc." + k, JSON.stringify(v)); } catch (e) {} }
  };
  let lang = ls.get("lang", (navigator.language || "en").toLowerCase().startsWith("pt") ? "pt" : "en");
  const t = (k, vars) => {
    let s = (T[lang] && T[lang][k]) || T.en[k] || k;
    if (vars) for (const v in vars) s = s.replace("{" + v + "}", vars[v]);
    return s;
  };
  const esc = s => String(s == null ? "" : s).replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const today = () => { const d = new Date(); d.setMinutes(d.getMinutes() - d.getTimezoneOffset()); return d.toISOString().slice(0, 10); };
  const uid = () => Date.now().toString(36) + "-" + Math.random().toString(36).slice(2, 8);
  const configured = () => CFG.ENDPOINT && /^https:\/\//.test(CFG.ENDPOINT);
  const itemName = (it, l = lang) => l === "pt" ? it.pt : it.en;
  function toast(msg, ms = 2600) {
    const el = $("#toast"); el.textContent = msg; el.classList.remove("hidden");
    clearTimeout(toast._t); toast._t = setTimeout(() => el.classList.add("hidden"), ms);
  }

  // ---------- IndexedDB (queue + drafts; holds photos) ----------
  const DB = (() => {
    let p;
    const open = () => p || (p = new Promise((res, rej) => {
      const r = indexedDB.open("vehicle-check", 1);
      r.onupgradeneeded = () => { r.result.createObjectStore("queue", { keyPath: "id" }); r.result.createObjectStore("drafts"); };
      r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error);
    }));
    const tx = (store, mode, fn) => open().then(db => new Promise((res, rej) => {
      const x = db.transaction(store, mode), s = x.objectStore(store); const out = fn(s);
      x.oncomplete = () => res(out && "result" in out ? out.result : undefined); x.onerror = () => rej(x.error);
    }));
    return {
      put: (store, val, key) => tx(store, "readwrite", s => s.put(val, key)),
      get: (store, key) => tx(store, "readonly", s => s.get(key)),
      del: (store, key) => tx(store, "readwrite", s => s.delete(key)),
      all: store => tx(store, "readonly", s => s.getAll())
    };
  })();

  // ---------- history (small list, no photos) ----------
  const history = {
    list: () => ls.get("history", []),
    add(h) { const l = history.list(); l.unshift(h); ls.set("history", l.slice(0, 30)); },
    mark(id, status) { const l = history.list(); const h = l.find(x => x.id === id); if (h) { h.status = status; ls.set("history", l); } }
  };
  // remembered names / vehicles for quick filling
  const prefs = {
    get: () => ls.get("prefs", { names: [], depts: [], vehicles: {} }),
    learn(f) {
      const p = prefs.get();
      const addTo = (arr, v) => { v = (v || "").trim(); if (!v) return arr; return [v, ...arr.filter(x => x.toUpperCase() !== v.toUpperCase())].slice(0, 25); };
      p.names = addTo(addTo(addTo(p.names, f.regularDriver), f.inspector2), f.inspector1);
      p.depts = addTo(p.depts, f.department);
      const m = (f.matricula || "").trim().toUpperCase();
      if (m) {
        const prev = p.vehicles[m] || {};
        p.vehicles[m] = Object.assign(prev, { vehicle: f.vehicle || prev.vehicle, inspExp: f.inspExp || prev.inspExp, licExp: f.licExp || prev.licExp,
          insExp: f.insExp || prev.insExp, nextService: f.nextService || prev.nextService, department: f.department || prev.department });
      }
      ls.set("prefs", p);
    }
  };

  // ---------- network banner ----------
  function updateNet() {
    const b = $("#netBanner");
    if (!configured()) { b.textContent = t("notConfigured"); b.classList.remove("hidden"); return; }
    if (!navigator.onLine) { b.textContent = t("offline"); b.classList.remove("hidden"); } else b.classList.add("hidden");
  }
  window.addEventListener("online", () => { updateNet(); sync(); });
  window.addEventListener("offline", updateNet);

  // ---------- sending ----------
  let syncing = false;
  async function sync(manual) {
    if (syncing || !configured()) return;
    syncing = true; refreshPending();
    let sentAny = false, failed = false;
    try {
      const q = (await DB.all("queue")).sort((a, b) => a.created.localeCompare(b.created));
      for (const rec of q) {
        try {
          const ctrl = new AbortController(); const to = setTimeout(() => ctrl.abort(), 90000);
          const res = await fetch(CFG.ENDPOINT, { method: "POST", body: JSON.stringify(buildPayload(rec)),
            headers: { "Content-Type": "text/plain;charset=utf-8" }, redirect: "follow", signal: ctrl.signal });
          clearTimeout(to);
          const j = await res.json();
          if (!j.ok) throw new Error(j.error || "server");
          await DB.del("queue", rec.id); history.mark(rec.id, "sent"); sentAny = true;
        } catch (e) { failed = true; console.warn("send failed", e); break; }
      }
    } finally { syncing = false; }
    if (manual && failed) toast(t("sendFail"));
    refreshPending(); if (sentAny && (location.hash === "" || location.hash === "#/")) renderHome();
    return !failed;
  }
  setInterval(() => { if (navigator.onLine) sync(); }, 60000);
  document.addEventListener("visibilitychange", () => { if (!document.hidden) sync(); });

  async function refreshPending() {
    const el = $("#pendingBox"); if (!el) return;
    const n = (await DB.all("queue")).length;
    el.classList.toggle("zero", n === 0);
    $(".n", el).textContent = n;
    $(".lbl", el).textContent = n ? t("pending") : t("pendingNone");
    const b = $("button", el); b.classList.toggle("hidden", n === 0); b.disabled = syncing; b.textContent = syncing ? t("sending") : t("sendNow");
  }

  // ---------- turning a saved record into the spreadsheet row ----------
  function expiryState(dateStr, ref) {
    if (!dateStr) return "";
    const d = new Date(dateStr), r = new Date(ref || today());
    if (d < r) return "expired";
    if ((d - r) / 864e5 <= 30) return "soon";
    return "";
  }
  const isBad = v => v === "bad" || v === "yes";
  function defectsOf(f) {
    return ITEMS.filter(it => f.items[it.id] && isBad(f.items[it.id].v)).map(it => ({ it, note: f.items[it.id].note || "" }));
  }
  function warningsOf(f) {
    const w = [];
    [["inspExp", "Inspection"], ["licExp", "Licence"], ["insExp", "Insurance"]].forEach(([k, n]) => {
      const s = expiryState(f[k], f.date); if (s) w.push(n + (s === "expired" ? " EXPIRED " : " expires ") + f[k]);
    });
    const ex = f.items.extinguisher && f.items.extinguisher.expiry, s = expiryState(ex, f.date);
    if (s) w.push("Fire extinguisher" + (s === "expired" ? " EXPIRED " : " expires ") + ex);
    return w;
  }
  const b64 = dataUrl => dataUrl.split(",")[1];
  function buildPayload(rec) {
    const f = rec.form, photos = [];
    const base = { key: CFG.APP_KEY, id: rec.id, kind: rec.type, version: CFG.VERSION };
    if (rec.type === "check") {
      const defects = defectsOf(f), warnings = warningsOf(f);
      const row = {
        "Submission ID": rec.id, "Check type": f.checkType === "preservice" ? "Before service" : "Weekly",
        "Date": f.date, "Department": f.department, "Vehicle": f.vehicle, "Matrícula": (f.matricula || "").toUpperCase(),
        "Km": f.km ? Number(f.km) : "", "Next service due": f.nextService,
        "Inspection expires": f.inspExp, "Licence expires": f.licExp, "Insurance expires": f.insExp,
        "Inspector": f.inspector1, "Second person (lights)": f.inspector2, "Regular driver consulted": f.regularDriver,
        "Result": defects.length ? "PROBLEMS" : "ALL OK", "Problems": defects.length,
        "Warnings": warnings.join("; ")
      };
      ITEMS.forEach(it => {
        const a = f.items[it.id] || {};
        row[it.en] = a.v === "ok" ? "✓" : a.v === "bad" ? "X" : a.v === "yes" ? "Yes" : a.v === "no" ? "No" : "";
      });
      row["Fire extinguisher expiry"] = (f.items.extinguisher || {}).expiry || "";
      row["Oil level (H/L)"] = (f.items.oil || {}).hl || "";
      row["Coolant level (H/L)"] = (f.items.coolant || {}).hl || "";
      row["Problem details"] = defects.map(d => d.it.en + ": " + d.note).join("\n");
      row["Remarks"] = f.remarks;
      row["Submitted on phone"] = rec.created; row["App language"] = rec.lang.toUpperCase();
      ITEMS.forEach(it => ((f.items[it.id] || {}).photos || []).forEach((p, i) => photos.push({ ref: it.id, name: it.en + " " + (i + 1), mime: "image/jpeg", data: b64(p) })));
      (f.remarkPhotos || []).forEach((p, i) => photos.push({ ref: "remarks", name: "Remarks " + (i + 1), mime: "image/jpeg", data: b64(p) }));
      return Object.assign(base, {
        sheet: "Checklists", row, photos,
        defects: defects.map(d => ({ ref: d.it.id, row: {
          "Submission ID": rec.id, "Date": f.date, "Vehicle": f.vehicle, "Matrícula": (f.matricula || "").toUpperCase(),
          "Km": f.km ? Number(f.km) : "", "Item": d.it.en, "Problem": d.note,
          "Reported by": [f.inspector1, f.inspector2].filter(Boolean).join(" / "), "Status": "Open", "Action taken": "", "Closed on": ""
        } }))
      });
    }
    // service record
    const row = {
      "Submission ID": rec.id, "Date": f.date, "Vehicle": f.vehicle, "Matrícula": (f.matricula || "").toUpperCase(),
      "Km": f.km ? Number(f.km) : "", "Workshop / mechanic": f.workshop, "Entered by": f.enteredBy
    };
    PARTS.forEach(p => { const a = f.parts[p.id] || {}; row[p.en] = a.v === "ok" ? "✓" : a.v === "bad" ? "X" : ""; row[p.en + " – qty & part no."] = a.pn || ""; });
    row["Remarks"] = f.remarks; row["Submitted on phone"] = rec.created; row["App language"] = rec.lang.toUpperCase();
    (f.remarkPhotos || []).forEach((p, i) => photos.push({ ref: "remarks", name: "Service " + (i + 1), mime: "image/jpeg", data: b64(p) }));
    return Object.assign(base, { sheet: "Service Records", row, photos, defects: [] });
  }

  // ---------- photos ----------
  function compress(file) {
    return new Promise((res, rej) => {
      const url = URL.createObjectURL(file), img = new Image();
      img.onload = () => {
        const max = 1280, sc = Math.min(1, max / Math.max(img.width, img.height));
        const c = document.createElement("canvas"); c.width = Math.round(img.width * sc); c.height = Math.round(img.height * sc);
        c.getContext("2d").drawImage(img, 0, 0, c.width, c.height); URL.revokeObjectURL(url);
        res(c.toDataURL("image/jpeg", 0.7));
      };
      img.onerror = () => { URL.revokeObjectURL(url); rej(new Error("img")); };
      img.src = url;
    });
  }
  const photoStrip = (arr, key) =>
    `<div class="photos">${(arr || []).map((p, i) => `<div class="ph"><img src="${p}" alt=""><button type="button" data-delphoto="${key}" data-i="${i}">✕</button></div>`).join("")}
     <label class="addph">📷 ${t("addPhoto")}<input type="file" accept="image/*" capture="environment" data-photo="${key}"></label></div>`;

  // ---------- routing ----------
  let form = null, lastDone = null;
  function route() {
    const h = location.hash.replace(/^#\/?/, "");
    $("#backBtn").classList.toggle("hidden", !h);
    $("#langBtn").textContent = lang === "pt" ? "EN" : "PT";
    $("#hdrTitle").textContent = t("appName");
    document.documentElement.lang = lang;
    updateNet();
    if (h === "check") return openForm("check");
    if (h === "service") return openForm("service");
    if (h === "done" && lastDone) return renderDone();
    form = null; renderHome();
  }
  window.addEventListener("hashchange", route);
  $("#backBtn").onclick = () => { location.hash = "#/"; };
  $("#langBtn").onclick = () => { lang = lang === "pt" ? "en" : "pt"; ls.set("lang", lang); const y = scrollY; route(); scrollTo(0, y); };

  // ---------- home ----------
  function renderHome() {
    view.classList.add("is-home");
    const hist = history.list().slice(0, 8);
    view.innerHTML = `<div class="home">
      <div class="logo-top"><img src="icons/lbh-logo.png" alt="LBH Group"></div>
      <button class="bigbtn" data-go="check"><span class="ico"><img src="icons/vehicle.png" alt=""></span><span><b>${t("newCheck")}</b><small>${t("newCheckSub")}</small></span></button>
      <button class="bigbtn" data-go="service"><span class="ico">🔧</span><span><b>${t("newService")}</b><small>${t("newServiceSub")}</small></span></button>
      <div class="card pending" id="pendingBox"><div><div class="n">0</div><div class="lbl"></div></div><button class="btn sec">${t("sendNow")}</button></div>
      ${hist.length ? `<h2>${t("recent")}</h2><div class="card list">${hist.map(h => `
        <div class="row"><span><b>${esc(h.label)}</b><br><small>${esc(h.date)} · ${h.type === "service" ? "🔧" : "🚐"}${h.defects ? ` · <span style="color:var(--bad)">${t("defectsFound", { n: h.defects })}</span>` : ""}</small></span>
        <span class="tag ${h.status === "sent" ? "ok" : "q"}">${h.status === "sent" ? t("sent") : t("notSent")}</span></div>`).join("")}</div>` : ""}
      <div class="logo-foot"><img src="icons/logo.png" alt="DM27"></div></div>`;
    $$("[data-go]").forEach(b => b.onclick = () => { location.hash = "#/" + b.dataset.go; });
    $("#pendingBox button").onclick = () => sync(true);
    refreshPending();
  }

  // ---------- forms ----------
  function blankCheck() {
    const p = prefs.get();
    return { type: "check", checkType: "weekly", department: p.depts[0] || "", inspector1: ls.get("me", ""), inspector2: "", regularDriver: "",
      date: today(), vehicle: "", matricula: "", km: "", nextService: "", inspExp: "", licExp: "", insExp: "",
      items: {}, remarks: "", remarkPhotos: [] };
  }
  function blankService() {
    return { type: "service", date: today(), vehicle: "", matricula: "", km: "", workshop: "", enteredBy: ls.get("me", ""), parts: {}, remarks: "", remarkPhotos: [] };
  }
  async function openForm(type) {
    if (!form || form.type !== type) {
      const d = await DB.get("drafts", type);
      if (d) { form = d; setTimeout(() => { if (location.hash === "#/" + type) toast(t("resume")); }, 300); } else form = type === "check" ? blankCheck() : blankService();
    }
    type === "check" ? renderCheck() : renderService();
  }
  let saveT;
  const saveDraft = () => { clearTimeout(saveT); saveT = setTimeout(() => form && DB.put("drafts", form, form.type), 500); };

  const field = (k, label, type = "text", extra = "") =>
    `<label class="f"><span>${label}</span><input name="${k}" type="${type}" value="${esc(form[k])}" ${extra}><div class="exp-warn" data-warn="${k}"></div></label>`;

  function vehicleDatalists() {
    const p = prefs.get();
    return `<datalist id="dl-names">${p.names.map(n => `<option value="${esc(n)}">`).join("")}</datalist>
      <datalist id="dl-depts">${p.depts.map(n => `<option value="${esc(n)}">`).join("")}</datalist>
      <datalist id="dl-mat">${Object.keys(p.vehicles).map(m => `<option value="${esc(m)}">${esc(p.vehicles[m].vehicle || "")}</option>`).join("")}</datalist>
      <datalist id="dl-veh">${[...new Set(Object.values(p.vehicles).map(v => v.vehicle).filter(Boolean))].map(n => `<option value="${esc(n)}">`).join("")}</datalist>`;
  }

  function renderCheck() {
    view.classList.remove("is-home");
    const f = form;
    view.innerHTML = `
      ${vehicleDatalists()}
      <div id="errBox"></div>
      <h2>${t("checkType")}</h2>
      <div class="seg" data-seg="checkType">
        <button type="button" data-val="weekly" class="${f.checkType === "weekly" ? "on" : ""}">${t("weekly")}</button>
        <button type="button" data-val="preservice" class="${f.checkType === "preservice" ? "on" : ""}">${t("preService")}</button>
      </div>
      <h2>${t("details")}</h2>
      <div class="card">
        <div class="two">${field("matricula", t("matricula"), "text", 'list="dl-mat" autocapitalize="characters" autocomplete="off"')}${field("km", t("km"), "number", 'inputmode="numeric" min="0"')}</div>
        ${field("vehicle", t("vehicle"), "text", 'list="dl-veh" autocomplete="off"')}
        <div class="two">${field("date", t("date"), "date")}${field("department", t("department"), "text", 'list="dl-depts" autocomplete="off"')}</div>
        ${field("inspector1", t("inspector1"), "text", 'list="dl-names" autocapitalize="characters" autocomplete="off"')}
        ${field("inspector2", t("inspector2"), "text", 'list="dl-names" autocapitalize="characters" autocomplete="off"')}
        <div class="help" style="margin:-8px 0 12px">💡 ${t("inspector2Help")}</div>
        ${field("regularDriver", t("regularDriver"), "text", 'list="dl-names" autocapitalize="characters" autocomplete="off"')}
        <div class="two">${field("inspExp", t("inspExp"), "date")}${field("licExp", t("licExp"), "date")}</div>
        <div class="two">${field("insExp", t("insExp"), "date")}${field("nextService", t("nextService"), "text")}</div>
      </div>
      <div class="progress"><div class="bar"><i id="pbar"></i></div><small id="ptext"></small></div>
      <h2 style="margin-top:6px">${t("checklist")}</h2>
      ${ITEMS.map((it, i) => itemHtml(it, i)).join("")}
      <h2>${t("remarks")}</h2>
      <div class="card"><textarea name="remarks" rows="4" placeholder="${t("remarksPh")}">${esc(f.remarks)}</textarea>${photoStrip(f.remarkPhotos, "remarks")}</div>
      <div class="stop">⛔ ${t("stopWork")}</div>
      <div style="text-align:center"><button class="linkbtn" id="clearBtn">${t("clear")}</button></div>
      <div class="footer"><button class="btn" id="submitBtn">${t("submit")}</button></div>`;
    bindCommon();
    $$("[data-item]").forEach(el => bindItem(el));
    $("#submitBtn").onclick = submitCheck;
    updateProgress(); ["inspExp", "licExp", "insExp"].forEach(showExpiry);
  }

  function itemHtml(it, i) {
    const a = form.items[it.id] || {}, L = it.kind === "yesno";
    const cls = a.v ? (isBad(a.v) ? "isbad" : "isok") : "";
    const goodVal = L ? "no" : "ok", badVal = L ? "yes" : "bad";
    const goodTxt = L ? t("no") : "✓ " + t("ok"), badTxt = L ? t("yes") : "✕ " + t("notOk");
    const noteLabel = it.id === "noises" ? t("describeNoise") : it.id === "body_damage" ? t("describeDamage") : t("describe");
    return `<div class="item ${cls}" data-item="${it.id}" id="it-${it.id}">
      <div class="head"><span class="num">${i + 1}</span><span class="name">${esc(itemName(it))}${it.lights ? `<span class="lightsTag">👥 2</span>` : ""}</span>
        <button type="button" class="info" aria-label="${t("whatToLook")}">ⓘ</button></div>
      <div class="guide hidden"><b>${t("whatToLook")}:</b> ${esc(lang === "pt" ? it.gpt : it.gen)}</div>
      <div class="choice">
        ${L ? `<button type="button" class="bad ${a.v === badVal ? "on" : ""}" data-v="${badVal}">${badTxt}</button><button type="button" class="good ${a.v === goodVal ? "on" : ""}" data-v="${goodVal}">${goodTxt}</button>`
            : `<button type="button" class="good ${a.v === goodVal ? "on" : ""}" data-v="${goodVal}">${goodTxt}</button><button type="button" class="bad ${a.v === badVal ? "on" : ""}" data-v="${badVal}">${badTxt}</button>`}
      </div>
      ${it.extra === "hl" ? `<div class="hl">${t("level")}: <button type="button" data-hl="H" class="${a.hl === "H" ? "on" : ""}">H</button><button type="button" data-hl="L" class="${a.hl === "L" ? "on" : ""}">L</button></div>` : ""}
      ${it.extra === "expiry" ? `<label class="f sub" style="margin-bottom:0"><span>${t("expiry")}</span><input type="date" data-expiry value="${esc(a.expiry || "")}"><div class="exp-warn" data-warn="ext"></div></label>` : ""}
      <div class="sub badbox ${isBad(a.v) ? "" : "hidden"}">
        <textarea rows="2" data-note placeholder="${esc(noteLabel)}">${esc(a.note || "")}</textarea>
        ${photoStrip(a.photos, it.id)}
      </div>
    </div>`;
  }

  function bindItem(el) {
    const id = el.dataset.item, it = ITEMS.find(x => x.id === id);
    const a = () => (form.items[id] = form.items[id] || {});
    $(".info", el).onclick = () => $(".guide", el).classList.toggle("hidden");
    $$("[data-v]", el).forEach(b => b.onclick = () => {
      a().v = b.dataset.v;
      $$("[data-v]", el).forEach(x => x.classList.toggle("on", x === b));
      el.classList.toggle("isbad", isBad(a().v)); el.classList.toggle("isok", !isBad(a().v)); el.classList.remove("err");
      $(".badbox", el).classList.toggle("hidden", !isBad(a().v));
      if (isBad(a().v)) setTimeout(() => $("[data-note]", el).focus(), 50);
      updateProgress(); saveDraft();
    });
    $$("[data-hl]", el).forEach(b => b.onclick = () => { a().hl = b.dataset.hl; $$("[data-hl]", el).forEach(x => x.classList.toggle("on", x === b)); saveDraft(); });
    const ex = $("[data-expiry]", el);
    if (ex) { ex.onchange = ex.oninput = () => { a().expiry = ex.value; showExpiry("ext"); saveDraft(); }; showExpiry("ext"); }
    $("[data-note]", el).oninput = e => { a().note = e.target.value; e.target.classList.remove("err"); saveDraft(); };
    void it;
  }

  function showExpiry(k) {
    const el = $(`[data-warn="${k}"]`); if (!el) return;
    const v = k === "ext" ? (form.items.extinguisher || {}).expiry : form[k];
    const s = expiryState(v, form.date);
    el.className = "exp-warn " + (s === "expired" ? "bad" : s === "soon" ? "soon" : "");
    el.textContent = s === "expired" ? "⚠ " + t("expired") : s === "soon" ? "⚠ " + t("expiresSoon") : "";
  }

  function updateProgress() {
    const n = ITEMS.filter(it => form.items[it.id] && form.items[it.id].v).length;
    $("#pbar").style.width = (100 * n / ITEMS.length) + "%";
    $("#ptext").textContent = t("progress", { n, t: ITEMS.length });
  }

  // fields, photos, segments – shared by both forms
  function bindCommon() {
    $$("input[name], textarea[name]").forEach(inp => {
      inp.oninput = inp.onchange = () => {
        form[inp.name] = inp.value; inp.classList.remove("err");
        if (inp.name === "matricula") fillFromVehicle(inp.value);
        if (["inspExp", "licExp", "insExp"].includes(inp.name)) showExpiry(inp.name);
        if (inp.name === "date") { ["inspExp", "licExp", "insExp", "ext"].forEach(showExpiry); }
        saveDraft();
      };
    });
    $$("[data-seg]").forEach(seg => $$("button", seg).forEach(b => b.onclick = () => {
      form[seg.dataset.seg] = b.dataset.val; $$("button", seg).forEach(x => x.classList.toggle("on", x === b)); saveDraft();
    }));
    const cb = $("#clearBtn");
    if (cb) cb.onclick = async () => { if (!confirm(t("discard"))) return; await DB.del("drafts", form.type); form = null; openForm(location.hash.includes("service") ? "service" : "check"); scrollTo(0, 0); };
  }
  function fillFromVehicle(m) {
    const v = prefs.get().vehicles[(m || "").trim().toUpperCase()]; if (!v) return;
    ["vehicle", "inspExp", "licExp", "insExp", "nextService", "department"].forEach(k => {
      if (k in form && !form[k] && v[k]) { form[k] = v[k]; const i = $(`[name="${k}"]`); if (i) i.value = v[k]; }
    });
    ["inspExp", "licExp", "insExp"].forEach(showExpiry);
  }
  const photoArr = key => {
    if (key === "remarks") return form.remarkPhotos = form.remarkPhotos || [];
    const a = form.items[key] = form.items[key] || {}; return a.photos = a.photos || [];
  };
  async function onPhoto(e) {
    const inp = e.target; if (!inp.dataset || !inp.dataset.photo) return;
    const files = Array.from(inp.files || []); if (!files.length) return;
    const key = inp.dataset.photo;
    for (const file of files) { try { photoArr(key).push(await compress(file)); } catch (er) { toast(t("photoFail")); } }
    rerenderPhotos(key, inp.closest(".photos")); saveDraft();
  }
  function onDelPhoto(e) {
    const b = e.target.closest("[data-delphoto]"); if (!b) return;
    const key = b.dataset.delphoto; photoArr(key).splice(+b.dataset.i, 1);
    rerenderPhotos(key, b.closest(".photos")); saveDraft();
  }
  function rerenderPhotos(key, strip) {
    const tmp = document.createElement("div"); tmp.innerHTML = photoStrip(photoArr(key), key);
    strip.replaceWith(tmp.firstElementChild);
  }

  // ---------- validation + submit ----------
  function showErrors(list, firstEl) {
    $("#errBox").innerHTML = `<div class="errors"><b>${t("missing")}</b><ul>${list.map(x => `<li>${esc(x)}</li>`).join("")}</ul></div>`;
    (firstEl || $("#errBox")).scrollIntoView({ behavior: "smooth", block: "center" });
  }
  function req(names, errs, labels) {
    let first = null;
    names.forEach(n => { if (!String(form[n] || "").trim()) { const i = $(`[name="${n}"]`); i && i.classList.add("err"); first = first || i; errs.push(labels[n]); } });
    return first;
  }
  async function submitCheck() {
    const f = form, errs = [];
    let first = req(["matricula", "km", "vehicle", "date", "inspector1", "inspector2"], errs,
      { matricula: t("matricula"), km: t("km"), vehicle: t("vehicle"), date: t("date"), inspector1: t("inspector1"), inspector2: t("inspector2") });
    const unanswered = ITEMS.filter(it => !(f.items[it.id] && f.items[it.id].v));
    unanswered.forEach(it => { const el = $("#it-" + it.id); el.classList.add("err"); first = first || el; });
    if (unanswered.length) errs.push(t("mItems", { n: unanswered.length }) + " (" + unanswered.map(it => ITEMS.indexOf(it) + 1).join(", ") + ")");
    ITEMS.forEach(it => {
      const a = f.items[it.id];
      if (a && isBad(a.v) && !String(a.note || "").trim()) { const n = $("#it-" + it.id + " [data-note]"); n.classList.add("err"); first = first || n; errs.push(t("mComment", { item: itemName(it) })); }
      if (a && a.v === "ok" && it.extra === "expiry" && !a.expiry) { const n = $("#it-" + it.id + " [data-expiry]"); n.classList.add("err"); first = first || n; errs.push(itemName(it) + " – " + t("expiry")); }
      if (a && a.v === "ok" && it.extra === "hl" && !a.hl) { const n = $("#it-" + it.id); n.classList.add("err"); first = first || n; errs.push(itemName(it) + " – " + t("level") + " H / L"); }
    });
    if (errs.length) return showErrors(errs, first);
    await finish();
  }
  async function finish() {
    const f = form, rec = { id: uid(), type: f.type, created: new Date().toISOString(), lang, form: f };
    $("#submitBtn").disabled = true;
    await DB.put("queue", rec);
    await DB.del("drafts", f.type);
    if (f.type === "check") prefs.learn(f);
    if (f.inspector1 || f.enteredBy) ls.set("me", f.inspector1 || f.enteredBy);
    const defects = f.type === "check" ? defectsOf(f).length : 0;
    history.add({ id: rec.id, type: f.type, label: [f.vehicle, (f.matricula || "").toUpperCase()].filter(Boolean).join(" · "), date: f.date, defects, status: "queued" });
    lastDone = { rec, status: "queued" }; form = null;
    location.hash = "#/done";
    if (navigator.onLine) { const ok = await sync(); if (lastDone && lastDone.rec.id === rec.id && ok && !(await DB.get("queue", rec.id))) { lastDone.status = "sent"; if (location.hash === "#/done") renderDone(); } }
  }

  // ---------- service record ----------
  function renderService() {
    view.classList.remove("is-home");
    const f = form;
    view.innerHTML = `
      ${vehicleDatalists()}
      <div id="errBox"></div>
      <h2>${t("details")}</h2>
      <div class="card">
        <div class="two">${field("matricula", t("matricula"), "text", 'list="dl-mat" autocapitalize="characters" autocomplete="off"')}${field("km", t("km"), "number", 'inputmode="numeric" min="0"')}</div>
        ${field("vehicle", t("vehicle"), "text", 'list="dl-veh" autocomplete="off"')}
        <div class="two">${field("date", t("date"), "date")}${field("workshop", t("workshop"))}</div>
        ${field("enteredBy", t("yourName"), "text", 'list="dl-names" autocapitalize="characters" autocomplete="off"')}
      </div>
      <h2>${t("checklist")}</h2>
      <div class="card"><div class="help" style="margin-bottom:6px">${t("serviceIntro")}</div>
      ${PARTS.map(p => { const a = f.parts[p.id] || {}; return `<div class="svc" data-part="${p.id}">
          <b>${esc(lang === "pt" ? p.pt : p.en)}</b>
          <button type="button" class="tick ${a.v === "ok" ? "on" : a.v === "bad" ? "x" : ""}">${a.v === "ok" ? "✓" : a.v === "bad" ? "✕" : "–"}</button>
          <input class="pn" placeholder="${t("qtyPart")}" value="${esc(a.pn || "")}"></div>`; }).join("")}
      </div>
      <h2>${t("remarks")}</h2>
      <div class="card"><textarea name="remarks" rows="4" placeholder="${t("remarksPh")}">${esc(f.remarks)}</textarea>${photoStrip(f.remarkPhotos, "remarks")}</div>
      <div style="text-align:center"><button class="linkbtn" id="clearBtn">${t("clear")}</button></div>
      <div class="footer"><button class="btn" id="submitBtn">${t("submitService")}</button></div>`;
    bindCommon();
    $$("[data-part]").forEach(row => {
      const id = row.dataset.part, a = () => (f.parts[id] = f.parts[id] || {});
      const tb = $(".tick", row);
      tb.onclick = () => {
        a().v = a().v === "ok" ? "bad" : a().v === "bad" ? "" : "ok";
        tb.className = "tick " + (a().v === "ok" ? "on" : a().v === "bad" ? "x" : ""); tb.textContent = a().v === "ok" ? "✓" : a().v === "bad" ? "✕" : "–"; saveDraft();
      };
      $(".pn", row).oninput = e => { a().pn = e.target.value; saveDraft(); };
    });
    $("#submitBtn").onclick = async () => {
      const errs = [];
      const first = req(["matricula", "km", "date", "enteredBy"], errs, { matricula: t("matricula"), km: t("km"), date: t("date"), enteredBy: t("yourName") });
      if (errs.length) return showErrors(errs, first);
      await finish();
    };
  }

  // ---------- done + WhatsApp ----------
  function shareText(rec) {
    const f = rec.form, L = rec.lang, tt = k => T[L][k];
    const lines = [];
    if (rec.type === "check") {
      lines.push(tt("waHeader") + (f.checkType === "preservice" ? " (" + tt("preService") + ")" : ""));
      lines.push(`${f.vehicle} · ${(f.matricula || "").toUpperCase()}`);
      lines.push(`${tt("date")}: ${f.date} · ${tt("km")}: ${f.km}`);
      lines.push(`${tt("inspector1")}: ${f.inspector1} / ${f.inspector2}`);
      const d = defectsOf(f), w = warningsOf(f);
      lines.push("");
      if (!d.length) lines.push(tt("waAllOk"));
      else { lines.push(tt("waDefects")); d.forEach(x => lines.push(`• ${itemName(x.it, L)}: ${x.note}`)); }
      const o = (f.items.oil || {}).hl, c = (f.items.coolant || {}).hl;
      if (o || c) lines.push(`🛢 ${L === "pt" ? "Óleo" : "Oil"} ${o || "-"} · ${L === "pt" ? "Refrigerante" : "Coolant"} ${c || "-"}`);
      if (w.length) { lines.push(""); w.forEach(x => lines.push("⚠ " + x)); }
    } else {
      lines.push(tt("waService"));
      lines.push(`${f.vehicle || ""} · ${(f.matricula || "").toUpperCase()}`);
      lines.push(`${tt("date")}: ${f.date} · ${tt("km")}: ${f.km}`);
      PARTS.forEach(p => { const a = f.parts[p.id]; if (a && a.v) lines.push(`${a.v === "ok" ? "✓" : "✕"} ${L === "pt" ? p.pt : p.en}${a.pn ? " – " + a.pn : ""}`); });
    }
    if (f.remarks) { lines.push(""); lines.push(f.remarks); }
    return lines.join("\n");
  }
  function allPhotos(rec) {
    const f = rec.form, out = [...(f.remarkPhotos || [])];
    if (f.items) Object.values(f.items).forEach(a => (a.photos || []).forEach(p => out.push(p)));
    return out;
  }
  async function share(rec) {
    const text = shareText(rec), photos = allPhotos(rec);
    if (photos.length && navigator.canShare) {
      try {
        const files = await Promise.all(photos.map(async (p, i) => new File([await (await fetch(p)).blob()], `photo-${i + 1}.jpg`, { type: "image/jpeg" })));
        if (navigator.canShare({ files })) { await navigator.share({ files, text }); return; }
      } catch (e) { if (e && e.name === "AbortError") return; }
    }
    location.href = "https://wa.me/?text=" + encodeURIComponent(text);
  }
  function renderDone() {
    view.classList.remove("is-home");
    const { rec, status } = lastDone, f = rec.form;
    const d = rec.type === "check" ? defectsOf(f) : [];
    view.innerHTML = `<div class="done">
      <div class="big">${d.length ? "⚠️" : "✅"}</div>
      <h1>${t("savedTitle")}</h1>
      <p style="color:var(--muted)">${status === "sent" ? t("savedSent") : t("savedQueued")}</p>
      <div class="card summary">
        <b>${esc(f.vehicle || "")} · ${esc((f.matricula || "").toUpperCase())}</b><br><small>${esc(f.date)} · ${esc(f.km)} km</small>
        ${rec.type === "check" ? `<p style="margin:10px 0 0;font-weight:700;color:${d.length ? "var(--bad)" : "var(--ok)"}">${d.length ? t("defectsFound", { n: d.length }) : t("allOk")}</p>
          ${d.length ? `<ul style="margin:6px 0 0 18px;padding:0">${d.map(x => `<li>${esc(itemName(x.it))}: ${esc(x.note)}</li>`).join("")}</ul>` : ""}` : ""}
      </div>
      ${d.length ? `<div class="stop">⛔ ${t("stopWork")}</div>` : ""}
      <button class="btn wa" id="waBtn">💬 ${t("share")}</button>
      <button class="btn sec" id="doneBtn">${t("done")}</button></div>`;
    $("#waBtn").onclick = () => share(rec);
    $("#doneBtn").onclick = () => { lastDone = null; location.hash = "#/"; };
    $("#backBtn").classList.add("hidden");
  }

  // ---------- start ----------
  view.addEventListener("change", e => form && onPhoto(e));
  view.addEventListener("click", e => form && onDelPhoto(e));
  if ("serviceWorker" in navigator && location.protocol !== "file:") navigator.serviceWorker.register("sw.js").catch(() => {});
  if (location.hash === "#/done") location.hash = "#/";
  route();
  sync();
  // test hook
  window.__vc = { buildPayload, DB, sync, shareText };
})();
