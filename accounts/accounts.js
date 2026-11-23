// Fleet Accounts page (PC): enter costs for workshop jobs, record running costs (fuel, insurance, Imposto…),
// see every vehicle's costs by site, and produce vehicle / site / fleet reports with Excel export.
// Talks to the same Apps Script web app as the phones and the controller page.
(function () {
  "use strict";
  const CFG = window.APP_CONFIG;
  const $ = (s, el = document) => el.querySelector(s), $$ = (s, el = document) => [...el.querySelectorAll(s)];
  const store = {
    get: (k, d) => { try { const v = localStorage.getItem("acc." + k); return v === null ? d : JSON.parse(v); } catch (e) { return d; } },
    set: (k, v) => { try { localStorage.setItem("acc." + k, JSON.stringify(v)); } catch (e) {} },
    del: k => { try { localStorage.removeItem("acc." + k); } catch (e) {} }
  };
  const esc = s => String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const money = n => n == null || n === "" || isNaN(n) ? "" : Number(n).toLocaleString("en-GB", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const int = n => n == null || n === "" || isNaN(n) ? "" : Number(n).toLocaleString("en-GB", { maximumFractionDigits: 1 });
  const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const longDate = v => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v || ""); return m ? m[3] + " " + MONTHS[+m[2] - 1] + " " + m[1] : ""; };
  const iso = d => d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
  const today = () => iso(new Date());
  const uid = () => "RC-" + Date.now().toString(36) + "-" + Math.random().toString(36).slice(2, 6);
  const cmp = (x, y) => String(x || "").localeCompare(String(y || ""), "en", { sensitivity: "base", numeric: true });
  const firstLine = s => String(s || "").split("\n")[0];
  let DATA = null, CATS = [];
  const state = { from: "", to: "", quick: store.get("quick", "year"), vehicle: null };

  function toast(msg, ms = 3500) { const t = $("#toast"); t.textContent = msg; t.classList.remove("hidden"); clearTimeout(t._h); t._h = setTimeout(() => t.classList.add("hidden"), ms); }

  // ---------- date boxes shown as "21 Nov 2026" (same as the controller page) ----------
  const desc = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value");
  function dateBox(inp) {
    const wrap = document.createElement("span"); wrap.className = "datebox";
    inp.parentNode.insertBefore(wrap, inp); wrap.appendChild(inp);
    const shown = document.createElement("span"); shown.className = "datetext"; wrap.appendChild(shown);
    const show = () => { const v = desc.get.call(inp); shown.textContent = longDate(v) || inp.dataset.empty || "";
      shown.classList.toggle("ph", !v); wrap.classList.toggle("has", !!shown.textContent); };
    Object.defineProperty(inp, "value", { get() { return desc.get.call(this); }, set(v) { desc.set.call(this, v); show(); } });
    ["input", "change"].forEach(ev => inp.addEventListener(ev, show));
    inp.form && inp.form.addEventListener("reset", () => setTimeout(show));
    show();
  }
  $$('input[type="date"]').forEach(dateBox);

  // ---------- loading ----------
  async function load() {
    const akey = store.get("akey", "");
    if (!akey) return showGate();
    $("#loading").classList.remove("hidden"); $("#loading").textContent = "Loading costs…";
    try {
      const res = await fetch(CFG.ENDPOINT + "?action=accounts&akey=" + encodeURIComponent(akey), { redirect: "follow" });
      const j = await res.json();
      if (!j.ok) { if (/key/.test(j.error)) { store.del("akey"); return showGate("That key wasn't accepted. Check the Settings tab of the Sheet."); } throw new Error(j.error); }
      DATA = j; CATS = ["Workshop"].concat(j.categories);
      $("#loading").classList.add("hidden"); $("#gate").classList.add("hidden");
      $("#tabs").classList.remove("hidden"); $("#reload").classList.remove("hidden");
      fillLists(); renderAll();
      showTab(location.hash.replace("#", "") || "needs");
    } catch (e) {
      $("#loading").innerHTML = '<span class="err">Could not reach the Sheet: ' + esc(e.message) + '</span> <button class="btn sec" onclick="location.reload()">Try again</button>';
    }
  }
  function showGate(msg) { $("#loading").classList.add("hidden"); $("#gate").classList.remove("hidden"); $("#gateMsg").textContent = msg || ""; }
  $("#gateForm").addEventListener("submit", e => { e.preventDefault(); store.set("akey", $("#akey").value.trim()); $("#gate").classList.add("hidden"); load(); });
  $("#reload").addEventListener("click", () => load());

  // ---------- tabs ----------
  const TABS = ["needs", "vehicles", "running", "reports"];
  function showTab(t) {
    if (!TABS.includes(t)) t = "needs";
    $$("#tabs button").forEach(b => b.classList.toggle("on", b.dataset.tab === t));
    TABS.forEach(x => $("#tab-" + x).classList.toggle("hidden", x !== t));
    if (t === "running" && !$("#runForm").by.value) $("#runForm").by.value = store.get("me", "");
    history.replaceState(null, "", "#" + t);
  }
  $$("#tabs button").forEach(b => b.addEventListener("click", () => showTab(b.dataset.tab)));

  // ---------- lists ----------
  const sites = () => [...new Set(DATA.fleet.map(v => v.site).concat(DATA.jobs.map(j => j.site), DATA.running.map(c => c.site)).filter(Boolean))].sort(cmp);
  const vehiclesOf = site => DATA.fleet.filter(v => v.site === site).sort((a, b) => cmp(a.make, b.make) || cmp(a.reg, b.reg));
  function fillLists() {
    const ss = sites(), siteOpts = '<option value="">All sites</option>' + ss.map(s => `<option>${esc(s)}</option>`).join("");
    ["#needSite", "#vehSite"].forEach(id => { const el = $(id), cur = el.value; el.innerHTML = siteOpts; el.value = cur; });
    const regOpts = ss.map(s => `<optgroup label="${esc(s)}">` + vehiclesOf(s).map(v => `<option value="${esc(v.reg)}">${esc(v.reg)} – ${esc(v.make)}</option>`).join("") + "</optgroup>").join("");
    const rf = $("#runForm").reg, cur = rf.value; rf.innerHTML = '<option value="">– choose vehicle –</option>' + regOpts; rf.value = cur;
    const cf = $("#runForm").category, cc = cf.value; cf.innerHTML = DATA.categories.map(c => `<option>${esc(c)}</option>`).join(""); if (cc) cf.value = cc;
    const sc = $("#repScope"), scv = sc.value;
    sc.innerHTML = '<option value="fleet">Whole fleet</option><optgroup label="Site">' + ss.map(s => `<option value="site:${esc(s)}">${esc(s)}</option>`).join("") +
      "</optgroup>" + ss.map(s => `<optgroup label="${esc(s)} vehicles">` + vehiclesOf(s).map(v => `<option value="veh:${esc(v.reg)}">${esc(v.reg)} – ${esc(v.make)}</option>`).join("") + "</optgroup>").join("");
    sc.value = scv || "fleet";
    $("#suppliers").innerHTML = DATA.suppliers.map(s => `<option value="${esc(s)}">`).join("");
    $("#workshops").innerHTML = DATA.workshops.map(s => `<option value="${esc(s)}">`).join("");
    $$(".cur").forEach(el => el.textContent = DATA.currency);
    const me = store.get("me", ""); if (me) { $("#runForm").by.value = $("#runForm").by.value || me; }
    if (!$("#runForm").date.value) $("#runForm").date.value = today();
  }

  // ---------- period ----------
  const QUICK = [["month", "This month"], ["lastmonth", "Last month"], ["year", "This year"], ["lastyear", "Last year"], ["12m", "Last 12 months"], ["all", "All time"]];
  function quickRange(q) {
    const d = new Date(), y = d.getFullYear(), m = d.getMonth();
    if (q === "month") return [iso(new Date(y, m, 1)), today()];
    if (q === "lastmonth") return [iso(new Date(y, m - 1, 1)), iso(new Date(y, m, 0))];
    if (q === "year") return [iso(new Date(y, 0, 1)), today()];
    if (q === "lastyear") return [iso(new Date(y - 1, 0, 1)), iso(new Date(y - 1, 11, 31))];
    if (q === "12m") return [iso(new Date(y, m - 11, 1)), today()];
    return ["", ""];
  }
  function setQuick(q) { state.quick = q; store.set("quick", q); [state.from, state.to] = quickRange(q); renderAll(); }
  function renderPeriod() {
    $$("[data-period]").forEach(box => {
      box.innerHTML = `<b>Period</b><label>From <input type="date" class="pf" data-empty="any date" value="${esc(state.from)}"></label><label>to <input type="date" class="pt" data-empty="today" value="${esc(state.to)}"></label>` +
        QUICK.map(([k, t]) => `<button type="button" class="q ${state.quick === k ? "on" : ""}" data-q="${k}">${t}</button>`).join("") +
        `<span class="muted small">${state.from || state.to ? (longDate(state.from) || "start") + " – " + (longDate(state.to) || "today") : "all dates"}</span>`;
      $$('input[type="date"]', box).forEach(dateBox);
      $$(".q", box).forEach(b => b.onclick = () => setQuick(b.dataset.q));
      $(".pf", box).onchange = e => { state.from = e.target.value; state.quick = ""; renderAll(); };
      $(".pt", box).onchange = e => { state.to = e.target.value; state.quick = ""; renderAll(); };
    });
  }
  const inPeriod = d => (!state.from || (d && d >= state.from)) && (!state.to || (d && d <= state.to));
  const periodText = () => state.from || state.to ? (longDate(state.from) || "start") + " to " + (longDate(state.to) || longDate(today())) : "All dates";

  // ---------- cost lines (costed jobs + running costs) ----------
  function lines(filter) {
    const out = [];
    DATA.jobs.forEach(j => { if (j.total && inPeriod(j.date) && filter(j)) out.push({ kind: "job", id: j.id, date: j.date, site: j.site, reg: j.reg, make: j.make,
      cat: "Workshop", detail: j.type, desc: firstLine(j.work) || firstLine(j.parts), supplier: j.workshop, invNo: j.invNo, invoice: j.invoice,
      parts: j.partsCost, labour: j.labour, other: j.other, litres: null, km: j.km, amount: j.total, by: j.costedBy || j.enteredBy }); });
    DATA.running.forEach(c => { if (inPeriod(c.date) && filter(c)) out.push({ kind: "run", id: c.id, date: c.date, site: c.site, reg: c.reg, make: c.make,
      cat: c.category, detail: "", desc: c.desc, supplier: c.supplier, invNo: c.invNo, invoice: c.invoice, parts: null, labour: null, other: null,
      litres: c.litres, km: c.km, amount: c.amount, by: c.enteredBy }); });
    return out.sort((a, b) => cmp(b.date, a.date));
  }
  const byCat = ls => { const o = {}; CATS.forEach(c => o[c] = 0); ls.forEach(l => o[l.cat] = (o[l.cat] || 0) + (l.amount || 0)); o.total = ls.reduce((t, l) => t + (l.amount || 0), 0);
    o.litres = ls.reduce((t, l) => t + (l.cat === "Fuel" ? l.litres || 0 : 0), 0); return o; };
  const pending = filter => DATA.jobs.filter(j => j.needsCost && filter(j)).sort((a, b) => cmp(a.date, b.date));
  const scopeFilter = sc => sc === "fleet" ? (() => true) : sc.startsWith("site:") ? (x => x.site === sc.slice(5)) : (x => x.reg === sc.slice(4));
  const scopeName = sc => sc === "fleet" ? "Whole fleet" : sc.startsWith("site:") ? sc.slice(5) + " site" : (v => v ? v.reg + " – " + v.make : sc.slice(4))(DATA.fleet.find(v => v.reg === sc.slice(4)));

  function renderAll() { if (!DATA) return; renderPeriod(); renderNeeds(); renderVehicles(); renderRunRecent(); renderReport(); }

  // ---------- needs costs ----------
  function renderNeeds() {
    const site = $("#needSite").value, only = $("#needOnly").checked;
    const all = pending(() => true);
    $("#needBadge").textContent = all.length; $("#needBadge").classList.toggle("hidden", !all.length);
    const shown = sites().filter(s => !site || s === site);
    $("#needSummary").innerHTML = [[all.length, "jobs waiting for costs"]].concat(sites().map(s => [all.filter(j => j.site === s).length, s]))
      .map(x => `<div class="${x[0] ? "warnbox" : ""}"><b>${x[0]}</b><span>${esc(x[1])}</span></div>`).join("");
    let html = "";
    shown.forEach(s => {
      const vs = vehiclesOf(s), sp = all.filter(j => j.site === s);
      const extra = [...new Set(sp.map(j => j.reg))].filter(r => !vs.some(v => v.reg === r)).map(r => ({ reg: r, make: (sp.find(j => j.reg === r) || {}).make || "", site: s }));
      const list = vs.concat(extra).filter(v => !only || sp.some(j => j.reg === v.reg));
      if (only && !list.length) return;
      html += `<div class="site-h"><b>${esc(s)}</b><span>${sp.length ? sp.length + " job" + (sp.length > 1 ? "s" : "") + " to cost" : "all costed ✓"}</span></div>`;
      list.forEach(v => {
        const js = sp.filter(j => j.reg === v.reg);
        html += `<div class="vrow ${js.length ? "need" : ""}"><b>${esc(v.reg)}</b><span>${esc(v.make)}</span><span class="cnt">${js.length ? js.length + " job" + (js.length > 1 ? "s" : "") + " need costs" : "nothing to cost"}</span><span></span></div>`;
        if (js.length) html += '<div class="jobs">' + js.map(j => `<div class="job"><span>${esc(longDate(j.date))}<br><small>${j.km ? int(j.km) + " km" : ""}</small></span>
          <span><b>${esc(j.type)}</b><br><small>${esc(j.source)}</small></span><span class="w">${esc(firstLine(j.work) || firstLine(j.parts) || "—")}<br><small>${esc(j.workshop)}${j.enteredBy ? " · entered by " + esc(j.enteredBy) : ""}</small></span>
          <span><small>${esc(j.invNo ? "Invoice " + j.invNo : "")}</small></span><button class="btn slim" data-job="${esc(j.id)}">Enter costs</button></div>`).join("") + "</div>";
      });
    });
    $("#needList").innerHTML = html || '<p class="muted">Nothing waiting for costs ✓</p>';
    $$("#needList [data-job]").forEach(b => b.onclick = () => openJob(b.dataset.job));
  }
  $("#needSite").addEventListener("change", renderNeeds);
  $("#needOnly").addEventListener("change", renderNeeds);

  // ---------- vehicles ----------
  const catHead = () => CATS.map(c => `<th class="num">${esc(c)}</th>`).join("") + '<th class="num">Total</th>';
  const catCells = o => CATS.map(c => `<td class="num ${o[c] ? "" : "zero"}">${o[c] ? money(o[c]) : "–"}</td>`).join("") + `<td class="num"><b>${money(o.total)}</b></td>`;
  function renderVehicles() {
    if (state.vehicle) return renderVehicle(state.vehicle);
    const site = $("#vehSite").value;
    const all = lines(() => true), pend = pending(() => true);
    let html = `<thead><tr><th>Site</th><th>Matrícula</th><th>Make</th>${catHead()}<th class="num">Fuel litres</th><th>To cost</th></tr></thead><tbody>`;
    sites().filter(s => !site || s === site).forEach((s, si) => {
      const vs = vehiclesOf(s);
      vs.forEach((v, i) => {
        const o = byCat(all.filter(l => l.reg === v.reg)), np = pend.filter(j => j.reg === v.reg).length;
        html += `<tr data-reg="${esc(v.reg)}" class="${i === 0 && si ? "site-start" : ""} ${np ? "needrow" : ""}"><td>${esc(s)}</td><td><b>${esc(v.reg)}</b></td><td>${esc(v.make)}</td>${catCells(o)}<td class="num">${o.litres ? int(o.litres) : ""}</td><td>${np ? `<span class="tag need">${np} job${np > 1 ? "s" : ""}</span>` : ""}</td></tr>`;
      });
      const so = byCat(all.filter(l => l.site === s));
      html += `<tr class="tot"><td colspan="3">${esc(s)} total</td>${catCells(so)}<td class="num">${so.litres ? int(so.litres) : ""}</td><td></td></tr>`;
    });
    if (!site) { const fo = byCat(all); html += `<tr class="tot"><td colspan="3">Whole fleet</td>${catCells(fo)}<td class="num">${fo.litres ? int(fo.litres) : ""}</td><td></td></tr>`; }
    $("#vehTable").innerHTML = html + "</tbody>";
    $$("#vehTable tbody tr[data-reg]").forEach(tr => tr.onclick = () => { state.vehicle = tr.dataset.reg; renderVehicle(state.vehicle); scrollTo(0, 0); });
    $("#vehListCard").classList.remove("hidden"); $("#vehDetailCard").classList.add("hidden");
  }
  $("#vehSite").addEventListener("change", renderVehicles);
  function lineRows(ls, withVeh) {
    return ls.map(l => `<tr data-kind="${l.kind}" data-id="${esc(l.id)}"><td>${esc(longDate(l.date))}</td>${withVeh ? `<td>${esc(l.site)}</td><td><b>${esc(l.reg)}</b></td>` : ""}
      <td><span class="tag ${esc(l.cat)}">${esc(l.cat)}</span>${l.detail ? " " + esc(l.detail) : ""}</td><td>${esc(l.desc)}</td><td>${esc(l.supplier)}</td>
      <td>${esc(l.invNo)}${l.invoice ? ` <a class="inv" href="${esc(l.invoice)}" target="_blank" rel="noopener">📄 view</a>` : ""}</td>
      <td class="num">${l.litres ? int(l.litres) + " L" : ""}</td><td class="num"><b>${money(l.amount)}</b></td><td><button class="btn sec slim">Edit</button></td></tr>`).join("");
  }
  function renderVehicle(reg) {
    const v = DATA.fleet.find(x => x.reg === reg) || { reg, make: "", site: "" };
    const ls = lines(l => l.reg === reg), o = byCat(ls), pend = pending(j => j.reg === reg);
    $("#vehTitle").textContent = `${v.reg} – ${v.make} · ${v.site}`;
    $("#vehSummary").innerHTML = [[money(o.total), "total " + DATA.currency]].concat(CATS.filter(c => o[c]).map(c => [money(o[c]), c]))
      .concat(o.litres ? [[int(o.litres), "fuel litres"]] : []).concat([[pend.length, "jobs waiting for costs"]])
      .map(x => `<div><b>${x[0]}</b><span>${esc(x[1])}</span></div>`).join("");
    $("#vehLines").innerHTML = `<thead><tr><th>Date</th><th>Category</th><th>Description</th><th>Supplier / workshop</th><th>Invoice</th><th class="num">Litres</th><th class="num">Amount (${esc(DATA.currency)})</th><th></th></tr></thead><tbody>` +
      (ls.length ? lineRows(ls, false) : `<tr><td colspan="8" class="muted">No costs in this period.</td></tr>`) +
      (pend.length ? pend.map(j => `<tr class="needrow" data-kind="job" data-id="${esc(j.id)}"><td>${esc(longDate(j.date))}</td><td><span class="tag need">Needs costs</span> ${esc(j.type)}</td><td>${esc(firstLine(j.work) || firstLine(j.parts))}</td><td>${esc(j.workshop)}</td><td>${esc(j.invNo)}</td><td></td><td></td><td><button class="btn slim">Enter costs</button></td></tr>`).join("") : "") +
      `<tr class="tot"><td colspan="6">Total for the period</td><td class="num">${money(o.total)}</td><td></td></tr></tbody>`;
    $$("#vehLines tbody tr[data-id]").forEach(tr => tr.onclick = () => tr.dataset.kind === "job" ? openJob(tr.dataset.id) : editRunning(tr.dataset.id));
    $("#vehListCard").classList.add("hidden"); $("#vehDetailCard").classList.remove("hidden");
  }
  $("#vehBack").onclick = () => { state.vehicle = null; renderVehicles(); };
  $("#vehExport").onclick = () => exportReport("veh:" + state.vehicle);
  $("#vehAddRun").onclick = () => { newRunning(); $("#runForm").reg.value = state.vehicle; showTab("running"); };

  // ---------- job costs dialog ----------
  let curJob = null;
  const dlg = $("#jobDlg"), jf = $("#jobForm");
  function openJob(id) {
    const j = DATA.jobs.find(x => x.id === id); if (!j) return;
    curJob = j; jf.reset(); $("#jobMsg").textContent = "";
    $("#jobTitle").textContent = (j.needsCost ? "Enter costs – " : "Edit costs – ") + j.reg + " · " + j.make;
    $("#jobInfo").innerHTML = `<b>${esc(longDate(j.date))}</b> · ${esc(j.type)}${j.km ? " · " + int(j.km) + " km" : ""} · ${esc(j.site)}` +
      (j.work ? "\n<b>Work:</b> " + esc(j.work) : "") + (j.parts ? "\n<b>Parts:</b> " + esc(j.parts) : "") +
      "\n<small>" + esc([j.source, j.enteredBy ? "entered by " + j.enteredBy : "", j.costedBy ? "costed by " + j.costedBy : ""].filter(Boolean).join(" · ")) + "</small>" +
      (j.invoice ? `\n<a href="${esc(j.invoice)}" target="_blank" rel="noopener">📄 Invoice already attached</a>` : "");
    jf.parts.value = j.partsCost ?? ""; jf.labour.value = j.labour ?? ""; jf.other.value = j.other ?? "";
    if (!j.partsCost && !j.labour && !j.other && j.total) jf.other.value = j.total;
    jf.workshop.value = j.workshop || ""; jf.invNo.value = j.invNo || ""; jf.by.value = store.get("me", "");
    jobTotal(); dlg.showModal();
  }
  function jobTotal() { $("#jobTotal").textContent = money((+jf.parts.value || 0) + (+jf.labour.value || 0) + (+jf.other.value || 0)) || "0.00"; }
  ["parts", "labour", "other"].forEach(n => jf[n].addEventListener("input", jobTotal));
  $("#jobCancel").onclick = () => dlg.close();
  $("#jobSave").onclick = async () => {
    const total = (+jf.parts.value || 0) + (+jf.labour.value || 0) + (+jf.other.value || 0);
    jf.by.classList.toggle("invalid", !jf.by.value.trim());
    if (!total) return $("#jobMsg").textContent = "Enter at least one cost.";
    if (!jf.by.value.trim()) return $("#jobMsg").textContent = "Enter your name.";
    const inv = await fileData(jf.invoice.files[0]); if (inv === false) return;
    $("#jobSave").disabled = true; $("#jobMsg").textContent = "Saving…";
    try {
      const j = await post({ kind: "jobcost", id: curJob.id, by: jf.by.value.trim(), invoice: inv,
        costs: { parts: jf.parts.value, labour: jf.labour.value, other: jf.other.value, invNo: jf.invNo.value.trim(), workshop: jf.workshop.value.trim(), notes: jf.notes.value.trim() } });
      store.set("me", jf.by.value.trim());
      Object.assign(curJob, { partsCost: +jf.parts.value || null, labour: +jf.labour.value || null, other: +jf.other.value || null, total: j.total,
        invNo: jf.invNo.value.trim(), workshop: jf.workshop.value.trim() || curJob.workshop, needsCost: false, costedBy: jf.by.value.trim() });
      if (inv) curJob.invoice = curJob.invoice || "#saved";
      dlg.close(); renderAll(); toast("Costs saved – " + money(j.total) + " " + DATA.currency + ". Press Refresh to load the invoice link.");
    } catch (e) { $("#jobMsg").textContent = "Not saved: " + e.message; }
    finally { $("#jobSave").disabled = false; }
  };

  // ---------- running costs ----------
  const rf = $("#runForm");
  let editingRun = null;
  function fuelToggle() { $(".fuelOnly").classList.toggle("off", rf.category.value !== "Fuel"); }
  rf.category.addEventListener("change", fuelToggle);
  function newRunning() {
    editingRun = null; rf.reset(); rf.date.value = today(); rf.by.value = store.get("me", "");
    $("#runTitle").textContent = "Add a running cost"; $("#runSave").textContent = "Save running cost";
    $("#runCancel").classList.add("hidden"); $("#runDelete").classList.add("hidden"); $("#runMsg").textContent = ""; fuelToggle();
  }
  function editRunning(id) {
    const c = DATA.running.find(x => x.id === id); if (!c) return;
    newRunning(); editingRun = c;
    rf.reg.value = c.reg; rf.date.value = c.date; rf.category.value = c.category; rf.amount.value = c.amount ?? ""; rf.litres.value = c.litres ?? "";
    rf.km.value = c.km ?? ""; rf.supplier.value = c.supplier; rf.desc.value = c.desc; rf.invNo.value = c.invNo; rf.by.value = c.enteredBy || store.get("me", ""); rf.notes.value = c.notes || "";
    $("#runTitle").textContent = "Edit running cost – " + c.reg; $("#runSave").textContent = "Save changes";
    $("#runCancel").classList.remove("hidden"); $("#runDelete").classList.remove("hidden"); fuelToggle();
    showTab("running"); scrollTo(0, 0);
  }
  $("#runCancel").onclick = newRunning;
  rf.addEventListener("submit", async e => {
    e.preventDefault();
    const need = ["reg", "date", "category", "amount", "by"].filter(n => !String(rf[n].value).trim());
    $$("#runForm .invalid").forEach(x => x.classList.remove("invalid")); need.forEach(n => rf[n].classList.add("invalid"));
    if (need.length) return $("#runMsg").textContent = "Fill in the boxes marked in red.";
    const inv = await fileData(rf.invoice.files[0]); if (inv === false) return;
    const id = editingRun ? editingRun.id : uid();
    const row = { "Date": rf.date.value, "Matrícula": rf.reg.value, "Category": rf.category.value, "Amount": rf.amount.value,
      "Litres": rf.category.value === "Fuel" ? rf.litres.value : "", "Km": rf.km.value, "Supplier": rf.supplier.value.trim(),
      "Description": rf.desc.value.trim(), "Invoice no.": rf.invNo.value.trim(), "Entered by": rf.by.value.trim(), "Notes": rf.notes.value.trim() };
    $("#runSave").disabled = true; $("#runMsg").textContent = "Saving…";
    try {
      await post({ kind: "running", id, row, invoice: inv });
      store.set("me", rf.by.value.trim());
      const v = DATA.fleet.find(x => x.reg === rf.reg.value) || {};
      const rec = { id, date: row.Date, reg: row["Matrícula"], make: v.make || "", site: v.site || "", category: row.Category, desc: row.Description,
        amount: +row.Amount, litres: row.Litres ? +row.Litres : null, km: row.Km ? +row.Km : null, supplier: row.Supplier, invNo: row["Invoice no."],
        invoice: editingRun ? editingRun.invoice : (inv ? "#saved" : ""), enteredBy: row["Entered by"], notes: row.Notes };
      if (editingRun) Object.assign(editingRun, rec); else DATA.running.push(rec);
      if (row.Supplier && !DATA.suppliers.includes(row.Supplier)) { DATA.suppliers.push(row.Supplier); fillLists(); }
      toast((editingRun ? "Changes saved – " : "Saved – ") + rec.reg + " " + rec.category + " " + money(rec.amount) + " " + DATA.currency);
      const keep = { reg: rf.reg.value, date: rf.date.value, category: rf.category.value };
      newRunning(); rf.reg.value = keep.reg; rf.date.value = keep.date; rf.category.value = keep.category; fuelToggle();
      renderAll();
    } catch (err) { $("#runMsg").textContent = "Not saved: " + err.message; }
    finally { $("#runSave").disabled = false; }
  });
  $("#runDelete").onclick = async () => {
    if (!editingRun || !confirm(`Delete this ${editingRun.category} cost of ${money(editingRun.amount)} ${DATA.currency} on ${editingRun.reg}?`)) return;
    try {
      await post({ kind: "running", id: editingRun.id, del: true });
      DATA.running = DATA.running.filter(c => c !== editingRun); toast("Deleted."); newRunning(); renderAll();
    } catch (e) { $("#runMsg").textContent = "Not deleted: " + e.message; }
  };
  function renderRunRecent() {
    const ls = DATA.running.slice().sort((a, b) => cmp(b.date, a.date)).slice(0, 30);
    $("#runRecent").innerHTML = `<thead><tr><th>Date</th><th>Site</th><th>Matrícula</th><th>Category</th><th>Description</th><th>Supplier</th><th class="num">Litres</th><th class="num">Amount (${esc(DATA.currency)})</th></tr></thead><tbody>` +
      (ls.length ? ls.map(c => `<tr data-id="${esc(c.id)}"><td>${esc(longDate(c.date))}</td><td>${esc(c.site)}</td><td><b>${esc(c.reg)}</b></td><td><span class="tag ${esc(c.category)}">${esc(c.category)}</span></td><td>${esc(c.desc)}</td><td>${esc(c.supplier)}</td><td class="num">${c.litres ? int(c.litres) : ""}</td><td class="num">${money(c.amount)}</td></tr>`).join("")
        : '<tr><td colspan="8" class="muted">No running costs recorded yet.</td></tr>') + "</tbody>";
    $$("#runRecent tbody tr[data-id]").forEach(tr => tr.onclick = () => editRunning(tr.dataset.id));
  }

  // ---------- reports ----------
  function reportData(sc) {
    const f = scopeFilter(sc), ls = lines(f), pend = pending(f).filter(j => inPeriod(j.date));
    const vehs = sc.startsWith("veh:") ? DATA.fleet.filter(v => v.reg === sc.slice(4)) : sites().filter(s => sc === "fleet" || "site:" + s === sc).flatMap(vehiclesOf);
    const months = [...new Set(ls.map(l => l.date.slice(0, 7)))].sort();
    return { sc, name: scopeName(sc), ls, pend, vehs, months, all: byCat(ls),
      bySite: sc === "fleet" ? sites().map(s => ({ site: s, o: byCat(ls.filter(l => l.site === s)), n: vehiclesOf(s).length })) : [],
      byVeh: vehs.map(v => ({ v, o: byCat(ls.filter(l => l.reg === v.reg)) })),
      byMonth: months.map(m => ({ m, o: byCat(ls.filter(l => l.date.slice(0, 7) === m)) })) };
  }
  const monthName = m => MONTHS[+m.slice(5, 7) - 1] + " " + m.slice(0, 4);
  function renderReport() {
    const sc = $("#repScope").value || "fleet", r = reportData(sc);
    $("#repSummary").innerHTML = [[money(r.all.total), "total " + DATA.currency], [money(r.all.Workshop), "workshop jobs"],
      [money(r.all.total - r.all.Workshop), "running costs"], [r.all.litres ? int(r.all.litres) : "0", "fuel litres"], [r.pend.length, "jobs still to cost"]]
      .map((x, i) => `<div class="${i === 4 && x[0] ? "warnbox" : ""}"><b>${x[0]}</b><span>${esc(x[1])}</span></div>`).join("");
    $("#repCat").innerHTML = `<thead><tr><th>Category</th><th class="num">Amount (${esc(DATA.currency)})</th><th class="num">Share</th></tr></thead><tbody>` +
      CATS.map(c => `<tr><td><span class="tag ${esc(c)}">${esc(c)}</span></td><td class="num ${r.all[c] ? "" : "zero"}">${r.all[c] ? money(r.all[c]) : "–"}</td><td class="num">${r.all.total ? Math.round(100 * r.all[c] / r.all.total) + "%" : ""}</td></tr>`).join("") +
      `<tr class="tot"><td>Total</td><td class="num">${money(r.all.total)}</td><td></td></tr></tbody>`;
    $("#repSiteBox").classList.toggle("hidden", !r.bySite.length);
    $("#repSite").innerHTML = `<thead><tr><th>Site</th><th class="num">Vehicles</th>${catHead()}</tr></thead><tbody>` +
      r.bySite.map(x => `<tr><td><b>${esc(x.site)}</b></td><td class="num">${x.n}</td>${catCells(x.o)}</tr>`).join("") + `<tr class="tot"><td>Whole fleet</td><td class="num">${DATA.fleet.length}</td>${catCells(r.all)}</tr></tbody>`;
    $("#repVehBox").classList.toggle("hidden", r.byVeh.length < 2);
    $("#repVeh").innerHTML = `<thead><tr><th>Site</th><th>Matrícula</th><th>Make</th>${catHead()}</tr></thead><tbody>` +
      r.byVeh.map((x, i) => `<tr class="${i && r.byVeh[i - 1].v.site !== x.v.site ? "site-start" : ""}"><td>${esc(x.v.site)}</td><td><b>${esc(x.v.reg)}</b></td><td>${esc(x.v.make)}</td>${catCells(x.o)}</tr>`).join("") +
      `<tr class="tot"><td colspan="3">Total</td>${catCells(r.all)}</tr></tbody>`;
    $("#repMonth").innerHTML = `<thead><tr><th>Month</th>${catHead()}</tr></thead><tbody>` +
      (r.byMonth.length ? r.byMonth.map(x => `<tr><td>${monthName(x.m)}</td>${catCells(x.o)}</tr>`).join("") : `<tr><td colspan="${CATS.length + 2}" class="muted">No costs in this period.</td></tr>`) +
      `<tr class="tot"><td>Total</td>${catCells(r.all)}</tr></tbody>`;
    $("#repNote").textContent = r.pend.length ? `${r.pend.length} job(s) in this period have no costs yet and are not included in these totals – see "Needs costs".` : "";
  }
  $("#repScope").addEventListener("change", renderReport);
  $("#repExport").onclick = () => exportReport($("#repScope").value || "fleet");

  // ---------- Excel export ----------
  const H = t => ({ v: t, s: "head" }), M = n => ({ v: n || 0, s: "money" }), TM = n => ({ v: n || 0, s: "totalmoney" }), B = t => ({ v: t, s: "bold" });
  const catRow = o => CATS.map(c => M(o[c])).concat([TM(o.total)]);
  const catRowT = o => CATS.map(c => TM(o[c])).concat([TM(o.total)]);
  const fileName = s => s.replace(/[\\/:*?"<>|]+/g, "-");
  function exportReport(sc) {
    const r = reportData(sc), cur = DATA.currency, stamp = longDate(today());
    const head = [[{ v: "LBH Vehicles – Fleet costs – " + r.name, s: "title" }], [{ v: "Period: " + periodText() + " · exported " + stamp + " · amounts in " + cur, s: "sub" }], []];
    const sheets = [];
    sheets.push({ name: "Summary", cols: [26, 18, 10], rows: head.concat([[H("Category"), H("Amount (" + cur + ")"), H("Share")]],
      CATS.map(c => [c, M(r.all[c]), r.all.total ? { v: Math.round(1000 * r.all[c] / r.all.total) / 10 + "%" } : ""]),
      [[B("Total"), TM(r.all.total), B("")], [], ["Fuel litres", { v: r.all.litres || 0, s: "dec" }],
        ["Cost lines", { v: r.ls.length, s: "int" }], ["Jobs still to cost (not included)", { v: r.pend.length, s: "int" }]]) });
    if (r.bySite.length) sheets.push({ name: "By site", freeze: 4, cols: [16, 10].concat(CATS.map(() => 13), [15]), rows: head.concat(
      [[H("Site"), H("Vehicles")].concat(CATS.map(H), [H("Total")])], r.bySite.map(x => [x.site, { v: x.n, s: "int" }].concat(catRow(x.o))),
      [[B("Whole fleet"), { v: DATA.fleet.length, s: "totalint" }].concat(catRowT(r.all))]) });
    if (r.byVeh.length) sheets.push({ name: "By vehicle", freeze: 4, filter: 4, cols: [14, 14, 20].concat(CATS.map(() => 13), [15, 12]), rows: head.concat(
      [[H("Site"), H("Matrícula"), H("Make")].concat(CATS.map(H), [H("Total"), H("Fuel litres")])],
      r.byVeh.map(x => [x.v.site, x.v.reg, x.v.make].concat(catRow(x.o), [{ v: x.o.litres || 0, s: "dec" }])),
      [[B("Total"), B(""), B("")].concat(catRowT(r.all), [{ v: r.all.litres || 0, s: "totalint" }])]) });
    sheets.push({ name: "By month", freeze: 4, cols: [14].concat(CATS.map(() => 13), [15]), rows: head.concat(
      [[H("Month")].concat(CATS.map(H), [H("Total")])], r.byMonth.map(x => [monthName(x.m)].concat(catRow(x.o))), [[B("Total")].concat(catRowT(r.all))]) });
    sheets.push({ name: "All costs", freeze: 4, filter: 4, cols: [13, 12, 13, 18, 12, 14, 36, 24, 14, 12, 12, 12, 10, 10, 14, 18, 40], rows: head.concat(
      [["Date", "Site", "Matrícula", "Make", "Category", "Type", "Description", "Supplier / workshop", "Invoice no.", "Parts", "Labour", "Other", "Litres", "Km", "Amount (" + cur + ")", "Entered / costed by", "Invoice link"].map(H)],
      r.ls.slice().sort((a, b) => cmp(a.date, b.date)).map(l => [{ v: l.date, s: "date" }, l.site, l.reg, l.make, l.cat, l.detail, l.desc, l.supplier, l.invNo,
        l.parts != null ? M(l.parts) : "", l.labour != null ? M(l.labour) : "", l.other != null ? M(l.other) : "", l.litres ? { v: l.litres, s: "dec" } : "",
        l.km ? { v: l.km, s: "int" } : "", M(l.amount), l.by, l.invoice && l.invoice !== "#saved" ? l.invoice : ""]),
      [[B("Total")].concat(Array(13).fill(B("")), [TM(r.all.total), B(""), B("")])]) });
    sheets.push({ name: "Awaiting costs", freeze: 4, filter: 4, cols: [13, 12, 13, 18, 16, 24, 44, 18, 24], rows: head.concat(
      [["Date", "Site", "Matrícula", "Make", "Type", "Workshop", "Work done", "Entered by", "Source"].map(H)],
      r.pend.map(j => [{ v: j.date, s: "date" }, j.site, j.reg, j.make, j.type, j.workshop, { v: j.work || j.parts, s: "wrap" }, j.enteredBy, j.source])) });
    XLSX_LITE.download({ sheets }, fileName(`Fleet costs – ${r.name} – ${periodText()}.xlsx`));
  }
  $("#needExport").onclick = () => {
    const site = $("#needSite").value, ps = pending(j => !site || j.site === site);
    const head = [[{ v: "LBH Vehicles – Jobs waiting for costs" + (site ? " – " + site : ""), s: "title" }], [{ v: "Exported " + longDate(today()), s: "sub" }], []];
    XLSX_LITE.download({ sheets: [{ name: "Waiting for costs", freeze: 4, filter: 4, cols: [13, 12, 13, 18, 16, 24, 44, 18, 24, 14], rows: head.concat(
      [["Date", "Site", "Matrícula", "Make", "Type", "Workshop", "Work done", "Entered by", "Source", "Km"].map(H)],
      ps.map(j => [{ v: j.date, s: "date" }, j.site, j.reg, j.make, j.type, j.workshop, { v: j.work || j.parts, s: "wrap" }, j.enteredBy, j.source, j.km ? { v: j.km, s: "int" } : ""])) }] },
      fileName(`Jobs waiting for costs${site ? " – " + site : ""} – ${longDate(today())}.xlsx`));
  };

  // ---------- helpers ----------
  async function post(body) {
    const res = await fetch(CFG.ENDPOINT, { method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" }, redirect: "follow",
      body: JSON.stringify(Object.assign({ key: CFG.APP_KEY, akey: store.get("akey", "") }, body)) });
    const j = await res.json();
    if (!j.ok) throw new Error(j.error || "server error");
    return j;
  }
  function fileData(f) {
    if (!f) return Promise.resolve(null);
    if (f.size > 15e6) { toast("That file is over 15 MB – please use a smaller scan or photo."); return Promise.resolve(false); }
    return new Promise(res => { const r = new FileReader(); r.onload = () => res({ name: f.name, mime: f.type, data: String(r.result).split(",")[1] }); r.readAsDataURL(f); });
  }

  [state.from, state.to] = quickRange(state.quick);
  newRunning();
  load();
  window.__acc = { exportReport, reportData };
})();
