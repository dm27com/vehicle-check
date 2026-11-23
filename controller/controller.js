// Fleet Controller page (PC): records work done, parts and costs into the Vehicle Checks Sheet,
// closes defects, and shows the fleet's service status. Talks to the same Apps Script web app as the phones.
(function () {
  "use strict";
  const CFG = window.APP_CONFIG;
  const $ = (s, el = document) => el.querySelector(s), $$ = (s, el = document) => [...el.querySelectorAll(s)];
  const store = {
    get: (k, d) => { try { const v = localStorage.getItem("ctl." + k); return v === null ? d : JSON.parse(v); } catch (e) { return d; } },
    set: (k, v) => { try { localStorage.setItem("ctl." + k, JSON.stringify(v)); } catch (e) {} },
    del: k => { try { localStorage.removeItem("ctl." + k); } catch (e) {} }
  };
  const esc = s => String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const fmt = n => n == null || n === "" || isNaN(n) ? "" : Number(n).toLocaleString("en-GB", { maximumFractionDigits: 2 });
  const today = () => { const d = new Date(); return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0"); };
  let DATA = null;
  const form = $("#workForm");

  function toast(msg, ms = 3500) { const t = $("#toast"); t.textContent = msg; t.classList.remove("hidden"); clearTimeout(t._h); t._h = setTimeout(() => t.classList.add("hidden"), ms); }

  // ---------- date boxes shown as "21 Nov 2026" ----------
  // The browser's own date box always uses the PC's regional format (21/11/2026), so the
  // date is shown as text over it; the calendar button and picking a date work as normal.
  const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const longDate = v => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(v || ""); return m ? m[3] + " " + MONTHS[+m[2] - 1] + " " + m[1] : ""; };
  const desc = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value");
  $$('input[type="date"]').forEach(inp => {
    const wrap = document.createElement("span"); wrap.className = "datebox";
    inp.parentNode.insertBefore(wrap, inp); wrap.appendChild(inp);
    const shown = document.createElement("span"); shown.className = "datetext"; wrap.appendChild(shown);
    const show = () => { shown.textContent = longDate(desc.get.call(inp)); wrap.classList.toggle("has", !!shown.textContent); };
    // also catch dates set by the page itself (new job, draft restored, form cleared)
    Object.defineProperty(inp, "value", { get() { return desc.get.call(this); }, set(v) { desc.set.call(this, v); show(); } });
    ["input", "change"].forEach(ev => inp.addEventListener(ev, show));
    inp.form && inp.form.addEventListener("reset", () => setTimeout(show));
    show();
  });

  // ---------- loading ----------
  async function load() {
    const ckey = store.get("ckey", "");
    if (!ckey) return showGate();
    $("#loading").classList.remove("hidden");
    try {
      const res = await fetch(CFG.ENDPOINT + "?action=fleet&ckey=" + encodeURIComponent(ckey), { redirect: "follow" });
      const j = await res.json();
      if (!j.ok) { if (/key/.test(j.error)) { store.del("ckey"); return showGate("That key wasn't accepted. Check the Settings tab."); } throw new Error(j.error); }
      DATA = j;
      $("#loading").classList.add("hidden"); $("#gate").classList.add("hidden");
      $("#tabs").classList.remove("hidden"); $("#reload").classList.remove("hidden");
      fillLists(); renderFleet(); renderRecent(); vehicleChanged();
      showTab(location.hash.replace("#", "") || "work");
    } catch (e) {
      $("#loading").innerHTML = '<span class="err">Could not reach the Sheet: ' + esc(e.message) + '</span> <button class="btn sec" onclick="location.reload()">Try again</button>';
    }
  }
  function showGate(msg) {
    $("#loading").classList.add("hidden"); $("#gate").classList.remove("hidden"); $("#gateMsg").textContent = msg || "";
  }
  $("#gateForm").addEventListener("submit", e => { e.preventDefault(); store.set("ckey", $("#ckey").value.trim()); $("#gate").classList.add("hidden"); load(); });
  $("#reload").addEventListener("click", () => load());

  // ---------- tabs ----------
  function showTab(t) {
    if (!["work", "fleet", "recent"].includes(t)) t = "work";
    $$("#tabs button").forEach(b => b.classList.toggle("on", b.dataset.tab === t));
    ["work", "fleet", "recent"].forEach(x => $("#tab-" + x).classList.toggle("hidden", x !== t));
    history.replaceState(null, "", "#" + t);
  }
  $$("#tabs button").forEach(b => b.addEventListener("click", () => showTab(b.dataset.tab)));

  // ---------- lists ----------
  function fillLists() {
    const sites = [...new Set(DATA.fleet.map(v => v.site))].sort();
    const sel = form.reg, cur = sel.value || store.get("draft", {}).reg || "";
    sel.innerHTML = '<option value="">– choose vehicle –</option>' + sites.map(s =>
      `<optgroup label="${esc(s)}">` + DATA.fleet.filter(v => v.site === s).sort((a, b) => a.reg.localeCompare(b.reg))
        .map(v => `<option value="${esc(v.reg)}">${esc(v.reg)} – ${esc(v.make)}${v.active ? "" : " (inactive)"}</option>`).join("") + "</optgroup>").join("");
    sel.value = cur;
    form.type.innerHTML = '<option value="">– choose –</option>' + DATA.types.map(t => `<option>${esc(t)}</option>`).join("");
    const parts = new Set((window.SERVICE_PARTS || []).map(p => p.en).concat(Object.keys(DATA.parts || {})));
    $("#partNames").innerHTML = [...parts].sort().map(p => `<option value="${esc(p)}">`).join("");
    $("#workshops").innerHTML = (DATA.workshops || []).map(w => `<option value="${esc(w)}">`).join("");
    $("#siteFilter").innerHTML = '<option value="">All sites</option>' + sites.map(s => `<option>${esc(s)}</option>`).join("");
    $$(".cur").forEach(el => el.textContent = DATA.currency);
    restoreDraft();
  }

  // ---------- fleet status ----------
  function renderFleet() {
    const site = $("#siteFilter").value, act = $("#onlyActive").checked;
    const rows = DATA.fleet.filter(v => (!site || v.site === site) && (!act || v.active));
    // grouped by site, then vehicle type (make/model), then registration
    const cmp = (x, y) => String(x || "").localeCompare(String(y || ""), "en", { sensitivity: "base", numeric: true });
    rows.sort((a, b) => cmp(a.site, b.site) || cmp(a.make, b.make) || cmp(a.reg, b.reg));
    const n = s => rows.filter(v => v.status === s).length;
    $("#summary").innerHTML = [[rows.length, "vehicles"], [n("OVERDUE"), "service overdue"], [n("Due soon"), "service due soon"],
      [rows.reduce((t, v) => t + v.openDefects, 0), "open defects"]].map(x => `<div><b>${x[0]}</b><span>${x[1]}</span></div>`).join("");
    $("#fleetTable").innerHTML = `<thead><tr><th>Site</th><th>Matrícula</th><th>Make</th><th>Service</th><th class="num">Latest km</th><th class="num">Next service km</th><th class="num">Km to go</th><th>Date limit</th><th>Last service</th><th>Last check</th><th class="num">Open defects</th></tr></thead><tbody>` +
      rows.map((v, i) => `<tr data-reg="${esc(v.reg)}"${i && rows[i - 1].site !== v.site ? ' class="site-start"' : ""}><td>${esc(v.site)}</td><td><b>${esc(v.reg)}</b></td><td>${esc(v.make)}</td>
        <td><span class="pill p-${esc(v.status.split(" ")[0])}">${esc(v.status)}</span></td><td class="num">${fmt(v.latestKm)}</td>
        <td class="num">${fmt(v.nextKm)}</td><td class="num">${fmt(v.kmTo)}</td><td>${esc(v.nextDate)}</td><td>${esc(v.lastService)}</td>
        <td>${esc(v.lastCheck) || '<span class="muted">never</span>'}</td><td class="num">${v.openDefects || ""}</td></tr>`).join("") + "</tbody>";
    $$("#fleetTable tbody tr").forEach(tr => tr.addEventListener("click", () => { form.reg.value = tr.dataset.reg; vehicleChanged(); saveDraft(); showTab("work"); scrollTo(0, 0); }));
  }
  $("#siteFilter").addEventListener("change", renderFleet);
  $("#onlyActive").addEventListener("change", renderFleet);

  function renderRecent() {
    $("#recentTable").innerHTML = `<thead><tr><th>Date</th><th>Matrícula</th><th>Type</th><th>Workshop</th><th>Work</th><th class="num">Total (${esc(DATA.currency)})</th></tr></thead><tbody>` +
      (DATA.recent.length ? DATA.recent.map(r => `<tr><td>${esc(r.date)}</td><td><b>${esc(r.reg)}</b></td><td>${esc(r.type)}</td><td>${esc(r.workshop)}</td><td>${esc(r.work)}</td><td class="num">${fmt(r.total)}</td></tr>`).join("")
        : '<tr><td colspan="6" class="muted">Nothing recorded yet.</td></tr>') + "</tbody>";
  }

  // ---------- vehicle info + its open defects ----------
  function vehicleChanged() {
    const v = DATA && DATA.fleet.find(x => x.reg === form.reg.value);
    $("#vehInfo").innerHTML = v ? [esc(v.site), esc(v.make), `<span class="pill p-${esc(v.status.split(" ")[0])}">${esc(v.status)}</span>`,
      v.latestKm ? "latest " + fmt(v.latestKm) + " km" : "", v.nextKm ? "next service " + fmt(v.nextKm) + " km" + (v.kmTo != null ? " (" + fmt(v.kmTo) + " to go)" : "") : "",
      v.nextDate ? "date limit " + esc(v.nextDate) : "", v.interval ? "every " + fmt(v.interval) + " km" : ""].filter(Boolean).join(" · ") : "";
    const ds = DATA ? DATA.defects.filter(d => d.reg === form.reg.value) : [];
    const saved = store.get("draft", {}).close || {};
    $("#defectList").innerHTML = !v ? "Choose a vehicle." : !ds.length ? '<span class="muted">No open defects on this vehicle.</span>' :
      ds.map(d => { const k = d.sid + "|" + d.item, s = saved[k] || {};
        return `<div class="defect" data-k="${esc(k)}" data-r="${d.r}" data-sid="${esc(d.sid)}" data-item="${esc(d.item)}">
          <div class="what"><b>${esc(d.item)}</b>: ${esc(d.problem)}<br><small>${esc(d.date)} · ${esc(d.by)} · now ${esc(d.status)}</small></div>
          <select><option value="">Leave as is</option><option ${s.status === "In progress" ? "selected" : ""}>In progress</option><option ${s.status === "Closed" ? "selected" : ""}>Closed</option></select>
          <input placeholder="Action taken" value="${esc(s.action || "")}"></div>`; }).join("");
    $$("#defectList select, #defectList input").forEach(el => el.addEventListener("input", saveDraft));
  }
  form.reg.addEventListener("change", vehicleChanged);

  // ---------- parts ----------
  function addPart(p = {}) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td><input class="p-part" list="partNames" value="${esc(p.part || "")}"></td><td><input class="p-no" value="${esc(p.partNo || "")}"></td>
      <td><input class="p-qty" type="number" min="0" step="any" value="${esc(p.qty || 1)}"></td><td><input class="p-unit" type="number" min="0" step="0.01" value="${esc(p.unit || "")}"></td>
      <td class="num p-line">0</td><td><button type="button" class="btn x" title="Remove">×</button></td>`;
    $("#partRows").appendChild(tr);
    $(".x", tr).addEventListener("click", () => { tr.remove(); totals(); saveDraft(); });
    $(".p-part", tr).addEventListener("change", e => {   // remember last part no. / price for this part
      const known = (DATA && DATA.parts || {})[e.target.value];
      if (known) { if (!$(".p-no", tr).value) $(".p-no", tr).value = known.partNo || ""; if (!$(".p-unit", tr).value) $(".p-unit", tr).value = known.unit || ""; }
      totals();
    });
    totals();
  }
  $("#addPart").addEventListener("click", () => { addPart(); $("#partRows tr:last-child .p-part").focus(); });
  const partsOf = () => $$("#partRows tr").map(tr => ({ part: $(".p-part", tr).value.trim(), partNo: $(".p-no", tr).value.trim(),
    qty: Number($(".p-qty", tr).value) || 1, unit: Number($(".p-unit", tr).value) || 0 })).filter(p => p.part);
  function totals() {
    let parts = 0;
    $$("#partRows tr").forEach(tr => { const l = (Number($(".p-qty", tr).value) || 0) * (Number($(".p-unit", tr).value) || 0); $(".p-line", tr).textContent = fmt(l); parts += l; });
    $("#partsTotal").textContent = fmt(parts) + " " + (DATA ? DATA.currency : "");
    const total = parts + (Number(form.labour.value) || 0) + (Number(form.other.value) || 0);
    $("#grandTotal").textContent = fmt(total) + " " + (DATA ? DATA.currency : "");
    return total;
  }
  form.addEventListener("input", () => { totals(); saveDraft(); });

  // ---------- draft (so nothing is lost if the page is closed) ----------
  function saveDraft() {
    const f = new FormData(form), d = {};
    ["reg", "date", "km", "type", "workshop", "invoiceNo", "work", "labour", "other", "followups", "notes"].forEach(k => d[k] = f.get(k) || "");
    d.parts = $$("#partRows tr").map(tr => ({ part: $(".p-part", tr).value, partNo: $(".p-no", tr).value, qty: $(".p-qty", tr).value, unit: $(".p-unit", tr).value }));
    d.close = {};
    $$("#defectList .defect").forEach(el => { const s = $("select", el).value, a = $("input", el).value; if (s || a) d.close[el.dataset.k] = { status: s, action: a }; });
    store.set("draft", d);
    if (form.enteredBy.value) store.set("me", form.enteredBy.value);
  }
  function restoreDraft() {
    const d = store.get("draft", null);
    form.date.value = (d && d.date) || today();
    form.enteredBy.value = store.get("me", "");
    $("#partRows").innerHTML = "";
    if (d) {
      ["reg", "km", "type", "workshop", "invoiceNo", "work", "labour", "other", "followups", "notes"].forEach(k => { if (form[k] && d[k] != null) form[k].value = d[k]; });
      (d.parts || []).forEach(p => addPart(p));
    }
    if (!$$("#partRows tr").length) addPart();
    totals();
  }
  $("#clearForm").addEventListener("click", () => {
    if (!confirm("Clear the whole form?")) return;
    store.del("draft"); form.reset(); form.invoice.value = ""; restoreDraft(); vehicleChanged(); $("#formMsg").textContent = "";
  });

  // ---------- submit ----------
  const readFile = file => new Promise((ok, bad) => { const r = new FileReader(); r.onload = () => ok(String(r.result).split(",")[1]); r.onerror = bad; r.readAsDataURL(file); });
  form.addEventListener("submit", async e => {
    e.preventDefault();
    $$(".invalid", form).forEach(el => el.classList.remove("invalid"));
    const errs = [];
    ["reg", "date", "type", "work", "enteredBy"].forEach(k => { if (!String(form[k].value).trim()) { form[k].classList.add("invalid"); errs.push(k); } });
    const v = DATA.fleet.find(x => x.reg === form.reg.value);
    const km = Number(form.km.value);
    if (km && v && v.latestKm && km < v.latestKm - 50 && !confirm(`The km (${fmt(km)}) is lower than the latest known for this vehicle (${fmt(v.latestKm)}). Save anyway?`)) return;
    const file = form.invoice.files[0];
    if (file && file.size > 15 * 1024 * 1024) { errs.push("invoice"); form.invoice.classList.add("invalid"); }
    if (errs.length) { $("#formMsg").textContent = "Please complete the highlighted fields" + (errs.includes("invoice") ? " (invoice must be under 15 MB)" : "") + "."; return; }
    $("#formMsg").textContent = "";
    const btn = $("#submitBtn"); btn.disabled = true; btn.textContent = "Saving…";
    try {
      const close = $$("#defectList .defect").map(el => ({ r: el.dataset.r, sid: el.dataset.sid, item: el.dataset.item, status: $("select", el).value, action: $("input", el).value.trim() }))
        .filter(c => c.status);
      const payload = {
        key: CFG.APP_KEY, kind: "work", ckey: store.get("ckey", ""), id: "W-" + Date.now().toString(36) + "-" + Math.random().toString(36).slice(2, 6),
        row: { "Date": form.date.value, "Matrícula": form.reg.value, "Type": form.type.value, "Km": form.km.value, "Workshop": form.workshop.value.trim(),
          "Work done": form.work.value.trim(), "Invoice no.": form.invoiceNo.value.trim(), "Labour cost": form.labour.value, "Other cost": form.other.value,
          "Entered by": form.enteredBy.value.trim(), "Notes": form.notes.value.trim() },
        parts: partsOf(), close,
        followups: form.followups.value.split("\n").map(s => s.trim()).filter(Boolean),
        invoice: file ? { name: file.name, mime: file.type, data: await readFile(file) } : null
      };
      const res = await fetch(CFG.ENDPOINT, { method: "POST", body: JSON.stringify(payload), headers: { "Content-Type": "text/plain;charset=utf-8" }, redirect: "follow" });
      const j = await res.json();
      if (!j.ok) throw new Error(j.error || "server error");
      store.del("draft"); const me = form.enteredBy.value;
      form.reset(); form.enteredBy.value = me; form.invoice.value = "";
      toast(`Saved – ${payload.row["Matrícula"]}, total ${fmt(j.total)} ${DATA.currency}`, 5000);
      await load();
    } catch (err) {
      $("#formMsg").textContent = "Not saved: " + err.message + ". Your entries are kept – try again.";
    } finally { btn.disabled = false; btn.textContent = "Save work record"; }
  });

  if (!CFG || !CFG.ENDPOINT) { $("#loading").textContent = "config.js is missing the database address."; return; }
  load();
})();
