/* ตรวจทรัพย์สิน — PWA ที่อ่านข้อมูลจาก OneDrive ผ่าน Microsoft Graph
   - ข้อมูลหลัก: AssetDB/AssetDB.xlsx (อ่านอย่างเดียว แก้ข้อมูลทรัพย์สินใน Excel)
   - ผลตรวจ:    AssetDB/audit/<ปีงบ>.json (แอปเขียน)
   - รูป:       AssetDB/images/<รหัสสินทรัพย์>.jpg
   - สรุป:      AssetDB/reports/ผลตรวจทรัพย์สิน_<ปีงบ>.xlsx (ตอนกดส่งออก) */
(() => {
  "use strict";
  const CFG = window.APP_CONFIG;
  const SCOPES = ["User.Read", "Files.ReadWrite"];
  const GRAPH = "https://graph.microsoft.com/v1.0";

  const $ = (s) => document.querySelector(s);
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const store = {
    get(k, d) { try { const v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch { return d; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { console.warn(e); } }
  };

  const state = {
    assets: [], history: [], thumbs: {}, records: {},
    year: null, filter: "all", q: "", account: null
  };

  const fiscalYear = (d = new Date()) => (d.getMonth() >= 9 ? d.getFullYear() + 1 : d.getFullYear()) + 543;
  const tone = (label) => !label ? "pending" : (CFG.statuses.find((s) => s.label === label)?.tone || "warn");
  const fmtDate = (iso) => { const d = new Date(iso); return isNaN(d) ? "" : d.toLocaleString("th-TH", { dateStyle: "medium", timeStyle: "short" }); };
  const findAsset = (id) => state.assets.find((a) => a.AssetID === String(id).trim());

  /* ---------- toast ---------- */
  let toastTimer;
  function toast(text, ms = 2600) {
    const t = $("#toast"); t.textContent = text; t.classList.add("show");
    clearTimeout(toastTimer); toastTimer = setTimeout(() => t.classList.remove("show"), ms);
  }

  /* ---------- screens ---------- */
  function showMessage(title, text, btnLabel, onClick) {
    $("#screen-list").hidden = true; $("#scanBtn").hidden = true;
    $("#screen-msg").hidden = false;
    $("#msgTitle").textContent = title; $("#msgText").textContent = text || "";
    const b = $("#msgBtn"); b.hidden = !btnLabel; b.textContent = btnLabel || ""; b.onclick = onClick || null;
  }
  function showList() {
    $("#screen-msg").hidden = true; $("#screen-list").hidden = false;
    $("#scanBtn").hidden = false; $("#topActions").hidden = false;
  }

  /* ---------- auth (MSAL) ---------- */
  let pca;
  async function initAuth() {
    pca = new msal.PublicClientApplication({
      auth: { clientId: CFG.clientId, authority: CFG.authority, redirectUri: location.origin + location.pathname },
      cache: { cacheLocation: "localStorage" }
    });
    const res = await pca.handleRedirectPromise();
    const acc = res?.account || pca.getAllAccounts()[0] || null;
    if (acc) pca.setActiveAccount(acc);
    return acc;
  }
  async function getToken() {
    const account = pca.getActiveAccount();
    try {
      return (await pca.acquireTokenSilent({ scopes: SCOPES, account })).accessToken;
    } catch (e) {
      if (e instanceof msal.InteractionRequiredAuthError) await pca.acquireTokenRedirect({ scopes: SCOPES });
      throw e;
    }
  }

  /* ---------- Graph helpers ---------- */
  const encPath = (p) => p.split("/").map(encodeURIComponent).join("/");
  const item = (rel) => `/me/drive/root:/${encPath(CFG.folder + "/" + rel)}:`;
  async function g(path, opts = {}) {
    const token = await getToken();
    return fetch(path.startsWith("http") ? path : GRAPH + path, {
      ...opts, headers: { Authorization: "Bearer " + token, ...(opts.headers || {}) }
    });
  }
  async function gjson(path, opts) {
    const r = await g(path, opts);
    if (!r.ok) throw new Error(`Graph ${r.status}`);
    return r.json();
  }

  /* ---------- data loading ---------- */
  async function loadWorkbook() {
    const r = await g(item(CFG.workbook) + "/content");
    if (r.status === 404) throw new Error(`ไม่พบไฟล์ ${CFG.folder}/${CFG.workbook} ใน OneDrive\nอัปโหลดโฟลเดอร์ ${CFG.folder} ไว้ที่ root ของ OneDrive แล้วลองอีกครั้ง`);
    if (!r.ok) throw new Error(`โหลดไฟล์ Excel ไม่สำเร็จ (รหัส ${r.status})`);
    const wb = XLSX.read(await r.arrayBuffer(), { type: "array" });
    const rows = (name) => wb.Sheets[name] ? XLSX.utils.sheet_to_json(wb.Sheets[name], { raw: true, defval: "" }) : [];
    const assets = rows("Assets");
    if (!assets.length) throw new Error("ไม่พบชีท Assets หรือชีทว่าง");
    state.assets = assets
      .filter((a) => a.AssetID !== "")
      .map((a) => ({ ...a, AssetID: String(a.AssetID).trim() }));
    state.history = rows("AuditLog").map((h) => ({ ...h, AssetID: String(h.AssetID).trim() }));
  }

  async function loadThumbs() {
    let url = item("images") + "/children?$select=id,name&$expand=thumbnails&$top=200";
    const map = {};
    while (url) {
      const j = await gjson(url);
      for (const f of j.value) {
        const t = f.thumbnails?.[0];
        if (t) map[f.name.toLowerCase()] = { s: t.medium?.url || t.small?.url, l: t.large?.url || t.medium?.url };
      }
      url = j["@odata.nextLink"];
    }
    state.thumbs = map;
  }
  function photoOf(a) {
    const byId = state.thumbs[(a.AssetID + ".jpg").toLowerCase()];
    if (byId) return byId;
    const f = String(a["รูปภาพ"] || "").split("/").pop().toLowerCase();
    return f ? state.thumbs[f] : null;
  }

  async function readAudit(year) {
    const r = await g(item(`audit/${year}.json`));
    if (r.status === 404) return { etag: null, data: {} };
    if (!r.ok) throw new Error(`อ่านผลตรวจไม่สำเร็จ (รหัส ${r.status})`);
    const meta = await r.json();
    const c = await fetch(meta["@microsoft.graph.downloadUrl"], { cache: "no-store" });
    return { etag: meta.eTag, data: c.ok ? await c.json() : {} };
  }

  // ผลตรวจที่มากับ Excel (ชีท AuditLog) เป็นค่าตั้งต้นของปีนั้น
  function baseRecords(year) {
    const out = {};
    for (const h of state.history) {
      if (String(h["ปีงบ"]) !== String(year) || !h["ผลตรวจ"]) continue;
      out[h.AssetID] = { status: h["ผลตรวจ"], note: h["หมายเหตุ"] || "", by: h["ผู้ตรวจ"] || "", at: "", from: "excel" };
    }
    return out;
  }
  const pending = () => store.get("pending", []);
  function pendingMap(year) {
    const m = {};
    for (const p of pending()) if (String(p.year) === String(year)) m[p.id] = { ...p.rec, pending: true };
    return m;
  }
  function composeRecords(remote) {
    state.records = { ...baseRecords(state.year), ...remote, ...pendingMap(state.year) };
  }
  async function loadYear() {
    try {
      const { data } = await readAudit(state.year);
      store.set("rec:" + state.year, data);
      composeRecords(data);
    } catch (e) {
      composeRecords(store.get("rec:" + state.year, {}));
      throw e;
    }
  }

  /* ---------- saving (offline queue + eTag) ---------- */
  function enqueue(year, id, rec) {
    const q = pending().filter((p) => !(String(p.year) === String(year) && p.id === id));
    q.push({ year, id, rec });
    store.set("pending", q);
  }
  let flushing = false;
  async function flush() {
    if (flushing || !navigator.onLine) return updateSync();
    const q = pending();
    if (!q.length) return updateSync();
    flushing = true; updateSync();
    try {
      for (const y of [...new Set(q.map((p) => String(p.year)))]) {
        const batch = q.filter((p) => String(p.year) === y);
        let saved = false;
        for (let attempt = 0; attempt < 4 && !saved; attempt++) {
          const { etag, data } = await readAudit(y);
          for (const p of batch) {
            const cur = data[p.id];
            if (!cur || (cur.at || "") <= p.rec.at) data[p.id] = p.rec;
          }
          const r = await g(item(`audit/${y}.json`) + "/content", {
            method: "PUT",
            headers: { "Content-Type": "application/json", ...(etag ? { "If-Match": etag } : {}) },
            body: JSON.stringify(data, null, 1)
          });
          if (r.ok) {
            saved = true;
            store.set("pending", pending().filter((x) => !batch.some((b) => String(b.year) === String(x.year) && b.id === x.id && b.rec.at === x.rec.at)));
            store.set("rec:" + y, data);
            if (y === String(state.year)) composeRecords(data);
          } else if (r.status === 412 || r.status === 409) {
            await sleep(400 * (attempt + 1)); // มีคนบันทึกพร้อมกัน — อ่านใหม่แล้วรวม
          } else {
            throw new Error(`บันทึกขึ้น OneDrive ไม่สำเร็จ (รหัส ${r.status})`);
          }
        }
      }
    } catch (e) {
      console.warn(e);
    } finally {
      flushing = false; updateSync(); render();
    }
  }
  function updateSync() {
    const n = pending().length;
    $("#syncNote").textContent = !n ? "" :
      flushing ? `กำลังส่ง ${n} รายการขึ้น OneDrive…` :
      navigator.onLine ? `รอส่งขึ้น OneDrive ${n} รายการ` : `ออฟไลน์: เก็บไว้ในเครื่อง ${n} รายการ จะส่งเมื่อต่อเน็ต`;
  }

  /* ---------- list ---------- */
  function filtered() {
    const q = state.q.trim().toLowerCase();
    return state.assets.filter((a) => {
      const rec = state.records[a.AssetID];
      const t = tone(rec?.status);
      if (state.filter === "todo" && rec) return false;
      if (state.filter === "done" && !rec) return false;
      if (state.filter === "issue" && t !== "warn" && t !== "bad") return false;
      if (!q) return true;
      return [a.AssetID, a["ชื่อสินทรัพย์"], a["สถานที่"], a["อาคาร"], a["ผู้ครอบครอง"]].join(" ").toLowerCase().includes(q);
    });
  }
  function render() {
    if (!state.assets.length) return;
    const done = state.assets.filter((a) => state.records[a.AssetID]).length;
    $("#doneCount").textContent = done;
    $("#totalCount").textContent = state.assets.length;
    $("#bar").style.width = (100 * done / state.assets.length) + "%";
    const items = filtered();
    $("#empty").hidden = items.length > 0;
    $("#list").innerHTML = items.map((a) => {
      const rec = state.records[a.AssetID];
      const t = tone(rec?.status);
      const ph = photoOf(a);
      const loc = [a["สถานที่"], a["อาคาร"]].filter(Boolean).join(", ");
      return `<li><button class="row ${t}" data-id="${esc(a.AssetID)}">
        ${ph?.s ? `<img class="thumb" loading="lazy" alt="" src="${esc(ph.s)}">` : `<span class="thumb"></span>`}
        <span class="row-main">
          <span class="rid">${esc(a.AssetID)}</span>
          <span class="rname">${esc(a["ชื่อสินทรัพย์"])}</span>
          <span class="rloc">${esc(loc)}</span>
        </span>
        <span class="rstat">${esc(rec?.status || "ยังไม่ตรวจ")}${rec?.pending ? "<br>รอส่ง" : ""}</span>
      </button></li>`;
    }).join("");
  }

  /* ---------- sheets + back button ---------- */
  const stack = [];
  function openSheet(el, replace) {
    if (replace && stack.length) { closeTop(true); history.replaceState({ sheet: el.id }, ""); }
    else history.pushState({ sheet: el.id }, "");
    el.hidden = false; document.body.style.overflow = "hidden";
    stack.push(el);
  }
  function closeTop(silent) {
    const el = stack.pop(); if (!el) return;
    el.hidden = true;
    if (el.id === "scanner") stopScanner();
    if (!stack.length) document.body.style.overflow = "";
    if (!silent) render();
  }
  window.addEventListener("popstate", () => closeTop());
  const back = () => history.back();

  /* ---------- detail ---------- */
  let draft = null;
  function openDetail(id, fromScan, replace) {
    const a = findAsset(id); if (!a) return toast(`ไม่พบรหัส ${id} ในฐานข้อมูล`);
    const rec = state.records[a.AssetID];
    draft = { id: a.AssetID, status: rec?.status || "", photo: null, fromScan };
    const ph = photoOf(a);
    const lat = parseFloat(a.Lat), lng = parseFloat(a.Lng);
    const fact = (k, v) => v === "" || v == null ? "" : `<dt>${k}</dt><dd>${v}</dd>`;
    const el = $("#detail");
    el.innerHTML = `
      <div class="sheet-bar"><button class="back" data-close>‹ กลับ</button><span class="rid">${esc(a.AssetID)}</span></div>
      <div class="sheet-body">
        <figure class="photo" id="photoBox">${ph?.l ? `<img alt="รูปทรัพย์สิน" src="${esc(ph.l)}">` : "ยังไม่มีรูป"}</figure>
        <h2>${esc(a["ชื่อสินทรัพย์"])}</h2>
        <dl class="facts">
          ${fact("ประเภท", esc([a["รหัสประเภท"], a["ชื่อประเภท"]].filter(Boolean).join(" ")))}
          ${fact("จำนวน", esc(a["จำนวน"]))}
          ${fact("ผู้ครอบครอง", esc(a["ผู้ครอบครอง"]))}
          ${fact("สถานที่", esc(a["สถานที่"]))}
          ${fact("อาคาร", esc(a["อาคาร"]))}
          ${fact("สถานะ", esc(a["สถานะ"]))}
          ${fact("หมายเหตุ", esc(a["หมายเหตุ"]))}
          ${isFinite(lat) && isFinite(lng) ? fact("ตำแหน่ง", `<a href="https://www.google.com/maps?q=${lat},${lng}" target="_blank" rel="noopener">เปิดในแผนที่</a>`) : ""}
        </dl>
        <section class="audit">
          <h3>ผลตรวจปีงบ ${esc(state.year)}</h3>
          <p class="last">${rec ? `ล่าสุด: ${esc(rec.status)}${rec.by ? " โดย " + esc(rec.by) : ""}${rec.at ? " " + esc(fmtDate(rec.at)) : ""}${rec.pending ? " (รอส่ง)" : ""}` : "ยังไม่ได้ตรวจ"}</p>
          <div class="seg" role="radiogroup" aria-label="ผลตรวจ">
            ${CFG.statuses.map((s) => `<button type="button" role="radio" class="${s.tone}" aria-checked="${s.label === draft.status}" data-status="${esc(s.label)}">${esc(s.label)}</button>`).join("")}
          </div>
          <label class="note">หมายเหตุการตรวจ<textarea id="note">${esc(rec?.note || "")}</textarea></label>
          <p class="photo-note" id="photoNote" hidden>จะอัปโหลดรูปใหม่เมื่อกดบันทึก</p>
          <div class="actions">
            <label class="btn ghost">ถ่ายรูปใหม่<input type="file" id="photoInput" accept="image/*" capture="environment" hidden></label>
            ${fromScan ? `<button class="btn ghost" id="saveOnly">บันทึก</button>` : ""}
            <button class="btn primary" id="save">${fromScan ? "บันทึกและสแกนต่อ" : "บันทึกผลตรวจ"}</button>
          </div>
        </section>
      </div>`;
    el.scrollTop = 0;
    openSheet(el, replace);
  }

  $("#detail").addEventListener("click", (e) => {
    if (e.target.closest("[data-close]")) return back();
    const s = e.target.closest("[data-status]");
    if (s) {
      draft.status = s.dataset.status;
      $("#detail").querySelectorAll("[data-status]").forEach((b) => b.setAttribute("aria-checked", String(b === s)));
      return;
    }
    if (e.target.id === "save") return save(draft.fromScan);
    if (e.target.id === "saveOnly") return save(false);
  });
  $("#detail").addEventListener("change", (e) => {
    if (e.target.id !== "photoInput" || !e.target.files[0]) return;
    draft.photo = e.target.files[0];
    $("#photoBox").innerHTML = `<img alt="รูปใหม่" src="${URL.createObjectURL(draft.photo)}">`;
    $("#photoNote").hidden = false;
  });

  function getPosition() {
    return new Promise((res) => {
      if (!navigator.geolocation) return res(null);
      navigator.geolocation.getCurrentPosition(
        (p) => res({ lat: +p.coords.latitude.toFixed(6), lng: +p.coords.longitude.toFixed(6) }),
        () => res(null), { timeout: 5000, maximumAge: 60000 });
    });
  }
  async function resize(file, max = 1600) {
    const bmp = await createImageBitmap(file);
    const k = Math.min(1, max / Math.max(bmp.width, bmp.height));
    const c = document.createElement("canvas");
    c.width = Math.round(bmp.width * k); c.height = Math.round(bmp.height * k);
    c.getContext("2d").drawImage(bmp, 0, 0, c.width, c.height);
    return new Promise((r) => c.toBlob(r, "image/jpeg", 0.85));
  }
  async function uploadPhoto(id, file) {
    const blob = await resize(file);
    const name = `${id}.jpg`;
    const r = await g(item("images/" + name) + "/content", { method: "PUT", headers: { "Content-Type": "image/jpeg" }, body: blob });
    if (!r.ok) throw new Error(`อัปโหลดรูปไม่สำเร็จ (รหัส ${r.status})`);
    const url = URL.createObjectURL(blob);
    state.thumbs[name.toLowerCase()] = { s: url, l: url };
  }

  async function save(scanNext) {
    if (!draft.status) return toast("เลือกผลตรวจก่อนบันทึก");
    const btns = $("#detail").querySelectorAll(".actions button");
    btns.forEach((b) => (b.disabled = true));
    const pos = await getPosition();
    const rec = {
      status: draft.status,
      note: $("#note").value.trim(),
      by: state.account?.name || state.account?.username || "",
      at: new Date().toISOString(),
      ...(pos || {})
    };
    enqueue(state.year, draft.id, rec);
    state.records[draft.id] = { ...rec, pending: true };
    let msg = "บันทึกผลตรวจแล้ว";
    if (draft.photo) {
      if (!navigator.onLine) msg = "บันทึกผลตรวจแล้ว แต่ต้องต่อเน็ตเพื่ออัปโหลดรูป";
      else {
        try { await uploadPhoto(draft.id, draft.photo); msg = "บันทึกผลตรวจและรูปแล้ว"; }
        catch (e) { msg = "บันทึกผลตรวจแล้ว แต่ " + e.message; }
      }
    }
    toast(msg);
    flush();
    if (scanNext) openScanner(true); else back();
  }

  /* ---------- scanner ---------- */
  let qr, lastMiss = { t: "", at: 0 };
  function parseCode(text) {
    let t = String(text).trim();
    try { const u = new URL(t); t = u.searchParams.get("data") || u.hash.slice(1) || t; } catch { /* not a URL */ }
    const m = t.match(/\d{8,}/);
    return m ? m[0] : t;
  }
  async function openScanner(replace) {
    $("#scanErr").textContent = ""; $("#manualId").value = "";
    openSheet($("#scanner"), replace);
    qr = qr || new Html5Qrcode("reader", { verbose: false });
    try {
      await qr.start({ facingMode: "environment" },
        { fps: 10, qrbox: (w, h) => { const s = Math.floor(Math.min(w, h) * 0.7); return { width: s, height: s }; } },
        onScan, () => {});
    } catch (e) {
      $("#scanErr").textContent = "เปิดกล้องไม่ได้ ให้อนุญาตการใช้กล้องสำหรับเว็บนี้ในการตั้งค่าเบราว์เซอร์ หรือพิมพ์รหัสด้านล่าง";
    }
  }
  async function stopScanner() {
    try { if (qr && qr.isScanning) await qr.stop(); } catch { /* ignore */ }
  }
  async function onScan(text) {
    const id = parseCode(text);
    if (!findAsset(id)) {
      if (lastMiss.t !== id || Date.now() - lastMiss.at > 4000) toast(`ไม่พบรหัส ${id} ในฐานข้อมูล`);
      lastMiss = { t: id, at: Date.now() };
      return;
    }
    await stopScanner();
    if (navigator.vibrate) navigator.vibrate(60);
    openDetail(id, true, true);
  }
  $("#scanner").addEventListener("click", (e) => { if (e.target.closest("[data-close]")) back(); });
  $("#manual").addEventListener("submit", async (e) => {
    e.preventDefault();
    const id = $("#manualId").value.trim();
    if (!findAsset(id)) return toast(`ไม่พบรหัส ${id} ในฐานข้อมูล`);
    await stopScanner();
    openDetail(id, true, true);
  });
  $("#scanBtn").addEventListener("click", () => openScanner(false));

  /* ---------- export ---------- */
  async function exportXlsx() {
    const rows = state.assets.map((a) => {
      const r = state.records[a.AssetID] || {};
      return {
        "รหัสสินทรัพย์": a.AssetID, "ชื่อสินทรัพย์": a["ชื่อสินทรัพย์"], "ผู้ครอบครอง": a["ผู้ครอบครอง"],
        "สถานที่": a["สถานที่"], "อาคาร": a["อาคาร"], "ผลตรวจ": r.status || "ยังไม่ตรวจ",
        "หมายเหตุการตรวจ": r.note || "", "ผู้ตรวจ": r.by || "", "วันที่ตรวจ": r.at ? fmtDate(r.at) : ""
      };
    });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "ผลตรวจ " + state.year);
    const name = `ผลตรวจทรัพย์สิน_${state.year}.xlsx`;
    XLSX.writeFile(wb, name);
    try {
      const buf = XLSX.write(wb, { type: "array", bookType: "xlsx" });
      const r = await g(item("reports/" + name) + "/content", {
        method: "PUT", body: buf,
        headers: { "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }
      });
      if (r.ok) toast(`ดาวน์โหลดแล้ว และเก็บสำเนาไว้ที่ ${CFG.folder}/reports`);
    } catch { toast("ดาวน์โหลดแล้ว (เก็บสำเนาใน OneDrive ไม่ได้เพราะออฟไลน์)"); }
  }

  /* ---------- controls ---------- */
  $("#q").addEventListener("input", (e) => { state.q = e.target.value; render(); });
  $("#chips").addEventListener("click", (e) => {
    const b = e.target.closest("[data-f]"); if (!b) return;
    state.filter = b.dataset.f;
    $("#chips").querySelectorAll("button").forEach((x) => x.classList.toggle("on", x === b));
    render();
  });
  $("#list").addEventListener("click", (e) => {
    const b = e.target.closest("[data-id]"); if (b) openDetail(b.dataset.id, false);
  });
  $("#year").addEventListener("change", async (e) => {
    state.year = e.target.value; store.set("year", state.year);
    try { await loadYear(); } catch { toast("ออฟไลน์: แสดงผลตรวจที่เคยโหลดไว้"); }
    render();
  });
  const menu = $("#menu"), menuBtn = $("#menuBtn");
  menuBtn.addEventListener("click", (e) => {
    e.stopPropagation(); menu.hidden = !menu.hidden; menuBtn.setAttribute("aria-expanded", String(!menu.hidden));
  });
  document.addEventListener("click", () => { menu.hidden = true; menuBtn.setAttribute("aria-expanded", "false"); });
  menu.addEventListener("click", async (e) => {
    const act = e.target.dataset.act;
    if (act === "refresh") refresh();
    if (act === "export") exportXlsx();
    if (act === "logout") {
      if (pending().length && !confirm(`ยังมีผลตรวจ ${pending().length} รายการที่ยังไม่ได้ส่งขึ้น OneDrive ออกจากระบบต่อหรือไม่`)) return;
      pca.logoutRedirect({ account: pca.getActiveAccount() });
    }
  });
  window.addEventListener("online", () => { updateSync(); flush(); });
  window.addEventListener("offline", updateSync);

  function fillYears() {
    const years = new Set([String(fiscalYear()), String(state.year)]);
    state.history.forEach((h) => h["ปีงบ"] && years.add(String(h["ปีงบ"])));
    $("#year").innerHTML = [...years].sort().reverse()
      .map((y) => `<option ${y === String(state.year) ? "selected" : ""}>${esc(y)}</option>`).join("");
  }

  async function refresh() {
    document.body.classList.add("busy");
    try {
      await loadWorkbook();
      store.set("cache", { assets: state.assets, history: state.history });
      fillYears();
      await loadYear();
      showList(); render();
      loadThumbs().then(render).catch((e) => console.warn("thumbs", e));
      flush();
    } catch (e) {
      console.warn(e);
      if (state.assets.length) { showList(); render(); toast(navigator.onLine ? e.message : "ออฟไลน์: แสดงข้อมูลล่าสุดที่เคยโหลด", 4000); }
      else showMessage("โหลดข้อมูลไม่สำเร็จ", e.message, "ลองอีกครั้ง", refresh);
    } finally {
      document.body.classList.remove("busy");
    }
  }

  /* ---------- start ---------- */
  async function start() {
    if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(() => {});
    if (!CFG.clientId || !/^[0-9a-f-]{36}$/i.test(CFG.clientId)) {
      return showMessage("ยังไม่ได้ตั้งค่าแอป", "ใส่ Application (client) ID ในไฟล์ config.js ตามขั้นตอนใน README");
    }
    let acc;
    try { acc = await initAuth(); }
    catch (e) { return showMessage("เข้าสู่ระบบไม่สำเร็จ", e.message, "ลองอีกครั้ง", () => location.reload()); }
    if (!acc) {
      return showMessage("เข้าสู่ระบบ", "ใช้บัญชี Microsoft ที่เก็บโฟลเดอร์ " + CFG.folder + " ไว้ใน OneDrive",
        "เข้าสู่ระบบด้วย Microsoft", () => pca.loginRedirect({ scopes: SCOPES }));
    }
    state.account = acc;
    $("#who").textContent = acc.name || acc.username;
    state.year = String(store.get("year", fiscalYear()));
    const cache = store.get("cache", null);
    if (cache?.assets?.length) {
      state.assets = cache.assets; state.history = cache.history || [];
      fillYears(); composeRecords(store.get("rec:" + state.year, {}));
      showList(); render(); updateSync();
    }
    refresh();
  }
  start();
})();
