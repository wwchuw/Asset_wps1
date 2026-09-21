/* ตรวจทรัพย์สิน — PWA ที่ใช้ Google Sheet ผ่าน Apps Script
   - ข้อมูลทรัพย์สิน: Google Sheet เดิม (ชีท Sheet1)
   - ผลตรวจ: คอลัมน์ ตรวจทรัพสินย์งบ_<ปีงบ> + ประวัติในชีท AuditLog
   - รูป: โฟลเดอร์รูปเดิมใน Google Drive */
(() => {
  "use strict";
  const CFG = window.APP_CONFIG;

  const $ = (s) => document.querySelector(s);
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) =>
    ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const store = {
    get(k, d) { try { const v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch { return d; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { console.warn(e); } },
    del(k) { try { localStorage.removeItem(k); } catch { /* ignore */ } }
  };


  const fiscalYear = (d = new Date()) => (d.getMonth() >= 9 ? d.getFullYear() + 1 : d.getFullYear()) + 543;
  const tone = (label) => !label ? "pending" : (CFG.statuses.find((s) => s.label === label)?.tone || "warn");
  const fmtDate = (iso) => { const d = new Date(iso); return !iso || isNaN(d) ? "" : d.toLocaleString("th-TH", { dateStyle: "medium", timeStyle: "short" }); };
  const state = {
    assets: [], server: {}, records: {}, sheetUrl: "",
    year: String(store.get("year", fiscalYear())), building: store.get("building", ""), filter: "all", q: "", session: store.get("session", null)
  };
  // อายุการใช้งาน: เทียบปีงบที่เปลี่ยนได้กับปีงบปัจจุบัน (ไม่ใช่ปีงบที่เลือกดูผลตรวจ)
  const disposed = (a) => a.status === "จำหน่ายคืน";
  function ageInfo(a) {
    if (disposed(a) || !a.replaceFY) return null;
    const now = fiscalYear();
    if (a.replaceFY <= now) return { tone: "bad", short: "ครบอายุ", text: `ครบอายุแล้ว (ตั้งแต่ปีงบ ${a.replaceFY})` };
    if (a.replaceFY === now + 1) return { tone: "warn", short: "ครบปีงบหน้า", text: `ครบอายุปีงบหน้า (${a.replaceFY})` };
    return { tone: "ok", short: "", text: `ปีงบ ${a.replaceFY} (อีก ${a.replaceFY - now} ปี)` };
  }
  const isDue = (a) => { const g = ageInfo(a); return !!g && g.tone !== "ok"; };
  const findAsset = (id) => state.assets.find((a) => a.id === String(id).trim());
  const thumb = (a, size) => a.imageId ? `https://drive.google.com/thumbnail?id=${encodeURIComponent(a.imageId)}&sz=w${size}` : "";

  /* ---------- toast ---------- */
  let toastTimer;
  function toast(text, ms = 2600) {
    const t = $("#toast"); t.textContent = text; t.classList.add("show");
    clearTimeout(toastTimer); toastTimer = setTimeout(() => t.classList.remove("show"), ms);
  }

  /* ---------- screens ---------- */
  function showMessage(title, text, btnLabel, onClick) {
    $("#screen-list").hidden = true; $("#scanBtn").hidden = true; $("#topActions").hidden = true;
    $("#screen-msg").hidden = false; $("#login").hidden = true;
    $("#msgTitle").textContent = title; $("#msgText").textContent = text || "";
    const b = $("#msgBtn"); b.hidden = !btnLabel; b.textContent = btnLabel || ""; b.onclick = onClick || null;
  }
  function showLogin(error) {
    showMessage("เข้าใช้งาน", error || "ใส่รหัสทีมที่ได้รับจากผู้ดูแล และชื่อของคุณ");
    $("#login").hidden = false;
    if (!$("#loginName").value) $("#loginName").value = store.get("lastName", "");
  }
  function showList() {
    $("#screen-msg").hidden = true; $("#screen-list").hidden = false;
    $("#scanBtn").hidden = false; $("#topActions").hidden = false;
  }

  /* ---------- API (Apps Script) ---------- */
  class ApiError extends Error {}
  async function api(action, payload = {}, key = state.session?.key) {
    let r;
    try {
      // text/plain ป้องกัน preflight ที่ Apps Script ไม่รองรับ
      r = await fetch(CFG.apiUrl, {
        method: "POST", redirect: "follow",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({ action, key, ...payload })
      });
    } catch (e) {
      throw new Error("เชื่อมต่อไม่ได้ ตรวจสอบอินเทอร์เน็ต");
    }
    if (!r.ok) throw new Error(`เซิร์ฟเวอร์ตอบกลับผิดปกติ (รหัส ${r.status})`);
    let j;
    try { j = await r.json(); } catch { throw new Error("อ่านคำตอบจากเซิร์ฟเวอร์ไม่ได้ ตรวจสอบ apiUrl และการ Deploy"); }
    if (!j.ok) throw new ApiError(j.error || "เกิดข้อผิดพลาด");
    return j;
  }

  /* ---------- data ---------- */
  const pending = () => store.get("pending", []);
  function pendingMap(year) {
    const m = {};
    for (const p of pending()) if (String(p.year) === String(year)) m[p.id] = { ...p.rec, pending: true };
    return m;
  }
  function composeRecords() {
    state.records = { ...(state.server[state.year] || {}), ...pendingMap(state.year) };
  }
  function fillYears() {
    const years = new Set([String(fiscalYear()), String(state.year), ...Object.keys(state.server)]);
    $("#year").innerHTML = [...years].filter(Boolean).sort().reverse()
      .map((y) => `<option ${y === String(state.year) ? "selected" : ""}>${esc(y)}</option>`).join("");
  }
  async function refresh() {
    document.body.classList.add("busy");
    try {
      const j = await api("load");
      state.assets = j.assets; state.server = j.records || {}; state.sheetUrl = j.sheetUrl || "";
      store.set("cache", { assets: state.assets, server: state.server, sheetUrl: state.sheetUrl });
      fillYears(); fillBuildings(); composeRecords(); showList(); render();
      flush();
    } catch (e) {
      if (e instanceof ApiError && /รหัสทีม/.test(e.message)) return logout(e.message);
      if (state.assets.length) { showList(); render(); toast(navigator.onLine ? e.message : "ออฟไลน์: แสดงข้อมูลล่าสุดที่เคยโหลด", 4000); }
      else showMessage("โหลดข้อมูลไม่สำเร็จ", e.message, "ลองอีกครั้ง", refresh);
    } finally {
      document.body.classList.remove("busy");
    }
  }

  /* ---------- saving (offline queue) ---------- */
  function enqueue(year, id, rec) {
    const q = pending().filter((p) => !(String(p.year) === String(year) && p.id === id));
    q.push({ year: String(year), id, rec });
    store.set("pending", q);
  }
  let flushing = false;
  async function flush() {
    if (flushing || !navigator.onLine) return updateSync();
    const q = pending();
    if (!q.length) return updateSync();
    flushing = true; updateSync();
    try {
      const { saved } = await api("audit", {
        records: q.map((p) => ({ year: p.year, id: p.id, ...p.rec }))
      });
      const ok = new Set(saved);
      const done = q.filter((p) => ok.has(p.id + "|" + p.rec.at));
      for (const p of done) (state.server[p.year] = state.server[p.year] || {})[p.id] = p.rec;
      store.set("pending", pending().filter((x) => !done.some((d) => d.year === x.year && d.id === x.id && d.rec.at === x.rec.at)));
      store.set("cache", { assets: state.assets, server: state.server, sheetUrl: state.sheetUrl });
      composeRecords();
    } catch (e) {
      console.warn(e);
      if (e instanceof ApiError) toast("ส่งผลตรวจไม่สำเร็จ: " + e.message, 4000);
    } finally {
      flushing = false; updateSync(); render();
    }
  }
  function updateSync() {
    const n = pending().length;
    $("#syncNote").textContent = !n ? "" :
      flushing ? `กำลังส่ง ${n} รายการขึ้น Google Sheet…` :
      navigator.onLine ? `รอส่งขึ้น Google Sheet ${n} รายการ` : `ออฟไลน์: เก็บไว้ในเครื่อง ${n} รายการ จะส่งเมื่อต่อเน็ต`;
  }

  /* ---------- list ---------- */
  const inBuilding = (a) => !state.building || (a.building || "ไม่ระบุอาคาร") === state.building;
  function fillBuildings() {
    const count = {};
    for (const a of state.assets) { const b = a.building || "ไม่ระบุอาคาร"; count[b] = (count[b] || 0) + 1; }
    const names = Object.keys(count).sort((x, y) => x.localeCompare(y, "th"));
    if (state.building && !count[state.building]) { state.building = ""; store.set("building", ""); }
    $("#building").innerHTML = `<option value="">ทุกอาคาร (${state.assets.length})</option>` +
      names.map((b) => `<option value="${esc(b)}" ${b === state.building ? "selected" : ""}>${esc(b)} (${count[b]})</option>`).join("");
  }
  const FILTERS = {
    all: () => true,
    todo: (a) => !state.records[a.id],
    done: (a) => !!state.records[a.id],
    issue: (a) => { const t = tone(state.records[a.id]?.status); return t === "warn" || t === "bad"; },
    due: (a) => isDue(a)
  };
  function renderCounts(scope) {
    for (const k of Object.keys(FILTERS)) {
      const el = document.querySelector(`[data-n="${k}"]`);
      if (el) el.textContent = scope.filter(FILTERS[k]).length;
    }
    let over = 0, next = 0;
    for (const a of scope) { const g = ageInfo(a); if (g?.tone === "bad") over++; else if (g?.tone === "warn") next++; }
    const box = $("#ageSummary");
    box.hidden = !(over || next);
    box.innerHTML = `<span class="age-bad">ครบอายุแล้ว ${over} รายการ</span>` +
      (next ? `<span class="age-warn">ครบปีงบหน้า ${next} รายการ</span>` : "");
  }
  function filtered() {
    const q = state.q.trim().toLowerCase();
    return state.assets.filter((a) => {
      if (!inBuilding(a)) return false;
      if (!FILTERS[state.filter](a)) return false;
      if (!q) return true;
      return [a.id, a.name, a.room, a.building, a.owner].join(" ").toLowerCase().includes(q);
    });
  }
  function render() {
    if (!state.assets.length) return;
    const scope = state.assets.filter(inBuilding);
    const done = scope.filter((a) => state.records[a.id]).length;
    $("#doneCount").textContent = done;
    $("#totalCount").textContent = scope.length;
    $("#scope").textContent = state.building ? ` ใน${state.building}` : "";
    renderCounts(scope);
    $("#bar").style.width = (scope.length ? 100 * done / scope.length : 0) + "%";
    const items = filtered();
    $("#empty").hidden = items.length > 0;
    $("#list").innerHTML = items.map((a) => {
      const rec = state.records[a.id];
      const t = tone(rec?.status);
      const src = a.localPhoto || thumb(a, 200);
      const loc = [a.room, a.building].filter(Boolean).join(", ");
      return `<li><button class="row ${t}" data-id="${esc(a.id)}">
        ${src ? `<img class="thumb" loading="lazy" alt="" referrerpolicy="no-referrer" src="${esc(src)}">` : `<span class="thumb"></span>`}
        <span class="row-main">
          <span class="rid">${esc(a.id)}</span>
          <span class="rname">${esc(a.name)}</span>
          <span class="rloc">${esc(loc)}</span>
          ${(() => { const g = ageInfo(a); return g && g.short ? `<span class="rage ${g.tone}">${esc(g.short)} ${esc(a.replaceFY)}</span>` : ""; })()}
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

  /* ---------- QR code ของทรัพย์สิน (เข้ารหัสเฉพาะรหัสสินทรัพย์ เหมือนสติ๊กเกอร์เดิม) ---------- */
  function qrDataUrl(text) {
    if (typeof qrcode !== "function") return "";
    const q = qrcode(0, "M");
    q.addData(String(text));
    q.make();
    return q.createDataURL(8, 4);
  }

  /* ---------- detail ---------- */
  let draft = null;
  function openDetail(id, fromScan, replace) {
    const a = findAsset(id); if (!a) return toast(`ไม่พบรหัส ${id} ในฐานข้อมูล`);
    const rec = state.records[a.id];
    draft = { id: a.id, status: rec?.status || "", photo: null, fromScan };
    const big = a.localPhoto || thumb(a, 1200);
    const fact = (k, v) => v === "" || v == null ? "" : `<dt>${k}</dt><dd>${v}</dd>`;
    const el = $("#detail");
    el.innerHTML = `
      <div class="sheet-bar"><button class="back" data-close>‹ กลับ</button><span class="rid">${esc(a.id)}</span></div>
      <div class="sheet-body">
        <figure class="photo" id="photoBox">${big ? `<img alt="รูปทรัพย์สิน" referrerpolicy="no-referrer" src="${esc(big)}">` : "ยังไม่มีรูป"}</figure>
        <h2>${esc(a.name)}</h2>
        <dl class="facts">
          ${fact("ประเภท", esc([a.typeCode, a.typeName].filter(Boolean).join(" ")))}
          ${fact("จำนวน", esc(a.qty))}
          ${fact("ผู้ครอบครอง", esc(a.owner))}
          ${fact("สถานที่", esc(a.room))}
          ${fact("อาคาร", esc(a.building))}
          ${fact("สถานะ", esc(a.status))}
          ${fact("วันที่ได้มา", esc(a.acquired))}
          ${fact("อายุการใช้งาน", a.life ? esc(a.life) + " ปี" : (a.acquired ? "ไม่กำหนด" : ""))}
          ${(() => { const g = ageInfo(a); return g ? fact("เปลี่ยนได้", `<span class="age-${g.tone}">${esc(g.text)}</span>`) : ""; })()}
          ${fact("ขอทดแทนแล้ว", a.replaceReq ? "ปี " + esc(a.replaceReq) : "")}
          ${fact("หมายเหตุ", esc(a.note))}
          ${a.lat != null && a.lng != null ? fact("ตำแหน่ง", `<a href="https://www.google.com/maps?q=${a.lat},${a.lng}" target="_blank" rel="noopener">เปิดในแผนที่</a>`) : ""}
        </dl>
        ${(() => { const qr = qrDataUrl(a.id); return qr ? `
        <section class="qr-box">
          <img src="${qr}" alt="QR code รหัส ${esc(a.id)}">
          <div>
            <p class="qr-id">${esc(a.id)}</p>
            <p class="qr-hint">ใช้แทนสติ๊กเกอร์ที่สแกนไม่ติด กดค้างที่รูปเพื่อบันทึกหรือพิมพ์ใหม่</p>
            <a class="btn ghost" download="QR_${esc(a.id)}.gif" href="${qr}">ดาวน์โหลด QR</a>
          </div>
        </section>` : ""; })()}
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
  async function toJpegBase64(file, max = 1600) {
    const bmp = await createImageBitmap(file);
    const k = Math.min(1, max / Math.max(bmp.width, bmp.height));
    const c = document.createElement("canvas");
    c.width = Math.round(bmp.width * k); c.height = Math.round(bmp.height * k);
    c.getContext("2d").drawImage(bmp, 0, 0, c.width, c.height);
    const url = c.toDataURL("image/jpeg", 0.85);
    return { url, b64: url.split(",")[1] };
  }

  async function save(scanNext) {
    if (!draft.status) return toast("เลือกผลตรวจก่อนบันทึก");
    $("#detail").querySelectorAll(".actions button").forEach((b) => (b.disabled = true));
    const pos = await getPosition();
    const rec = {
      status: draft.status,
      note: $("#note").value.trim(),
      by: state.session?.name || "",
      at: new Date().toISOString(),
      ...(pos || {})
    };
    enqueue(state.year, draft.id, rec);
    state.records[draft.id] = { ...rec, pending: true };
    let msg = "บันทึกผลตรวจแล้ว";
    if (draft.photo) {
      if (!navigator.onLine) msg = "บันทึกผลตรวจแล้ว แต่ต้องต่อเน็ตเพื่ออัปโหลดรูป";
      else {
        try {
          const { url, b64 } = await toJpegBase64(draft.photo);
          const j = await api("photo", { id: draft.id, data: b64 });
          const a = findAsset(draft.id);
          a.imageId = j.imageId; a.localPhoto = url;
          msg = "บันทึกผลตรวจและรูปแล้ว";
        } catch (e) { msg = "บันทึกผลตรวจแล้ว แต่อัปโหลดรูปไม่สำเร็จ: " + e.message; }
      }
    }
    toast(msg, 3000);
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

  /* ---------- login / logout ---------- */
  $("#login").addEventListener("submit", async (e) => {
    e.preventDefault();
    const key = $("#loginKey").value.trim(), name = $("#loginName").value.trim();
    const btn = $("#login button"); btn.disabled = true;
    try {
      await api("ping", {}, key);
      state.session = { key, name };
      store.set("session", state.session); store.set("lastName", name);
      $("#loginKey").value = "";
      $("#who").textContent = name;
      showMessage("กำลังโหลดข้อมูล…", "");
      refresh();
    } catch (err) {
      showLogin(err.message);
    } finally { btn.disabled = false; }
  });
  function logout(reason) {
    store.del("session"); store.del("cache");
    state.session = null; state.assets = []; state.server = {}; state.records = {};
    $("#who").textContent = "";
    showLogin(reason);
  }

  /* ---------- controls ---------- */
  $("#building").addEventListener("change", (e) => {
    state.building = e.target.value; store.set("building", state.building); render();
    window.scrollTo({ top: 0 });
  });
  $("#ageSummary").addEventListener("click", () => {
    state.filter = "due";
    $("#chips").querySelectorAll("button").forEach((x) => x.classList.toggle("on", x.dataset.f === "due"));
    render();
  });
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
  $("#year").addEventListener("change", (e) => {
    state.year = e.target.value; store.set("year", state.year);
    composeRecords(); render();
  });
  const menu = $("#menu"), menuBtn = $("#menuBtn");
  menuBtn.addEventListener("click", (e) => {
    e.stopPropagation(); menu.hidden = !menu.hidden; menuBtn.setAttribute("aria-expanded", String(!menu.hidden));
  });
  document.addEventListener("click", () => { menu.hidden = true; menuBtn.setAttribute("aria-expanded", "false"); });
  menu.addEventListener("click", (e) => {
    const act = e.target.dataset.act;
    if (act === "refresh") refresh();
    if (act === "sheet") { if (state.sheetUrl) window.open(state.sheetUrl, "_blank", "noopener"); else toast("ยังไม่ได้โหลดข้อมูล"); }
    if (act === "logout") {
      if (pending().length && !confirm(`ยังมีผลตรวจ ${pending().length} รายการที่ยังไม่ได้ส่ง ถ้าออกจากระบบจะยังเก็บไว้ในเครื่องนี้ และส่งเมื่อเข้าใช้งานครั้งถัดไป ต้องการออกหรือไม่`)) return;
      logout();
    }
  });
  window.addEventListener("online", () => { updateSync(); flush(); });
  window.addEventListener("offline", updateSync);

  /* ---------- start ---------- */
  function start() {
    if ("serviceWorker" in navigator) navigator.serviceWorker.register("sw.js").catch(() => {});
    if (!CFG.apiUrl || !/^https:\/\/script\.google\.com\/.+\/exec$/.test(CFG.apiUrl)) {
      return showMessage("ยังไม่ได้ตั้งค่าแอป", "ใส่ Web app URL ของ Apps Script ในไฟล์ config.js ตามขั้นตอนใน README");
    }
    if (!state.session?.key) return showLogin();
    $("#who").textContent = state.session.name;
    const cache = store.get("cache", null);
    if (cache?.assets?.length) {
      state.assets = cache.assets; state.server = cache.server || {}; state.sheetUrl = cache.sheetUrl || "";
      fillYears(); fillBuildings(); composeRecords(); showList(); render(); updateSync();
    } else {
      showMessage("กำลังโหลดข้อมูล…", "");
    }
    refresh();
  }
  start();
})();
