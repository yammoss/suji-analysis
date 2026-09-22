// 수지분석 웹 - 화면 동작. 공유폴더 파일을 읽어 워커(파이썬)에 넘기고 결과 엑셀을 받는다.
"use strict";
const $ = (id) => document.getElementById(id);
const worker = new Worker("worker.js");
let seq = 0;
const waiting = new Map();
let folder = null;          // FileSystemDirectoryHandle
let pickedFiles = null;     // 폴더 선택을 못 쓰는 브라우저용 (File 목록)
const sent = new Map();     // 파일명 -> "size_lastModified" (바뀐 파일만 다시 넘긴다)
let costSig = "";
let info = null;
let ready = false;

// ── 워커 호출 ─────────────────────────────────────────────────────────
worker.onmessage = (ev) => {
  const m = ev.data;
  if (m.type === "log") return appendLog(m.text);
  if (m.type === "status") return setStatus(m.text, true);
  const w = waiting.get(m.id);
  if (!w) return;
  waiting.delete(m.id);
  m.type === "error" ? w.reject(new Error(m.text)) : w.resolve(m.result);
};
function rpc(cmd, args = {}, transfer = []) {
  const id = ++seq;
  return new Promise((resolve, reject) => {
    waiting.set(id, { resolve, reject });
    worker.postMessage({ id, cmd, args }, transfer);
  });
}

// ── 화면 표시 ─────────────────────────────────────────────────────────
function appendLog(text) {
  const el = $("log");
  el.textContent += text + "\n";
  el.scrollTop = el.scrollHeight;
}
function setStatus(text, busy = false) {
  $("statusline").textContent = text;
  $("statusline").classList.toggle("busy", busy);
}
function setRunnable(on) {
  for (const b of ["run-compare", "preview-compare", "run-cost"]) $(b).disabled = !on;
}

// ── 폴더 기억 (IndexedDB) ────────────────────────────────────────────
function idb() {
  return new Promise((ok, fail) => {
    const r = indexedDB.open("suji-web", 1);
    r.onupgradeneeded = () => r.result.createObjectStore("kv");
    r.onsuccess = () => ok(r.result);
    r.onerror = () => fail(r.error);
  });
}
async function kv(key, value) {
  const db = await idb();
  return new Promise((ok, fail) => {
    const tx = db.transaction("kv", value === undefined ? "readonly" : "readwrite");
    const st = tx.objectStore("kv");
    const r = value === undefined ? st.get(key) : st.put(value, key);
    r.onsuccess = () => ok(r.result);
    r.onerror = () => fail(r.error);
  });
}

// ── 데이터 파일 읽기 ─────────────────────────────────────────────────
const wanted = (name) => /\.xlsx$/i.test(name) && !name.startsWith("~$") && !name.includes("발매");

async function listFiles() {
  if (pickedFiles) return pickedFiles.filter((f) => wanted(f.name));
  const out = [];
  for await (const entry of folder.values()) {
    if (entry.kind === "file" && wanted(entry.name)) out.push(await entry.getFile());
  }
  return out;
}

async function syncFiles() {
  const files = await listFiles();
  const changed = files.filter((f) => sent.get(f.name) !== `${f.size}_${f.lastModified}`);
  if (changed.length) {
    setStatus(`데이터 파일 읽는 중 (${changed.map((f) => f.name).join(", ")})`, true);
    const payload = [];
    for (const f of changed) payload.push({ name: f.name, lastModified: f.lastModified, buffer: await f.arrayBuffer() });
    await rpc("files", { files: payload }, payload.map((p) => p.buffer));
    for (const f of changed) sent.set(f.name, `${f.size}_${f.lastModified}`);
  }
  info = await rpc("status");
  renderFiles(files);
  if (!info.cost || !info.actuals) throw new Error("필요한 파일이 폴더에 없습니다 (위 목록 확인)");
  const sig = sent.get(info.cost);
  if (sig !== costSig) {
    setStatus("비용파일에서 기준데이터 만드는 중 (파일이 바뀌었을 때만, 1분 안팎)", true);
    appendLog(await rpc("dataset"));
    costSig = sig;
  }
}

function renderFiles(files) {
  const has = (n) => files.some((f) => f.name === n);
  const items = [
    [info && info.cost, `비용 추정용 파일${info && info.cost ? " : " + info.cost : " (이름에 '비용' 이 든 엑셀)"}`],
    [has("과거실적 DATA.xlsx"), "과거실적 DATA.xlsx"],
    [has("공휴일 DATA.xlsx"), "공휴일 DATA.xlsx (일자 매칭용, 없어도 됨)"],
  ];
  $("files").innerHTML = items.map(([ok, t]) => `<li class="${ok ? "ok" : "miss"}">${t}</li>`).join("");
  const plan = info && info.plan_tabs && info.plan_tabs.length;
  $("basis-box").hidden = !plan;
  $("plantabs").textContent = plan ? info.plan_tabs.join(", ") : "";
  updateDailyBox();
}

function updateDailyBox() {
  const past = !$("basis-box").hidden ? document.querySelector("input[name=basis]:checked").value === "past" : true;
  $("daily-box").hidden = !(info && info.daily && past);
}
document.querySelectorAll("input[name=basis]").forEach((r) => r.addEventListener("change", updateDailyBox));

async function connect(handle) {
  folder = handle;
  pickedFiles = null;
  $("foldername").textContent = "연결된 폴더 : " + handle.name;
  $("regrant").hidden = true;
  await afterConnect();
}

async function afterConnect() {
  if (!ready) return;
  try {
    setRunnable(false);
    await syncFiles();
    setStatus("준비 완료 - 분석할 내용을 넣고 계산하세요");
    setRunnable(true);
  } catch (e) {
    setStatus("확인 필요 : " + e.message);
  }
}

$("pick").onclick = async () => {
  if (window.showDirectoryPicker) {
    try {
      const h = await window.showDirectoryPicker({ id: "suji-data", mode: "read" });
      await kv("folder", h);
      await connect(h);
    } catch (e) {
      if (e.name !== "AbortError") setStatus("폴더 선택 실패 : " + e.message);
    }
  } else {
    const inp = document.createElement("input");
    inp.type = "file";
    inp.multiple = true;
    inp.webkitdirectory = true;
    inp.onchange = async () => {
      pickedFiles = [...inp.files].filter((f) => !f.webkitRelativePath.split("/").slice(1, -1).length);
      folder = null;
      $("foldername").textContent = "선택한 폴더 (이 브라우저는 폴더를 기억하지 못해 매번 선택)";
      await afterConnect();
    };
    inp.click();
  }
};

$("regrant").onclick = async () => {
  const h = await kv("folder");
  if (h && (await h.requestPermission({ mode: "read" })) === "granted") await connect(h);
};

// ── 탭 ────────────────────────────────────────────────────────────────
document.querySelectorAll(".tabs button").forEach((b) => {
  b.onclick = () => {
    document.querySelectorAll(".tabs button").forEach((x) => x.classList.toggle("on", x === b));
    document.querySelectorAll(".pane").forEach((p) => (p.hidden = p.id !== "pane-" + b.dataset.tab));
  };
});

// ── 실행 ──────────────────────────────────────────────────────────────
async function runTool(script, argv, inputText) {
  setRunnable(false);
  $("download").innerHTML = "";
  $("log").textContent = "";
  try {
    await syncFiles();                                   // 공유폴더 파일이 바뀌었으면 다시 읽음
    if (inputText !== undefined) {
      const buf = new TextEncoder().encode(inputText).buffer;
      await rpc("files", { files: [{ name: "입력.txt", lastModified: Date.now(), buffer: buf }] }, [buf]);
    }
    setStatus("계산 중...", true);
    const res = await rpc("run", { script, argv });
    if (res && res.bytes) {
      const url = URL.createObjectURL(new Blob([res.bytes], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
      const a = document.createElement("a");
      a.href = url;
      a.download = res.name;
      a.className = "dl";
      a.textContent = "⬇ " + res.name;
      $("download").appendChild(a);
      a.click();
      setStatus("완료 - 엑셀을 내려받았습니다 (안 받아졌으면 아래 버튼)");
    } else {
      setStatus("완료");
    }
  } catch (e) {
    setStatus("오류 : " + e.message);
    appendLog("\n[오류] " + e.message);
  } finally {
    setRunnable(true);
  }
}

async function compareArgs() {
  let text, period;
  if (formMode()) {
    const g = genForm();
    if (g.errors.length) { setStatus("폼 확인 : " + g.errors[0]); return null; }
    if (!g.lines.length) { setStatus("노선을 입력하세요"); return null; }
    text = g.lines.join("\n");
    period = $("period").value.trim();
  } else {
    text = $("input").value.trim();
    if (!text) { setStatus("분석할 내용을 입력하세요"); return null; }
    period = (await rpc("period", { text })) || $("period").value.trim();   // 입력에 적힌 기간이 우선
  }
  if (!period) { setStatus("기간을 알 수 없습니다 - 기간 칸에 W26 / 27.01~27.02 처럼 적으세요"); return null; }
  const argv = ["입력.txt", "--period", period, "--yes",
    "--fixed-alloc", document.querySelector("input[name=alloc]:checked").value];
  const planShown = !$("basis-box").hidden;
  const past = !planShown || document.querySelector("input[name=basis]:checked").value === "past";
  if (planShown && past) argv.push("--no-plan");
  if (!$("daily-box").hidden && document.querySelector("input[name=daily]:checked").value === "daily") argv.push("--daily-match");
  return { argv, text: text + "\n" };
}

$("run-compare").onclick = async () => {
  const a = await compareArgs();
  if (a) await runTool("run.py", a.argv, a.text);
};

// ── 인식 결과 미리보기 (계산 전 확인) ─────────────────────────────────
const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
const CASK_MARK = { EXACT: ["OK", "ok"], SIBLING_PROXY: ["대체", "warn"], BT_PROXY: ["대체", "warn"],
  TYPE_PROXY: ["대체", "warn"], NONE: ["없음", "bad"] };

function renderPreview(d) {
  const box = $("preview");
  if (!d) {
    box.innerHTML = `<div class="pv-head bad">입력을 읽지 못했습니다 - 아래 기록 창의 메시지를 확인하세요</div>`;
    return;
  }
  const cask = new Map(d.cask.map((c) => [c.route + "|" + c.aircraft, c]));
  let bad = 0;
  const rows = d.rows.map((r) => {
    const issues = [];
    if (!r.route) issues.push(`노선 인식 실패 (${esc(r.route_raw || "?")})`);
    if (!r.aircraft) issues.push("기종 없음");
    if (r.rt_failed) issues.push("스케줄 인식 실패");
    const c = cask.get(r.route + "|" + r.aircraft);
    const [mark, cls] = c ? CASK_MARK[c.level] || [c.level, ""] : ["-", ""];
    if (c && c.level === "NONE") issues.push("원가 없음");
    if (issues.length) bad++;
    return `<tr class="${issues.length ? "err" : ""}">
      <td>${r.no}</td><td>${esc(r.side)}</td><td>${esc(r.route || r.route_raw)}</td><td>${esc(r.aircraft || "-")}</td>
      <td>${esc(r.period)}</td>
      <td class="num">${r.rt.toLocaleString("ko-KR", { maximumFractionDigits: 1 })}<div class="sub">${esc(r.rt_src)}</div></td>
      <td class="num">${r.lf == null ? "-" : (r.lf * 100).toFixed(1) + "%"}</td>
      <td class="num">${r.ar == null ? "-" : Math.round(r.ar).toLocaleString("ko-KR")}</td>
      <td><span class="tag ${cls}" title="${esc(c && c.note)}">${mark}</span></td>
      <td class="issue">${issues.join("<br>")}</td></tr>`;
  }).join("");
  const warn = d.problems.length;
  const head = bad ? `<div class="pv-head bad">인식 못 한 항목 ${bad}건 - 빨간 줄을 고친 뒤 다시 확인하세요</div>`
    : warn ? `<div class="pv-head warn">인식은 정상 · 참고사항 ${warn}건 (아래) - 확인 후 계산하세요</div>`
    : `<div class="pv-head ok">모두 정상 인식 - 계산해도 됩니다</div>`;
  box.innerHTML = `${head}
    <div class="pv-meta">기간 ${esc(d.period)}${d.grid ? " · 환율 x 유가 민감도 모드" : ""}</div>
    <div class="pv-scroll"><table class="pv">
      <tr><th>#</th><th>구분</th><th>노선</th><th>기종</th><th>기간</th><th>왕복</th><th>L/F</th><th>A/R</th><th>원가</th><th></th></tr>
      ${rows}</table></div>
    ${warn ? `<ul class="pv-warn">${d.problems.map((p) => `<li>${esc(p)}</li>`).join("")}</ul>` : ""}
    ${d.basis.length ? `<details><summary>L/F · A/R · 화물 근거</summary><ul class="pv-basis">${d.basis.map((b) => `<li>${esc(b)}</li>`).join("")}</ul></details>` : ""}`;
}

$("preview-compare").onclick = async () => {
  const a = await compareArgs();
  if (!a) return;
  setRunnable(false);
  $("log").textContent = "";
  $("preview").innerHTML = "";
  try {
    await syncFiles();
    const buf = new TextEncoder().encode(a.text).buffer;
    await rpc("files", { files: [{ name: "입력.txt", lastModified: Date.now(), buffer: buf }] }, [buf]);
    setStatus("입력 확인 중...", true);
    renderPreview(await rpc("preview", { argv: a.argv }));
    setStatus("인식 결과를 확인하세요");
  } catch (e) {
    setStatus("오류 : " + e.message);
    appendLog("\n[오류] " + e.message);
  } finally {
    setRunnable(true);
  }
};
$("input").addEventListener("input", () => { $("preview").innerHTML = ""; });

document.querySelectorAll("button.ex").forEach((b) => {
  b.onclick = () => {
    $("input").value = b.dataset.ex;
    $("preview").innerHTML = "";
    $("input").focus();
  };
});

// ── 수지비교 폼 입력 ──────────────────────────────────────────────────
// 폼은 입력 문장을 만들어 줄 뿐이고, 계산은 그 문장을 기존 파서가 그대로 읽는다.
const TYPES = [
  ["ac", "기종 변경"], ["sc", "스케줄 변경"], ["both", "기종 + 스케줄"],
  ["swap", "노선 간 기종 맞바꿈"], ["cut", "감편 · 운휴"],
];
const DAYS = ["월", "화", "수", "목", "금", "토", "일"];
const formMode = () => !$("mode-form").hidden;

const newPart = () => ({ until: "", days: [1, 1, 1, 1, 1, 1, 1], x2: false });
const newSide = () => ({ ac: "", parts: [newPart()] });
let cards = [];

function partCode(p) {
  const on = p.days.map((v, i) => (v ? i + 1 : "")).join("");
  if (on.length === 7) return p.x2 ? "DAILY x2" : "DAILY";
  if (!on.length) return "비운항";
  return `${on.length}/W D${on}`;
}
function schedText(side) {
  const ps = side.parts;
  if (ps.length === 1) return partCode(ps[0]);
  const head = ps.slice(0, -1).map((p) => `${p.until.trim()}까지 ${partCode(p)}`);
  return head.join(", ") + ", 이후 " + partCode(ps[ps.length - 1]);
}

function genForm() {
  const lines = [], errors = [];
  cards.forEach((c, i) => {
    const n = i + 1, r = c.route.trim().toUpperCase().replace(/\s+/g, ""), r2 = c.route2.trim().toUpperCase().replace(/\s+/g, "");
    const A = c.a.ac.trim().toUpperCase(), B = c.b.ac.trim().toUpperCase();
    if (!r && !A && !B) return;                                   // 빈 카드는 건너뜀
    if (!r) errors.push(`비교 ${n} : 노선을 입력하세요`);
    if (!A) errors.push(`비교 ${n} : ${c.type === "swap" ? "노선 A" : "기존(안)"} 기종을 입력하세요`);
    if ((c.type === "ac" || c.type === "both" || c.type === "swap") && !B)
      errors.push(`비교 ${n} : ${c.type === "swap" ? "노선 B" : "변경(안)"} 기종을 입력하세요`);
    if (c.type === "swap" && !r2) errors.push(`비교 ${n} : 노선 B 를 입력하세요`);
    for (const [label, side] of [["기존(안)", c.a], ["변경(안)", c.b]])
      side.parts.slice(0, -1).forEach((p) => {
        if (!/^\d{1,2}[./]\d{1,2}$/.test(p.until.trim())) errors.push(`비교 ${n} ${label} : 구간 끝 날짜를 12/18 처럼 적으세요`);
      });
    const sa = schedText(c.a), sb = schedText(c.b);
    if (c.type === "ac") lines.push(`${r} ${A} ${sa} vs ${B} ${sa}`);
    else if (c.type === "sc") lines.push(`${r} ${A} ${sa} vs ${A} ${sb}`);
    else if (c.type === "both") lines.push(`${r} ${A} ${sa} vs ${B} ${sb}`);
    else if (c.type === "cut") lines.push(`${r} ${A} ${sa} vs 비운항`);
    else lines.push(`${r} ${A} ${sa} + ${r2} ${B} ${sb} vs ${r} ${B} ${sa} + ${r2} ${A} ${sb}`);
  });
  return { lines, errors };
}

function renderGen() {
  const g = genForm();
  $("gen").textContent = g.lines.length ? g.lines.join("\n") : "(노선·기종을 넣으면 여기에 문장이 만들어집니다)";
  $("gen-err").textContent = g.errors.join(" · ");
  $("preview").innerHTML = "";
}

function partRow(side, idx, last, onChange) {
  const p = side.parts[idx];
  const row = document.createElement("div");
  row.className = "part";
  if (!last) {
    row.innerHTML = `<span class="plabel">~ <input type="text" class="until" placeholder="12/18"> 까지</span>`;
    const u = row.querySelector(".until");
    u.value = p.until;
    u.oninput = () => { p.until = u.value; renderGen(); };
  } else {
    row.innerHTML = `<span class="plabel">${side.parts.length > 1 ? "이후" : "전체 기간"}</span>`;
  }
  DAYS.forEach((d, i) => {
    const b = document.createElement("button");
    b.type = "button";
    b.className = "chip" + (p.days[i] ? " on" : "");
    b.textContent = d;
    b.onclick = () => { p.days[i] = p.days[i] ? 0 : 1; onChange(); };
    row.appendChild(b);
  });
  if (p.days.every(Boolean)) {
    const l = document.createElement("label");
    l.className = "opt";
    l.innerHTML = `<input type="checkbox"> x2`;
    l.querySelector("input").checked = p.x2;
    l.querySelector("input").onchange = (e) => { p.x2 = e.target.checked; onChange(); };
    row.appendChild(l);
  }
  const code = document.createElement("span");
  code.className = "code";
  code.textContent = partCode(p);
  row.appendChild(code);
  if (!last) {
    const x = document.createElement("button");
    x.type = "button";
    x.className = "x";
    x.textContent = "✕";
    x.title = "이 구간 삭제";
    x.onclick = () => { side.parts.splice(idx, 1); onChange(); };
    row.appendChild(x);
  }
  return row;
}

function sideBox(c, key, title, opts) {
  const side = c[key];
  const box = document.createElement("div");
  box.className = "side";
  box.innerHTML = `<h4>${title}</h4>`;
  const redraw = () => renderCards();
  if (opts.acLocked) {
    box.insertAdjacentHTML("beforeend", `<div class="locked">기종 : 기존(안)과 같음 (${esc(c.a.ac || "-")})</div>`);
  } else {
    const l = document.createElement("label");
    l.innerHTML = `기종 <input type="text" size="10" list="ac-list" placeholder="B738">`;
    const inp = l.querySelector("input");
    inp.value = side.ac;
    inp.oninput = () => { side.ac = inp.value; if (key === "a") syncLocked(box.parentElement, c); renderGen(); };
    box.appendChild(l);
  }
  if (opts.cut) {
    box.insertAdjacentHTML("beforeend", `<div class="locked">스케줄 : 비운항 (0왕복)</div>`);
  } else if (opts.schedLocked) {
    box.insertAdjacentHTML("beforeend", `<div class="locked">스케줄 : 기존(안)과 같음 (${esc(schedText(c.a))})</div>`);
  } else {
    side.parts.forEach((_, i) => box.appendChild(partRow(side, i, i === side.parts.length - 1, redraw)));
    const add = document.createElement("button");
    add.type = "button";
    add.className = "ghost small";
    add.textContent = "+ 구간 나누기";
    add.title = "기간 중간에 스케줄이 바뀔 때 (예: 12/18까지 주4회, 이후 DAILY)";
    add.onclick = () => { side.parts.splice(side.parts.length - 1, 0, newPart()); redraw(); };
    box.appendChild(add);
  }
  return box;
}

function syncLocked(sidesEl, c) {       // 기존(안) 기종을 칠 때 '같음' 표시만 갱신 (입력칸 포커스 유지)
  sidesEl.querySelectorAll(".locked").forEach((el) => {
    if (el.textContent.startsWith("기종")) el.textContent = `기종 : 기존(안)과 같음 (${c.a.ac || "-"})`;
  });
}

function renderCards() {
  const wrap = $("cards");
  wrap.innerHTML = "";
  cards.forEach((c, i) => {
    const el = document.createElement("div");
    el.className = "card2";
    const head = document.createElement("div");
    head.className = "card2-head";
    head.innerHTML = `<span class="no">비교 ${i + 1}</span>`;
    for (const [t, label] of TYPES) {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "tbtn" + (c.type === t ? " on" : "");
      b.textContent = label;
      b.onclick = () => { c.type = t; renderCards(); };
      head.appendChild(b);
    }
    const del = document.createElement("button");
    del.type = "button";
    del.className = "del";
    del.textContent = "✕";
    del.title = "이 비교 삭제";
    del.onclick = () => { cards.splice(i, 1); if (!cards.length) cards.push(newCard()); renderCards(); };
    head.appendChild(del);
    el.appendChild(head);

    const rrow = document.createElement("div");
    rrow.className = "row";
    rrow.innerHTML = `<label>${c.type === "swap" ? "노선 A" : "노선"} <input type="text" class="r1" size="10" placeholder="ICNDAD"></label>`
      + (c.type === "swap" ? `<label>노선 B <input type="text" class="r2" size="10" placeholder="ICNNRT"></label>` : "");
    const r1 = rrow.querySelector(".r1");
    r1.value = c.route;
    r1.oninput = () => { c.route = r1.value; renderGen(); };
    const r2 = rrow.querySelector(".r2");
    if (r2) { r2.value = c.route2; r2.oninput = () => { c.route2 = r2.value; renderGen(); }; }
    el.appendChild(rrow);

    const sides = document.createElement("div");
    sides.className = "sides";
    if (c.type === "swap") {
      sides.appendChild(sideBox(c, "a", "노선 A (지금 기종 · 스케줄)", {}));
      sides.appendChild(sideBox(c, "b", "노선 B (지금 기종 · 스케줄)", {}));
    } else {
      sides.appendChild(sideBox(c, "a", "기존(안)", {}));
      sides.appendChild(sideBox(c, "b", "변경(안)", {
        acLocked: c.type === "sc" || c.type === "cut",
        schedLocked: c.type === "ac",
        cut: c.type === "cut",
      }));
    }
    el.appendChild(sides);
    if (c.type === "swap")
      el.insertAdjacentHTML("beforeend", `<p class="muted">변경(안)은 두 노선의 기종을 서로 바꾸고, 스케줄은 각 노선 그대로 둡니다.</p>`);
    wrap.appendChild(el);
  });
  renderGen();
}
const newCard = () => ({ type: "ac", route: "", route2: "", a: newSide(), b: newSide() });

$("card-add").onclick = () => { cards.push(newCard()); renderCards(); };
$("to-text").onclick = () => {
  const g = genForm();
  const per = $("period").value.trim();
  $("input").value = (per ? per + "\n" : "") + g.lines.join("\n");
  setMode("text");
};
function setMode(m) {
  document.querySelectorAll(".modes button").forEach((b) => b.classList.toggle("on", b.dataset.mode === m));
  $("mode-form").hidden = m !== "form";
  $("mode-text").hidden = m !== "text";
  $("period-hint").textContent = m === "form" ? "W26 / S27 / 27.01~27.02 / 26.12.20~27.02.28"
    : "입력에 W26 처럼 기간을 적었으면 그게 우선합니다";
  $("preview").innerHTML = "";
}
document.querySelectorAll(".modes button").forEach((b) => (b.onclick = () => setMode(b.dataset.mode)));
cards.push(newCard());
renderCards();

// ── 비용 탭 ───────────────────────────────────────────────────────────
function addCostRow(route = "", ac = "") {
  const row = document.createElement("div");
  row.className = "cost-row";
  row.innerHTML = `<label>노선 <input type="text" class="c-route" size="10" placeholder="ICNNRT"></label>
    <label>기종 <input type="text" class="c-ac" size="13" list="ac-list" placeholder="B738, A333"></label>
    <label class="opt"><input type="checkbox" class="c-all"> 보유 기종 전부</label>
    <label>스케줄 <input type="text" class="c-sched" size="17" placeholder="(선택) 4/W D3467"></label>
    <button type="button" class="del" title="삭제">✕</button>`;
  row.querySelector(".c-route").value = route;
  row.querySelector(".c-ac").value = ac;
  const all = row.querySelector(".c-all"), acIn = row.querySelector(".c-ac");
  all.onchange = () => { acIn.disabled = all.checked; };
  row.querySelector(".del").onclick = () => {
    if ($("cost-rows").children.length > 1) row.remove();
    else { row.querySelector(".c-route").value = ""; acIn.value = ""; row.querySelector(".c-sched").value = ""; }
  };
  $("cost-rows").appendChild(row);
}
function addIndexBox(listId, value = "") {
  const inp = document.createElement("input");
  inp.type = "text";
  inp.inputMode = "decimal";
  inp.placeholder = listId === "fx-list" ? "1,350" : "300";
  inp.value = value;
  inp.oninput = updateIndexCount;
  $(listId).appendChild(inp);
}
const indexValues = (listId) => [...$(listId).querySelectorAll("input")]
  .map((i) => i.value.replace(/[,\s]/g, "")).filter((v) => v && !isNaN(+v));
function updateIndexCount() {
  const fx = indexValues("fx-list").length, fuel = indexValues("fuel-list").length;
  $("idx-count").textContent = fx || fuel
    ? `환율 ${fx || "INDEX 평균 1"}개 x 유가 ${fuel || "INDEX 평균 1"}개 = ${(fx || 1) * (fuel || 1)}가지 경우를 계산합니다`
    : "";
}
$("cost-add").onclick = () => addCostRow();
document.querySelectorAll("button[data-add]").forEach((b) => (b.onclick = () => addIndexBox(b.dataset.add)));
addCostRow();
for (let i = 0; i < 3; i++) { addIndexBox("fx-list"); addIndexBox("fuel-list"); }

$("run-cost").onclick = () => {
  const lines = [];
  for (const row of $("cost-rows").children) {
    const route = row.querySelector(".c-route").value.trim().toUpperCase();
    if (!route) continue;
    const sched = row.querySelector(".c-sched").value.trim();
    const tail = (sched ? sched + " " : "") + "비용";
    if (row.querySelector(".c-all").checked) { lines.push(`${route} 기종별 ${tail}`); continue; }
    const acs = row.querySelector(".c-ac").value.split(/[,\s/]+/).map((a) => a.trim().toUpperCase()).filter(Boolean);
    if (!acs.length) return setStatus(`${route} 의 기종을 입력하세요 (또는 '보유 기종 전부')`);
    for (const ac of acs) lines.push(`${route} ${ac} ${tail}`);
  }
  if (!lines.length) return setStatus("노선을 입력하세요");
  const fx = indexValues("fx-list"), fuel = indexValues("fuel-list");
  const head = [fx.length ? `환율 : ${fx.join("/")}` : "", fuel.length ? `유가 : ${fuel.join("/")}` : ""].filter(Boolean).join(", ");
  const text = [head, ...lines].filter(Boolean).join("\n") + "\n";
  const period = $("c-period").value.trim() || "W26";
  runTool("run.py", ["입력.txt", "--period", period, "--yes",
    "--fixed-alloc", document.querySelector("input[name=calloc]:checked").value], text);
};

// ── 시작 ──────────────────────────────────────────────────────────────
(async () => {
  const saved = await kv("folder").catch(() => null);
  rpc("init").then(async () => {
    ready = true;
    setStatus("계산 환경 준비 완료 - 데이터 폴더를 연결하세요");
    if (folder || pickedFiles) await afterConnect();
  }).catch((e) => setStatus("계산 환경을 불러오지 못했습니다 : " + e.message));
  if (saved) {
    $("foldername").textContent = "이전에 연결한 폴더 : " + saved.name;
    if ((await saved.queryPermission({ mode: "read" })) === "granted") await connect(saved);
    else $("regrant").hidden = false;
  }
})();
