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
  for (const b of ["run-compare", "run-breakeven"]) $(b).disabled = !on;
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

$("run-compare").onclick = async () => {
  const text = $("input").value.trim();
  if (!text) return setStatus("분석할 내용을 입력하세요");
  let period = $("period").value.trim();
  if (!period) period = await rpc("period", { text });
  if (!period) return setStatus("기간을 알 수 없습니다 - 기간 칸이나 입력에 W26 / 27.01~27.02 처럼 적으세요");
  const argv = ["입력.txt", "--period", period, "--yes",
    "--fixed-alloc", document.querySelector("input[name=alloc]:checked").value];
  const planShown = !$("basis-box").hidden;
  const past = !planShown || document.querySelector("input[name=basis]:checked").value === "past";
  if (planShown && past) argv.push("--no-plan");
  if (!$("daily-box").hidden && document.querySelector("input[name=daily]:checked").value === "daily") argv.push("--daily-match");
  await runTool("run.py", argv, text + "\n");
};

document.querySelectorAll("button.ex").forEach((b) => {
  b.onclick = () => {
    $("input").value = b.dataset.ex;
    $("period").value = "";
    $("input").focus();
  };
});

$("run-breakeven").onclick = () => {
  const route = $("b-route").value.trim(), ac = $("b-ac").value.trim();
  if (!route || !ac) return setStatus("노선과 기종을 입력하세요");
  const argv = [route, ac, "--period", $("b-period").value.trim() || "W26",
    "--fixed-alloc", document.querySelector("input[name=balloc]:checked").value];
  if ($("b-ar").value.trim()) argv.push("--ar", $("b-ar").value.trim().replace(/,/g, ""));
  if ($("b-lf").value.trim()) argv.push("--lf", $("b-lf").value.trim().replace("%", ""));
  runTool("손익분기_계산기.py", argv);
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
