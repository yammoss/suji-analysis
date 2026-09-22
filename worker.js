// 수지분석 웹 - 파이썬(Pyodide) 실행 워커. 화면이 멈추지 않게 계산은 여기서만 한다.
// 데이터 파일은 화면(app.js)이 공유폴더에서 읽어 넘겨준다. 이 워커는 인터넷으로 아무것도 보내지 않는다
// (받아오는 것은 Pyodide 실행환경과 이 사이트의 파이썬 코드뿐).
const PYODIDE = "https://cdn.jsdelivr.net/pyodide/v0.27.2/full/";
importScripts(PYODIDE + "pyodide.js");

let py = null;
const log = (text, kind = "log") => postMessage({ type: kind, text });

async function syncfs(populate) {
  return new Promise((ok, fail) => py.FS.syncfs(populate, (e) => (e ? fail(e) : ok())));
}

async function init() {
  log("계산 환경 준비 중 (처음 한 번은 10~20초)...", "status");
  py = await loadPyodide({ indexURL: PYODIDE });
  py.setStdout({ batched: (s) => log(s) });
  py.setStderr({ batched: (s) => log(s) });
  try {
    await py.loadPackage(["openpyxl"]);
  } catch (e) {
    await py.loadPackage("micropip");
    await py.runPythonAsync("import micropip; await micropip.install('openpyxl')");
  }
  py.FS.mkdirTree("/cache");
  py.FS.mount(py.FS.filesystems.IDBFS, {}, "/cache");
  await syncfs(true);
  py.FS.mkdirTree("/work/app/profit_tool");
  py.FS.mkdirTree("/work/app/data");
  py.FS.mkdirTree("/work/app/output");
  const files = await (await fetch("py/manifest.json", { cache: "no-cache" })).json();
  for (const f of files.concat(["../web_entry.py"])) {
    const src = await (await fetch("py/" + f, { cache: "no-cache" })).text();
    const dest = "/work/app/" + f.replace("../", "");
    py.FS.writeFile(dest, src);
  }
  await py.runPythonAsync("import sys; sys.path.insert(0, '/work/app'); import web_entry");
  log("계산 환경 준비 완료", "status");
}

function writeData(files) {
  for (const f of files) {
    const path = "/work/app/" + f.name;
    py.FS.writeFile(path, new Uint8Array(f.buffer));
    const t = f.lastModified;
    py.FS.utime(path, t, t);
  }
}

async function call(expr) {
  return await py.runPythonAsync(expr);
}

onmessage = async (ev) => {
  const { id, cmd, args } = ev.data;
  try {
    let result = null;
    if (cmd === "init") {
      await init();
    } else if (cmd === "files") {
      writeData(args.files);
    } else if (cmd === "status") {
      result = JSON.parse(await call("web_entry.status()"));
    } else if (cmd === "dataset") {
      result = await call("web_entry.ensure_dataset()");
      await syncfs(false);
    } else if (cmd === "period") {
      py.globals.set("_txt", args.text);
      result = await call("web_entry.detect_period(_txt)");
    } else if (cmd === "run") {
      py.globals.set("_argv", py.toPy(args.argv));
      const out = await call(`web_entry.run_script(${JSON.stringify(args.script)}, _argv)`);
      if (out) {
        const bytes = py.FS.readFile(out);
        result = { name: out.split("/").pop(), bytes };
        postMessage({ id, type: "done", result }, [bytes.buffer]);
        return;
      }
    }
    postMessage({ id, type: "done", result });
  } catch (e) {
    postMessage({ id, type: "error", text: String(e && e.message ? e.message : e) });
  }
};
