// HTML Frame: スライドに置いた枠の中で HTML や Web ページを表示する
"use strict";

const KEY = "htmlFrame";
const DEFAULTS = {
  mode: "html",
  html: "",
  url: "",
  designWidth: 0,
  pixelRatio: 0, // 0 なら画面の倍率のまま
  showFps: false,
  sameOrigin: false,
};

// 枠の左上に 1 秒あたりの描画回数を出す。0.5 秒ごとに数え直す
const FPS_SCRIPT = `(function () {
  var el, n = 0, t0 = performance.now();
  function tick(t) {
    n++;
    if (t - t0 >= 500) {
      if (!el && document.body) {
        el = document.createElement("div");
        el.style.cssText = "position:fixed;left:4px;top:4px;z-index:2147483647;padding:2px 6px;border-radius:4px;background:rgba(0,0,0,.6);color:#fff;font:12px/1.4 monospace;pointer-events:none";
        document.body.appendChild(el);
      }
      if (el) el.textContent = Math.round((n * 1000) / (t - t0)) + " FPS";
      n = 0;
      t0 = t;
    }
    requestAnimationFrame(tick);
  }
  requestAnimationFrame(tick);
})();`;

const $ = (id) => document.getElementById(id);
const form = $("editor");

let store;
let inOffice = false;
let config = null; // 保存済みの内容。null なら未設定
let view = "edit"; // "edit"（標準表示）または "read"（スライドショー・閲覧表示）
let editing = false;
let shown = null; // いま iframe が表示している config

// 設定はアドインの枠ごとにプレゼンテーションの中へ保存される
function officeStore() {
  const settings = Office.context.document.settings;
  return {
    load: () => settings.get(KEY),
    save(cfg) {
      settings.set(KEY, cfg);
      return new Promise((resolve, reject) => {
        settings.saveAsync((r) =>
          r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(r.error)
        );
      });
    },
  };
}

// Office の外（ブラウザ）で開いたときの確認用。?view=read でスライドショー中の表示になる
function browserStore() {
  return {
    load() {
      try {
        return JSON.parse(localStorage.getItem(KEY));
      } catch {
        return null;
      }
    },
    save(cfg) {
      try {
        localStorage.setItem(KEY, JSON.stringify(cfg));
      } catch {}
      return Promise.resolve();
    },
  };
}

function setView(v) {
  view = v === "read" ? "read" : "edit";
  refresh();
}

function refresh() {
  const showEditor = editing && view === "edit";
  form.hidden = !showEditor;
  $("viewer").hidden = showEditor || !config;
  $("empty").hidden = showEditor || !!config;
  // 「ブラウザで開く」は URL を開くときだけ、スライドショー中も出す。ほかのボタンは編集画面だけ
  const isUrl = !!config && config.mode === "url";
  $("btn-open").hidden = !isUrl;
  $("btn-reload").hidden = view !== "edit";
  $("btn-edit").hidden = view !== "edit";
  $("toolbar").hidden = view !== "edit" && !isUrl;
  if (!$("viewer").hidden && shown !== config) render();
}

// YouTube の再生ページ、短縮 URL、ショート、ライブの URL を、枠に表示できる埋め込み用の URL に直す
function embedUrl(url) {
  let u;
  try {
    u = new URL(url);
  } catch {
    return url;
  }
  const host = u.hostname.replace(/^(www|m)\./, "");
  let id = null;
  if (host === "youtu.be") {
    id = u.pathname.split("/")[1];
  } else if (host === "youtube.com") {
    const path = u.pathname.match(/^\/(shorts|live)\/([^/]+)/);
    if (u.pathname === "/watch") id = u.searchParams.get("v");
    else if (path) id = path[2];
  }
  if (!id || !/^[\w-]{11}$/.test(id)) return url;
  const start = toSeconds(u.searchParams.get("t") || u.searchParams.get("start"));
  return `https://www.youtube.com/embed/${id}` + (start ? `?start=${start}` : "");
}

// "90"、"90s"、"1m30s"、"1h2m3s" を秒にする
function toSeconds(t) {
  if (!t) return 0;
  if (/^\d+$/.test(t)) return Number(t);
  const m = t.match(/^(?:(\d+)h)?(?:(\d+)m)?(?:(\d+)s)?$/);
  return m ? Number(m[1] || 0) * 3600 + Number(m[2] || 0) * 60 + Number(m[3] || 0) : 0;
}

// Office の中では既定のブラウザで開く。使えない環境では新しいウィンドウで開く
function openInBrowser() {
  if (!config || !config.url) return;
  if (inOffice) {
    try {
      if (Office.context.requirements.isSetSupported("OpenBrowserWindowApi", "1.1")) {
        Office.context.ui.openBrowserWindow(config.url);
        return;
      }
    } catch {}
  }
  window.open(config.url, "_blank", "noopener");
}

function nextFrame() {
  return new Promise((resolve) => requestAnimationFrame(resolve));
}

// 枠の大きさが iframe の中へ伝わるまで待つ。すぐ読み込むと中の innerWidth が 0 のまま始まる
function settle() {
  const frames = nextFrame().then(nextFrame).then(nextFrame);
  return Promise.race([frames, new Promise((resolve) => setTimeout(resolve, 300))]);
}

async function render() {
  const stage = $("stage");
  stage.textContent = "";
  const cfg = (shown = config);
  const frame = document.createElement("iframe");
  const flags = [
    "allow-scripts",
    "allow-forms",
    "allow-modals",
    "allow-popups",
    "allow-pointer-lock",
    "allow-downloads",
    "allow-presentation",
  ];
  // URL のページは別オリジンなので本体には触れられない。HTML を書くときは指定があるときだけ許す
  if (cfg.mode === "url" || cfg.sameOrigin) flags.push("allow-same-origin");
  frame.setAttribute("sandbox", flags.join(" "));
  frame.setAttribute("allow", "fullscreen; autoplay; clipboard-write; gamepad; xr-spatial-tracking");
  stage.appendChild(frame);
  layout();
  await settle();
  if (cfg.mode === "url") frame.src = embedUrl(cfg.url);
  else frame.srcdoc = withPrelude(cfg.html, cfg);
}

// 「HTML を書く」の HTML の先頭に差し込み、利用者のスクリプトより先に動かすスクリプト
function prelude(cfg) {
  let js = "";
  if (cfg.pixelRatio > 0) {
    // キャンバスや WebGL の大きさを devicePixelRatio で決める HTML なら、描くピクセル数が減る
    js += `try { Object.defineProperty(window, "devicePixelRatio", { get: function () { return ${cfg.pixelRatio}; }, configurable: true }); } catch (e) {}\n`;
  }
  if (cfg.showFps) js += FPS_SCRIPT;
  return js ? `<script>\n${js}</script>` : "";
}

function withPrelude(html, cfg) {
  const head = prelude(cfg);
  if (!head) return html;
  // <!doctype> より前に置くと互換モードになるので、doctype のすぐ後ろに入れる
  const doctype = html.match(/^\s*<!doctype[^>]*>/i);
  return doctype ? doctype[0] + head + html.slice(doctype[0].length) : head + html;
}

// 基準の幅があれば、その幅で描いた画面を枠の幅に合わせて拡大縮小する
function layout() {
  const frame = $("stage").firstElementChild;
  if (!frame || !shown) return;
  const w = shown.designWidth;
  if (w > 0) {
    const scale = window.innerWidth / w;
    frame.style.width = w + "px";
    frame.style.height = window.innerHeight / scale + "px";
    frame.style.transform = `scale(${scale})`;
  } else {
    frame.style.width = "100%";
    frame.style.height = "100%";
    frame.style.transform = "";
  }
}

function status(text) {
  $("status").textContent = text;
}

function syncPanes() {
  const mode = form.elements.mode.value;
  for (const pane of form.querySelectorAll(".pane")) pane.hidden = pane.dataset.mode !== mode;
  for (const el of form.querySelectorAll(".html-only")) el.hidden = mode !== "html";
}

async function loadSample() {
  const res = await fetch("sample.html", { cache: "no-store" });
  $("html").value = await res.text();
}

async function openEditor() {
  const c = config || DEFAULTS;
  form.elements.mode.value = c.mode;
  $("html").value = c.html;
  $("url").value = c.url;
  $("design-width").value = c.designWidth;
  $("pixel-ratio").value = String(c.pixelRatio);
  $("show-fps").checked = c.showFps;
  $("same-origin").checked = c.sameOrigin;
  $("dpr-now").textContent = String(Math.round(window.devicePixelRatio * 100) / 100);
  $("btn-cancel").hidden = !config;
  status("");
  syncPanes();
  editing = true;
  refresh();
  if (!config) {
    try {
      await loadSample();
    } catch {}
  }
}

async function save(e) {
  e.preventDefault();
  const next = {
    mode: form.elements.mode.value,
    html: $("html").value,
    url: $("url").value.trim(),
    designWidth: Math.max(0, Math.round(Number($("design-width").value) || 0)),
    pixelRatio: Number($("pixel-ratio").value) || 0,
    showFps: $("show-fps").checked,
    sameOrigin: $("same-origin").checked,
  };
  if (next.mode === "url" && !/^https?:\/\//i.test(next.url)) {
    status("https:// で始まる URL を入れてください");
    return;
  }
  if (next.mode === "html" && !next.html.trim()) {
    status("HTML が空です");
    return;
  }
  status("保存中…");
  try {
    await store.save(next);
  } catch (err) {
    status("保存できませんでした: " + ((err && err.message) || err));
    return;
  }
  config = next;
  editing = false;
  status("");
  refresh();
}

function start(office) {
  inOffice = office;
  store = inOffice ? officeStore() : browserStore();
  const saved = store.load();
  config = saved ? { ...DEFAULTS, ...saved } : null;

  form.addEventListener("submit", save);
  form.addEventListener("change", (e) => {
    if (e.target.name === "mode") syncPanes();
  });
  $("btn-cancel").addEventListener("click", () => {
    editing = false;
    refresh();
  });
  $("btn-sample").addEventListener("click", () => loadSample().catch(() => status("サンプルを読み込めませんでした")));
  $("btn-file").addEventListener("click", () => $("file").click());
  $("file").addEventListener("change", async () => {
    const file = $("file").files[0];
    if (!file) return;
    $("html").value = await file.text();
    $("file").value = "";
    status(`${file.name} を読み込みました（同じフォルダの画像や JS は読み込まれません）`);
  });
  // Tab キーで字下げできるようにする
  $("html").addEventListener("keydown", (e) => {
    if (e.key !== "Tab" || e.shiftKey || e.ctrlKey || e.altKey) return;
    e.preventDefault();
    document.execCommand("insertText", false, "  ");
  });
  $("btn-edit").addEventListener("click", openEditor);
  $("btn-reload").addEventListener("click", render);
  $("btn-open").addEventListener("click", openInBrowser);
  window.addEventListener("resize", layout);

  if (inOffice) {
    const doc = Office.context.document;
    doc.getActiveViewAsync((r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) setView(r.value);
    });
    doc.addHandlerAsync(Office.EventType.ActiveViewChanged, (e) => setView(e.activeView));
  } else {
    view = new URLSearchParams(location.search).get("view") === "read" ? "read" : "edit";
  }

  if (config) refresh();
  else openEditor();
}

if (window.Office) Office.onReady((info) => start(!!info.host));
else start(false);
