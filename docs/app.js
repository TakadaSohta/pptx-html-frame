// HTML Frame: スライドに置いた枠の中で HTML や Web ページを表示する
"use strict";

const KEY = "htmlFrame";
const DEFAULTS = { mode: "html", html: "", url: "", designWidth: 0, sameOrigin: false };

const $ = (id) => document.getElementById(id);
const form = $("editor");

let store;
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
  $("toolbar").hidden = view !== "edit";
  if (!$("viewer").hidden && shown !== config) render();
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
  if (cfg.mode === "url") frame.src = cfg.url;
  else frame.srcdoc = cfg.html;
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
  $("same-origin").checked = c.sameOrigin;
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

function start(inOffice) {
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
