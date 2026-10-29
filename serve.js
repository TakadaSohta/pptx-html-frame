// docs/ を https://localhost:3000 で配信する開発用サーバー。
// `node serve.js --http` ならブラウザ確認用に http で配信する（PowerPoint は https でないと読み込まない）。
const fs = require("fs");
const http = require("http");
const https = require("https");
const os = require("os");
const path = require("path");

const PORT = Number(process.env.npm_package_config_dev_server_port) || 3000;
const ROOT = path.join(__dirname, "docs");
const TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
};

function handler(req, res) {
  let urlPath;
  try {
    urlPath = decodeURIComponent(new URL(req.url, "http://localhost").pathname);
  } catch {
    res.writeHead(400);
    return res.end();
  }
  let file = path.normalize(path.join(ROOT, urlPath));
  if (file !== ROOT && !file.startsWith(ROOT + path.sep)) {
    res.writeHead(403);
    return res.end();
  }
  if (urlPath.endsWith("/")) file = path.join(file, "index.html");
  fs.readFile(file, (err, data) => {
    if (err) {
      res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
      return res.end("Not found");
    }
    res.writeHead(200, {
      "Content-Type": TYPES[path.extname(file)] || "application/octet-stream",
      "Cache-Control": "no-store",
    });
    res.end(data);
  });
}

if (process.argv.includes("--http")) {
  http.createServer(handler).listen(PORT, () => {
    console.log(`http://localhost:${PORT}/ で配信中（ブラウザ確認用）`);
  });
} else {
  const dir = path.join(os.homedir(), ".office-addin-dev-certs");
  const cert = path.join(dir, "localhost.crt");
  const key = path.join(dir, "localhost.key");
  if (!fs.existsSync(cert) || !fs.existsSync(key)) {
    console.error("開発用の証明書がありません。先に `npm run certs` を実行してください。");
    process.exit(1);
  }
  const options = { cert: fs.readFileSync(cert), key: fs.readFileSync(key) };
  https.createServer(options, handler).listen(PORT, () => {
    console.log(`https://localhost:${PORT}/ で配信中`);
  });
}
