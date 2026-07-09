const http = require("http");
const fs = require("fs");
const path = require("path");

const ROOT = __dirname;
const PORT = process.env.PORT || 3000;
const SITE_PASSWORD = process.env.SITE_PASSWORD || "";

const MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".png": "image/png",
  ".svg": "image/svg+xml",
};

const LOGIN_HTML = `<!doctype html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>企业微信看板 — 登录</title>
<style>
  * { box-sizing:border-box; }
  body { margin:0; display:flex; align-items:center; justify-content:center; min-height:100vh; background:#f6f8fb; font-family:-apple-system,"Microsoft YaHei",sans-serif; }
  .box { background:#fff; border-radius:14px; box-shadow:0 10px 30px rgba(15,23,42,.08); padding:32px 28px; width:360px; max-width:90vw; text-align:center; }
  h1 { margin:0 0 6px; font-size:20px; color:#172033; }
  p { margin:0 0 20px; font-size:13px; color:#64748b; }
  input { width:100%; padding:10px 14px; border:1px solid #d9e0ea; border-radius:8px; font-size:15px; text-align:center; margin-bottom:14px; }
  button { width:100%; padding:10px; border:0; border-radius:8px; background:#6d8f3c; color:#fff; font-size:15px; font-weight:600; cursor:pointer; }
  button:hover { opacity:.9; }
  #err { color:#b91c1c; font-size:13px; margin-top:8px; }
</style>
</head>
<body>
<div class="box">
  <h1>企业微信转单看板</h1>
  <p>请输入密码查看数据</p>
  <input type="password" id="pwd" placeholder="输入密码" autofocus>
  <button onclick="login()">进入看板</button>
  <div id="err"></div>
</div>
<script>
async function login() {
  var p = document.getElementById("pwd").value;
  if (!p) return;
  var r = await fetch("/api/auth", { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify({password:p}) });
  var d = await r.json();
  if (d.ok) { window.location.href = "/"; }
  else { document.getElementById("err").textContent = d.error || "密码错误"; }
}
document.getElementById("pwd").addEventListener("keydown", function(e) { if (e.key==="Enter") login(); });
</script>
</body>
</html>`;

function getCookie(cookieHeader, name) {
  if (!cookieHeader) return "";
  const match = cookieHeader.match(new RegExp("(?:^|;\\s*)" + name + "=([^;]*)"));
  return match ? decodeURIComponent(match[1]) : "";
}

function serveStatic(res, filePath) {
  const ext = path.extname(filePath).toLowerCase();
  fs.readFile(filePath, (error, content) => {
    if (error) {
      res.writeHead(error.code === "ENOENT" ? 404 : 500, { "Content-Type": "text/plain; charset=utf-8" });
      res.end(error.code === "ENOENT" ? "未找到" : "服务器错误");
      return;
    }
    res.writeHead(200, { "Content-Type": MIME_TYPES[ext] || "application/octet-stream" });
    res.end(content);
  });
}

const server = http.createServer((req, res) => {
  const urlPath = decodeURIComponent(req.url.split("?")[0]);

  // GET /api/auth — 检查认证状态
  if (urlPath === "/api/auth" && req.method === "GET") {
    const authenticated = !SITE_PASSWORD || getCookie(req.headers.cookie, "wecom_auth") === "1";
    res.writeHead(200, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ authenticated }));
    return;
  }

  // POST /api/auth — 密码验证
  if (urlPath === "/api/auth" && req.method === "POST") {
    let body = "";
    req.on("data", (chunk) => { body += chunk; });
    req.on("end", () => {
      try {
        const { password } = JSON.parse(body);
        if (SITE_PASSWORD && password === SITE_PASSWORD) {
          res.writeHead(200, {
            "Content-Type": "application/json",
            "Set-Cookie": "wecom_auth=1; Path=/; Max-Age=86400; HttpOnly; SameSite=Lax",
          });
          res.end(JSON.stringify({ success: true }));
          return;
        }
      } catch (e) {}
      res.writeHead(401, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ success: false, error: "密码错误" }));
    });
    return;
  }

  // index.html 始终放行（认证由客户端处理）
  if (urlPath === "/" || urlPath === "/index.html") {
    serveStatic(res, path.join(ROOT, "index.html"));
    return;
  }

  // 其他资源：检查 auth cookie
  if (SITE_PASSWORD && getCookie(req.headers.cookie, "wecom_auth") !== "1") {
    res.writeHead(401, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ error: "Unauthorized" }));
    return;
  }

  // 已认证 — 提供静态文件
  const relativePath = urlPath === "/" ? "index.html" : urlPath.replace(/^\/+/, "");
  const filePath = path.normalize(path.join(ROOT, relativePath));
  if (!filePath.startsWith(ROOT)) {
    res.writeHead(403);
    res.end("Forbidden");
    return;
  }
  serveStatic(res, filePath);
});

server.listen(PORT, () => {
  console.log(`企业微信看板已启动: http://0.0.0.0:${PORT}/`);
});
