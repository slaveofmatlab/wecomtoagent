const http = require("http");
const https = require("https");
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const { buildPageData, findFile } = require("./scripts/lib/page_logic");

const ROOT = __dirname;
const PORT = process.env.PORT || 3000;
const SITE_PASSWORD = process.env.SITE_PASSWORD || "";
const GITHUB_TOKEN = process.env.GITHUB_TOKEN || "";
const GITHUB_REPO = "slaveofmatlab/wecomtoagent";
const GITHUB_BRANCH = "master";

const MIME_TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  ".png": "image/png",
  ".svg": "image/svg+xml",
};

// ====== 内存状态 ======
let currentPageData = null;  // 最新一次上传/加载的 page_data
let currentTrends = null;    // trends.json 内容（日期 → 汇总）
let githubShas = {};         // { filename: sha } 用于 GitHub PUT

// ====== GitHub 工具 ======
function githubRequest(method, apiPath, body) {
  return new Promise((resolve, reject) => {
    const payload = body ? JSON.stringify(body) : null;
    const opts = {
      hostname: "api.github.com",
      path: apiPath,
      method,
      headers: {
        Authorization: "Bearer " + GITHUB_TOKEN,
        Accept: "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "wecomtoagent-server",
        ...(payload ? { "Content-Type": "application/json", "Content-Length": Buffer.byteLength(payload) } : {}),
      },
    };
    const req = https.request(opts, (res) => {
      let buf = "";
      res.on("data", (c) => (buf += c));
      res.on("end", () => {
        try { resolve({ status: res.statusCode, data: JSON.parse(buf) }); }
        catch (e) { resolve({ status: res.statusCode, data: buf }); }
      });
    });
    req.on("error", reject);
    if (payload) req.write(payload);
    req.end();
  });
}

async function githubGetFile(filename) {
  const res = await githubRequest("GET", `/repos/${GITHUB_REPO}/contents/data/${filename}?ref=${GITHUB_BRANCH}`);
  if (res.status === 404) return null;
  if (res.status !== 200) throw new Error(`GitHub GET ${filename}: HTTP ${res.status}`);
  return {
    data: JSON.parse(Buffer.from(res.data.content, "base64").toString("utf8")),
    sha: res.data.sha,
  };
}

async function githubPutFile(filename, data, sha) {
  const content = Buffer.from(JSON.stringify(data, null, 2)).toString("base64");
  const body = {
    message: `data: update ${filename} [skip render]`,
    content,
    branch: GITHUB_BRANCH,
    ...(sha ? { sha } : {}),
  };
  const res = await githubRequest("PUT", `/repos/${GITHUB_REPO}/contents/data/${filename}`, body);
  if (res.status !== 200 && res.status !== 201) {
    throw new Error(`GitHub PUT ${filename}: HTTP ${res.status} — ${JSON.stringify(res.data)}`);
  }
  return res.data.content.sha;
}

// ====== 启动时初始化 ======
async function init() {
  if (GITHUB_TOKEN) {
    console.log("从 GitHub 加载数据…");
    try {
      const pd = await githubGetFile("page_data.json");
      if (pd) { currentPageData = pd.data; githubShas["page_data.json"] = pd.sha; console.log("  page_data.json 已加载"); }
      const tr = await githubGetFile("trends.json");
      if (tr) { currentTrends = tr.data; githubShas["trends.json"] = tr.sha; console.log("  trends.json 已加载"); }
    } catch (e) {
      console.error("GitHub 加载失败，回退到磁盘:", e.message);
    }
  }
  // 磁盘兜底（本地开发 / GitHub 加载失败）
  if (!currentPageData) {
    const p = path.join(ROOT, "data", "page_data.json");
    if (fs.existsSync(p)) { try { currentPageData = JSON.parse(fs.readFileSync(p, "utf8")); console.log("  page_data.json 从磁盘加载"); } catch (e) {} }
  }
  if (!currentTrends) {
    const p = path.join(ROOT, "data", "trends.json");
    if (fs.existsSync(p)) { try { currentTrends = JSON.parse(fs.readFileSync(p, "utf8")); console.log("  trends.json 从磁盘加载"); } catch (e) {} }
  }
}

// ====== 静态文件 ======
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

function isAuthed(req) {
  return !SITE_PASSWORD || getCookie(req.headers.cookie, "wecom_auth") === "1";
}

function unauthorized(res) {
  res.writeHead(401, { "Content-Type": "application/json" });
  res.end(JSON.stringify({ error: "Unauthorized" }));
}

// ====== 上传处理 ======
async function handleUpload(req, res) {
  let body = Buffer.alloc(0);
  let size = 0;
  const MAX = 25 * 1024 * 1024; // 25MB

  await new Promise((resolve, reject) => {
    req.on("data", (chunk) => {
      size += chunk.length;
      if (size > MAX) { req.destroy(); reject(new Error("文件过大（上限 25MB）")); return; }
      body = Buffer.concat([body, chunk]);
    });
    req.on("end", resolve);
    req.on("error", reject);
  });

  const { sales, pending, progress, log, cutoff } = JSON.parse(body.toString("utf8"));
  if (!sales || !pending) throw new Error("缺少销售订单或待转单文件");
  if (!cutoff || cutoff.length !== 4) throw new Error("截止日期格式错误，应为 4 位 MMDD");

  const salesWb = XLSX.read(Buffer.from(sales, "base64"), { type: "buffer" });
  const pendingWb = XLSX.read(Buffer.from(pending, "base64"), { type: "buffer" });
  const progressWb = progress
    ? XLSX.read(Buffer.from(progress, "base64"), { type: "buffer" })
    : loadFallbackWorkbook("basicData", "企业微信AI转单推进表");
  const logWb = log ? XLSX.read(Buffer.from(log, "base64"), { type: "buffer" }) : null;

  if (!progressWb) throw new Error("缺少推进表文件");

  const data = buildPageData({
    salesWorkbook: salesWb,
    pendingWorkbook: pendingWb,
    progressWorkbook: progressWb,
    logWorkbook: logWb,
    cutoffDate: cutoff,
    sources: { salesPath: null, pendingPath: null, progressPath: null },
  });

  // 没有上传日志时，保留上一次的备注（备注来自消息日志，不随订单每天变）
  if (!logWb && currentPageData && currentPageData.companySummary) {
    const oldRemarks = {};
    (currentPageData.companySummary.rows || []).forEach((r) => {
      if (r.orderMethod && r.operationCompany) oldRemarks[r.operationCompany] = r.orderMethod;
    });
    data.companySummary.rows.forEach((r) => {
      if (!r.orderMethod && oldRemarks[r.operationCompany]) r.orderMethod = oldRemarks[r.operationCompany];
    });
  }

  // 更新内存
  currentPageData = data;
  if (!currentTrends) currentTrends = {};
  const key = cutoff.slice(0, 2) + "-" + cutoff.slice(2, 4);
  const t = data.companySummary.totals;
  currentTrends[key] = {
    cutoff,
    registered: t.registeredCount,
    itOk: t.itConfiguredCount,
    configRate: t.configRate,
    orderTotal: t.orderTotal,
    orderAi: t.orderAiCount,
    aiRate: t.aiRate,
    companies: data.companySummary.rows.map((r) => ({
      operationCompany: r.operationCompany,
      orderTotal: r.orderTotal,
      orderAiCount: r.orderAiCount,
      aiRate: r.aiRate,
    })),
    ts: new Date().toISOString(),
  };

  // 写磁盘
  const dataDir = path.join(ROOT, "data");
  fs.mkdirSync(dataDir, { recursive: true });
  fs.writeFileSync(path.join(dataDir, "page_data.json"), JSON.stringify(data, null, 2));
  fs.writeFileSync(path.join(dataDir, "trends.json"), JSON.stringify(currentTrends, null, 2));

  // 同步推 GitHub（等待完成，确保 Render 重启后数据不丢失）
  // 只推渲染所需字段，去掉 salesRows/pendingRows/progressRows 大数组，避免超 GitHub 1MB 限制
  let githubWarning = null;
  if (GITHUB_TOKEN) {
    const slimData = {
      generatedAt: data.generatedAt,
      cutoffDate: data.cutoffDate,
      sources: data.sources,
      pendingTotals: data.pendingTotals,
      companySummary: data.companySummary,
      logStats: data.logStats,
    };
    let ghErrors = [];
    const [pdSha, trSha] = await Promise.all([
      githubPutFile("page_data.json", slimData, githubShas["page_data.json"])
        .then((sha) => { githubShas["page_data.json"] = sha; return sha; })
        .catch((e) => { ghErrors.push("page_data: " + e.message); console.error("GitHub push page_data.json:", e.message); return null; }),
      githubPutFile("trends.json", currentTrends, githubShas["trends.json"])
        .then((sha) => { githubShas["trends.json"] = sha; return sha; })
        .catch((e) => { ghErrors.push("trends: " + e.message); console.error("GitHub push trends.json:", e.message); return null; }),
    ]);
    if (!pdSha || !trSha) {
      githubWarning = "GitHub 备份失败（" + ghErrors.join("；") + "）";
    }
  }

  return { success: true, data, trends: currentTrends, warning: githubWarning };
}

function loadFallbackWorkbook(dir, keyword) {
  const dirPath = path.join(ROOT, dir);
  if (!fs.existsSync(dirPath)) return null;
  const file = findFile(dirPath, keyword);
  return file ? XLSX.readFile(file) : null;
}

// ====== HTTP 服务器 ======
const server = http.createServer((req, res) => {
  const urlPath = decodeURIComponent(req.url.split("?")[0]);

  // GET /api/auth
  if (urlPath === "/api/auth" && req.method === "GET") {
    res.writeHead(200, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ authenticated: isAuthed(req) }));
    return;
  }

  // POST /api/auth
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

  // GET /api/data — 内存优先，磁盘兜底
  if (urlPath === "/api/data" && req.method === "GET") {
    if (!isAuthed(req)) { unauthorized(res); return; }
    const urlParams = new URLSearchParams(req.url.split("?")[1] || "");
    const fileName = urlParams.get("file") || "page_data.json";
    if (fileName.includes("..") || fileName.includes("/")) {
      res.writeHead(400, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ error: "Invalid file" }));
      return;
    }
    let memData = null;
    if (fileName === "page_data.json") memData = currentPageData;
    else if (fileName === "trends.json") memData = currentTrends;
    if (memData) {
      res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
      res.end(JSON.stringify(memData));
      return;
    }
    serveStatic(res, path.join(ROOT, "data", fileName));
    return;
  }

  // POST /api/upload
  if (urlPath === "/api/upload" && req.method === "POST") {
    if (!isAuthed(req)) { unauthorized(res); return; }
    handleUpload(req, res)
      .then((result) => {
        res.writeHead(200, { "Content-Type": "application/json" });
        res.end(JSON.stringify(result));
      })
      .catch((e) => {
        console.error("Upload error:", e.message);
        res.writeHead(500, { "Content-Type": "application/json" });
        res.end(JSON.stringify({ error: e.message }));
      });
    return;
  }

  // index.html 不需要认证（认证由客户端处理）
  if (urlPath === "/" || urlPath === "/index.html") {
    serveStatic(res, path.join(ROOT, "index.html"));
    return;
  }

  // vendor 静态库不含业务数据，无需认证
  if (urlPath.startsWith("/vendor/")) {
    const vPath = path.normalize(path.join(ROOT, urlPath.replace(/^\/+/, "")));
    if (vPath.startsWith(ROOT)) { serveStatic(res, vPath); return; }
  }

  // 其他资源需要认证
  if (!isAuthed(req)) { unauthorized(res); return; }

  const relativePath = urlPath.replace(/^\/+/, "");
  const filePath = path.normalize(path.join(ROOT, relativePath));
  if (!filePath.startsWith(ROOT)) { res.writeHead(403); res.end("Forbidden"); return; }
  serveStatic(res, filePath);
});

init().then(() => {
  server.listen(PORT, () => {
    console.log(`企业微信看板已启动: http://0.0.0.0:${PORT}/`);
  });
});
