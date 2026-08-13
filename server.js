const http = require("http");
const https = require("https");
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const Busboy = require("busboy");
const { buildPageData, findFile, buildOrderMethodReport, classifyGroupCategory, parsePendingWecom, parseWecomProgress, parseSalesFull, normalizeText } = require("./scripts/lib/page_logic");
const ExcelJS = require("exceljs");

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
let storedSalesByDate = {};  // { "0804": { salesRows, allItOkCodes, codeToCompany }, ... } 历史销售快照
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

async function githubGetSha(filename) {
  const res = await githubRequest("GET", `/repos/${GITHUB_REPO}/contents/data/${filename}?ref=${GITHUB_BRANCH}`);
  if (res.status === 404) return null; // 文件不存在，首次创建不需要 SHA
  if (res.status !== 200) {
    throw new Error(`GET SHA for ${filename} failed: HTTP ${res.status} — ${JSON.stringify(res.data).slice(0, 300)}`);
  }
  if (typeof res.data !== "object" || !res.data.sha) {
    throw new Error(`GET SHA for ${filename}: unexpected response (keys: ${typeof res.data === "object" ? Object.keys(res.data).join(",") : typeof res.data})`);
  }
  return res.data.sha;
}

async function githubGetFile(filename) {
  const res = await githubRequest("GET", `/repos/${GITHUB_REPO}/contents/data/${filename}?ref=${GITHUB_BRANCH}`);
  if (res.status === 404) return null;
  if (res.status !== 200) throw new Error(`GitHub GET ${filename}: HTTP ${res.status}`);
  let data = null;
  try {
    const raw = Buffer.from(res.data.content || "", "base64").toString("utf8");
    if (raw) data = JSON.parse(raw);
  } catch (e) {
    console.log(`  ${filename}: 内容过大无法内联读取，仅启动时磁盘数据有效`);
  }
  return { data, sha: res.data.sha };
}

async function githubPutFile(filename, data) {
  const content = Buffer.from(JSON.stringify(data, null, 2)).toString("base64");
  async function attempt() {
    const sha = await githubGetSha(filename);
    const body = {
      message: `data: update ${filename} [skip render]`,
      content,
      branch: GITHUB_BRANCH,
      ...(sha ? { sha } : {}),
    };
    return githubRequest("PUT", `/repos/${GITHUB_REPO}/contents/data/${filename}`, body);
  }
  let res = await attempt();
  // 409 = SHA conflict（并发写或 git push 改了 blob SHA），重取 SHA 重试一次
  if (res.status === 409) {
    console.log(`  ${filename}: 409 SHA conflict，重取 SHA 重试…`);
    res = await attempt();
  }
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
      if (pd && pd.data) { currentPageData = pd.data; console.log("  page_data.json 已加载"); }
      const tr = await githubGetFile("trends.json");
      if (tr && tr.data) { currentTrends = tr.data; console.log("  trends.json 已加载"); }
      const cp = await githubGetFile("pending_cumulative.json");
      if (cp && cp.data) { currentCumulativePending = cp.data; console.log("  pending_cumulative.json 已加载"); }
      const ss = await githubGetFile("sales_snapshots.json");
      if (ss && ss.data) { storedSalesByDate = ss.data; console.log("  sales_snapshots.json 已加载 (" + Object.keys(storedSalesByDate).length + " 天)"); }
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
  // 累积待转单始终从磁盘加载（GitHub 已覆盖则跳过）
  if (!Object.keys(currentCumulativePending).length) {
    currentCumulativePending = loadCumulativePending();
    if (Object.keys(currentCumulativePending).length) console.log("  pending_cumulative.json 从磁盘加载");
  }
  // 销售快照从磁盘加载
  if (!Object.keys(storedSalesByDate).length) {
    storedSalesByDate = loadSalesSnapshots();
    if (Object.keys(storedSalesByDate).length) console.log("  sales_snapshots.json 从磁盘加载 (" + Object.keys(storedSalesByDate).length + " 天)");
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

// ====== 累积待转单（跨天去重，解决待转单在前、销售订单在后的时差误判）======
let currentCumulativePending = {};

function loadCumulativePending() {
  const p = path.join(ROOT, "data", "pending_cumulative.json");
  if (fs.existsSync(p)) {
    try { return JSON.parse(fs.readFileSync(p, "utf8")); } catch (e) {}
  }
  return {};
}

function saveCumulativePending(map) {
  // 紧凑格式（无缩进），减少磁盘 I/O 和 GitHub push 体积
  fs.writeFileSync(
    path.join(ROOT, "data", "pending_cumulative.json"),
    JSON.stringify(map)
  );
}

function loadSalesSnapshots() {
  const p = path.join(ROOT, "data", "sales_snapshots.json");
  if (fs.existsSync(p)) {
    try { return JSON.parse(fs.readFileSync(p, "utf8")); } catch (e) {}
  }
  return {};
}

function saveSalesSnapshots(map) {
  fs.writeFileSync(
    path.join(ROOT, "data", "sales_snapshots.json"),
    JSON.stringify(map)
  );
}

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
  // 使用 busboy 解析 multipart/form-data（支持大文件二进制上传，无 base64 膨胀）
  const { files, fields } = await new Promise((resolve, reject) => {
    const busboy = Busboy({ headers: req.headers, limits: { fileSize: 100 * 1024 * 1024, files: 4 } });
    const files = {};
    const fields = {};
    busboy.on("file", (fieldname, stream, info) => {
      const chunks = [];
      stream.on("data", (chunk) => chunks.push(chunk));
      stream.on("end", () => { files[fieldname] = Buffer.concat(chunks); });
      stream.on("error", reject);
    });
    busboy.on("field", (fieldname, val) => { fields[fieldname] = val; });
    busboy.on("finish", () => resolve({ files, fields }));
    busboy.on("error", reject);
    req.pipe(busboy);
  });

  const cutoff = fields.cutoff;
  if (!files.sales || !files.pending) throw new Error("缺少销售订单或待转单文件");
  if (!cutoff || cutoff.length !== 4) throw new Error("截止日期格式错误，应为 4 位 MMDD");

  const salesWb = XLSX.read(files.sales, { type: "buffer" });
  const pendingWb = XLSX.read(files.pending, { type: "buffer" });
  const progressWb = files.progress
    ? XLSX.read(files.progress, { type: "buffer" })
    : loadFallbackWorkbook("basicData", "企业微信AI转单推进表");
  const logWb = files.log ? XLSX.read(files.log, { type: "buffer" }) : null;

  if (!progressWb) throw new Error("缺少推进表文件");

  // 预解析（供快照存储 + 重算使用）
  const salesRows = parseSalesFull(salesWb);
  const pendingRows = parsePendingWecom(pendingWb);
  const progressRows = parseWecomProgress(progressWb, cutoff);

  // 解析当天待转单，合并入累积 Map（紧凑格式：{s: salesOrderNo, t: transferStatus, d: [dates]}）
  // key = customerOrderNo（如果非空），否则 '_' + salesOrderNo（兜底匹配用）
  const todayDate = cutoff;
  for (const row of pendingRows) {
    if (row.createdBy !== '供应链管理员') continue;
    var custKey = normalizeText(row.customerOrderNo);
    var salesKey = normalizeText(row.salesOrderNo);
    if (!custKey && !salesKey) continue;
    var mapKey = custKey || ('_' + salesKey);
    var entry = currentCumulativePending[mapKey];
    if (entry) {
      entry.s = row.salesOrderNo || '';
      entry.t = row.transferStatus || '';
      if (!entry.d) entry.d = [];
      if (!entry.d.includes(todayDate)) entry.d.push(todayDate);
    } else {
      currentCumulativePending[mapKey] = {
        s: row.salesOrderNo || '',
        t: row.transferStatus || '',
        d: [todayDate],
      };
    }
  }
  saveCumulativePending(currentCumulativePending);

  const data = buildPageData({
    salesWorkbook: salesWb,
    pendingWorkbook: pendingWb,
    progressWorkbook: progressWb,
    logWorkbook: logWb,
    cutoffDate: cutoff,
    sources: { salesPath: null, pendingPath: null, progressPath: null },
  });

  // 没有上传日志时，保留上一次的备注和下单方式分布（用于按下单形式导出）
  if (!logWb && currentPageData) {
    // 保留 orderMethod 备注
    if (currentPageData.companySummary) {
      const oldRemarks = {};
      (currentPageData.companySummary.rows || []).forEach((r) => {
        if (r.orderMethod && r.operationCompany) oldRemarks[r.operationCompany] = r.orderMethod;
      });
      data.companySummary.rows.forEach((r) => {
        if (!r.orderMethod && oldRemarks[r.operationCompany]) r.orderMethod = oldRemarks[r.operationCompany];
      });
    }
    // 保留 logSummary（下单方式分布），否则按下单形式导出全是 0
    if (!data.logSummary || Object.keys(data.logSummary).length === 0) {
      if (currentPageData.logSummary && Object.keys(currentPageData.logSummary).length > 0) {
        data.logSummary = currentPageData.logSummary;
      }
    }
  }

  // 更新内存
  currentPageData = data;
  if (!currentTrends) currentTrends = {};

  // ---- 存储销售快照 + 进度信息（用于后续日期重算） ----
  try {
  const allItOkCodes = [];
  const codeToCompany = {};
  let registeredCount = 0;
  const registeredSet = new Set();
  for (const pr of progressRows) {
    if (!pr.operationCompanyKey || !pr.hotelCode) continue;
    registeredSet.add(pr.hotelCode);
    if (pr.itConfigured) {
      allItOkCodes.push(pr.hotelCode);
      codeToCompany[pr.hotelCode] = { key: pr.operationCompanyKey, name: pr.operationCompany };
    }
  }
  registeredCount = registeredSet.size;

  storedSalesByDate[cutoff] = {
    sales: salesRows.map(function (sr) { return {
      h: sr.hotelCode,
      c: normalizeText(sr.customerOrderNo),
      o: normalizeText(sr.orderNo),
      k: sr.operationCompanyKey,
    }; }),
    allItOkCodes: allItOkCodes,
    codeToCompany: codeToCompany,
    registeredCount: registeredCount,
    itOkCount: allItOkCodes.length,
    companies: data.companySummary.rows.map(function (r) { return {
      operationCompany: r.operationCompany,
      orderTotal: r.orderTotal,
      orderAiCount: r.orderAiCount,
      aiRate: r.aiRate,
    }; }),
  };

  // ---- 用累积待转单池重算所有历史日期 ----
  var MINI_PROGRAM = {};
  MINI_PROGRAM["KHXMD10235"] = true;

  // 构建匹配索引（从累积待转单池，而不是当天上传的待转单）
  // 根因修复：当天待转单文件可能是"增量/精简"的（已转单会被归档移出），缺少历史日期的
  // "已转"记录，导致重算历史日期时 orderAi 被错误覆盖成小数字。累积池跨天累积了所有
  // 历史待转单（pending_cumulative.json），用它匹配才能保证历史 orderAi 稳定不丢。
  var matchByCust = new Map();
  var matchBySales = new Map();
  for (var ck in currentCumulativePending) {
    var ce = currentCumulativePending[ck];
    if (!ce) continue;
    var status = ce.t || '';
    if (ck.charAt(0) === '_') {
      // key 是 '_销售订单号'（客户订单号为空的兜底）
      var sn = ck.slice(1);
      if (sn && !matchBySales.has(sn)) matchBySales.set(sn, status);
    } else {
      // key 是客户订单号
      if (ck && !matchByCust.has(ck)) matchByCust.set(ck, status);
    }
    // 销售订单号兜底索引（无论 key 类型都建立）
    if (ce.s && !matchBySales.has(ce.s)) matchBySales.set(ce.s, status);
  }

  var dateKeys = Object.keys(storedSalesByDate).sort();
  for (var di = 0; di < dateKeys.length; di++) {
    var dk = dateKeys[di];
    var stored = storedSalesByDate[dk];
    if (!stored || !stored.allItOkCodes || !stored.sales) { console.log("  跳过 " + dk + ": 数据格式不兼容"); continue; }
    var aitOk = {};
    for (var ai = 0; ai < stored.allItOkCodes.length; ai++) {
      aitOk[stored.allItOkCodes[ai]] = true;
    }

    var total = 0, aiDone = 0;
    var companyAi = {};  // 公司显示名 → AI 已转行数（与顶层同口径，保证明细和总计一致）
    for (var si = 0; si < stored.sales.length; si++) {
      var sr = stored.sales[si];
      if (!sr.h || !aitOk[sr.h]) continue;
      total++;
      var isAi = false;
      if (MINI_PROGRAM[sr.h]) { isAi = true; }
      else {
        var ps = matchByCust.get(sr.c) || matchBySales.get(sr.o);
        if (ps && ps.indexOf("已转") >= 0) isAi = true;
      }
      if (isAi) {
        aiDone++;
        var owner = stored.codeToCompany && stored.codeToCompany[sr.h];
        if (owner) companyAi[owner.name] = (companyAi[owner.name] || 0) + 1;
      }
    }

    var trendKey = dk.slice(0, 2) + "-" + dk.slice(2, 4);
    if (!currentTrends[trendKey]) {
      currentTrends[trendKey] = {
        cutoff: dk,
        registered: stored.registeredCount,
        itOk: stored.itOkCount,
        configRate: stored.registeredCount > 0 ? stored.itOkCount / stored.registeredCount : null,
        orderTotal: total,
        orderAi: aiDone,
        aiRate: total > 0 ? aiDone / total : null,
        companies: stored.companies || [],
        ts: new Date().toISOString(),
      };
    } else {
      // 保留首次写入的 registered/itOk/configRate，重算匹配相关字段
      currentTrends[trendKey].orderTotal = total;
      currentTrends[trendKey].orderAi = aiDone;
      currentTrends[trendKey].aiRate = total > 0 ? aiDone / total : null;
    }
    // 重算 companies 明细的 AI 数（同样用累积池口径），保持顶层 orderAi 与明细之和一致
    if (currentTrends[trendKey].companies) {
      currentTrends[trendKey].companies.forEach(function (c) {
        c.orderAiCount = companyAi[c.operationCompany] || 0;
        c.aiRate = c.orderTotal > 0 ? c.orderAiCount / c.orderTotal : null;
      });
    }
    currentTrends[trendKey].ts = new Date().toISOString();
  }
  console.log("  已刷新 " + dateKeys.length + " 个日期的趋势数据");

  } catch (snapshotErr) {
    console.error("销售快照/重算失败（不影响当前上传）:", snapshotErr.message);
  }

  // 写磁盘
  const dataDir = path.join(ROOT, "data");
  fs.mkdirSync(dataDir, { recursive: true });
  fs.writeFileSync(path.join(dataDir, "page_data.json"), JSON.stringify(data, null, 2));
  fs.writeFileSync(path.join(dataDir, "trends.json"), JSON.stringify(currentTrends, null, 2));
  saveSalesSnapshots(storedSalesByDate);

  // 异步推 GitHub（不阻塞响应，避免 Render 超时）
  let githubWarning = null;
  if (GITHUB_TOKEN) {
    const slimData = {
      generatedAt: data.generatedAt,
      cutoffDate: data.cutoffDate,
      sources: data.sources,
      pendingTotals: data.pendingTotals,
      companySummary: data.companySummary,
      groupSummary: data.groupSummary,
      logStats: data.logStats,
      logSummary: data.logSummary || {},
    };
    const trendsCopy = JSON.parse(JSON.stringify(currentTrends));
    const snapshotsCopy = JSON.parse(JSON.stringify(storedSalesByDate));
    const pendingCopy = JSON.parse(JSON.stringify(currentCumulativePending));
    // 不 await，让 GitHub push 在后台执行
    (async function () {
      let ghErrors = [];
      await githubPutFile("page_data.json", slimData)
        .catch((e) => { ghErrors.push("page_data: " + e.message); console.error("GitHub push page_data.json:", e.message); });
      await githubPutFile("trends.json", trendsCopy)
        .catch((e) => { ghErrors.push("trends: " + e.message); console.error("GitHub push trends.json:", e.message); });
      await githubPutFile("sales_snapshots.json", snapshotsCopy)
        .catch((e) => { ghErrors.push("sales_snapshots: " + e.message); console.error("GitHub push sales_snapshots.json:", e.message); });
      await githubPutFile("pending_cumulative.json", pendingCopy)
        .catch((e) => { ghErrors.push("pending_cumulative: " + e.message); console.error("GitHub push pending_cumulative.json:", e.message); });
      if (ghErrors.length) console.error("GitHub 后台同步失败: " + ghErrors.join("；"));
    })();
  }

  // 返回瘦身版数据（去掉 salesRows/pendingRows/progressRows 大数组），避免响应过大
  const slimResult = {
    generatedAt: data.generatedAt,
    cutoffDate: data.cutoffDate,
    sources: data.sources,
    pendingTotals: data.pendingTotals,
    companySummary: data.companySummary,
    groupSummary: data.groupSummary,
    logStats: data.logStats,
    logSummary: data.logSummary || {},
  };
  return { success: true, data: slimResult, trends: currentTrends, warning: githubWarning };
}

function loadFallbackWorkbook(dir, keyword) {
  const dirPath = path.join(ROOT, dir);
  if (!fs.existsSync(dirPath)) return null;
  const file = findFile(dirPath, keyword);
  return file ? XLSX.readFile(file) : null;
}

// ====== 趋势数据删除 ======
async function handleDeleteTrend(dateKey, res) {
  if (!currentTrends || !currentTrends[dateKey]) {
    res.writeHead(404, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ error: "日期 " + dateKey + " 不存在" }));
    return;
  }
  delete currentTrends[dateKey];

  // 写磁盘
  const trendsPath = path.join(ROOT, "data", "trends.json");
  fs.writeFileSync(trendsPath, JSON.stringify(currentTrends, null, 2));
  console.log("已删除趋势数据: " + dateKey);

  // 同步 GitHub
  let githubWarning = null;
  if (GITHUB_TOKEN) {
    try {
      await githubPutFile("trends.json", currentTrends);
      console.log("  GitHub 已同步");
    } catch (e) {
      githubWarning = "GitHub 备份失败: " + e.message;
      console.error("  GitHub push trends.json:", e.message);
    }
  }

  res.writeHead(200, { "Content-Type": "application/json" });
  res.end(JSON.stringify({ success: true, date: dateKey, trends: currentTrends, warning: githubWarning }));
}

// ====== 累积待转单导出（仅下载，不清理） ======
async function handleExportPendingBackup(req, res) {
  const entries = Object.values(currentCumulativePending);
  if (entries.length === 0) {
    res.writeHead(200, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ success: true, message: "累积待转单为空，无数据可导出", count: 0 }));
    return;
  }

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("累积待转单备份");
  const headers = ["customerOrderNo", "salesOrderNo", "transferStatus", "createdBy", "operationCompany", "hotelName", "appearedDates"];

  // 表头
  headers.forEach(function (h, i) {
    var cell = ws.getCell(1, i + 1);
    cell.value = h;
    cell.font = { bold: true };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFAFECEB" } };
    cell.border = { top: { style: "thin" }, bottom: { style: "thin" }, left: { style: "thin" }, right: { style: "thin" } };
  });

  // 数据行
  // 数据行
  Object.entries(currentCumulativePending).forEach(function (kv, ri) {
    var custNo = kv[0];
    var entry = kv[1];
    // 兼容多种格式
    var row = entry.row || entry;
    var salesNo = entry.s || row.salesOrderNo || '';
    var status = entry.t || row.transferStatus || '';
    var createdBy = row.createdBy || '供应链管理员';
    var dates = entry.d || entry.dates || [];

    var rowData = [
      custNo,
      salesNo,
      status,
      createdBy,
      row.operationCompany || '',
      row.hotelName || '',
      Array.isArray(dates) ? dates.join(', ') : '',
    ];
    rowData.forEach(function (v, ci) { ws.getCell(ri + 2, ci + 1).value = v; });
  });

  ws.getColumn(1).width = 22;
  ws.getColumn(2).width = 20;
  ws.getColumn(3).width = 12;
  ws.getColumn(4).width = 14;
  ws.getColumn(5).width = 20;
  ws.getColumn(6).width = 30;
  ws.getColumn(7).width = 30;

  var now = new Date();
  var stamp = now.getFullYear() + String(now.getMonth() + 1).padStart(2, "0") + String(now.getDate()).padStart(2, "0");
  var backupName = "pending_backup_" + stamp + ".xlsx";
  var buf = await wb.xlsx.writeBuffer();

  res.writeHead(200, {
    "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "Content-Disposition": "attachment; filename=\"" + encodeURIComponent(backupName) + "\"",
  });
  res.end(Buffer.from(buf));
}

// ====== 累积待转单清理（按日期/天数选择性删除） ======
async function handleClearPending(req, res) {
  var body = "";
  await new Promise(function (resolve) {
    req.on("data", function (c) { body += c; });
    req.on("end", resolve);
  });

  var opts = {};
  try { if (body) opts = JSON.parse(body); } catch (e) {}

  var action = opts.action || "all";
  var entries = Object.values(currentCumulativePending);
  var beforeCount = entries.length;
  var removed = 0;

  if (action === "all") {
    // 全部清空
    removed = beforeCount;
    currentCumulativePending = {};
  } else if (action === "keepLast") {
    // 保留最近 N 天（如 keepLast=7 保留最近7天出现过待转单的条目）
    var keepDays = opts.days || 7;
    // 从 cutoff 日期往前算 N 天，生成应保留的日期集合
    // 需要当前 cutoff 日期，从 trends 或 pageData 取
    var currentCutoff = (currentPageData && currentPageData.cutoffDate) || "";
    if (currentCutoff.length !== 4) {
      res.writeHead(400, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ error: "无法确定当前截止日期，请先上传数据" }));
      return;
    }
    var month = parseInt(currentCutoff.slice(0, 2), 10);
    var day = parseInt(currentCutoff.slice(2, 4), 10);
    var keepDates = new Set();
    for (var i = 0; i < keepDays; i++) {
      var d = new Date(2026, month - 1, day - i);
      var ds = String(d.getMonth() + 1).padStart(2, "0") + String(d.getDate()).padStart(2, "0");
      keepDates.add(ds);
    }

    var newMap = {};
    for (var key in currentCumulativePending) {
      var entry = currentCumulativePending[key];
      var dates = entry.d || entry.dates || [];
      // 如果该条目在保留日期范围内出现过，保留
      var shouldKeep = dates.some(function (d) { return keepDates.has(d); });
      if (shouldKeep) {
        newMap[key] = entry;
      } else {
        removed++;
      }
    }
    currentCumulativePending = newMap;
  } else if (action === "dates") {
    // 删除指定日期的条目
    var removeDates = opts.dates || [];
    if (!Array.isArray(removeDates) || removeDates.length === 0) {
      res.writeHead(400, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ error: "请提供要删除的日期列表，格式: {action:'dates', dates:['0725','0726']}" }));
      return;
    }
    var removeSet = new Set(removeDates);

    var newMap = {};
    for (var key in currentCumulativePending) {
      var entry = currentCumulativePending[key];
      var dates = entry.d || entry.dates || [];
      // 过滤掉要删除的日期
      var remaining = dates.filter(function (d) { return !removeSet.has(d); });
      if (remaining.length > 0) {
        if (entry.d) { entry.d = remaining; }
        else { entry.dates = remaining; }
        newMap[key] = entry;
      } else {
        removed++;
      }
    }
    currentCumulativePending = newMap;
  } else {
    res.writeHead(400, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ error: "未知的 action: " + action + "，可选: all / keepLast / dates" }));
    return;
  }

  saveCumulativePending(currentCumulativePending);

  // GitHub 同步
  var githubWarning = null;
  if (GITHUB_TOKEN) {
    try {
      await githubPutFile("pending_cumulative.json", currentCumulativePending);
      console.log("  GitHub pending_cumulative.json 已同步");
    } catch (e) {
      githubWarning = "GitHub 同步失败: " + e.message;
      console.error("GitHub push pending_cumulative.json (clear):", e.message);
    }
  }

  var afterCount = Object.keys(currentCumulativePending).length;
  console.log("累积待转单已清理: " + beforeCount + " → " + afterCount + " (删除 " + removed + " 条)");

  res.writeHead(200, { "Content-Type": "application/json" });
  res.end(JSON.stringify({
    success: true,
    before: beforeCount,
    after: afterCount,
    removed: removed,
    warning: githubWarning,
  }));
}

// ====== 按下单形式导出 Excel ======
async function handleExportOrderMethod(req, res) {
  try {
    if (!currentPageData) {
      res.writeHead(400, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ error: "还没有数据，请先上传文件" }));
      return;
    }

    const { cutoffDate } = currentPageData;

    // ---- 解析 URL 参数：?categories=机器人,图片下单,Excel下单 ----
    const urlParams = new URLSearchParams(req.url.split("?")[1] || "");
    const categoriesParam = (urlParams.get("categories") || "").trim();
    const selectedCategories = categoriesParam
      ? categoriesParam.split(",").map(function (s) { return s.trim(); }).filter(Boolean)
      : [];

    // ---- 加载 priority_groups.json ----
    var pgGroups = {};
    try {
      var pgPath = path.join(ROOT, "basicData", "priority_groups.json");
      if (fs.existsSync(pgPath)) {
        pgGroups = (JSON.parse(fs.readFileSync(pgPath, "utf8")).groups) || {};
      }
    } catch (e) { console.error("加载 priority_groups.json 失败:", e.message); }

    // ---- 从群维度汇总按分类聚合 ----
    var gsRows = (currentPageData.groupSummary && currentPageData.groupSummary.rows) || [];
    var ALL_CATEGORIES = ["机器人", "手写", "图片下单", "混合", "Excel下单", "PDF下单", "文本消息", "小程序", "其他"];
    var categories = selectedCategories.length > 0 ? selectedCategories : ALL_CATEGORIES;

    // categoryOrderLines[category] = { orderLines, aiLines }
    var catTotals = {};
    categories.forEach(function (c) { catTotals[c] = { orderLines: 0, aiLines: 0 }; });

    // companyKey → { name, cat → { orderLines, aiLines } }
    var companyData = {};

    for (var gi = 0; gi < gsRows.length; gi++) {
      var gr = gsRows[gi];
      var cat = classifyGroupCategory(gr.groupName, pgGroups);
      if (cat === null) cat = "混合"; // 不在配置里的群归入混合
      if (categories.indexOf(cat) < 0) continue; // 不在选中分类里，跳过

      var ck = gr.operationCompanyKey;
      if (!ck) continue;

      if (!companyData[ck]) {
        companyData[ck] = { name: gr.operationCompany, cats: {} };
        categories.forEach(function (c) { companyData[ck].cats[c] = { orderLines: 0, aiLines: 0 }; });
      }

      companyData[ck].cats[cat].orderLines += gr.orderTotal;
      companyData[ck].cats[cat].aiLines += (gr.orderAiCount || 0);

      catTotals[cat].orderLines += gr.orderTotal;
      catTotals[cat].aiLines += (gr.orderAiCount || 0);
    }

    // ---- 组装左表（按分类） ----
    var grandOrder = 0, grandAi = 0;
    var leftRows = categories.map(function (cat) {
      var t = catTotals[cat];
      grandOrder += t.orderLines;
      grandAi += t.aiLines;
      return {
        method: cat,
        orderLines: t.orderLines,
        aiLines: t.aiLines,
        aiRate: t.orderLines > 0 ? Math.round(t.aiLines / t.orderLines * 100) + "%" : "0%",
      };
    });
    var leftTotal = {
      method: "总计",
      orderLines: grandOrder,
      aiLines: grandAi,
      aiRate: grandOrder > 0 ? Math.round(grandAi / grandOrder * 100) + "%" : "0%",
    };

    // ---- 组装右表（公司 × 分类） ----
    var rightCompanies = [];
    var rightTotals = { summary: { orderLines: 0, aiLines: 0, aiRate: "0%" }, methods: {} };
    categories.forEach(function (c) { rightTotals.methods[c] = { orderLines: 0, aiLines: 0, aiRate: "0%" }; });

    var companyKeys = Object.keys(companyData).sort(function (a, b) {
      var ta = 0, tb = 0;
      categories.forEach(function (c) { ta += companyData[a].cats[c].orderLines; tb += companyData[b].cats[c].orderLines; });
      return tb - ta;
    });

    for (var ci = 0; ci < companyKeys.length; ci++) {
      var ck = companyKeys[ci];
      var cd = companyData[ck];
      var compOrder = 0, compAi = 0;
      var compMethods = {};
      categories.forEach(function (cat) {
        var d = cd.cats[cat];
        compOrder += d.orderLines;
        compAi += d.aiLines;
        var rate = d.orderLines > 0 ? Math.round(d.aiLines / d.orderLines * 100) + "%" : "0%";
        compMethods[cat] = { orderLines: d.orderLines, aiLines: d.aiLines, aiRate: rate };
        rightTotals.methods[cat].orderLines += d.orderLines;
        rightTotals.methods[cat].aiLines += d.aiLines;
      });
      rightTotals.summary.orderLines += compOrder;
      rightTotals.summary.aiLines += compAi;

      rightCompanies.push({
        company: cd.name,
        companyKey: ck,
        summary: {
          orderLines: compOrder,
          aiLines: compAi,
          aiRate: compOrder > 0 ? Math.round(compAi / compOrder * 100) + "%" : "0%",
        },
        methods: compMethods,
      });
    }

    rightTotals.summary.aiRate = rightTotals.summary.orderLines > 0
      ? Math.round(rightTotals.summary.aiLines / rightTotals.summary.orderLines * 100) + "%" : "0%";
    categories.forEach(function (cat) {
      var rt = rightTotals.methods[cat];
      rt.aiRate = rt.orderLines > 0 ? Math.round(rt.aiLines / rt.orderLines * 100) + "%" : "0%";
    });

    // --- 创建 Excel 工作簿（动态列数） ---
    var wb = new ExcelJS.Workbook();
    var ws = wb.addWorksheet("按下单形式统计");

    // 颜色常量
    var C_HEADER_BG = "FFAFECEB";
    var C_TOTAL_BG = "FFAFECEB";
    var C_LEFT_DATA_BG = "FFFFFFFF";
    var C_RIGHT_SUMMARY_BG = "FFFFFF00";
    var C_RIGHT_DATA_BG = "FFFCE4D6";
    var C_RIGHT_COMPANY_BG = "FF404040";
    var C_BORDER = "FFA6A6A6";
    var C_TEXT = "FF000000";
    var C_TEXT_WHITE = "FFFFFFFF";

    var RIGHT_START_COL = 6; // F列

    function applyCellStyle(cell, bg, fontColor, bold, fontName, fontSize, halign, borderStyle) {
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bg } };
      cell.font = {
        color: { argb: fontColor || C_TEXT },
        bold: bold || false,
        name: fontName || "微软雅黑",
        size: fontSize || 10,
      };
      cell.alignment = { horizontal: halign || "center", vertical: "middle", wrapText: false };
      cell.border = {
        top: { style: borderStyle || "thin", color: { argb: C_BORDER } },
        bottom: { style: borderStyle || "thin", color: { argb: C_BORDER } },
        left: { style: borderStyle || "thin", color: { argb: C_BORDER } },
        right: { style: borderStyle || "thin", color: { argb: C_BORDER } },
      };
    }

    function headerStyle(cell) { applyCellStyle(cell, C_HEADER_BG, C_TEXT, true); }
    function totalStyle(cell) { applyCellStyle(cell, C_TOTAL_BG, C_TEXT, true); }
    function rightCompanyStyle(cell) { applyCellStyle(cell, C_RIGHT_COMPANY_BG, C_TEXT_WHITE, false, "微软雅黑", 10, "left"); }
    function rightSummaryStyle(cell) { applyCellStyle(cell, C_RIGHT_SUMMARY_BG, C_TEXT, false); }
    function rightDataStyle(cell) { applyCellStyle(cell, C_RIGHT_DATA_BG, C_TEXT, false); }

    // 列宽
    ws.getColumn(1).width = 12;
    ws.getColumn(2).width = 10;
    ws.getColumn(3).width = 10;
    ws.getColumn(4).width = 10;
    ws.getColumn(5).width = 5;

    ws.getColumn(RIGHT_START_COL).width = 16;
    var rightCol = RIGHT_START_COL + 1;
    ws.getColumn(rightCol).width = 9; ws.getColumn(rightCol + 1).width = 9; ws.getColumn(rightCol + 2).width = 9;
    rightCol += 3;
    for (var mi = 0; mi < categories.length; mi++) {
      ws.getColumn(rightCol).width = 9; ws.getColumn(rightCol + 1).width = 9; ws.getColumn(rightCol + 2).width = 9;
      rightCol += 3;
    }

    // ---- 行 1：顶层表头 ----
    var row1 = ws.getRow(1);
    row1.height = 22;
    var leftHeaders = ["下单方式", "订单行数", "AI已转", "转单率"];
    for (var ci2 = 0; ci2 < leftHeaders.length; ci2++) { headerStyle(row1.getCell(ci2 + 1)); row1.getCell(ci2 + 1).value = leftHeaders[ci2]; }
    applyCellStyle(row1.getCell(5), C_HEADER_BG, C_TEXT, false);

    headerStyle(row1.getCell(RIGHT_START_COL));
    row1.getCell(RIGHT_START_COL).value = "所属公司";
    ws.mergeCells(1, RIGHT_START_COL + 1, 1, RIGHT_START_COL + 3);
    headerStyle(row1.getCell(RIGHT_START_COL + 1));
    row1.getCell(RIGHT_START_COL + 1).value = "汇总";

    rightCol = RIGHT_START_COL + 4;
    for (var mj = 0; mj < categories.length; mj++) {
      ws.mergeCells(1, rightCol, 1, rightCol + 2);
      headerStyle(row1.getCell(rightCol));
      row1.getCell(rightCol).value = categories[mj];
      rightCol += 3;
    }

    // ---- 行 2：子表头 ----
    var row2 = ws.getRow(2);
    row2.height = 20;
    for (var cj = 1; cj <= 5; cj++) applyCellStyle(row2.getCell(cj), C_HEADER_BG, C_TEXT, false);
    headerStyle(row2.getCell(RIGHT_START_COL));

    var subCol = RIGHT_START_COL + 1;
    var subHeaders = ["订单行数", "AI已转", "转单率"];
    for (var sk = 0; sk < subHeaders.length; sk++) { headerStyle(row2.getCell(subCol)); row2.getCell(subCol).value = subHeaders[sk]; subCol++; }
    for (var mk = 0; mk < categories.length; mk++) {
      for (var sl = 0; sl < subHeaders.length; sl++) { headerStyle(row2.getCell(subCol)); row2.getCell(subCol).value = subHeaders[sl]; subCol++; }
    }

    // ---- 左表数据行（从 row 3 开始） ----
    for (var ri = 0; ri < leftRows.length; ri++) {
      var lr = leftRows[ri];
      var row = ws.getRow(3 + ri);
      row.height = 20;
      applyCellStyle(row.getCell(1), C_LEFT_DATA_BG, C_TEXT, false, "微软雅黑", 10, "left");
      row.getCell(1).value = lr.method;
      applyCellStyle(row.getCell(2), C_LEFT_DATA_BG, C_TEXT, false); row.getCell(2).value = lr.orderLines;
      applyCellStyle(row.getCell(3), C_LEFT_DATA_BG, C_TEXT, false); row.getCell(3).value = lr.aiLines;
      applyCellStyle(row.getCell(4), C_LEFT_DATA_BG, C_TEXT, false); row.getCell(4).value = lr.aiRate;
      applyCellStyle(row.getCell(5), C_LEFT_DATA_BG, C_TEXT, false);
    }

    // 左表总计行
    var leftTotalRowNum = 3 + leftRows.length;
    var ltRow = ws.getRow(leftTotalRowNum);
    ltRow.height = 22;
    totalStyle(ltRow.getCell(1)); ltRow.getCell(1).value = leftTotal.method;
    totalStyle(ltRow.getCell(2)); ltRow.getCell(2).value = leftTotal.orderLines;
    totalStyle(ltRow.getCell(3)); ltRow.getCell(3).value = leftTotal.aiLines;
    totalStyle(ltRow.getCell(4)); ltRow.getCell(4).value = leftTotal.aiRate;
    applyCellStyle(ltRow.getCell(5), C_TOTAL_BG, C_TEXT, false);

    // ---- 右表数据行 ----
    for (var rj = 0; rj < rightCompanies.length; rj++) {
      var rc = rightCompanies[rj];
      var row2b = ws.getRow(3 + rj);

      rightCompanyStyle(row2b.getCell(RIGHT_START_COL));
      row2b.getCell(RIGHT_START_COL).value = rc.company;

      rightCol = RIGHT_START_COL + 1;
      rightSummaryStyle(row2b.getCell(rightCol)); row2b.getCell(rightCol).value = rc.summary.orderLines; rightCol++;
      rightSummaryStyle(row2b.getCell(rightCol)); row2b.getCell(rightCol).value = rc.summary.aiLines; rightCol++;
      rightSummaryStyle(row2b.getCell(rightCol)); row2b.getCell(rightCol).value = rc.summary.aiRate;

      rightCol = RIGHT_START_COL + 4;
      for (var mk2 = 0; mk2 < categories.length; mk2++) {
        var md = rc.methods[categories[mk2]] || { orderLines: 0, aiLines: 0, aiRate: "0%" };
        rightDataStyle(row2b.getCell(rightCol)); row2b.getCell(rightCol).value = md.orderLines; rightCol++;
        rightDataStyle(row2b.getCell(rightCol)); row2b.getCell(rightCol).value = md.aiLines; rightCol++;
        rightDataStyle(row2b.getCell(rightCol)); row2b.getCell(rightCol).value = md.aiRate; rightCol++;
      }

      // 超出左表的行填充白色
      if (3 + rj > leftTotalRowNum) {
        for (var fc = 1; fc <= 5; fc++) {
          applyCellStyle(row2b.getCell(fc), "FFFFFFFF", C_TEXT, false);
          row2b.getCell(fc).value = null;
        }
      }
      if (3 + rj < leftTotalRowNum) row2b.height = 20;
      else if (3 + rj > leftTotalRowNum) row2b.height = 20;
    }

    // 右表总计行
    var rightTotalRowNum = 3 + rightCompanies.length;
    var rtRow = ws.getRow(rightTotalRowNum);
    rtRow.height = 22;

    for (var fc2 = 1; fc2 <= 5; fc2++) {
      applyCellStyle(rtRow.getCell(fc2), "FFFFFFFF", C_TEXT, false);
      rtRow.getCell(fc2).value = null;
    }

    totalStyle(rtRow.getCell(RIGHT_START_COL));
    rtRow.getCell(RIGHT_START_COL).value = "总计";

    rightCol = RIGHT_START_COL + 1;
    totalStyle(rtRow.getCell(rightCol)); rtRow.getCell(rightCol).value = rightTotals.summary.orderLines; rightCol++;
    totalStyle(rtRow.getCell(rightCol)); rtRow.getCell(rightCol).value = rightTotals.summary.aiLines; rightCol++;
    totalStyle(rtRow.getCell(rightCol)); rtRow.getCell(rightCol).value = rightTotals.summary.aiRate;

    rightCol = RIGHT_START_COL + 4;
    for (var mn = 0; mn < categories.length; mn++) {
      var rtm = rightTotals.methods[categories[mn]] || { orderLines: 0, aiLines: 0, aiRate: "0%" };
      totalStyle(rtRow.getCell(rightCol)); rtRow.getCell(rightCol).value = rtm.orderLines; rightCol++;
      totalStyle(rtRow.getCell(rightCol)); rtRow.getCell(rightCol).value = rtm.aiLines; rightCol++;
      totalStyle(rtRow.getCell(rightCol)); rtRow.getCell(rightCol).value = rtm.aiRate; rightCol++;
    }

    // ---- 生成 buffer 并返回 ----
    var buffer = await wb.xlsx.writeBuffer();

    var fileName = "按下单形式统计_" + (cutoffDate || "data") + ".xlsx";
    res.writeHead(200, {
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": "attachment; filename=\"" + encodeURIComponent(fileName) + "\"",
      "Content-Length": buffer.length,
    });
    res.end(Buffer.from(buffer));
  } catch (e) {
    console.error("Export order method error:", e);
    res.writeHead(500, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ error: e.message }));
  }
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

  // DELETE /api/trends?date=MM-DD
  if (urlPath === "/api/trends" && req.method === "DELETE") {
    if (!isAuthed(req)) { unauthorized(res); return; }
    const tParams = new URLSearchParams(req.url.split("?")[1] || "");
    const dateKey = tParams.get("date");
    if (!dateKey || !/^\d{2}-\d{2}$/.test(dateKey)) {
      res.writeHead(400, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ error: "date 格式错误，应为 MM-DD（如 07-09）" }));
      return;
    }
    handleDeleteTrend(dateKey, res);
    return;
  }

  // GET /api/export-order-method — 按下单方式导出 Excel
  if (urlPath === "/api/export-order-method" && req.method === "GET") {
    if (!isAuthed(req)) { unauthorized(res); return; }
    handleExportOrderMethod(req, res);
    return;
  }

  // GET /api/export-pending-backup — 下载累积待转单 Excel（不清理）
  if (urlPath === "/api/export-pending-backup" && req.method === "GET") {
    if (!isAuthed(req)) { unauthorized(res); return; }
    handleExportPendingBackup(req, res);
    return;
  }

  // POST /api/clear-pending — 清理累积待转单（all / keepLast / dates）
  if (urlPath === "/api/clear-pending" && req.method === "POST") {
    if (!isAuthed(req)) { unauthorized(res); return; }
    handleClearPending(req, res);
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
