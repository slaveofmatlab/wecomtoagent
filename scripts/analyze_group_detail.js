/**
 * 群维度综合表：群名 + 下单方式 + 订单数 + AI转单
 * 用法: node scripts/analyze_group_detail.js MMDD
 * 输出: 群分析/企业微信群综合明细_MMDD.md
 */
const path = require("path");
const fs = require("fs");
const XLSX = require("xlsx");
const {
  parseSalesFull,
  parsePendingWecom,
  parseWecomProgress,
  normalizeText,
  findFile,
  findHeaderIndex,
  rowsToObjects,
  sheetRows,
} = require("./lib/page_logic");

const ROOT = path.join(__dirname, "..");
const REPORT_DIR = path.join(ROOT, "..", "群分析");

// ---- 下单方式判定 ----
function getFileExt(filename) {
  if (!filename || !filename.includes(".")) return null;
  return filename.split(".").pop().toLowerCase();
}
function determineOrderMethod(msgtype, filename) {
  const mt = (msgtype || "").toLowerCase().trim();
  if (mt === "image") return { label: "图片下单", isAI: true };
  if (mt === "mixed") return { label: "图文混发", isAI: true };
  if (mt === "text")  return { label: "文本消息", isAI: true };
  if (mt === "file") {
    const ext = getFileExt(filename);
    if (ext === "pdf") return { label: "PDF下单", isAI: false };
    if (ext === "xlsx" || ext === "xls" || ext === "xlsm") return { label: "Excel下单", isAI: false };
    if (ext === "doc" || ext === "docx") return { label: "Word下单", isAI: false };
    if (["jpg","jpeg","png","gif","jfif","bmp","webp"].includes(ext)) return { label: "图片文件", isAI: false };
    return { label: "文件(" + (ext || "未知") + ")", isAI: false };
  }
  return null;
}

// ---- 从日志分析下单方式（按群） ----
function analyzeLogByRoom(logWb) {
  if (!logWb) return {};
  const rows = sheetRows(logWb, "");
  const hIdx = findHeaderIndex(rows, ["room_name", "msgtype", "filename", "filter_status"]);
  if (hIdx < 0) return {};
  const records = rowsToObjects(rows, hIdx);

  const byRoom = {};
  for (const r of records) {
    const status = normalizeText(r["filter_status"]);
    if (status !== "ACCEPTED" && status !== "SKIPPED") continue;
    const room = normalizeText(r["room_name"]);
    if (!room) continue;
    const method = determineOrderMethod(r["msgtype"], r["filename"]);
    if (!method) continue;
    if (status === "SKIPPED" && normalizeText(r["msgtype"]) === "text") continue;
    if (!byRoom[room]) byRoom[room] = {};
    if (!byRoom[room][method.label]) byRoom[room][method.label] = { count: 0, isAI: method.isAI };
    byRoom[room][method.label].count += 1;
  }

  const result = {};
  for (const [room, methods] of Object.entries(byRoom)) {
    const entries = Object.entries(methods).sort((a, b) => b[1].count - a[1].count);
    const total = entries.reduce((s, [, m]) => s + m.count, 0);
    if (entries.length === 1 || entries[0][1].count / total > 0.8) {
      result[room] = entries[0][0];
    } else {
      const parts = entries.slice(0, 2).map(([l, m]) => l + " " + Math.round(m.count / total * 100) + "%");
      result[room] = parts.join("，");
    }
  }
  return result;
}

function main() {
  const cutoffDate = process.argv[2];
  if (!cutoffDate || cutoffDate.length !== 4) { console.error("用法: node scripts/analyze_group_detail.js MMDD"); process.exit(1); }

  const month = String(parseInt(cutoffDate.slice(0, 2), 10));
  const day = String(parseInt(cutoffDate.slice(2, 4), 10));
  const searchDir = path.join(ROOT, "示例数据", month + "月" + day + "日");
  const sampleDir = path.join(ROOT, "示例数据");

  const salesPath = findFile(searchDir, "销售订单") || findFile(sampleDir, "销售订单");
  const pendingPath = findFile(searchDir, "待转单") || findFile(sampleDir, "待转单");
  const progressPath = findFile(searchDir, "企业微信AI转单推进表") || findFile(path.join(ROOT, "basicData"), "企业微信AI转单推进表");
  const logPath = findFile(path.join(ROOT, "basicData"), "微信日志") || findFile(ROOT, "微信日志");

  console.error("销售: " + path.basename(salesPath));
  console.error("待转: " + path.basename(pendingPath));
  console.error("推进: " + path.basename(progressPath));
  console.error("日志: " + (logPath ? path.basename(logPath) : "(无)"));

  const salesRows = parseSalesFull(XLSX.readFile(salesPath));
  const pendingRows = parsePendingWecom(XLSX.readFile(pendingPath));
  const progressRows = parseWecomProgress(XLSX.readFile(progressPath), cutoffDate);
  const logWb = logPath ? XLSX.readFile(logPath) : null;

  // 下单方式
  const orderMethodMap = analyzeLogByRoom(logWb);

  // hotelCode → groupName
  const hotelToGroup = {};
  for (const row of progressRows) {
    const code = normalizeText(row.hotelCode);
    if (!code || !row.groupName) continue;
    hotelToGroup[code] = row.groupName;
  }

  // 按群聚合
  const groups = {};
  for (const row of progressRows) {
    const name = row.groupName || "(未命名群)";
    if (!groups[name]) {
      groups[name] = { groupName: name, operationCompany: row.operationCompany || "", hotelCodes: new Set(), itConfigured: false };
    }
    if (row.hotelCode) groups[name].hotelCodes.add(normalizeText(row.hotelCode));
    if (row.itConfigured) groups[name].itConfigured = true;
    if (row.operationCompany) groups[name].operationCompany = row.operationCompany;
  }

  // 全局 IT已配置集合
  const allItOkCodes = new Set();
  for (const row of progressRows) {
    if (row.itConfigured && row.hotelCode) allItOkCodes.add(normalizeText(row.hotelCode));
  }

  // 待转单匹配表
  const pendingMap = new Map();
  for (const row of pendingRows) {
    if (row.createdBy !== "供应链管理员") continue;
    const ck = normalizeText(row.customerOrderNo);
    if (ck && !pendingMap.has(ck)) pendingMap.set(ck, row.transferStatus);
  }

  // 按群统计订单
  const groupStats = [];
  for (const key of Object.keys(groups)) {
    const g = groups[key];
    const entry = {
      groupName: g.groupName,
      operationCompany: g.operationCompany,
      hotelCodeCount: g.hotelCodes.size,
      itConfigured: g.itConfigured,
      orderTotal: 0,
      orderAi: 0,
      orderAiTotal: 0,
      orderMethod: "",
    };

    // 下单方式（按群名匹配日志）
    entry.orderMethod = orderMethodMap[normalizeText(g.groupName)] || "";
    // 尝试部分匹配
    if (!entry.orderMethod) {
      for (const [room, method] of Object.entries(orderMethodMap)) {
        if (normalizeText(g.groupName).includes(room) || room.includes(normalizeText(g.groupName))) {
          entry.orderMethod = method;
          break;
        }
      }
    }

    if (g.itConfigured) {
      for (const sr of salesRows) {
        const hc = normalizeText(sr.hotelCode);
        if (!hc || !g.hotelCodes.has(hc)) continue;
        // 只统计全链路订单中属于该群所属公司的行
        if (g.operationCompany && normalizeText(sr.operationCompany || "").replace(/[（(]/g, "(").replace(/[）)]/g, ")") !== normalizeText(g.operationCompany).replace(/[（(]/g, "(").replace(/[）)]/g, ")")) continue;
        entry.orderTotal += 1;
        const custKey = normalizeText(sr.customerOrderNo);
        const pendingStatus = pendingMap.get(custKey);
        if (pendingStatus) {
          entry.orderAiTotal += 1;
          if (pendingStatus.includes("已转")) entry.orderAi += 1;
        }
      }
    }

    entry.aiRate = entry.orderTotal > 0 ? entry.orderAi / entry.orderTotal : null;
    entry.aiRateTotal = entry.orderTotal > 0 ? entry.orderAiTotal / entry.orderTotal : null;
    groupStats.push(entry);
  }

  // 排序：有订单的排前面
  groupStats.sort((a, b) => {
    if (a.orderTotal > 0 && b.orderTotal === 0) return -1;
    if (a.orderTotal === 0 && b.orderTotal > 0) return 1;
    return b.orderTotal - a.orderTotal || b.hotelCodeCount - a.hotelCodeCount;
  });

  // ==== 输出 Markdown ====
  const lines = [];
  lines.push("# 企业微信群综合明细 — 截止 " + cutoffDate);
  lines.push("");
  lines.push("> 数据来源：推进表 + 销售订单全链路 + 待转单 + 微信日志");
  lines.push("> 订单行数 = 该群项目点中、属于同公司且 IT 已配置的销售订单行数");
  lines.push("> AI转单 = 订单行中客户订单号匹配到 AI 待转单且已转单的数量");
  lines.push("");

  const withOrders = groupStats.filter(g => g.orderTotal > 0);
  const totalOrder = withOrders.reduce((s, g) => s + g.orderTotal, 0);
  const totalAi = withOrders.reduce((s, g) => s + g.orderAi, 0);

  lines.push("## 总览");
  lines.push("");
  lines.push(`| 指标 | 数值 |`);
  lines.push(`|------|------|`);
  lines.push(`| 推进表群总数 | ${groupStats.length} |`);
  lines.push(`| 有订单的群数 | ${withOrders.length} |`);
  lines.push(`| 合计订单行数 | ${totalOrder} |`);
  lines.push(`| 合计 AI 已转单 | ${totalAi} |`);
  lines.push(`| 整体 AI 转单占比 | ${totalOrder > 0 ? (totalAi / totalOrder * 100).toFixed(1) + "%" : "-"} |`);
  lines.push("");

  lines.push("## 群明细（按订单行数降序）");
  lines.push("");
  lines.push("| # | 群名称 | 所属公司 | 下单方式 | 项目点数 | 订单行数 | AI已转 | AI转单占比 |");
  lines.push("|---|--------|----------|---------|---------|---------|--------|-----------|");

  withOrders.forEach((g, i) => {
    const method = g.orderMethod || "-";
    const aiPct = g.aiRate !== null ? (g.aiRate * 100).toFixed(1) + "%" : "-";
    lines.push(`| ${i + 1} | ${g.groupName} | ${g.operationCompany} | ${method} | ${g.hotelCodeCount} | ${g.orderTotal} | ${g.orderAi} | ${aiPct} |`);
  });
  lines.push("");

  // 无订单但已配置的群
  const noOrderConfigured = groupStats.filter(g => g.orderTotal === 0 && g.itConfigured);
  if (noOrderConfigured.length > 0) {
    lines.push("## 无订单但 IT 已配置的群");
    lines.push("");
    lines.push("| 群名称 | 所属公司 | 下单方式 | 项目点数 |");
    lines.push("|--------|----------|---------|---------|");
    noOrderConfigured.forEach(g => {
      lines.push(`| ${g.groupName} | ${g.operationCompany} | ${g.orderMethod || "-"} | ${g.hotelCodeCount} |`);
    });
    lines.push("");
  }

  const reportFile = path.join(REPORT_DIR, "企业微信群综合明细_" + cutoffDate + ".md");
  fs.mkdirSync(REPORT_DIR, { recursive: true });
  const content = lines.join("\n");
  fs.writeFileSync(reportFile, content);
  console.log(content);
  console.error("\n已保存: " + reportFile);
}

main();
