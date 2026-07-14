/**
 * 按群统计下单方式（图片下单 / PDF下单 / Excel下单 等）
 * 用法: node scripts/analyze_order_method.js
 * 输出: 群分析/企业微信群下单方式分析.md
 */
const path = require("path");
const fs = require("fs");
const XLSX = require("xlsx");
const { sheetRows, findHeaderIndex, rowsToObjects, normalizeText, findFile } = require("./lib/page_logic");

const ROOT = path.join(__dirname, "..");
const REPORT_DIR = path.join(ROOT, "..", "群分析");

// 消息类型 → 下单方式
function determineOrderMethod(msgtype, filename) {
  const mt = (msgtype || "").toLowerCase().trim();
  if (mt === "image")  return { label: "图片下单",   isAI: true };
  if (mt === "mixed")  return { label: "图文混发",   isAI: true };
  if (mt === "text")   return { label: "文本消息",   isAI: true };
  if (mt === "file") {
    const ext = (filename || "").includes(".") ? filename.split(".").pop().toLowerCase() : null;
    if (ext === "pdf")                    return { label: "PDF下单",    isAI: false };
    if (ext === "xlsx" || ext === "xls" || ext === "xlsm") return { label: "Excel下单",  isAI: false };
    if (ext === "doc" || ext === "docx")  return { label: "Word下单",   isAI: false };
    if (["jpg","jpeg","png","gif","jfif","bmp","webp"].includes(ext)) return { label: "图片文件", isAI: false };
    return { label: "文件(" + (ext || "未知") + ")", isAI: false };
  }
  return null;
}

function main() {
  const logPath = findFile(path.join(ROOT, "basicData"), "微信日志") || findFile(ROOT, "微信日志");
  if (!logPath) { console.error("找不到微信日志.xlsx"); process.exit(1); }
  console.error("日志文件: " + path.basename(logPath));

  const logWb = XLSX.readFile(logPath);
  const rows = sheetRows(logWb, "");
  const LOG_HEADERS = ["room_name", "msgtype", "filename", "filter_status", "skip_reason"];
  const headerIndex = findHeaderIndex(rows, LOG_HEADERS);
  if (headerIndex < 0) { console.error("找不到日志表头"); process.exit(1); }

  const records = rowsToObjects(rows, headerIndex);

  // 按群统计
  const byRoom = {};  // room → { methodLabel → { count, isAI } }

  for (const r of records) {
    const status = normalizeText(r["filter_status"]);
    if (status !== "ACCEPTED" && status !== "SKIPPED") continue;

    const room = normalizeText(r["room_name"]);
    if (!room) continue;

    const method = determineOrderMethod(r["msgtype"], r["filename"]);
    if (!method) continue;

    // 过滤纯文本噪音（SKIPPED + text）
    if (status === "SKIPPED" && normalizeText(r["msgtype"]) === "text") continue;

    if (!byRoom[room]) byRoom[room] = {};
    if (!byRoom[room][method.label]) byRoom[room][method.label] = { count: 0, isAI: method.isAI };
    byRoom[room][method.label].count += 1;
  }

  // 格式化输出
  const lines = [];
  lines.push("# 企业微信群下单方式分析");
  lines.push("");
  lines.push("> 数据来源: 微信日志.xlsx（ACCEPTED + SKIPPED 消息，排除纯文本噪音）");
  lines.push("> AI可处理: 图片下单 / 图文混发 / 文本消息 ｜ 系统不支持: PDF下单 / Excel下单 / Word下单 / 图片文件等");
  lines.push("");

  // 总览
  let totalGroups = Object.keys(byRoom).length;
  let imageOnly = 0, hasImage = 0, hasNonAI = 0, nonAIOnly = 0;
  for (const [room, methods] of Object.entries(byRoom)) {
    const labels = Object.keys(methods);
    const hasImg = labels.some(l => l === "图片下单" || l === "图文混发");
    const hasNon = labels.some(l => !methods[l].isAI);
    if (hasImg) hasImage++;
    if (hasNon) hasNonAI++;
    if (labels.length === 1 && labels[0] === "图片下单") imageOnly++;
    if (labels.every(l => !methods[l].isAI)) nonAIOnly++;
  }

  lines.push("## 总览");
  lines.push("");
  lines.push(`| 指标 | 数值 |`);
  lines.push(`|------|------|`);
  lines.push(`| 日志中有消息的群数 | ${totalGroups} |`);
  lines.push(`| 使用图片下单的群 | ${hasImage} |`);
  lines.push(`| 纯图片下单群 | ${imageOnly} |`);
  lines.push(`| 含AI不支持的格式的群 | ${hasNonAI} |`);
  lines.push(`| 完全不经过AI的群 | ${nonAIOnly} |`);
  lines.push("");

  // 按群详细列表，按总消息数降序
  const roomList = Object.entries(byRoom).map(([room, methods]) => {
    let total = 0, aiTotal = 0, nonAiTotal = 0;
    for (const [label, info] of Object.entries(methods)) {
      total += info.count;
      if (info.isAI) aiTotal += info.count; else nonAiTotal += info.count;
    }
    const sorted = Object.entries(methods).sort((a, b) => b[1].count - a[1].count);
    const top = sorted[0];
    const mainMethod = top[0] + (top[1].count / total > 0.8 ? "" : " " + Math.round(top[1].count / total * 100) + "%");
    return { room, methods: sorted, total, aiTotal, nonAiTotal, mainMethod };
  }).sort((a, b) => b.total - a.total);

  lines.push("## 各群下单方式明细（按消息总数降序）");
  lines.push("");
  lines.push(`| # | 群名称 | 总消息数 | AI可处理 | 不支持 | 主要方式 |`);
  lines.push(`|---|--------|---------|---------|--------|---------|`);

  roomList.forEach((g, i) => {
    const rate = Math.round(g.aiTotal / g.total * 100);
    const emoji = rate >= 80 ? "✅" : rate >= 50 ? "⚠️" : "🔴";
    lines.push(`| ${i + 1} | ${g.room} | ${g.total} | ${g.aiTotal} (${rate}%) | ${g.nonAiTotal} | ${emoji} ${g.mainMethod} |`);
  });
  lines.push("");

  // 含非AI格式的群（重点关注）
  lines.push("## 🔴 有AI不支持格式的群（重点）");
  lines.push("");
  lines.push("| 群名称 | 总消息 | 不支持格式详情 |");
  lines.push("|--------|------|--------------|");

  for (const g of roomList) {
    const nonAI = g.methods.filter(([l, m]) => !m.isAI);
    if (nonAI.length === 0) continue;
    const details = nonAI.map(([l, m]) => l + " " + m.count + "条").join("，");
    lines.push(`| ${g.room} | ${g.total} | ${details} |`);
  }
  lines.push("");

  // 纯图片下单群（AI覆盖率最高）
  lines.push("## ✅ 纯图片下单群");
  lines.push("");
  const pureImage = roomList.filter(g => {
    const labels = g.methods.map(([l]) => l);
    return labels.every(l => l === "图片下单" || l === "图文混发" || l === "文本消息") &&
           labels.some(l => l === "图片下单" || l === "图文混发");
  });
  lines.push(`共 ${pureImage.length} 个群：`);
  pureImage.forEach(g => {
    const imgs = g.methods.filter(([l]) => l === "图片下单" || l === "图文混发");
    const imgTotal = imgs.reduce((s, [, m]) => s + m.count, 0);
    lines.push(`- ${g.room}（${imgTotal} 条图片消息）`);
  });
  lines.push("");

  // 写入文件
  const reportFile = path.join(REPORT_DIR, "企业微信群下单方式分析.md");
  fs.mkdirSync(REPORT_DIR, { recursive: true });
  const content = lines.join("\n");
  fs.writeFileSync(reportFile, content);
  console.log(content);
  console.error("\n已保存: " + reportFile);
}

main();
