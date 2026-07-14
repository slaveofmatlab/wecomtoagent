/**
 * 按群统计下单方式（图片下单 / PDF下单 / Excel下单 等），只看7月
 * 用法: node scripts/analyze_order_method.js
 * 输出: 群分析/企业微信群下单方式分析.md
 */
const path = require("path");
const fs = require("fs");
const XLSX = require("xlsx");
const { sheetRows, findHeaderIndex, rowsToObjects, normalizeText, findFile } = require("./lib/page_logic");

const ROOT = path.join(__dirname, "..");
const REPORT_DIR = path.join(ROOT, "..", "群分析");

function determineOrderMethod(msgtype, filename) {
  const mt = (msgtype || "").toLowerCase().trim();
  if (mt === "image")  return { label: "图片下单",   isAI: true };
  if (mt === "mixed")  return { label: "图文混发",   isAI: true };
  if (mt === "text")   return { label: "文本消息",   isAI: true };
  if (mt === "file") {
    const ext = (filename || "").includes(".") ? filename.split(".").pop().toLowerCase() : null;
    if (ext === "pdf") return { label: "PDF下单", isAI: false };
    if (ext === "xlsx" || ext === "xls" || ext === "xlsm") return { label: "Excel下单", isAI: false };
    if (ext === "doc" || ext === "docx") return { label: "Word下单", isAI: false };
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

  // 适配新日志格式
  const hasMsgSheet = logWb.SheetNames.some(s => s.includes("企微消息"));
  const idToName = {};
  if (hasMsgSheet) {
    try {
      const cfgRows = sheetRows(logWb, "配置清单");
      const cfgHdr = findHeaderIndex(cfgRows, ["roomid", "room_name"]);
      if (cfgHdr >= 0) rowsToObjects(cfgRows, cfgHdr).forEach(r => { idToName[r["roomid"]] = r["room_name"]; });
    } catch (e) {}
  }
  const rows = sheetRows(logWb, hasMsgSheet ? "企微消息" : "");
  const headerCandidates = hasMsgSheet
    ? ["roomid", "msgtype", "filename", "filter_status", "msgtime"]
    : ["room_name", "msgtype", "filename", "filter_status"];
  const headerIndex = findHeaderIndex(rows, headerCandidates);
  if (headerIndex < 0) { console.error("找不到日志表头"); process.exit(1); }

  const records = rowsToObjects(rows, headerIndex);
  const JULY_1_MS = 1782835200000;

  // 按群统计
  const byRoom = {};

  for (const r of records) {
    // 只看 7 月
    if (r["msgtime"] && Number(r["msgtime"]) < JULY_1_MS) continue;

    const status = normalizeText(r["filter_status"]);
    if (status !== "ACCEPTED" && status !== "SKIPPED") continue;

    const room = hasMsgSheet
      ? (idToName[normalizeText(r["roomid"])] || normalizeText(r["roomid"]))
      : normalizeText(r["room_name"]);
    if (!room) continue;

    const method = determineOrderMethod(r["msgtype"], r["filename"]);
    if (!method) continue;
    if (status === "SKIPPED" && normalizeText(r["msgtype"]) === "text") continue;

    if (!byRoom[room]) byRoom[room] = {};
    if (!byRoom[room][method.label]) byRoom[room][method.label] = { count: 0, isAI: method.isAI };
    byRoom[room][method.label].count += 1;
  }

  // 格式化输出
  const lines = [];
  lines.push("# 企业微信群下单方式分析（仅7月）");
  lines.push("");
  lines.push("> 数据来源: 微信日志.xlsx（ACCEPTED + SKIPPED，只看7月，排除纯文本噪音）");
  lines.push("> AI可处理: 图片下单 / 图文混发 / 文本消息 ｜ 系统不支持: PDF下单 / Excel下单 / Word下单 / 图片文件等");
  lines.push("");

  let totalGroups = Object.keys(byRoom).length;
  let imageOnly = 0, hasImage = 0, hasNonAI = 0, nonAIOnly = 0;
  for (const [room, methods] of Object.entries(byRoom)) {
    const labels = Object.keys(methods);
    if (labels.some(l => l === "图片下单" || l === "图文混发")) hasImage++;
    if (labels.some(l => !methods[l].isAI)) hasNonAI++;
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

  const roomList = Object.entries(byRoom).map(([room, methods]) => {
    let total = 0, aiTotal = 0, nonAiTotal = 0;
    for (const [label, info] of Object.entries(methods)) { total += info.count; if (info.isAI) aiTotal += info.count; else nonAiTotal += info.count; }
    const sorted = Object.entries(methods).sort((a, b) => b[1].count - a[1].count);
    const top = sorted[0];
    const mainMethod = top[0] + (top[1].count / total > 0.8 ? "" : " " + Math.round(top[1].count / total * 100) + "%");
    return { room, methods: sorted, total, aiTotal, nonAiTotal, mainMethod };
  }).sort((a, b) => b.total - a.total);

  lines.push("## 各群下单方式明细");
  lines.push("");
  lines.push(`| # | 群名称 | 总消息数 | 图片相关 | 不支持 | 主要方式 |`);
  lines.push(`|---|--------|---------|---------|--------|---------|`);
  roomList.forEach((g, i) => {
    const imgCount = g.methods.filter(([l]) => l === "图片下单" || l === "图文混发").reduce((s, [, m]) => s + m.count, 0);
    const rate = Math.round(g.aiTotal / g.total * 100);
    const emoji = rate >= 80 ? "✅" : rate >= 50 ? "⚠️" : "🔴";
    lines.push(`| ${i + 1} | ${g.room} | ${g.total} | ${imgCount} | ${g.nonAiTotal} | ${emoji} ${g.mainMethod} |`);
  });
  lines.push("");

  // 更新 priority_groups.json 中已有的 mainMethod
  const pgPath = path.join(ROOT, "basicData", "priority_groups.json");
  if (fs.existsSync(pgPath)) {
    const pg = JSON.parse(fs.readFileSync(pgPath, "utf8"));
    let updated = 0;
    for (const g of roomList) {
      if (pg.groups[g.room]) {
        pg.groups[g.room].mainMethod = g.mainMethod;
        updated++;
      }
    }
    pg._日志来源 = "7月14日拉取日志（仅7月消息）";
    pg.updatedAt = new Date().toISOString().slice(0, 10);
    fs.writeFileSync(pgPath, JSON.stringify(pg, null, 2) + "\n");
    lines.push(`> 已更新 priority_groups.json 中 ${updated} 个群的 mainMethod`);
  }

  const reportFile = path.join(REPORT_DIR, "企业微信群下单方式分析.md");
  fs.mkdirSync(REPORT_DIR, { recursive: true });
  fs.writeFileSync(reportFile, lines.join("\n"));
  console.log(lines.join("\n"));
  console.error("\n已保存: " + reportFile);
}

main();
