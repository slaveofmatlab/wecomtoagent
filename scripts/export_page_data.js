/**
 * 预解析 Excel，输出 data/page_data.json 供页面加载
 *
 * 用法:
 *   node scripts/export_page_data.js
 *   node scripts/export_page_data.js --sales 示例数据/销售订单全链路.xlsx --pending 示例数据/待转单-全量.xlsx --progress basicData/企业微信AI转单推进表.xlsx
 */
const fs = require("fs");
const path = require("path");
const {
  loadDefaultData,
  readWorkbookFromPath,
  buildPageData,
  DEFAULT_CUTOFF_DATE,
} = require("./lib/page_logic");

const ROOT = path.join(__dirname, "..");
const OUT_PATH = path.join(ROOT, "data", "page_data.json");

function parseArgs(argv) {
  const opts = {};
  for (let i = 2; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === "--sales") opts.sales = argv[++i];
    else if (arg === "--pending") opts.pending = argv[++i];
    else if (arg === "--progress") opts.progress = argv[++i];
    else if (arg === "--cutoff") opts.cutoff = argv[++i];
    else if (arg === "--out") opts.out = argv[++i];
    else if (arg === "--help" || arg === "-h") opts.help = true;
  }
  return opts;
}

function resolvePath(inputPath) {
  return path.isAbsolute(inputPath) ? inputPath : path.join(ROOT, inputPath);
}

function printHelp() {
  console.log(`用法: node scripts/export_page_data.js [选项]

选项:
  --sales <path>     销售订单全链路 xlsx（默认 示例数据/*销售订单全链路*）
  --pending <path>   待转单 xlsx（默认 示例数据/*待转单*）
  --progress <path>  企业微信AI转单推进表 xlsx（默认 basicData/*企业微信AI转单推进表*）
  --cutoff <MMDD>    统计截止时间，推进表里晚于这天的"OK-MMDD"确认状态不计入（默认 ${DEFAULT_CUTOFF_DATE}，对应文件名"7.2日"）
  --out <path>       输出 JSON（默认 data/page_data.json）
`);
}

function main() {
  const opts = parseArgs(process.argv);
  if (opts.help) {
    printHelp();
    return;
  }

  const cutoffDate = opts.cutoff || DEFAULT_CUTOFF_DATE;
  const loaded = loadDefaultData(ROOT, cutoffDate);
  const sources = { ...loaded.sources };

  let salesWorkbook = loaded.salesWorkbook;
  let pendingWorkbook = loaded.pendingWorkbook;
  let progressWorkbook = loaded.progressWorkbook;

  if (opts.sales) {
    const salesPath = resolvePath(opts.sales);
    if (!fs.existsSync(salesPath)) throw new Error(`销售订单全链路不存在: ${salesPath}`);
    salesWorkbook = readWorkbookFromPath(salesPath);
    sources.salesPath = salesPath;
  }
  if (opts.pending) {
    const pendingPath = resolvePath(opts.pending);
    if (!fs.existsSync(pendingPath)) throw new Error(`待转单不存在: ${pendingPath}`);
    pendingWorkbook = readWorkbookFromPath(pendingPath);
    sources.pendingPath = pendingPath;
  }
  if (opts.progress) {
    const progressPath = resolvePath(opts.progress);
    if (!fs.existsSync(progressPath)) throw new Error(`企业微信AI转单推进表不存在: ${progressPath}`);
    progressWorkbook = readWorkbookFromPath(progressPath);
    sources.progressPath = progressPath;
  }

  const data = buildPageData({ salesWorkbook, pendingWorkbook, progressWorkbook, cutoffDate, sources });

  const outPath = opts.out ? resolvePath(opts.out) : OUT_PATH;
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.writeFileSync(outPath, JSON.stringify(data, null, 2), "utf8");

  console.log("已导出:", outPath);
  console.log(`  统计截止时间: ${data.cutoffDate}（晚于该时间的推进表 OK-MMDD 确认状态不计入）`);
  console.log(`  销售订单全链路 ${data.salesRows.length} 行`);
  console.log(`  待转单 ${data.pendingRows.length} 行（已转 ${data.pendingTotals.transferred} / 未转 ${data.pendingTotals.notTransferred}）`);
  console.log(`  企业微信配置 ${data.progressRows.length} 条`);
  console.log(`  运营公司汇总 ${data.companySummary.rows.length} 家`);
  console.log(`  总计：已登记 ${data.companySummary.totals.registeredCount} / IT已配置 ${data.companySummary.totals.itConfiguredCount} / 订单总行数 ${data.companySummary.totals.orderTotal} / AI识别 ${data.companySummary.totals.orderAiCount}`);
}

try {
  main();
} catch (error) {
  console.error("导出失败:", error.message);
  process.exit(1);
}
