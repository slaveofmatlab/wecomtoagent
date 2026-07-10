/**
 * 预计算所有示例数据日期的汇总，输出 data/trends.json
 * 格式与折线图已有的 key-value 结构一致（日期 → 总量）。
 * 如果已有 trends.json，只更新/新增日期，保留已有的其他日期。
 *
 * 用法:
 *   node scripts/export_trend_data.js              # 全量扫描
 *   node scripts/export_trend_data.js --cutoff 0709  # 只更新某一天
 */
const fs = require("fs");
const path = require("path");
const {
  loadDefaultData,
  readWorkbookFromPath,
  buildPageData,
} = require("./lib/page_logic");

const ROOT = path.join(__dirname, "..");
const OUT_PATH = path.join(ROOT, "data", "trends.json");
const SAMPLE_DIR = path.join(ROOT, "示例数据");

// 从目录名提取 cutoff 日期: "7月2日" → "0702"
function extractCutoff(dirName) {
  const match = dirName.match(/^(\d{1,2})月(\d{1,2})日$/);
  if (!match) return null;
  return String(parseInt(match[1], 10)).padStart(2, "0") +
    String(parseInt(match[2], 10)).padStart(2, "0");
}

// cutoff "0702" → key "07-02"
function cutoffToKey(cutoff) {
  return cutoff.slice(0, 2) + "-" + cutoff.slice(2, 4);
}

// 在指定目录找匹配文件（不递归）
function findFileInDir(dir, pattern) {
  if (!fs.existsSync(dir)) return null;
  for (const name of fs.readdirSync(dir)) {
    if (name.startsWith("~$")) continue;
    const fullPath = path.join(dir, name);
    if (fs.statSync(fullPath).isDirectory()) continue;
    if (name.includes(pattern)) return fullPath;
  }
  return null;
}

function loadExisting() {
  if (fs.existsSync(OUT_PATH)) {
    try {
      return JSON.parse(fs.readFileSync(OUT_PATH, "utf8"));
    } catch (e) {
      console.log("现有 trends.json 解析失败，将重新生成。");
    }
  }
  return {};
}

function processDate(dirName, cutoffDate, dateDir, loaded) {
  const salesFile = findFileInDir(dateDir, "销售订单");
  const pendingFile = findFileInDir(dateDir, "待转单");
  const progressFile = findFileInDir(dateDir, "企业微信AI转单推进表");

  if (!salesFile) { console.log("  跳过: 未找到销售订单"); return null; }
  if (!pendingFile) { console.log("  跳过: 未找到待转单"); return null; }

  const salesWorkbook = readWorkbookFromPath(salesFile);
  const pendingWorkbook = readWorkbookFromPath(pendingFile);
  const progressWorkbook = progressFile ? readWorkbookFromPath(progressFile) : loaded.progressWorkbook;
  if (!progressFile) console.log("  提示: 该目录无推进表，使用 basicData 全局推进表");

  const data = buildPageData({
    salesWorkbook,
    pendingWorkbook,
    progressWorkbook,
    logWorkbook: loaded.logWorkbook,
    cutoffDate,
    sources: { salesPath: salesFile, pendingPath: pendingFile, progressPath: progressFile || null },
  });

  const t = data.companySummary.totals;
  const entry = {
    cutoff: cutoffDate,
    registered: t.registeredCount,
    itOk: t.itConfiguredCount,
    configRate: t.configRate,
    orderTotal: t.orderTotal,
    orderAi: t.orderAiCount,
    aiRate: t.aiRate,
    ts: new Date().toISOString(),
  };

  console.log(`  已登记 ${t.registeredCount} / IT已配置 ${t.itConfiguredCount} / 订单 ${t.orderTotal} 行 / AI ${t.orderAiCount} 单`);
  return entry;
}

function parseArgs(argv) {
  const opts = {};
  for (let i = 2; i < argv.length; i++) {
    if (argv[i] === "--cutoff") opts.cutoff = argv[++i];
    else if (argv[i] === "--help" || argv[i] === "-h") opts.help = true;
  }
  return opts;
}

function main() {
  const opts = parseArgs(process.argv);
  if (opts.help) {
    console.log("用法: node scripts/export_trend_data.js [--cutoff MMDD]");
    console.log("  --cutoff MMDD  只更新指定日期（如 0709），不指定则扫描全部");
    return;
  }

  const existing = loadExisting();
  const loaded = loadDefaultData(ROOT, "0702"); // cutoff 不影响 progress/log 加载

  if (opts.cutoff) {
    // 单日模式：查找对应日期的子目录
    const mm = String(parseInt(opts.cutoff.slice(0, 2), 10));
    const dd = String(parseInt(opts.cutoff.slice(2, 4), 10));
    const dirName = `${mm}月${dd}日`;
    const dateDir = path.join(SAMPLE_DIR, dirName);

    if (!fs.existsSync(dateDir)) {
      console.error("日期目录不存在:", dateDir);
      process.exit(1);
    }

    console.log(`更新 ${dirName}（cutoff=${opts.cutoff}）...`);
    const entry = processDate(dirName, opts.cutoff, dateDir, loaded);
    if (entry) {
      existing[cutoffToKey(opts.cutoff)] = entry;
    }
  } else {
    // 全量模式：扫描所有日期子目录
    if (!fs.existsSync(SAMPLE_DIR)) {
      console.error("示例数据目录不存在:", SAMPLE_DIR);
      process.exit(1);
    }

    const dateDirs = fs.readdirSync(SAMPLE_DIR)
      .filter(name => fs.statSync(path.join(SAMPLE_DIR, name)).isDirectory() && extractCutoff(name))
      .sort((a, b) => extractCutoff(a).localeCompare(extractCutoff(b)));

    if (dateDirs.length === 0) {
      console.error("示例数据目录下未找到日期子文件夹（如 7月2日）");
      process.exit(1);
    }

    console.log("扫描到 " + dateDirs.length + " 个日期: " + dateDirs.join(", "));
    let updated = 0;
    for (const dirName of dateDirs) {
      const cutoffDate = extractCutoff(dirName);
      const dateDir = path.join(SAMPLE_DIR, dirName);
      console.log(`\n处理 ${dirName}（cutoff=${cutoffDate}）...`);
      const entry = processDate(dirName, cutoffDate, dateDir, loaded);
      if (entry) {
        existing[cutoffToKey(cutoffDate)] = entry;
        updated++;
      }
    }
    console.log(`\n更新了 ${updated} 个日期`);
  }

  fs.mkdirSync(path.dirname(OUT_PATH), { recursive: true });
  fs.writeFileSync(OUT_PATH, JSON.stringify(existing, null, 2), "utf8");

  const keys = Object.keys(existing).sort();
  console.log(`已导出: ${OUT_PATH}（${keys.length} 天: ${keys.join(", ")}）`);
}

try {
  main();
} catch (error) {
  console.error("导出失败:", error.message);
  console.error(error.stack);
  process.exit(1);
}
