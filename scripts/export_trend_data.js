/**
 * 预计算所有示例数据日期的汇总，输出 data/trend_data.json
 * 用法: node scripts/export_trend_data.js
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
const OUT_PATH = path.join(ROOT, "data", "trend_data.json");
const SAMPLE_DIR = path.join(ROOT, "示例数据");

// 从目录名提取 cutoff 日期: "7月2日" → "0702"
function extractCutoff(dirName) {
  const match = dirName.match(/^(\d{1,2})月(\d{1,2})日$/);
  if (!match) return null;
  return String(parseInt(match[1], 10)).padStart(2, "0") +
    String(parseInt(match[2], 10)).padStart(2, "0");
}

function main() {
  if (!fs.existsSync(SAMPLE_DIR)) {
    console.error("示例数据目录不存在:", SAMPLE_DIR);
    process.exit(1);
  }

  // 扫描所有日期子目录
  const dateDirs = fs.readdirSync(SAMPLE_DIR)
    .filter(name => {
      const fullPath = path.join(SAMPLE_DIR, name);
      return fs.statSync(fullPath).isDirectory() && extractCutoff(name);
    })
    .sort((a, b) => extractCutoff(a).localeCompare(extractCutoff(b)));

  if (dateDirs.length === 0) {
    console.error("示例数据目录下未找到日期子文件夹（如 7月2日）");
    process.exit(1);
  }

  console.log("扫描到 " + dateDirs.length + " 个日期: " + dateDirs.join(", "));

  const days = [];

  for (const dirName of dateDirs) {
    const cutoffDate = extractCutoff(dirName);
    const dateDir = path.join(SAMPLE_DIR, dirName);

    console.log(`\n处理 ${dirName}（cutoff=${cutoffDate}）...`);

    // 加载默认数据（sales/pending 从 dateDir 找，progress/log 从 basicData）
    const loaded = loadDefaultData(ROOT, cutoffDate);
    const sources = { ...loaded.sources };

    // 覆盖销售订单和待转单路径为当前日期目录
    let salesWorkbook = null;
    let pendingWorkbook = null;

    // 在当前日期目录找销售订单
    const salesFile = findFileInDir(dateDir, "销售订单");
    if (salesFile) {
      salesWorkbook = readWorkbookFromPath(salesFile);
      sources.salesPath = salesFile;
    }

    // 在当前日期目录找待转单
    const pendingFile = findFileInDir(dateDir, "待转单");
    if (pendingFile) {
      pendingWorkbook = readWorkbookFromPath(pendingFile);
      sources.pendingPath = pendingFile;
    }

    if (!salesWorkbook) {
      console.log("  跳过: 未找到销售订单文件");
      continue;
    }
    if (!pendingWorkbook) {
      console.log("  跳过: 未找到待转单文件");
      continue;
    }

    const data = buildPageData({
      salesWorkbook,
      pendingWorkbook,
      progressWorkbook: loaded.progressWorkbook,
      logWorkbook: loaded.logWorkbook,
      cutoffDate,
      sources,
    });

    console.log(`  销售订单 ${data.salesRows.length} 行`);
    console.log(`  待转单 ${data.pendingRows.length} 行`);
    console.log(`  运营公司 ${data.companySummary.rows.length} 家`);

    days.push({
      date: cutoffDate,
      label: dirName,
      companySummary: data.companySummary,
      pendingTotals: data.pendingTotals,
    });
  }

  // 合并所有公司名（不同日期可能有不同公司）
  const allCompanies = new Set();
  for (const day of days) {
    for (const row of day.companySummary.rows) {
      allCompanies.add(row.operationCompany);
    }
  }

  const output = {
    generatedAt: new Date().toISOString(),
    companies: [...allCompanies].sort(),
    days,
  };

  fs.mkdirSync(path.dirname(OUT_PATH), { recursive: true });
  fs.writeFileSync(OUT_PATH, JSON.stringify(output, null, 2), "utf8");

  console.log(`\n已导出: ${OUT_PATH}`);
  console.log(`  ${output.days.length} 天 × ${output.companies.length} 家公司`);
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

try {
  main();
} catch (error) {
  console.error("导出失败:", error.message);
  console.error(error.stack);
  process.exit(1);
}
