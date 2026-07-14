/**
 * 按企业微信群维度分析 AI 转单情况
 * 用法: node scripts/analyze_groups.js 0709
 * 输出: Markdown 格式到 stdout
 */
const path = require("path");
const XLSX = require("xlsx");
const {
  parseSalesFull,
  parsePendingWecom,
  parseWecomProgress,
  normalizeText,
  findFile,
} = require("./lib/page_logic");

const ROOT = path.join(__dirname, "..");

function main() {
  const cutoffDate = process.argv[2];
  if (!cutoffDate || cutoffDate.length !== 4) {
    console.error("用法: node scripts/analyze_groups.js MMDD");
    process.exit(1);
  }

  const month = String(parseInt(cutoffDate.slice(0, 2), 10));
  const day = String(parseInt(cutoffDate.slice(2, 4), 10));
  const dateDirName = `${month}月${day}日`;
  const sampleDir = path.join(ROOT, "示例数据");
  const searchDir = path.join(sampleDir, dateDirName);

  const salesPath = findFile(searchDir, "销售订单") || findFile(sampleDir, "销售订单");
  const pendingPath = findFile(searchDir, "待转单") || findFile(sampleDir, "待转单");
  const progressPath = findFile(searchDir, "企业微信AI转单推进表") || findFile(path.join(ROOT, "basicData"), "企业微信AI转单推进表");

  if (!salesPath) { console.error("找不到销售订单文件"); process.exit(1); }
  if (!pendingPath) { console.error("找不到待转单文件"); process.exit(1); }
  if (!progressPath) { console.error("找不到推进表文件"); process.exit(1); }

  console.error(`销售订单: ${path.basename(salesPath)}`);
  console.error(`待转单:   ${path.basename(pendingPath)}`);
  console.error(`推进表:   ${path.basename(progressPath)}`);

  const salesWb = XLSX.readFile(salesPath);
  const pendingWb = XLSX.readFile(pendingPath);
  const progressWb = XLSX.readFile(progressPath);

  const salesRows = parseSalesFull(salesWb);
  const pendingRows = parsePendingWecom(pendingWb);
  const progressRows = parseWecomProgress(progressWb, cutoffDate);

  console.error(`销售订单 ${salesRows.length} 行, 待转单 ${pendingRows.length} 行, 推进表 ${progressRows.length} 行`);

  // 推进表 → hotelCode → { groupName, operationCompany }
  const hotelToGroup = {};
  for (const row of progressRows) {
    const code = normalizeText(row.hotelCode);
    if (!code) continue;
    hotelToGroup[code] = {
      groupName: row.groupName || "(未命名群)",
      operationCompany: row.operationCompany || "",
      itConfigured: row.itConfigured,
    };
  }

  // 按群聚合：groupName → { operationCompany, hotelCodes: Set, itConfigured: bool }
  const groups = {};
  for (const row of progressRows) {
    const name = row.groupName || "(未命名群)";
    if (!groups[name]) {
      groups[name] = {
        groupName: name,
        operationCompany: row.operationCompany || "",
        hotelCodes: new Set(),
        itConfigured: false,
      };
    }
    if (row.hotelCode) groups[name].hotelCodes.add(normalizeText(row.hotelCode));
    if (row.itConfigured) groups[name].itConfigured = true;
    // 保留最新的运营公司名
    if (row.operationCompany) groups[name].operationCompany = row.operationCompany;
  }

  // 全局 IT已配置集合
  const allItOkCodes = new Set();
  for (const row of progressRows) {
    if (row.itConfigured && row.hotelCode) allItOkCodes.add(normalizeText(row.hotelCode));
  }

  // 待转单匹配表：客户订单号 → 转单状态（只取"供应链管理员"创建的）
  const pendingMap = new Map();
  for (const row of pendingRows) {
    if (row.createdBy !== "供应链管理员") continue;
    const custKey = normalizeText(row.customerOrderNo);
    if (custKey && !pendingMap.has(custKey)) {
      pendingMap.set(custKey, row.transferStatus);
    }
  }

  // 按群统计
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
      orderAiTotal: 0, // 匹配到待转单的总数（含未转）
    };

    if (g.itConfigured) {
      // 遍历销售订单，匹配该项目点代码属于这个群
      for (const sr of salesRows) {
        const hc = normalizeText(sr.hotelCode);
        if (!hc || !g.hotelCodes.has(hc)) continue;
        entry.orderTotal += 1;

        const custKey = normalizeText(sr.customerOrderNo);
        const pendingStatus = pendingMap.get(custKey);
        if (pendingStatus) {
          entry.orderAiTotal += 1;
          if (pendingStatus.includes("已转")) {
            entry.orderAi += 1;
          }
        }
      }
    }

    entry.aiRate = entry.orderTotal > 0 ? entry.orderAi / entry.orderTotal : null;
    entry.aiRateTotal = entry.orderTotal > 0 ? entry.orderAiTotal / entry.orderTotal : null;

    groupStats.push(entry);
  }

  // 排序：有订单的行排前面，按订单行数降序；没订单的按项目点数降序
  groupStats.sort((a, b) => {
    if (a.orderTotal > 0 && b.orderTotal === 0) return -1;
    if (a.orderTotal === 0 && b.orderTotal > 0) return 1;
    if (a.orderTotal !== b.orderTotal) return b.orderTotal - a.orderTotal;
    return b.hotelCodeCount - a.hotelCodeCount;
  });

  // ===== 输出 Markdown =====
  console.log(`## 企业微信群 AI 转单分析 — 截止 ${cutoffDate}`);
  console.log();
  console.log(`> 数据来源：销售订单 ${salesRows.length} 行 | 待转单 ${pendingRows.length} 行 | 推进表覆盖 ${Object.keys(groups).length} 个群`);
  console.log();

  // 总览
  const withOrders = groupStats.filter(g => g.orderTotal > 0);
  const withoutOrders = groupStats.filter(g => g.orderTotal === 0);
  const configuredGroups = groupStats.filter(g => g.itConfigured);
  const totalOrderRows = withOrders.reduce((s, g) => s + g.orderTotal, 0);
  const totalAi = withOrders.reduce((s, g) => s + g.orderAi, 0);
  const totalAiAll = withOrders.reduce((s, g) => s + g.orderAiTotal, 0);

  console.log("### 总览");
  console.log();
  console.log(`| 指标 | 数值 |`);
  console.log(`|------|------|`);
  console.log(`| 推进表群总数 | ${groupStats.length} |`);
  console.log(`| IT 已配置群数 | ${configuredGroups.length} |`);
  console.log(`| 有订单数据的群数 | ${withOrders.length} |`);
  console.log(`| 无订单数据的群数 | ${withoutOrders.length} |`);
  console.log(`| 已配置群订单总行数 | ${totalOrderRows} |`);
  console.log(`| AI 已转单行数 | ${totalAi} |`);
  console.log(`| AI 识别总行数（含未转） | ${totalAiAll} |`);
  console.log(`| 整体 AI 转单占比 | ${totalOrderRows > 0 ? (totalAi / totalOrderRows * 100).toFixed(1) + "%" : "-"} |`);
  console.log();

  // 有订单的群明细
  console.log("### 有订单数据的群（按订单行数降序）");
  console.log();
  console.log(`| # | 群名称 | 所属公司 | 项目点数 | 订单行数 | AI已转 | AI识别 | AI转单占比 |`);
  console.log(`|---|--------|----------|---------|---------|--------|--------|-----------|`);

  withOrders.forEach((g, i) => {
    const aiPct = g.aiRate !== null ? (g.aiRate * 100).toFixed(1) + "%" : "-";
    console.log(`| ${i + 1} | ${g.groupName} | ${g.operationCompany} | ${g.hotelCodeCount} | ${g.orderTotal} | ${g.orderAi} | ${g.orderAiTotal} | ${aiPct} |`);
  });
  console.log();

  // AI 转单占比 < 50% 的群（重点关注）
  const lowAi = withOrders.filter(g => g.orderTotal >= 5 && g.aiRate !== null && g.aiRate < 0.5);
  if (lowAi.length > 0) {
    console.log("### ⚠ AI 转单占比偏低（< 50% 且订单 ≥ 5 行）");
    console.log();
    console.log(`| 群名称 | 所属公司 | 订单行数 | AI已转 | AI转单占比 |`);
    console.log(`|--------|----------|---------|--------|-----------|`);
    lowAi.forEach(g => {
      console.log(`| ${g.groupName} | ${g.operationCompany} | ${g.orderTotal} | ${g.orderAi} | ${(g.aiRate * 100).toFixed(1)}% |`);
    });
    console.log();
  }

  // AI 转单占比 = 0 的群
  const zeroAi = withOrders.filter(g => g.aiRate === 0);
  if (zeroAi.length > 0) {
    console.log("### 🔴 AI 转单占比 = 0% 的群");
    console.log();
    console.log(`| 群名称 | 所属公司 | 项目点数 | 订单行数 |`);
    console.log(`|--------|----------|---------|---------|`);
    zeroAi.forEach(g => {
      console.log(`| ${g.groupName} | ${g.operationCompany} | ${g.hotelCodeCount} | ${g.orderTotal} |`);
    });
    console.log();
  }

  // 按公司汇总
  console.log("### 按公司汇总");
  console.log();
  console.log(`| 公司 | 群数 | 有订单群数 | 总订单行数 | AI已转 | AI转单占比 |`);
  console.log(`|------|------|-----------|-----------|--------|-----------|`);

  const byCompany = {};
  for (const g of groupStats) {
    const ck = g.operationCompany || "(未知)";
    if (!byCompany[ck]) byCompany[ck] = { groupCount: 0, activeGroups: 0, orderTotal: 0, orderAi: 0 };
    byCompany[ck].groupCount += 1;
    if (g.orderTotal > 0) {
      byCompany[ck].activeGroups += 1;
      byCompany[ck].orderTotal += g.orderTotal;
      byCompany[ck].orderAi += g.orderAi;
    }
  }

  Object.entries(byCompany)
    .sort((a, b) => b[1].orderTotal - a[1].orderTotal)
    .forEach(([name, s]) => {
      const rate = s.orderTotal > 0 ? (s.orderAi / s.orderTotal * 100).toFixed(1) + "%" : "-";
      console.log(`| ${name} | ${s.groupCount} | ${s.activeGroups} | ${s.orderTotal} | ${s.orderAi} | ${rate} |`);
    });
}

main();
