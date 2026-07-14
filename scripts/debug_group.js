// 诊断脚本：列出指定群的项目点匹配情况
const path = require("path");
const XLSX = require("xlsx");
const {
  parseSalesFull,
  parseWecomProgress,
  normalizeText,
  findFile,
} = require("./lib/page_logic");

const ROOT = path.join(__dirname, "..");
const cutoffDate = process.argv[2] || "0713";
const groupKeyword = process.argv[3] || "浙江丰厨运营AI转单群";

const month = String(parseInt(cutoffDate.slice(0, 2), 10));
const day = String(parseInt(cutoffDate.slice(2, 4), 10));
const sampleDir = path.join(ROOT, "示例数据");
const searchDir = path.join(sampleDir, month + "月" + day + "日");

const salesPath = findFile(searchDir, "销售订单") || findFile(sampleDir, "销售订单");
const progressPath = findFile(searchDir, "企业微信AI转单推进表") || findFile(path.join(ROOT, "basicData"), "企业微信AI转单推进表");

console.log("销售: " + path.basename(salesPath));
console.log("推进: " + path.basename(progressPath));

const salesRows = parseSalesFull(XLSX.readFile(salesPath));
const progressRows = parseWecomProgress(XLSX.readFile(progressPath), cutoffDate);

// 找出目标群的项目点
const targetCodes = new Set();
let targetCompany = "";
progressRows
  .filter(r => r.groupName && normalizeText(r.groupName).includes(groupKeyword))
  .forEach(r => {
    if (r.hotelCode) targetCodes.add(normalizeText(r.hotelCode));
    if (r.operationCompany) targetCompany = r.operationCompany;
  });

console.log("\n=== " + groupKeyword + " ===");
console.log("所属公司: " + targetCompany);
console.log("推进表项目点数: " + targetCodes.size);
progressRows
  .filter(r => r.groupName && normalizeText(r.groupName).includes(groupKeyword))
  .forEach(r => {
    console.log("  " + r.hotelCode + " [" + r.hotelName + "] it=" + r.itConfigured + " joined=" + r.joined + (r.deleted ? " DELETED" : ""));
  });

// 统计销售订单中每个项目点的行数
const byCode = {};
salesRows.forEach(sr => {
  const hc = normalizeText(sr.hotelCode);
  if (!hc) return;
  if (!byCode[hc]) byCode[hc] = { count: 0, hotelName: sr.hotelName, companies: new Set() };
  byCode[hc].count += 1;
  if (sr.operationCompany) byCode[hc].companies.add(sr.operationCompany);
});

// 有订单数据的项目点
console.log("\n=== 销售订单匹配明细 ===");
let totalMatched = 0;
for (const [code, info] of Object.entries(byCode).sort((a,b) => b[1].count - a[1].count)) {
  if (targetCodes.has(code)) {
    totalMatched += info.count;
    const companies = Array.from(info.companies).join(", ");
    console.log("  ✓ " + code + " [" + info.hotelName + "]: " + info.count + " 行  |  所属公司: " + companies);
  }
}
console.log("\n该群合计: " + totalMatched + " 行");

// 该群项目点中，在销售订单里没出现的
console.log("\n=== 销售订单中未出现的项目点 ===");
targetCodes.forEach(c => {
  if (!byCode[c]) console.log("  ✗ " + c);
});
