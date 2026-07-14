/**
 * 企业微信看板 数据解析逻辑（Node 预解析）
 * 与 emilToAgent/scripts/lib/page_logic.js 的写法保持一致（normalizeText/findHeaderIndex/rowsToObjects 等通用工具直接照搬）。
 */
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

function normalizeText(value) {
  return String(value ?? "")
    .replace(/　/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

// 运营公司名称在不同表里全角/半角括号混用（如"成都/丰厨（成都）" vs "成都/丰厨(成都)"），
// 跨表 join（推进表 vs 全链路）前必须统一，否则同一家公司会被拆成两行。
function normalizeCompanyName(value) {
  return normalizeText(value)
    .replace(/[（(]/g, "(")
    .replace(/[）)]/g, ")");
}

function getFirst(row, names) {
  for (const name of names) {
    if (row[name] !== undefined && row[name] !== null && normalizeText(row[name]) !== "") {
      return normalizeText(row[name]);
    }
  }
  return "";
}

function findHeaderIndex(rows, candidates) {
  let bestIndex = -1;
  let bestScore = 0;
  for (let i = 0; i < Math.min(rows.length, 30); i += 1) {
    const cells = rows[i].map(normalizeText);
    const score = candidates.filter((candidate) => cells.includes(candidate)).length;
    if (score > bestScore) {
      bestScore = score;
      bestIndex = i;
    }
  }
  return bestScore > 0 ? bestIndex : -1;
}

function rowsToObjects(rows, headerIndex, dataStartIndex = headerIndex + 1) {
  if (headerIndex < 0) return [];
  const headers = rows[headerIndex].map((cell, index) => normalizeText(cell) || `列${index + 1}`);
  return rows.slice(dataStartIndex)
    .filter((row) => row.some((cell) => normalizeText(cell) !== ""))
    .map((row) => {
      const item = {};
      headers.forEach((header, index) => {
        item[header] = row[index];
      });
      return item;
    });
}

function readWorkbookFromPath(filePath) {
  return XLSX.readFile(filePath, { cellDates: false });
}

function getActualSheetRange(sheet) {
  const cellRefs = Object.keys(sheet).filter((key) => key[0] !== "!");
  if (cellRefs.length === 0) return sheet["!ref"];

  return cellRefs.reduce((range, ref) => {
    const decoded = XLSX.utils.decode_cell(ref);
    range.s.r = Math.min(range.s.r, decoded.r);
    range.s.c = Math.min(range.s.c, decoded.c);
    range.e.r = Math.max(range.e.r, decoded.r);
    range.e.c = Math.max(range.e.c, decoded.c);
    return range;
  }, {
    s: { r: Number.MAX_SAFE_INTEGER, c: Number.MAX_SAFE_INTEGER },
    e: { r: 0, c: 0 },
  });
}

function sheetRows(workbook, preferredName) {
  const sheetName = workbook.SheetNames.find((name) => name.includes(preferredName)) || workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: false,
    range: getActualSheetRange(sheet),
  });
}

// 销售订单全链路.xlsx → "销售订单商品明细导出" 明细行。
// 不再依赖"转单状态/已登记企业微信群/已上线配置"这些预 join 字段——7.2 的版本是利拉姐
// 手工 join 的，7.3 之后的系统原生导出是双层表头（row 0=大类分组, row 1=列名）且没有这些字段。
// 转单状态用客户订单号去待转单匹配，配置状态用项目点代码去推进表匹配。
const SALES_HEADER_CANDIDATES = [
  "项目点代码", "项目点名称", "运营公司", "客户订单号", "订单号", "客户名称",
];

function parseSalesFull(workbook) {
  const rows = sheetRows(workbook, "商品明细导出");
  const headerIndex = findHeaderIndex(rows, SALES_HEADER_CANDIDATES);
  return rowsToObjects(rows, headerIndex).map((row, index) => {
    const operationCompany = getFirst(row, ["运营公司"]);
    return {
      rowIndex: index + 1,
      orderNo: getFirst(row, ["订单号"]),
      customerOrderNo: getFirst(row, ["客户订单号"]),
      operationCompany,
      operationCompanyKey: normalizeCompanyName(operationCompany),
      hotelCode: getFirst(row, ["项目点代码"]),
      hotelName: getFirst(row, ["项目点名称"]),
      customerName: getFirst(row, ["客户名称"]),
    };
  }).filter((row) => row.operationCompany || row.hotelCode);
}

// 待转单-全量.xlsx → "销售订单待转单导出"：客户订单号可与全链路的客户订单号匹配，
// 目前仅用于展示待转单总量/已转/未转的辅助统计，不参与公司汇总表主口径。
const PENDING_HEADER_CANDIDATES = ["*商户名", "转单状态", "客户订单号", "运营公司"];

function parsePendingWecom(workbook) {
  const rows = sheetRows(workbook, "待转单导出");
  const headerIndex = findHeaderIndex(rows, PENDING_HEADER_CANDIDATES);
  return rowsToObjects(rows, headerIndex).map((row, index) => {
    const transferStatus = getFirst(row, ["转单状态"]);
    return {
      rowIndex: index + 1,
      transferStatus,
      isTransferred: transferStatus.includes("已转"),
      customerOrderNo: getFirst(row, ["客户订单号"]),
      salesOrderNo: getFirst(row, ["销售订单号"]),
      createdBy: getFirst(row, ["创建人"]),
      operationCompany: getFirst(row, ["运营公司"]),
      hotelName: getFirst(row, ["*商户名"]),
    };
  }).filter((row) => row.hotelName || row.customerOrderNo || row.salesOrderNo);
}

// 企业微信AI转单推进表.xlsx → 工作表1：项目点级别的"是否加群"/"IT是否配置完成"状态，
// 用于统计各运营公司的已登记企业微信群项目点数、IT已配置数、配置率。
const PROGRESS_HEADER_CANDIDATES = [
  "企业微信群名称", "运营公司", "项目点代码", "项目点名称", "是否加群-张利拉", "IT是否配置完成-邓虎", "群ID",
];

// 表里的状态值是 "OK" 或 "OK-MMDD"（如 "OK-0703"）。看板是某一天的快照（默认 7.2），
// 如果状态确认日期晚于快照日期，说明是快照之后才更新的，不应该算进当天的数字里。
const DEFAULT_CUTOFF_DATE = "0702";

function isConfirmedByCutoff(status, cutoffMMDD) {
  if (!status || !status.startsWith("OK")) return false;
  const match = status.match(/^OK-(\d{4})$/);
  if (!match) return true; // 裸 "OK"，没有日期后缀，视为历史基线，始终计入
  return match[1] <= cutoffMMDD;
}

function parseWecomProgress(workbook, cutoffDate = DEFAULT_CUTOFF_DATE) {
  const rows = sheetRows(workbook, "工作表1");
  const headerIndex = findHeaderIndex(rows, PROGRESS_HEADER_CANDIDATES);

  // "群ID"后面紧邻的那一列（表头为空）是删除标记列：单元格含"删除"的项目点已被业务方废弃，
  // 整行剔除、不进任何统计。rowsToObjects 给空表头列自动命名为 列${位置+1}，据此定位。
  const headerCells = headerIndex >= 0 ? rows[headerIndex].map(normalizeText) : [];
  const groupIdPos = headerCells.indexOf("群ID");
  const deleteMarkPos = groupIdPos >= 0 ? groupIdPos + 1 : -1;
  const deleteMarkKey = deleteMarkPos >= 0
    ? (headerCells[deleteMarkPos] || `列${deleteMarkPos + 1}`)
    : null;

  return rowsToObjects(rows, headerIndex).map((row, index) => {
    const operationCompany = getFirst(row, ["运营公司"]);
    const joinStatus = getFirst(row, ["是否加群-张利拉", "是否加群"]);
    const itStatus = getFirst(row, ["IT是否配置完成-邓虎", "IT是否配置完成"]);
    const deleteMark = deleteMarkKey ? normalizeText(row[deleteMarkKey]) : "";
    return {
      rowIndex: index + 1,
      groupName: getFirst(row, ["企业微信群名称"]),
      groupId: getFirst(row, ["群ID"]),
      operationCompany,
      operationCompanyKey: normalizeCompanyName(operationCompany),
      hotelCode: getFirst(row, ["项目点代码"]),
      hotelName: getFirst(row, ["项目点名称"]),
      joined: isConfirmedByCutoff(joinStatus, cutoffDate),
      itConfigured: isConfirmedByCutoff(itStatus, cutoffDate),
      deleted: deleteMark.includes("删除"),
    };
  })
    .filter((row) => !row.deleted)
    .filter((row) => row.operationCompany || row.hotelCode);
}

// ---- 微信日志 → 下单方式（备注列） ----

const LOG_HEADER_CANDIDATES = ["room_name", "msgtype", "filename", "filter_status", "skip_reason"];

// 只有系统当前能处理的消息类型才标记为"已处理"
const PROCESSABLE_MSGTYPES = new Set(["image", "text", "mixed"]);

function getFileExt(filename) {
  if (!filename || !filename.includes(".")) return null;
  return filename.split(".").pop().toLowerCase();
}

function determineOrderMethod(msgtype, filename) {
  var mt = (msgtype || "").toLowerCase().trim();
  if (mt === "image") return { label: "图片下单", processed: true };
  if (mt === "mixed") return { label: "图文混发", processed: true };
  if (mt === "text") return { label: "文本消息", processed: true };
  if (mt === "file") {
    var ext = getFileExt(filename);
    if (ext === "pdf") return { label: "PDF下单", processed: false };
    if (ext === "xlsx" || ext === "xls" || ext === "xlsm") return { label: "Excel下单", processed: false };
    if (ext === "doc" || ext === "docx") return { label: "Word下单", processed: false };
    if (ext === "jpg" || ext === "jpeg" || ext === "png" || ext === "gif" || ext === "jfif" || ext === "bmp" || ext === "webp") return { label: "图片文件", processed: false };
    return { label: "文件下单(." + (ext || "未知") + ")", processed: false };
  }
  return null; // 非下单相关消息类型（revoke/voice/video等）
}

// 推进表 → room_name → companyKey 映射
function buildRoomCompanyMap(progressRows) {
  var map = {};
  for (var i = 0; i < progressRows.length; i++) {
    var row = progressRows[i];
    var groupName = normalizeText(row.groupName);
    if (groupName && row.operationCompanyKey) {
      // 同一个群名可能在不同项目点出现，取最后一个（一般唯一）
      map[groupName] = row.operationCompanyKey;
    }
  }
  return map;
}

function parseWecomLogForSummary(workbook, progressRows) {
  if (!workbook) return null;

  // 检测日志格式：新格式有"企微消息"sheet，用"配置清单"做 roomid→room_name 映射
  var hasMsgSheet = workbook.SheetNames.some(function (s) { return s.includes("企微消息"); });
  var idToName = {};
  var preferredSheet = "";
  if (hasMsgSheet) {
    // 新格式：配置清单 sheet → roomid → room_name
    try {
      var cfgRows = sheetRows(workbook, "配置清单");
      var cfgHdr = findHeaderIndex(cfgRows, ["roomid", "room_name"]);
      if (cfgHdr >= 0) {
        var cfgRecs = rowsToObjects(cfgRows, cfgHdr);
        cfgRecs.forEach(function (r) { idToName[r["roomid"]] = r["room_name"]; });
      }
    } catch (e) {}
    preferredSheet = "企微消息";
  }

  var rows = sheetRows(workbook, preferredSheet);
  // 新格式用 roomid，旧格式用 room_name
  var headerCandidates = hasMsgSheet
    ? ["roomid", "msgtype", "filename", "filter_status"]
    : LOG_HEADER_CANDIDATES;
  var headerIndex = findHeaderIndex(rows, headerCandidates);
  if (headerIndex < 0) return null;

  var records = rowsToObjects(rows, headerIndex);
  var roomCompanyMap = buildRoomCompanyMap(progressRows);

  // 按 companyKey 聚合下单方式
  var byCompany = {};

  for (var i = 0; i < records.length; i++) {
    var r = records[i];

    var status = normalizeText(r["filter_status"]);
    if (status !== "ACCEPTED" && status !== "SKIPPED") continue;

    var msgtype = normalizeText(r["msgtype"]);
    var filename = normalizeText(r["filename"]);
    var method = determineOrderMethod(msgtype, filename);
    if (!method) continue;

    if (status === "SKIPPED" && msgtype === "text") continue;

    // 新格式：roomid → room_name 映射；旧格式：直接读 room_name
    var room = hasMsgSheet
      ? (idToName[normalizeText(r["roomid"])] || normalizeText(r["roomid"]))
      : normalizeText(r["room_name"]);
    var companyKey = roomCompanyMap[room];
    if (!companyKey) continue;

    if (!byCompany[companyKey]) byCompany[companyKey] = {};
    if (!byCompany[companyKey][method.label]) {
      byCompany[companyKey][method.label] = { count: 0, processed: method.processed };
    }
    byCompany[companyKey][method.label].count += 1;
  }

  // 格式化备注字符串
  var result = {};
  for (var ck in byCompany) {
    var methods = byCompany[ck];
    var total = 0;
    var entries = [];
    for (var label in methods) {
      var m = methods[label];
      total += m.count;
      entries.push({ label: label, count: m.count, processed: m.processed });
    }
    entries.sort(function (a, b) { return b.count - a.count; });

    if (entries.length === 0) {
      result[ck] = { methodStats: [], remark: "" };
      continue;
    }

    // top 1-2 方法
    var topEntries = entries.slice(0, 2);
    var parts = topEntries.map(function (e) {
      var pct = Math.round(e.count / total * 100);
      return e.label + " " + pct + "%";
    });

    // 如果只有一个方法且占比 > 80%，简化显示
    var remark;
    if (entries.length === 1 || (entries[0].count / total > 0.8)) {
      remark = entries[0].label;
    } else {
      remark = parts.join("，");
    }

    result[ck] = { methodStats: entries, remark: remark };
  }

  return result;
}

function calcPendingTotals(pendingRows) {
  let transferred = 0;
  let notTransferred = 0;
  for (const row of pendingRows) {
    if (row.isTransferred) transferred += 1;
    else notTransferred += 1;
  }
  return { total: pendingRows.length, transferred, notTransferred };
}

// 核心汇总：按运营公司 join 推进表（配置率）+ 待转单（转单状态）+ 销售明细（订单行数）。
//
// 三张表的关联方式（不依赖销售订单里任何预 join 字段）：
// - 配置率：销售订单.项目点代码 ↔ 推进表.项目点代码 → 判断该项目点是否已登记/IT已配置
// - AI转单率：销售订单.客户订单号 ↔ 待转单.客户订单号 → 判断是否有匹配且已转单
// - 分母（已配置项目点订单总行数）：销售订单中，项目点代码在 IT已配置集合里的所有行
function buildCompanySummary(salesRows, pendingRows, progressRows, logSummary) {
  const byCompany = new Map();

  const ensure = (key, displayName) => {
    if (!byCompany.has(key)) {
      byCompany.set(key, {
        operationCompany: displayName,
        operationCompanyKey: key,
        registeredCodes: new Set(),
        itOkCodes: new Set(),
        orderTotal: 0,
        orderAi: 0,
        orderAiTotal: 0,
      });
    }
    return byCompany.get(key);
  };

  // 推进表 → 按公司分组的已登记 / IT已配置项目点代码集合
  for (const row of progressRows) {
    if (!row.operationCompanyKey || !row.hotelCode) continue;
    const entry = ensure(row.operationCompanyKey, row.operationCompany);
    entry.registeredCodes.add(row.hotelCode);
    if (row.itConfigured) entry.itOkCodes.add(row.hotelCode);
  }

  // 构建全局 IT已配置 集合（跨公司——项目点代码是全局唯一的）
  const allItOkCodes = new Set();
  for (const [key, entry] of byCompany) {
    for (const code of entry.itOkCodes) allItOkCodes.add(code);
  }

	// 待转单 → 客户订单号 → 转单状态（主匹配）；销售订单号 → 转单状态（兜底匹配）
	const pendingMap = new Map();
	const pendingBySalesNo = new Map();
		for (const row of pendingRows) {
		  // 只有创建人为"供应链管理员"的待转单才是AI转单
		  if (row.createdBy !== "供应链管理员") continue;
		  const custKey = normalizeText(row.customerOrderNo);
		  if (custKey && !pendingMap.has(custKey)) {
		    pendingMap.set(custKey, row.transferStatus);
		  }
		  const salesKey = normalizeText(row.salesOrderNo);
		  if (salesKey && !pendingBySalesNo.has(salesKey)) {
		    pendingBySalesNo.set(salesKey, row.transferStatus);
		  }
		}
	
	// 销售订单 → 匹配 IT已配置集合 + 待转单状态（两层匹配）
	for (const row of salesRows) {
	  if (!row.operationCompanyKey) continue;
	  if (!row.hotelCode || !allItOkCodes.has(row.hotelCode)) continue;
	
	  const entry = ensure(row.operationCompanyKey, row.operationCompany);
	  entry.orderTotal += 1;
	
	  // 先按客户订单号匹配，再按销售订单号兜底
	  const custKey = normalizeText(row.customerOrderNo);
	  let pendingStatus = pendingMap.get(custKey);
	  if (!pendingStatus) {
	    const salesKey = normalizeText(row.orderNo);
	    pendingStatus = pendingBySalesNo.get(salesKey);
	  }
	  if (pendingStatus) {
	    entry.orderAiTotal += 1;  // 匹配到AI待转单就算（已转+未转）
	    if (pendingStatus.includes("已转")) {
	      entry.orderAi += 1;     // 已转单
	    }
	  }
	}

  const summary = [...byCompany.values()].map((entry) => {
    const registered = entry.registeredCodes.size;
    const itOk = entry.itOkCodes.size;
    const remark = logSummary && logSummary[entry.operationCompanyKey]
      ? logSummary[entry.operationCompanyKey].remark
      : "";
    return {
      operationCompany: entry.operationCompany,
      operationCompanyKey: entry.operationCompanyKey,
      registeredCount: registered,
      itConfiguredCount: itOk,
      configRate: registered > 0 ? itOk / registered : null,
      orderTotal: entry.orderTotal,
      orderAiCount: entry.orderAi,
      orderAiTotal: entry.orderAiTotal,
      aiRate: entry.orderTotal > 0 ? entry.orderAi / entry.orderTotal : null,
      aiRateTotal: entry.orderTotal > 0 ? entry.orderAiTotal / entry.orderTotal : null,
      orderMethod: remark,
    };
  }).sort((a, b) => b.orderTotal - a.orderTotal || b.registeredCount - a.registeredCount);

  const totals = summary.reduce((acc, row) => {
    acc.registeredCount += row.registeredCount;
    acc.itConfiguredCount += row.itConfiguredCount;
    acc.orderTotal += row.orderTotal;
    acc.orderAiCount += row.orderAiCount;
    acc.orderAiTotal += row.orderAiTotal;
    return acc;
  }, { registeredCount: 0, itConfiguredCount: 0, orderTotal: 0, orderAiCount: 0, orderAiTotal: 0 });
  totals.configRate = totals.registeredCount > 0 ? totals.itConfiguredCount / totals.registeredCount : null;
  totals.aiRate = totals.orderTotal > 0 ? totals.orderAiCount / totals.orderTotal : null;
  totals.aiRateTotal = totals.orderTotal > 0 ? totals.orderAiTotal / totals.orderTotal : null;

  return { rows: summary, totals };
}

// 群维度汇总：按企业微信群名称聚合，用于看板「重点群监控」
function buildGroupSummary(salesRows, pendingRows, progressRows, logSummary) {
  const groups = new Map();

  // 从推进表收集群信息
  for (const row of progressRows) {
    const name = row.groupName;
    if (!name) continue;
    if (!groups.has(name)) {
      groups.set(name, {
        groupName: name,
        operationCompany: row.operationCompany || "",
        hotelCodes: new Set(),
        itConfigured: false,
      });
    }
    const g = groups.get(name);
    if (row.hotelCode) g.hotelCodes.add(normalizeText(row.hotelCode));
    if (row.itConfigured) g.itConfigured = true;
    if (row.operationCompany) g.operationCompany = row.operationCompany;
  }

  // 全局 IT已配置集合
  const allItOkCodes = new Set();
  for (const row of progressRows) {
    if (row.itConfigured && row.hotelCode) allItOkCodes.add(normalizeText(row.hotelCode));
  }

  // 待转单匹配：客户订单号（主）+ 销售订单号（兜底）
  const pendingMap = new Map();
  const pendingBySalesNo = new Map();
  for (const row of pendingRows) {
    if (row.createdBy !== "供应链管理员") continue;
    const ck = normalizeText(row.customerOrderNo);
    if (ck && !pendingMap.has(ck)) pendingMap.set(ck, row.transferStatus);
    const sk = normalizeText(row.salesOrderNo);
    if (sk && !pendingBySalesNo.has(sk)) pendingBySalesNo.set(sk, row.transferStatus);
  }

  // 统计每个群的订单
  for (const [name, g] of groups) {
    let orderTotal = 0, orderAi = 0, orderAiTotal = 0;
    if (g.itConfigured) {
      for (const sr of salesRows) {
        const hc = normalizeText(sr.hotelCode);
        if (!hc || !g.hotelCodes.has(hc)) continue;
        if (g.operationCompany) {
          const srCompany = normalizeText(sr.operationCompany || "").replace(/[（(]/g, "(").replace(/[）)]/g, ")");
          const grCompany = normalizeText(g.operationCompany).replace(/[（(]/g, "(").replace(/[）)]/g, ")");
          if (srCompany !== grCompany) continue;
        }
        orderTotal += 1;
        // 两层匹配：先客户订单号，再销售订单号兜底
        const custKey = normalizeText(sr.customerOrderNo);
        let pendingStatus = pendingMap.get(custKey);
        if (!pendingStatus) {
          const sk = normalizeText(sr.orderNo);
          pendingStatus = pendingBySalesNo.get(sk);
        }
        if (pendingStatus) {
          orderAiTotal += 1;
          if (pendingStatus.includes("已转")) orderAi += 1;
        }
      }
    }
    g.orderTotal = orderTotal;
    g.orderAiCount = orderAi;
    g.orderAiTotal = orderAiTotal;
    g.aiRate = orderTotal > 0 ? orderAi / orderTotal : null;
    g.aiRateTotal = orderTotal > 0 ? orderAiTotal / orderTotal : null;
    // 下单方式备注
    const companyKey = normalizeText(g.operationCompany).replace(/[（(]/g, "(").replace(/[）)]/g, ")");
    g.orderMethod = logSummary && logSummary[companyKey] ? logSummary[companyKey].remark : "";
    // 清理 Set（不可序列化）
    g.hotelCodeCount = g.hotelCodes.size;
    delete g.hotelCodes;
  }

  const rows = Array.from(groups.values())
    .sort((a, b) => b.orderTotal - a.orderTotal || b.hotelCodeCount - a.hotelCodeCount);

  return { rows };
}

function findFile(dir, pattern) {
  if (!fs.existsSync(dir)) return null;
  // 先找根目录（跳过临时文件和 Zone.Identifier）
  let best = null;
  for (const name of fs.readdirSync(dir)) {
    if (name.startsWith("~$") || name.includes(":Zone.Identifier")) continue;
    const fullPath = path.join(dir, name);
    if (fs.statSync(fullPath).isDirectory()) continue;
    if (name.includes(pattern)) best = fullPath;
  }
  // 再找子目录里匹配的（取最后一个，即按文件名字面序最大的，通常对应最新日期）
  for (const entry of fs.readdirSync(dir).sort()) {
    if (entry.startsWith("~$") || entry.includes(":Zone.Identifier")) continue;
    const fullPath = path.join(dir, entry);
    if (!fs.statSync(fullPath).isDirectory()) continue;
    const subMatch = findFile(fullPath, pattern);
    if (subMatch) best = subMatch;
  }
  return best;
}

function loadDefaultData(rootDir, cutoffDate) {
  const root = rootDir || path.join(__dirname, "..", "..");
  const basicDir = path.join(root, "basicData");
  const sampleDir = path.join(root, "示例数据");

  // 优先在匹配 cutoff 日期的子文件夹里找（如 --cutoff 0703 → 7月3日/）
  const month = String(parseInt(cutoffDate.slice(0, 2), 10));
  const day = String(parseInt(cutoffDate.slice(2, 4), 10));
  const dateDirName = `${month}月${day}日`;
  const dateSubDir = path.join(sampleDir, dateDirName);
  const searchDir = (fs.existsSync(dateSubDir)) ? dateSubDir : sampleDir;

  // 文件名按内容匹配（不要求精确日期），以便 7月2日/ 和 7月3日/ 子文件夹都能找到
  const salesPath = findFile(searchDir, "销售订单") || findFile(sampleDir, "销售订单");
  const pendingPath = findFile(searchDir, "待转单") || findFile(sampleDir, "待转单");
  const progressPath = findFile(searchDir, "企业微信AI转单推进表") || findFile(basicDir, "企业微信AI转单推进表");

  // 日志文件优先从 basicData 找，其次从根目录
  const logPath = findFile(basicDir, "微信日志") || findFile(root, "微信日志");

  const result = {
    root,
    sources: { salesPath, pendingPath, progressPath, logPath },
    salesWorkbook: null,
    pendingWorkbook: null,
    progressWorkbook: null,
    logWorkbook: null,
  };

  if (salesPath && fs.existsSync(salesPath)) result.salesWorkbook = readWorkbookFromPath(salesPath);
  if (pendingPath && fs.existsSync(pendingPath)) result.pendingWorkbook = readWorkbookFromPath(pendingPath);
  if (progressPath && fs.existsSync(progressPath)) result.progressWorkbook = readWorkbookFromPath(progressPath);
  if (logPath && fs.existsSync(logPath)) result.logWorkbook = readWorkbookFromPath(logPath);

  return result;
}

function buildPageData({ salesWorkbook, pendingWorkbook, progressWorkbook, logWorkbook, cutoffDate = DEFAULT_CUTOFF_DATE, sources = {} }) {
  const salesRows = salesWorkbook ? parseSalesFull(salesWorkbook) : [];
  const pendingRows = pendingWorkbook ? parsePendingWecom(pendingWorkbook) : [];
  const progressRows = progressWorkbook ? parseWecomProgress(progressWorkbook, cutoffDate) : [];

  const logSummary = parseWecomLogForSummary(logWorkbook || null, progressRows);
  const companySummary = buildCompanySummary(salesRows, pendingRows, progressRows, logSummary);
  const groupSummary = buildGroupSummary(salesRows, pendingRows, progressRows, logSummary);

  return {
    generatedAt: new Date().toISOString(),
    cutoffDate,
    sources,
    salesRows,
    pendingRows,
    progressRows,
    pendingTotals: calcPendingTotals(pendingRows),
    companySummary,
    groupSummary,
  };
}

module.exports = {
  normalizeText,
  normalizeCompanyName,
  getFirst,
  findHeaderIndex,
  rowsToObjects,
  readWorkbookFromPath,
  getActualSheetRange,
  sheetRows,
  DEFAULT_CUTOFF_DATE,
  isConfirmedByCutoff,
  parseSalesFull,
  parsePendingWecom,
  parseWecomProgress,
  calcPendingTotals,
  determineOrderMethod,
  buildRoomCompanyMap,
  parseWecomLogForSummary,
  buildCompanySummary,
  buildGroupSummary,
  findFile,
  loadDefaultData,
  buildPageData,
};
