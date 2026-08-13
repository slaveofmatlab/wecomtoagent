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

  const records = rowsToObjects(rows, headerIndex);

  // 容错：推进表有时"项目点代码"列表头丢失（表格里那一格是数据值而非标题），
  // rowsToObjects 会把那个数据值当成列名。检测首行的值形如 KHXMD/Biz 的列，自动修正。
  if (records.length > 0 && !records.some(r => normalizeText(r["项目点代码"] || ""))) {
    const firstRow = records[0];
    const guessedKey = Object.keys(firstRow).find(k =>
      k !== "项目点代码" && /^[A-Z]{2,}\d+$/i.test(normalizeText(firstRow[k]))
    );
    if (guessedKey) {
      records.forEach(r => { r["项目点代码"] = r[guessedKey]; });
    }
  }

  // 暂时不纳入统计口径的群，自动排除，不用每天手动标"删除"
  const PERMANENTLY_EXCLUDED_GROUPS = new Set([
    "粮贝铁路生鲜订单群",
  ]);

  return records.map((row, index) => {
    const operationCompany = getFirst(row, ["运营公司"]);
    const joinStatus = getFirst(row, ["是否加群-张利拉", "是否加群"]);
    const itStatus = getFirst(row, ["IT是否配置完成-邓虎", "IT是否配置完成"]);
    const deleteMark = deleteMarkKey ? normalizeText(row[deleteMarkKey]) : "";
    const groupName = getFirst(row, ["企业微信群名称"]);
    return {
      rowIndex: index + 1,
      groupName,
      groupId: getFirst(row, ["群ID"]),
      operationCompany,
      operationCompanyKey: normalizeCompanyName(operationCompany),
      hotelCode: getFirst(row, ["项目点代码"]),
      hotelName: getFirst(row, ["项目点名称"]),
      joined: isConfirmedByCutoff(joinStatus, cutoffDate),
      itConfigured: isConfirmedByCutoff(itStatus, cutoffDate),
      deleted: deleteMark.includes("删除") || PERMANENTLY_EXCLUDED_GROUPS.has(groupName),
    };
  })
    .filter((row) => !row.deleted)
    .filter((row) => row.operationCompany || row.hotelCode);
}

// 小程序下单的项目点：订单不存在失败，统计时算100% AI成功
const MINI_PROGRAM_CODES = new Set([
  "KHXMD10235", // 大兴工厂（北京丰厨转单群）
]);

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

/**
 * 群 → 下单方式分类（9 种），与 index.html priorityCategory() 完全一致。
 * 返回 null 表示该群不在 priority_groups.json 配置中（仅"全部"视图出现）。
 * @param {string} groupName 群名
 * @param {object} groupsConfig priority_groups.json 的 groups 字段
 * @returns {string|null} 机器人|手写|图片下单|混合|Excel下单|PDF下单|文本消息|小程序|其他|null
 */
function classifyGroupCategory(groupName, groupsConfig) {
  if ((groupName || "").indexOf("机器人接单群") >= 0) return "机器人";
  const c = groupsConfig[groupName];
  if (!c) return null;
  if (c.category) return c.category;
  if (c.tag === "非标准") return "手写";
  if (c.tag !== "标准") return null;
  const m = c.mainMethod || "";
  if (!m) return "其他";
  if (m.indexOf("图片下单") === 0) return "图片下单";
  if (m.indexOf("图文混发") === 0 || /\d+%/.test(m)) return "混合";
  const known = ["Excel下单", "PDF下单", "文本消息"];
  if (known.indexOf(m) >= 0) return m;
  return "混合";
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

  // 构建全局 IT已配置 集合 + 项目点 → 推进表运营公司 映射（项目点代码全局唯一）。
  // 同一项目点在销售订单表和推进表里的运营公司可能填得不一致（如"郑州/丰厨(上海)"
  // vs "上海/丰厨(上海)"），订单归属以推进表为准，与群维度同口径，两边合计才对得上。
  const allItOkCodes = new Set();
  const codeToCompany = new Map();
  for (const row of progressRows) {
    if (!row.operationCompanyKey || !row.hotelCode || !row.itConfigured) continue;
    allItOkCodes.add(row.hotelCode);
    codeToCompany.set(row.hotelCode, { key: row.operationCompanyKey, name: row.operationCompany });
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
	
	// 销售订单 → 匹配 IT已配置集合 + 待转单状态（两层匹配）。
	// 归属公司用推进表登记的运营公司（codeToCompany），不用销售表里的写法。
	for (const row of salesRows) {
	  if (!row.operationCompanyKey) continue;
	  if (!row.hotelCode || !allItOkCodes.has(row.hotelCode)) continue;

	  const owner = codeToCompany.get(row.hotelCode);
	  const entry = ensure(owner.key, owner.name);
	  entry.orderTotal += 1;

	  // 小程序项目点的订单不存在失败，算100% AI成功
	  if (MINI_PROGRAM_CODES.has(row.hotelCode)) {
	    entry.orderAiTotal += 1;
	    entry.orderAi += 1;
	  } else {
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

// 群维度汇总：按企业微信群名称聚合，用于看板「重点群监控」。
// 订单归属完全按项目点代码（推进表登记），与 buildCompanySummary 同口径、同条件
// （IT已配置 + 推进表有运营公司），保证「全部群」合计与公司汇总总计严格相等。
function buildGroupSummary(salesRows, pendingRows, progressRows, logSummary) {
  const groups = new Map();

  const ensureGroup = (name) => {
    if (!groups.has(name)) {
      groups.set(name, {
        groupName: name,
        operationCompany: "",
        operationCompanyKey: "",
        hotelCodes: new Set(),    // 全部项目点（展示用）
        itConfigured: false,
        orderTotal: 0,
        orderAiCount: 0,
        orderAiTotal: 0,
      });
    }
    return groups.get(name);
  };

  // 从推进表收集群信息
  for (const row of progressRows) {
    if (!row.groupName) continue;
    const g = ensureGroup(row.groupName);
    if (row.hotelCode) g.hotelCodes.add(row.hotelCode);
    if (row.itConfigured) g.itConfigured = true;
    if (row.operationCompany) { g.operationCompany = row.operationCompany; g.operationCompanyKey = row.operationCompanyKey || ""; }
  }

  // 每个 IT已配置项目点唯一归属到一个群（同一项目点登记在多个群时取推进表里最后
  // 出现的），避免同一订单行在多个群里重复计数。不校验销售表运营公司与群登记公司
  // 是否一致——项目点代码全局唯一，公司名两表可能填得不同（如郑州/上海丰厨）。
  // 条件与 buildCompanySummary 的 allItOkCodes 完全一致；推进表没填群名的配置行
  // 归入兜底桶，保证不漏行。
  const UNNAMED_GROUP = "（推进表未填群名）";
  const codeToGroup = new Map();
  for (const row of progressRows) {
    if (!row.operationCompanyKey || !row.hotelCode || !row.itConfigured) continue;
    codeToGroup.set(row.hotelCode, row.groupName || UNNAMED_GROUP);
  }

  // 待转单匹配：客户订单号（主）+ 销售订单号（兜底），与公司汇总相同
  const pendingMap = new Map();
  const pendingBySalesNo = new Map();
  for (const row of pendingRows) {
    if (row.createdBy !== "供应链管理员") continue;
    const ck = normalizeText(row.customerOrderNo);
    if (ck && !pendingMap.has(ck)) pendingMap.set(ck, row.transferStatus);
    const sk = normalizeText(row.salesOrderNo);
    if (sk && !pendingBySalesNo.has(sk)) pendingBySalesNo.set(sk, row.transferStatus);
  }

  // 单次遍历销售订单，按项目点代码归到对应群
  for (const sr of salesRows) {
    if (!sr.operationCompanyKey) continue;
    const groupName = sr.hotelCode ? codeToGroup.get(sr.hotelCode) : undefined;
    if (!groupName) continue;
    const g = ensureGroup(groupName);
    g.orderTotal += 1;
    // 小程序项目点的订单算100% AI成功
    if (MINI_PROGRAM_CODES.has(sr.hotelCode)) {
      g.orderAiTotal += 1;
      g.orderAiCount += 1;
    } else {
      // 两层匹配：先客户订单号，再销售订单号兜底
      const custKey = normalizeText(sr.customerOrderNo);
      let pendingStatus = pendingMap.get(custKey);
      if (!pendingStatus) {
        pendingStatus = pendingBySalesNo.get(normalizeText(sr.orderNo));
      }
      if (pendingStatus) {
        g.orderAiTotal += 1;
        if (pendingStatus.includes("已转")) g.orderAiCount += 1;
      }
    }
  }

  for (const g of groups.values()) {
    g.aiRate = g.orderTotal > 0 ? g.orderAiCount / g.orderTotal : null;
    g.aiRateTotal = g.orderTotal > 0 ? g.orderAiTotal / g.orderTotal : null;
    // 下单方式备注
    const companyKey = normalizeCompanyName(g.operationCompany);
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
  // 先找根目录（跳过临时文件和 Zone.Identifier）；按文件名排序取最后一个匹配，
  // 不依赖系统 readdir 顺序，保证多份同名变体（如"- 副本"）时结果确定
  let best = null;
  for (const name of fs.readdirSync(dir).sort()) {
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
  const month = parseInt(cutoffDate.slice(0, 2), 10);
  const day = parseInt(cutoffDate.slice(2, 4), 10);
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

  // 日志统计（含完整 methodStats，供导出 Excel 用）
  const logStats = logSummary ? buildLogStats(logSummary, progressRows) : null;

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
    logStats,
    logSummary: logSummary || {},
  };
}

/**
 * 按下单方式 × 运营公司 生成交叉统计报表数据。
 * 数据来源：微信日志（消息类型分布）+ 销售订单（订单行数/AI转单）。
 * 订单行数按各公司消息类型比例分配到各下单方式。
 */
function buildOrderMethodReport(salesRows, pendingRows, progressRows, logSummary, companySummaryRows) {
  // 所有下单方式（固定顺序，与参考模板一致）
  var ALL_METHODS = ["图片下单", "PDF下单", "文本消息", "Excel下单", "混合", "手写"];

  // 将日志里的方法标签归一到 ALL_METHODS
  function normalizeMethod(label) {
    if (!label) return null;
    if (label === "图文混发") return "混合";
    if (ALL_METHODS.indexOf(label) >= 0) return label;
    // Word下单、图片文件、文件下单(.xxx)等→归入混合
    if (label.indexOf("Word") >= 0 || label.indexOf("文件下单") >= 0 || label.indexOf("图片文件") >= 0) return "混合";
    return null; // 不归入任何已知类
  }

  // 从 orderMethod 备注字符串反推下单方式分布（如 "图片下单 75%，PDF下单 19%"）
  // 用于 logSummary 缺失时的 fallback
  function parseRemarkToMethods(remark) {
    var result = {};
    if (!remark) return result;
    var parts = remark.split(/[，,]/);
    for (var i = 0; i < parts.length; i++) {
      var m = parts[i].match(/(.+?)\s*(\d+)%?/);
      if (m) {
        var method = normalizeMethod(m[1].trim());
        var pct = parseInt(m[2], 10);
        if (method && pct > 0) result[method] = pct;
      }
    }
    // 如果只有一种方法且没带百分比，直接100%
    if (Object.keys(result).length === 0 && remark.trim()) {
      var singleMethod = normalizeMethod(remark.trim());
      if (singleMethod) result[singleMethod] = 100;
    }
    return result;
  }

  // 从 companySummary 构建 fallback method 分布
  var remarkFallback = {};
  if (companySummaryRows && companySummaryRows.length) {
    for (var ri = 0; ri < companySummaryRows.length; ri++) {
      var cr = companySummaryRows[ri];
      if (cr.orderMethod && cr.operationCompanyKey) {
        remarkFallback[cr.operationCompanyKey] = parseRemarkToMethods(cr.orderMethod);
      }
    }
  }

  // --- 1. 从 logSummary 提取各公司消息类型分布 ---
  // companyKey → { methodLabel → messageCount }
  var companyMethodMsgs = {};
  var companyTotalMsgs = {};
  for (var ck in logSummary) {
    var stats = logSummary[ck].methodStats || [];
    if (stats.length === 0) {
      // fallback: 从 orderMethod 备注反推
      var fb = remarkFallback[ck];
      if (fb && Object.keys(fb).length > 0) {
        companyMethodMsgs[ck] = {};
        companyTotalMsgs[ck] = 0;
        for (var m in fb) {
          companyMethodMsgs[ck][m] = fb[m];
          companyTotalMsgs[ck] += fb[m];
        }
        continue;
      }
    }
    companyMethodMsgs[ck] = {};
    companyTotalMsgs[ck] = 0;
    for (var i = 0; i < stats.length; i++) {
      var label = normalizeMethod(stats[i].label);
      if (!label) continue;
      var count = stats[i].count;
      companyMethodMsgs[ck][label] = (companyMethodMsgs[ck][label] || 0) + count;
      companyTotalMsgs[ck] += count;
    }
  }

  // --- 2. 从销售数据 + 推进表 计算各公司订单行数和 AI 转单数 ---
  // 构建 IT已配置项目点集合 + 项目点→公司映射（与 buildCompanySummary 同口径）
  var allItOkCodes = new Set();
  var codeToCompany = new Map();
  for (var pi = 0; pi < progressRows.length; pi++) {
    var pr = progressRows[pi];
    if (!pr.operationCompanyKey || !pr.hotelCode || !pr.itConfigured) continue;
    allItOkCodes.add(pr.hotelCode);
    codeToCompany.set(pr.hotelCode, { key: pr.operationCompanyKey, name: pr.operationCompany });
  }

  // 待转单匹配
  var pendingMap = new Map();
  var pendingBySalesNo = new Map();
  for (var pi2 = 0; pi2 < pendingRows.length; pi2++) {
    var pe = pendingRows[pi2];
    if (pe.createdBy !== "供应链管理员") continue;
    var ck2 = normalizeText(pe.customerOrderNo);
    if (ck2 && !pendingMap.has(ck2)) pendingMap.set(ck2, pe.transferStatus);
    var sk = normalizeText(pe.salesOrderNo);
    if (sk && !pendingBySalesNo.has(sk)) pendingBySalesNo.set(sk, pe.transferStatus);
  }

  // 按公司累计：总订单行数、AI已转行数
  var companyOrders = {};  // ck → { total, ai }
  for (var si = 0; si < salesRows.length; si++) {
    var sr = salesRows[si];
    if (!sr.operationCompanyKey) continue;
    if (!sr.hotelCode || !allItOkCodes.has(sr.hotelCode)) continue;
    var owner = codeToCompany.get(sr.hotelCode);
    if (!owner) continue;
    var ck = owner.key;
    if (!companyOrders[ck]) companyOrders[ck] = { total: 0, ai: 0 };
    companyOrders[ck].total += 1;
    if (MINI_PROGRAM_CODES.has(sr.hotelCode)) {
      companyOrders[ck].ai += 1;
    } else {
      var custKey = normalizeText(sr.customerOrderNo);
      var ps = pendingMap.get(custKey);
      if (!ps) ps = pendingBySalesNo.get(normalizeText(sr.orderNo));
      if (ps && ps.includes("已转")) companyOrders[ck].ai += 1;
    }
  }

  // --- 3. 按比例分配订单行数到各下单方式 ---
  // companyKey → { method → { orderLines, aiLines } }
  var companyMethodOrders = {};
  var methodTotals = {};  // 汇总：各下单方式的总订单行数/AI
  for (var mi = 0; mi < ALL_METHODS.length; mi++) {
    methodTotals[ALL_METHODS[mi]] = { orderLines: 0, aiLines: 0 };
  }
  var grandOrderLines = 0;
  var grandAiLines = 0;

  // 收集所有有订单的公司（按订单行数排序，含无日志的公司）
  var allCompanyKeys = Object.keys(companyOrders).sort(function (a, b) {
    return companyOrders[b].total - companyOrders[a].total;
  });

  for (var ai2 = 0; ai2 < allCompanyKeys.length; ai2++) {
    var ck3 = allCompanyKeys[ai2];
    var ord = companyOrders[ck3];
    var msgs = companyMethodMsgs[ck3];
    var totalMsgs = companyTotalMsgs[ck3] || 0;
    companyMethodOrders[ck3] = {};

    if (!msgs || totalMsgs === 0) {
      // 尝试从 orderMethod 备注反推（无需日志文件）
      var fb = remarkFallback[ck3];
      if (fb && Object.keys(fb).length > 0) {
        msgs = fb;
        totalMsgs = 0;
        for (var fk in fb) { totalMsgs += fb[fk]; }
        // 更新 companyMethodMsgs 以便后续使用
        companyMethodMsgs[ck3] = fb;
        companyTotalMsgs[ck3] = totalMsgs;
      } else {
        // 无任何数据：各下单方式列为 0
        for (var mj = 0; mj < ALL_METHODS.length; mj++) {
          companyMethodOrders[ck3][ALL_METHODS[mj]] = { orderLines: 0, aiLines: 0 };
        }
        continue;
      }
    }

    for (var mj2 = 0; mj2 < ALL_METHODS.length; mj2++) {
      var method = ALL_METHODS[mj2];
      var msgCount = msgs[method] || 0;
      var proportion = totalMsgs > 0 ? msgCount / totalMsgs : 0;
      var orderLines = Math.round(ord.total * proportion);
      var aiLines = Math.round(ord.ai * proportion);

      companyMethodOrders[ck3][method] = { orderLines: orderLines, aiLines: aiLines };
      methodTotals[method].orderLines += orderLines;
      methodTotals[method].aiLines += aiLines;
    }
    grandOrderLines += ord.total;
    grandAiLines += ord.ai;
  }

  // --- 4. 公司名（展示用）---
  var companyNames = {};
  for (var pi3 = 0; pi3 < progressRows.length; pi3++) {
    var pr2 = progressRows[pi3];
    if (pr2.operationCompanyKey && pr2.operationCompany) {
      companyNames[pr2.operationCompanyKey] = pr2.operationCompany;
    }
  }

  // --- 5. 组装返回数据 ---
  // 左表：按方法汇总
  var leftRows = [];
  for (var mk = 0; mk < ALL_METHODS.length; mk++) {
    var m = ALL_METHODS[mk];
    var mt = methodTotals[m];
    var aiRate = mt.orderLines > 0 ? Math.round(mt.aiLines / mt.orderLines * 100) + "%" : "0%";
    leftRows.push({
      method: m,
      orderLines: mt.orderLines,
      aiLines: mt.aiLines,
      aiRate: aiRate,
    });
  }
  var leftTotal = {
    method: "总计",
    orderLines: grandOrderLines,
    aiLines: grandAiLines,
    aiRate: grandOrderLines > 0 ? Math.round(grandAiLines / grandOrderLines * 100) + "%" : "0%",
  };

  // 右表：公司 × 下单方式
  var rightCompanies = [];
  var rightTotals = { summary: { orderLines: 0, aiLines: 0 }, methods: {} };
  for (var mn = 0; mn < ALL_METHODS.length; mn++) {
    rightTotals.methods[ALL_METHODS[mn]] = { orderLines: 0, aiLines: 0 };
  }
  for (var ai3 = 0; ai3 < allCompanyKeys.length; ai3++) {
    var ck4 = allCompanyKeys[ai3];
    var ord2 = companyOrders[ck4];
    var methodData = companyMethodOrders[ck4] || {};
    var displayName = companyNames[ck4] || ck4;
    var compAiRate = ord2.total > 0 ? Math.round(ord2.ai / ord2.total * 100) + "%" : "0%";

    var compEntry = {
      company: displayName,
      companyKey: ck4,
      summary: {
        orderLines: ord2.total,
        aiLines: ord2.ai,
        aiRate: compAiRate,
      },
      methods: {},
    };

    // 各方法
    var compMethodTotal = 0;
    var compMethodAiTotal = 0;
    for (var mp = 0; mp < ALL_METHODS.length; mp++) {
      var m2 = ALL_METHODS[mp];
      var md = methodData[m2] || { orderLines: 0, aiLines: 0 };
      var rate = md.orderLines > 0 ? Math.round(md.aiLines / md.orderLines * 100) + "%" : "0%";
      compEntry.methods[m2] = {
        orderLines: md.orderLines,
        aiLines: md.aiLines,
        aiRate: rate,
      };
      compMethodTotal += md.orderLines;
      compMethodAiTotal += md.aiLines;
      rightTotals.methods[m2].orderLines += md.orderLines;
      rightTotals.methods[m2].aiLines += md.aiLines;
    }
    rightTotals.summary.orderLines += ord2.total;
    rightTotals.summary.aiLines += ord2.ai;

    rightCompanies.push(compEntry);
  }
  rightTotals.summary.aiRate = rightTotals.summary.orderLines > 0
    ? Math.round(rightTotals.summary.aiLines / rightTotals.summary.orderLines * 100) + "%"
    : "0%";
  for (var mq = 0; mq < ALL_METHODS.length; mq++) {
    var m3 = ALL_METHODS[mq];
    var rt = rightTotals.methods[m3];
    rt.aiRate = rt.orderLines > 0 ? Math.round(rt.aiLines / rt.orderLines * 100) + "%" : "0%";
  }

  return {
    methods: ALL_METHODS,
    leftRows: leftRows,
    leftTotal: leftTotal,
    rightCompanies: rightCompanies,
    rightTotals: rightTotals,
  };
}

/**
 * 从 logSummary 提取统计摘要（消息类型分布汇总）。
 */
function buildLogStats(logSummary, progressRows) {
  var methodTotalMsgs = {};
  var totalMsgs = 0;
  var totalAccepted = 0;
  for (var ck in logSummary) {
    var stats = logSummary[ck].methodStats || [];
    for (var i = 0; i < stats.length; i++) {
      var label = stats[i].label;
      var count = stats[i].count;
      if (!methodTotalMsgs[label]) methodTotalMsgs[label] = 0;
      methodTotalMsgs[label] += count;
      totalMsgs += count;
      if (stats[i].processed) totalAccepted += count;
    }
  }
  return {
    methodCounts: methodTotalMsgs,
    totalMessages: totalMsgs,
    totalAccepted: totalAccepted,
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
  classifyGroupCategory,
  determineOrderMethod,
  buildRoomCompanyMap,
  parseWecomLogForSummary,
  buildCompanySummary,
  buildGroupSummary,
  findFile,
  loadDefaultData,
  buildPageData,
  buildOrderMethodReport,
  buildLogStats,
};
