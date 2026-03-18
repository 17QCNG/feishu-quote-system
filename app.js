// app.js
(function () {
  function setMsg(text, type) {
    var el = document.getElementById("message");
    if (!el) return;
    el.textContent = text;
    el.className = "msg " + type;
  }

  function clearMsg() {
    var el = document.getElementById("message");
    if (!el) return;
    el.textContent = "";
    el.className = "msg";
  }

  function toPlainText(v) {
    if (v == null) return "";
    if (typeof v === "string") return v;
    if (typeof v === "number") return String(v);

    if (Array.isArray(v)) {
      var parts = [];
      for (var i = 0; i < v.length; i++) {
        var x = v[i];
        if (x && typeof x === "object" && typeof x.text === "string")
          parts.push(x.text);
        else if (x != null) parts.push(String(x));
      }
      return parts.join("");
    }

    if (typeof v === "object") {
      if (typeof v.text === "string") return v.text;

      if (v && (typeof v.value === "string" || typeof v.value === "number"))
        return String(v.value);
      if (v && typeof v.name === "string") return v.name;

      try {
        return JSON.stringify(v);
      } catch (e) {
        return "";
      }
    }

    return String(v);
  }

  function toNumber(v) {
    if (v == null) return 0;
    if (typeof v === "number") return v;
    var s = toPlainText(v);
    var n = Number(s);
    return isFinite(n) ? n : 0;
  }

  function getBitableMaybe() {
    return window.bitable || (window.lark && window.lark.bitable) || null;
  }

  function waitForBitable(maxMs) {
    return new Promise(function (resolve, reject) {
      var start = Date.now();
      var timer = setInterval(function () {
        var bt = getBitableMaybe();
        if (bt && bt.base) {
          clearInterval(timer);
          resolve(bt);
          return;
        }
        if (Date.now() - start > maxMs) {
          clearInterval(timer);
          reject(
            new Error(
              "飞书 SDK 未加载或不可用。window.bitable=" +
                (window.bitable ? "yes" : "no") +
                ", window.lark.bitable=" +
                (window.lark && window.lark.bitable ? "yes" : "no")
            )
          );
        }
      }, 80);
    });
  }

  function normalizeName(s) {
    return String(s || "").trim();
  }

  function parseTableName(fullName) {
    var s = String(fullName || "").trim();
    if (!s) return { major: "", supplier: "" };

    // 支持：物料搭建（建桥） 或 物料搭建(建桥)
    var m = s.match(/^(.+?)\s*[（(]\s*(.+?)\s*[）)]\s*$/);
    if (m)
      return {
        major: String(m[1] || "").trim(),
        supplier: String(m[2] || "").trim(),
      };

    return { major: s, supplier: "" };
  }

  function escapeHtml(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  async function getFieldByAnyName(table, candidates, required) {
    var tried = [];
    for (var i = 0; i < candidates.length; i++) {
      var n = normalizeName(candidates[i]);
      if (!n) continue;
      tried.push(n);
      try {
        var f = await table.getFieldByName(n);
        if (f) return { field: f, picked: n };
      } catch (e) {}
    }
    if (required) {
      throw new Error("字段不存在：" + tried.join("/") + "（请检查字段名是否一致）");
    }
    return { field: null, picked: "" };
  }

  async function getAllRecords(table, pageSize) {
    var all = [];
    var token = undefined;
    var ps = Math.min(Math.max(Number(pageSize) || 200, 1), 500);

    for (;;) {
      var res = await table.getRecords({ pageSize: ps, pageToken: token });
      var batch = res && res.records ? res.records : [];
      all = all.concat(batch);

      if (!res || !res.hasMore) break;
      token = res.pageToken;
      if (!token) break;
    }

    return all;
  }

  async function getAllTables(bitable) {
    if (bitable.base && typeof bitable.base.getTableMetaList === "function") {
      var metaList = await bitable.base.getTableMetaList();
      return metaList || [];
    }

    var list = await bitable.base.getTableList();
    var metas = [];
    for (var i = 0; i < list.length; i++) {
      var t = list[i];
      var name = "";
      if (t && typeof t.getName === "function") name = await t.getName();
      metas.push({ id: t.id, name: name || t.id });
    }
    return metas;
  }

  function downloadBlob(filename, blob) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function normalizeNewlines(s) {
    return String(s || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  }

  function safeFileName(name) {
    var s = String(name || "").trim();
    if (!s) return "报价单";
    s = s.replace(/[\\/:*?"<>|]+/g, "_").trim();
    return s || "报价单";
  }

  function getQuoteNameInput() {
    return document.getElementById("quoteFileName");
  }

  function loadQuoteName() {
    try {
      return localStorage.getItem("quote_filename") || "";
    } catch (e) {
      return "";
    }
  }

  function saveQuoteName(v) {
    try {
      localStorage.setItem("quote_filename", String(v || ""));
    } catch (e) {}
  }

  // 中文为主：把 CJK 大致按 2 个“宽度单位”估算，更接近 Excel 换行效果
  function visualUnits(s) {
    var str = String(s || "");
    var units = 0;
    for (var i = 0; i < str.length; i++) {
      var ch = str.charAt(i);
      if (/[\u2E80-\u9FFF\uF900-\uFAFF]/.test(ch)) units += 2;
      else units += 1;
    }
    return units;
  }

  function estimateLines(text, colWidthChars) {
    var s = normalizeNewlines(text);
    if (!s) return 1;

    var width = Math.max(1, Math.floor(Number(colWidthChars) || 9));
    var parts = s.split("\n");
    var total = 0;

    for (var i = 0; i < parts.length; i++) {
      var p = parts[i] || "";
      var lenUnits = visualUnits(p);
      total += Math.max(1, Math.ceil(lenUnits / width));
    }

    return Math.max(1, total);
  }

  function rowHeightByLines(maxLines) {
    if (maxLines <= 1) return 20;
    if (maxLines === 2) return 40;
    return 50;
  }

  function toCnIndex(n) {
    var num = Number(n) || 0;
    if (num <= 0) return "";
    var d = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"];
    if (num < 10) return d[num];
    if (num === 10) return "十";
    if (num < 20) return "十" + d[num % 10];
    var tens = Math.floor(num / 10);
    var ones = num % 10;
    return d[tens] + "十" + (ones ? d[ones] : "");
  }

  function lower(s) {
    return String(s || "").toLowerCase();
  }

  function splitTokens(q) {
    var s = normalizeName(q);
    if (!s) return [];
    var parts = s
      .replace(/[，,；;、|/]+/g, " ")
      .split(/\s+/g)
      .filter(Boolean);
    return parts;
  }

  function displaySupplierName(opt) {
    var sup = opt && opt.supplier ? String(opt.supplier) : "";
    return sup ? sup : "默认供应商";
  }

  // 保留旧函数（不再作为“跟随主视觉”的限定条件）
  function isFollowMajor(major) {
    var m = normalizeName(major);
    return m === "物料搭建" || m === "印刷制作";
  }

  function scoreTemplateByQuery(tpl, tokens) {
    if (!tokens || tokens.length === 0) return 1;
    var name = lower(tpl && tpl.name ? tpl.name : "");
    var sizeDays = lower(tpl && tpl.sizeDays ? tpl.sizeDays : "");
    var unit = lower(tpl && tpl.unit ? tpl.unit : "");
    var major = lower(tpl && tpl.major ? tpl.major : "");
    var desc = "";
    var suppliersText = "";

    if (tpl && tpl.options && tpl.options.length) {
      for (var i = 0; i < tpl.options.length; i++) {
        var opt = tpl.options[i];
        if (opt && opt.desc) desc += " " + lower(opt.desc);
        suppliersText += " " + lower(displaySupplierName(opt));
        if (desc.length > 1200) break;
      }
    }

    var hayName = " " + name + " ";
    var hayMeta = " " + major + " " + sizeDays + " " + unit + " ";
    var hayDesc = " " + desc + " ";
    var haySup = " " + suppliersText + " ";

    var total = 0;
    for (var t = 0; t < tokens.length; t++) {
      var token = lower(tokens[t]);
      if (!token) continue;

      var hit = 0;
      if (hayName.indexOf(token) >= 0) hit = Math.max(hit, 6);
      if (hayMeta.indexOf(token) >= 0) hit = Math.max(hit, 3);
      if (haySup.indexOf(token) >= 0) hit = Math.max(hit, 2);
      if (hayDesc.indexOf(token) >= 0) hit = Math.max(hit, 1);

      total += hit;
    }

    return total;
  }

  function safeSessionGet(key) {
    try {
      return sessionStorage.getItem(key);
    } catch (e) {
      return null;
    }
  }

  function safeSessionSet(key, val) {
    try {
      sessionStorage.setItem(key, val);
    } catch (e) {}
  }

  function safeSessionRemove(key) {
    try {
      sessionStorage.removeItem(key);
    } catch (e) {}
  }

  function runWithConcurrency(items, worker, limit) {
    var idx = 0;
    var inFlight = 0;
    var results = [];
    limit = Math.max(1, Number(limit) || 4);

    return new Promise(function (resolve) {
      function next() {
        while (inFlight < limit && idx < items.length) {
          (function (curIndex) {
            var item = items[curIndex];
            idx++;
            inFlight++;

            Promise.resolve()
              .then(function () {
                return worker(item, curIndex);
              })
              .then(function (r) {
                results[curIndex] = r;
              })
              .catch(function (e) {
                results[curIndex] = { __error: e };
              })
              .finally(function () {
                inFlight--;
                if (idx >= items.length && inFlight === 0) resolve(results);
                else next();
              });
          })(idx);
        }
      }
      next();
    });
  }

  function hashStr(s) {
    var str = String(s || "");
    var h = 0;
    for (var i = 0; i < str.length; i++) {
      h = (h * 31 + str.charCodeAt(i)) >>> 0;
    }
    return h >>> 0;
  }

  function majorToColor(major) {
    var palette = [
      "#3370FF",
      "#00B42A",
      "#FF7D00",
      "#F53F3F",
      "#86909C",
      "#722ED1",
      "#14C9C9",
      "#F7BA1E",
      "#1D2129",
      "#2B5DD4",
    ];
    var idx = hashStr(major) % palette.length;
    return palette[idx];
  }

  function buildProductTemplates(allItems) {
    var map = new Map();

    for (var i = 0; i < allItems.length; i++) {
      var it = allItems[i];

      // 需求5：产品选择中始终隐藏“主视觉喷绘”
      if (normalizeName(it.name) === "主视觉喷绘") continue;

      var k =
        normalizeName(it.name) +
        "||" +
        normalizeName(it.sizeDays) +
        "||" +
        normalizeName(it.unit);

      var tpl = map.get(k);
      if (!tpl) {
        tpl = {
          key: k,
          name: it.name || "未命名产品",
          sizeDays: it.sizeDays || "",
          unit: it.unit || "",
          major: it.sourceTableName || "",
          options: [],
        };
        map.set(k, tpl);
      }

      if (!tpl.major) tpl.major = it.sourceTableName || "";

      tpl.options.push({
        supplier: it.supplier || "",
        cost: Number(it.cost || 0),
        price: Number(it.price || 0),
        desc: it.desc || "",
        sourceTableName: it.sourceTableName || "",
      });
    }

    var arr = [];
    map.forEach(function (v) {
      arr.push(v);
    });

    arr.sort(function (a, b) {
      var an = normalizeName(a.name);
      var bn = normalizeName(b.name);
      if (an !== bn) return an.localeCompare(bn, "zh");
      return normalizeName(a.sizeDays).localeCompare(
        normalizeName(b.sizeDays),
        "zh"
      );
    });

    for (var j = 0; j < arr.length; j++) {
      arr[j].options.sort(function (x, y) {
        return normalizeName(x.supplier).localeCompare(
          normalizeName(y.supplier),
          "zh"
        );
      });
    }

    return arr;
  }

async function exportXlsx(selected, quoteName) {
  if (!window.ExcelJS || !window.ExcelJS.Workbook) {
    throw new Error(
      "ExcelJS 未加载：请检查 index.html 是否已引入 exceljs.min.js"
    );
  }

  var fileTitle = safeFileName(quoteName || "报价单");

  var wb = new window.ExcelJS.Workbook();
  wb.creator = "Feishu Quote Plugin";
  wb.created = new Date();

  var ws = wb.addWorksheet("报价单", {
    properties: { defaultRowHeight: 20 },
  });

  // 新需求：产品名称前加“类目”
  var headers = [
    "科目",
    "类目",
    "产品名称",
    "尺寸/天数",
    "数量",
    "单位",
    "成本单价",
    "成本总价",
    "单价",
    "总价",
    "产品描述",
  ];

  var colWidths = [8, 14, 30, 24, 9, 9, 9, 9, 9, 9, 35];
  for (var c = 1; c <= colWidths.length; c++) {
    ws.getColumn(c).width = colWidths[c - 1];
  }

  ws.addRow([fileTitle]);
  ws.mergeCells(1, 1, 1, headers.length);

  var titleCell = ws.getCell(1, 1);
  titleCell.fill = {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFED7D31" },
  };
  titleCell.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 16 };
  titleCell.alignment = {
    vertical: "middle",
    horizontal: "center",
    wrapText: true,
  };
  ws.getRow(1).height = 30;

  ws.addRow(headers);
  var headerRow = ws.getRow(2);
  headerRow.font = { bold: true };
  headerRow.alignment = {
    vertical: "middle",
    horizontal: "center",
    wrapText: true,
  };
  headerRow.height = 20;

  // 仍高亮“单价/总价”
  for (var hc = 9; hc <= 10; hc++) {
    ws.getCell(2, hc).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFFF00" },
    };
  }

  // 新需求：屏幕舞台/灯光/音响归为“多媒体搭建”
  var typeOrder = [
    "物料搭建",
    "印刷制作",
    "线上小程序",
    "系统设计",
    "系统适配",
    "人员执行",
    "硬件设备",
    "多媒体搭建",
    "摄影摄像",
    "视觉设计",
    "项目管理",
    "其他",
  ];
  var typeOrderMap = {};
  for (var oi = 0; oi < typeOrder.length; oi++) typeOrderMap[typeOrder[oi]] = oi;

  function canonicalType(rawType) {
    var t0 = normalizeName(rawType);
    var p = parseTableName(t0);
    var t = normalizeName(p.major || t0);

    // ✅ 需求2：归并
    if (t === "屏幕舞台" || t === "灯光" || t === "音响") return "多媒体搭建";

    if (Object.prototype.hasOwnProperty.call(typeOrderMap, t)) return t;
    return "其他";
  }

  // “类目列(B)”显示/合并所用：读取数据表的大类（不走归并）
  function rawMajorLabel(it) {
    var t0 = normalizeName(it && it.sourceTableName ? it.sourceTableName : "");
    if (!t0) return "其他";
    var p = parseTableName(t0);
    return normalizeName(p.major || t0) || "其他";
  }

  // 导出排序规则：主视觉喷绘（物料搭建）必须排在“物料搭建”大类最上方
  function isMainVisualExportItem(it) {
    if (!it) return false;
    var name = normalizeName(it.name || "");
    if (name !== "主视觉喷绘") return false;
    var t = canonicalType(it.sourceTableName || "");
    return t === "物料搭建";
  }

  var selectedArr = [];
  if (selected && typeof selected.forEach === "function") {
    selected.forEach(function (v) {
      selectedArr.push(v);
    });
  } else if (Array.isArray(selected)) {
    selectedArr = selected.slice();
  }

  // 需求2：同一大类内排序改为“先选择的在前”（__pickSeq 越小越靠前）
  // 新需求：主视觉喷绘在“物料搭建”大类内永远置顶（且按添加顺序）
  selectedArr.sort(function (a, b) {
    var at = canonicalType(a && a.sourceTableName ? a.sourceTableName : "");
    var bt = canonicalType(b && b.sourceTableName ? b.sourceTableName : "");

    var ai = Object.prototype.hasOwnProperty.call(typeOrderMap, at)
      ? typeOrderMap[at]
      : typeOrder.length;
    var bi = Object.prototype.hasOwnProperty.call(typeOrderMap, bt)
      ? typeOrderMap[bt]
      : typeOrderMap["其他"];
    if (ai !== bi) return ai - bi;

    // ✅ 同一大类内：主视觉喷绘（物料搭建）排在最前
    var amv = isMainVisualExportItem(a);
    var bmv = isMainVisualExportItem(b);
    if (amv !== bmv) return amv ? -1 : 1;

    // ✅ 主视觉喷绘之间：按添加顺序（__mvSeq 越小越靠前）
    if (amv && bmv) {
      var ams = a && typeof a.__mvSeq === "number" ? a.__mvSeq : 0;
      var bms = b && typeof b.__mvSeq === "number" ? b.__mvSeq : 0;
      if (ams !== bms) return ams - bms;
    }

    var as = a && typeof a.__pickSeq === "number" ? a.__pickSeq : 0;
    var bs = b && typeof b.__pickSeq === "number" ? b.__pickSeq : 0;
    if (as && bs && as !== bs) return as - bs;
    if (as && !bs) return -1;
    if (!as && bs) return 1;

    // 兜底：保持原来的稳定排序（名称/编号）
    var an = normalizeName(a && a.name ? a.name : "");
    var bn = normalizeName(b && b.name ? b.name : "");
    if (an !== bn) return an.localeCompare(bn, "zh");

    var ac = normalizeName(a && a.code ? a.code : "");
    var bc = normalizeName(b && b.code ? b.code : "");
    return ac.localeCompare(bc, "zh");
  });

  var rowNum = 3;

  var groupType = null;
  var groupIdx = 0;
  var groupSerial = 0;
  var groupStartRow = 0;
  var groupEndRow = 0;
  var subtotalRows = [];

  // 类目（B列）按“读取数据表大类”在组内分段合并
  var catLabel = null;
  var catStartRow = 0;
  var catEndRow = 0;

  function styleRowAllCols(rn, fillArgb, bold) {
    for (var cc = 1; cc <= headers.length; cc++) {
      var cell = ws.getCell(rn, cc);
      if (fillArgb) {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: fillArgb },
        };
      }
      if (bold) {
        cell.font = Object.assign({}, cell.font || {}, { bold: true });
      }
      cell.alignment = Object.assign({}, cell.alignment || {}, {
        vertical: "middle",
        horizontal: "center",
        wrapText: true,
      });
    }
  }

  function addGroupHeader(typeLabel) {
    ws.getRow(rowNum).height = 20;

    ws.getCell(rowNum, 1).value = toCnIndex(groupIdx);
    ws.mergeCells(rowNum, 2, rowNum, headers.length);
    ws.getCell(rowNum, 2).value = typeLabel;

    styleRowAllCols(rowNum, "FFBFBFBF", true);
    rowNum++;
  }

  // 合并 B 列某一段，并填上段标题（读取数据表的大类名称）
  function mergeCategoryCells(startRn, endRn, label) {
    if (!startRn || !endRn) return;
    if (startRn > endRn) return;

    ws.mergeCells(startRn, 2, endRn, 2);

    var cell = ws.getCell(startRn, 2);
    cell.value = label || "";
    cell.alignment = {
      vertical: "middle",
      horizontal: "center",
      wrapText: true,
    };
  }

  function addGroupSubtotal(startRn, endRn) {
    var subtotalRowNum = rowNum;

    ws.getRow(subtotalRowNum).height = 20;

    // 新增“类目”列后，"小计" 合并到 成本单价(第7列) 为止
    ws.mergeCells(subtotalRowNum, 1, subtotalRowNum, 7);
    ws.getCell(subtotalRowNum, 1).value = "小计";

    // 成本总价 = 第8列(H)，总价 = 第10列(J)
    ws.getCell(subtotalRowNum, 8).value = {
      formula: "SUM(H" + startRn + ":H" + endRn + ")",
    };
    ws.getCell(subtotalRowNum, 10).value = {
      formula: "SUM(J" + startRn + ":J" + endRn + ")",
    };

    styleRowAllCols(subtotalRowNum, null, true);

    ws.getCell(subtotalRowNum, 8).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFFF00" },
    };
    ws.getCell(subtotalRowNum, 10).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFFF00" },
    };

    subtotalRows.push(subtotalRowNum);
    rowNum++;
  }

  function addGrandTotalWithTax() {
    ws.getRow(rowNum).height = 20;

    ws.mergeCells(rowNum, 1, rowNum, 7);
    ws.getCell(rowNum, 1).value = "含税合计（税价6%）";

    if (subtotalRows.length === 0) {
      ws.getCell(rowNum, 8).value = 0;
      ws.getCell(rowNum, 10).value = 0;
    } else {
      var hRefs = [];
      var jRefs = [];
      for (var i = 0; i < subtotalRows.length; i++) {
        hRefs.push("H" + subtotalRows[i]);
        jRefs.push("J" + subtotalRows[i]);
      }

      ws.getCell(rowNum, 8).value = {
        formula: "SUM(" + hRefs.join(",") + ")",
      };
      ws.getCell(rowNum, 10).value = {
        formula: "SUM(" + jRefs.join(",") + ")*1.06",
      };
    }

    styleRowAllCols(rowNum, null, true);
    rowNum++;
  }

  for (var si = 0; si < selectedArr.length; si++) {
    var it = selectedArr[si];
    var typeLabel = canonicalType(it.sourceTableName || "");

    if (typeLabel !== groupType) {
      if (groupType != null) {
        // 结束上一科目组：先合并该组内最后一段“类目”，再写小计
        mergeCategoryCells(catStartRow, catEndRow, catLabel);
        addGroupSubtotal(groupStartRow, groupEndRow);
      }

      groupType = typeLabel;
      groupIdx++;
      groupSerial = 0;

      addGroupHeader(groupType);

      groupStartRow = rowNum;
      groupEndRow = rowNum - 1;

      // 新科目组开始：重置类目分段游标
      catLabel = null;
      catStartRow = 0;
      catEndRow = 0;
    }

    groupSerial++;

    // 额外需求：导出“产品描述”不追加供应商（保持纯描述）
    var desc = normalizeNewlines(it.desc || "");

    var rn = rowNum;
    var row = ws.getRow(rn);

    // 列结构：
    // A 科目(序号) | B 类目(按读取数据表大类分段合并写) | C 产品名称 | D 尺寸/天数 | E 数量 | F 单位 | G 成本单价 | H 成本总价 | I 单价 | J 总价 | K 产品描述
    row.getCell(1).value = groupSerial;
    row.getCell(2).value = ""; // 由 mergeCategoryCells 分段合并后写入
    row.getCell(3).value = it.name || "";
    row.getCell(4).value = it.sizeDays || "";
    row.getCell(5).value = Math.max(1, Number(it.qty) || 1);
    row.getCell(6).value = it.unit || "";
    row.getCell(7).value = Number(it.cost || 0);
    row.getCell(8).value = { formula: "G" + rn + "*E" + rn };
    row.getCell(9).value = Number(it.price || 0);
    row.getCell(10).value = { formula: "I" + rn + "*E" + rn };
    row.getCell(11).value = desc;

    // ===== 关键：B列类目按“读取数据表大类”分段合并 =====
    var curCat = rawMajorLabel(it);
    if (catLabel == null) {
      catLabel = curCat;
      catStartRow = rn;
      catEndRow = rn;
    } else if (curCat === catLabel) {
      catEndRow = rn;
    } else {
      mergeCategoryCells(catStartRow, catEndRow, catLabel);
      catLabel = curCat;
      catStartRow = rn;
      catEndRow = rn;
    }
    // ===============================================

    var maxLines = 1;
    for (var ci = 1; ci <= headers.length; ci++) {
      var cell = ws.getCell(rn, ci);
      cell.alignment = {
        vertical: "middle",
        horizontal: "center",
        wrapText: true,
      };

      var v = cell.value;
      if (v && typeof v === "object" && v.formula) continue;
      if (typeof v === "number") continue;

      var t = v == null ? "" : String(v);
      var w = ws.getColumn(ci).width || 9;
      maxLines = Math.max(maxLines, estimateLines(t, w));
    }

    row.height = rowHeightByLines(Math.min(3, maxLines));
    groupEndRow = rn;
    rowNum++;
  }

  if (groupType != null && groupStartRow <= groupEndRow) {
    // 收尾最后一组：合并最后一段类目 + 小计
    mergeCategoryCells(catStartRow, catEndRow, catLabel);
    addGroupSubtotal(groupStartRow, groupEndRow);
  }

  addGrandTotalWithTax();

  var lastRow = rowNum - 1;
  var borderThin = {
    top: { style: "thin" },
    left: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: "thin" },
  };

  for (var r = 1; r <= lastRow; r++) {
    for (var cc = 1; cc <= headers.length; cc++) {
      var ccell = ws.getCell(r, cc);
      ccell.border = borderThin;
      ccell.alignment = Object.assign({}, ccell.alignment || {}, {
        vertical: "middle",
        horizontal: "center",
        wrapText: true,
      });
    }
  }

  var buf = await wb.xlsx.writeBuffer();
  var blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });

  downloadBlob(fileTitle + ".xlsx", blob);
}

  async function main() {
    clearMsg();

    var productList = document.getElementById("productList");
    var exportBtn = document.getElementById("exportXlsx");
    var selectAllBtn = document.getElementById("selectAllTables");
    var clearAllBtn = document.getElementById("clearAllTables");
    var reloadBtn = document.getElementById("reloadProducts");
    var quoteNameInput = getQuoteNameInput();

    var viewPoolBtn = document.getElementById("viewPool");
    var viewSelectedBtn = document.getElementById("viewSelected");
    var productSearchInput = document.getElementById("productSearch");

    var needMainVisualCb = document.getElementById("needMainVisual");
    var mvBox = document.getElementById("mainVisualConfig");
    var mvSupplier = document.getElementById("mainVisualSupplier");
    var mvSize = document.getElementById("mainVisualSize");
    var mvQty = document.getElementById("mainVisualQty");
    var mvQtyUnit = document.getElementById("mainVisualQtyUnit");
    var mvAdd = document.getElementById("addMainVisual");
    var mvList = document.getElementById("mainVisualList");

    var majorDropdownBtn = document.getElementById("majorDropdownBtn");
    var majorDropdownPanel = document.getElementById("majorDropdownPanel");

    var selectAllVisibleProductsBtn = document.getElementById(
      "selectAllVisibleProducts"
    );
    var clearAllSelectedProductsBtn = document.getElementById(
      "clearAllSelectedProducts"
    );

    if (
      !productList ||
      !exportBtn ||
      !majorDropdownBtn ||
      !majorDropdownPanel
    ) {
      setMsg("页面元素缺失：请确认 index.html 与 app.js 已正确上传。", "err");
      return;
    }

    if (quoteNameInput) {
      var cached = loadQuoteName();
      if (cached) quoteNameInput.value = cached;
      quoteNameInput.addEventListener("input", function () {
        saveQuoteName(this.value || "");
      });
    }

    var bitable;
    try {
      bitable = await waitForBitable(5000);
    } catch (e) {
      setMsg(
        "飞书 SDK 未加载：请在【飞书多维表格】里通过【扩展/自定义插件】打开本页面，不要直接在浏览器访问。",
        "err"
      );
      return;
    }

    // === 单次操作临时暂存（sessionStorage）：已选产品 ===
    var DRAFT_KEY = "quote_selected_draft_v1";
    var selected = new Map();

    // 需求2：记录选择顺序（同一大类内按先选先排）
    var pickSeqCounter = 0;
    function nextPickSeq() {
      pickSeqCounter++;
      return pickSeqCounter;
    }
    function bumpPickSeqCounterFromSelected() {
      var max = 0;
      selected.forEach(function (v) {
        var s = v && typeof v.__pickSeq === "number" ? v.__pickSeq : 0;
        if (s > max) max = s;
      });
      pickSeqCounter = Math.max(pickSeqCounter, max);
    }

    // === 单次操作临时暂存（sessionStorage）：待选库预设（未勾选前的供应商/数量） ===
    var POOL_DRAFT_KEY = "quote_pool_draft_v1";
    var poolDraft = new Map();

    function persistPoolDraft() {
      var arr = [];
      poolDraft.forEach(function (v, key) {
        if (!key || !v) return;

        var qty = Math.max(1, Number(v.qty) || 1);
        var supplierTouched = !!v.supplierTouched;

        // 只存“有意义”的草稿，避免 sessionStorage 过大
        if (qty === 1 && !supplierTouched) return;

        var out = {
          key: String(key),
          qty: qty,
          supplierTouched: supplierTouched,
        };
        if (supplierTouched)
          out.supplierIndex = Math.max(0, Number(v.supplierIndex) || 0);
        arr.push(out);
      });

      if (arr.length === 0) safeSessionRemove(POOL_DRAFT_KEY);
      else safeSessionSet(POOL_DRAFT_KEY, JSON.stringify(arr));
    }

    function restorePoolDraft() {
      var raw = safeSessionGet(POOL_DRAFT_KEY);
      if (!raw) return;

      try {
        var arr = JSON.parse(raw);
        if (!Array.isArray(arr)) return;

        for (var i = 0; i < arr.length; i++) {
          var x = arr[i];
          if (!x || !x.key) continue;

          // 主视觉喷绘不走产品池
          if (String(x.key).indexOf("主视觉喷绘||") === 0) continue;

          var d = {
            qty: Math.max(1, Number(x.qty) || 1),
            supplierTouched: !!x.supplierTouched,
          };
          if (d.supplierTouched)
            d.supplierIndex = Math.max(0, Number(x.supplierIndex) || 0);

          poolDraft.set(String(x.key), d);
        }
      } catch (e) {}
    }

    function upsertPoolDraftQty(templateKey, qtyValue) {
      var d = poolDraft.get(templateKey) || { qty: 1, supplierTouched: false };
      d.qty = Math.max(1, Number(qtyValue) || 1);
      poolDraft.set(templateKey, d);
      persistPoolDraft();
    }

    function upsertPoolDraftSupplier(templateKey, supplierIndex) {
      var d = poolDraft.get(templateKey) || { qty: 1, supplierTouched: false };
      d.supplierTouched = true;
      d.supplierIndex = Math.max(0, Number(supplierIndex) || 0);
      poolDraft.set(templateKey, d);
      persistPoolDraft();
    }

    function resetPoolDraftAndPersist(templateKey) {
      if (!templateKey) return;
      poolDraft.delete(templateKey);
      persistPoolDraft();
    }

    function clearPoolDraftForSelectedAndPersist() {
      // 清空“已选产品”的预设（因为需求：取消勾选/清空时重置预设）
      selected.forEach(function (_, key) {
        poolDraft.delete(key);
      });
      persistPoolDraft();
    }

    function persistSelectedDraft() {
      var arr = [];
      selected.forEach(function (v, key) {
        arr.push({
          key: key,
          qty: v && v.qty != null ? Number(v.qty) : 1,
          supplier: v && v.supplier ? String(v.supplier) : "",
          __optIndex: v && typeof v.__optIndex === "number" ? v.__optIndex : 0,
          __supplierManual: !!(v && v.__supplierManual),
          __pickSeq: v && typeof v.__pickSeq === "number" ? v.__pickSeq : 0,
        });
      });
      if (arr.length === 0) safeSessionRemove(DRAFT_KEY);
      else safeSessionSet(DRAFT_KEY, JSON.stringify(arr));
    }

    function restoreSelectedDraft() {
      var raw = safeSessionGet(DRAFT_KEY);
      if (!raw) return [];

      try {
        var arr = JSON.parse(raw);
        if (!Array.isArray(arr)) return [];

        for (var i = 0; i < arr.length; i++) {
          var x = arr[i];
          if (!x || !x.key) continue;

          // 主视觉喷绘不走产品池
          if (String(x.key).indexOf("主视觉喷绘||") === 0) continue;

          var seq =
            typeof x.__pickSeq === "number" &&
            isFinite(x.__pickSeq) &&
            x.__pickSeq > 0
              ? x.__pickSeq
              : i + 1; // 兼容旧草稿（没有 seq 时按存储顺序补齐）

          selected.set(String(x.key), {
            name: "",
            sizeDays: "",
            unit: "",
            cost: 0,
            price: 0,
            desc: "",
            qty: Math.max(1, Number(x.qty) || 1),
            sourceTableName: "",
            supplier: x.supplier || "",
            __optIndex: Math.max(0, Number(x.__optIndex) || 0),
            __supplierManual: !!x.__supplierManual,
            __pickSeq: seq,
            __placeholder: true,
          });
        }

        bumpPickSeqCounterFromSelected();
        return arr;
      } catch (e) {
        return [];
      }
    }

    function setSelected(templateKey, item) {
      selected.set(templateKey, item);
      persistSelectedDraft();
    }

    function deleteSelected(templateKey) {
      selected.delete(templateKey);
      persistSelectedDraft();

      // 需求：取消勾选移除后，待选库预设要重置
      resetPoolDraftAndPersist(templateKey);
    }

    function clearSelected() {
      // 需求：清空已选后，待选库预设也重置
      clearPoolDraftForSelectedAndPersist();

      selected.clear();
      pickSeqCounter = 0;
      persistSelectedDraft();
    }

    // === 主视觉喷绘 & 供应商跟随逻辑 ===
    var mainVisualItems = []; // { supplier, size, qty }
    var mainVisualPriceBook = null; // Map supplier -> { cost, price, unit, desc, sourceTableName }
    var mainVisualDefaultSupplier = "";

    function updateMainVisualQtyUnit() {
      if (!mvQtyUnit) return;

      var unitText = "平方";
      if (mainVisualPriceBook && mvSupplier) {
        var sup = mvSupplier.value || "默认供应商";
        var pb = mainVisualPriceBook.get(sup);
        if (pb && pb.unit) unitText = String(pb.unit);
      }
      mvQtyUnit.textContent = unitText;
    }

    // --- 新规则：只要某产品“存在与主视觉相同的供应商选项”，就跟随主视觉供应商 ---
    function findOptionIndexBySupplierStrict(tplRef, supplierName) {
      if (!tplRef || !tplRef.options || tplRef.options.length === 0) return -1;

      var target = normalizeName(supplierName);
      if (!target) return -1;

      for (var i = 0; i < tplRef.options.length; i++) {
        var opt = tplRef.options[i];
        var disp = normalizeName(displaySupplierName(opt));
        if (disp === target) return i;
      }

      for (var j = 0; j < tplRef.options.length; j++) {
        var opt2 = tplRef.options[j];
        var raw = normalizeName(opt2 && opt2.supplier ? opt2.supplier : "");
        if (raw === target) return j;
      }

      if (target === "默认供应商") {
        for (var k = 0; k < tplRef.options.length; k++) {
          var opt3 = tplRef.options[k];
          if (!normalizeName(opt3 && opt3.supplier ? opt3.supplier : ""))
            return k;
        }
      }

      return -1;
    }

    // 兼容旧调用：找不到时返回 0
    function findOptionIndexBySupplier(tplRef, supplierName) {
      var idx = findOptionIndexBySupplierStrict(tplRef, supplierName);
      return idx >= 0 ? idx : 0;
    }

    function getFollowSupplierIndexForTemplate(tplRef) {
      var defSup = normalizeName(mainVisualDefaultSupplier || "");
      if (!defSup) return null;

      var idx = findOptionIndexBySupplierStrict(tplRef, defSup);
      if (idx < 0) return null;

      return idx;
    }

    function applyPick(
      templateKey,
      tplRef,
      supplierIndex,
      qtyValue,
      supplierManualFlag
    ) {
      var idx = Math.max(0, Number(supplierIndex) || 0);
      var opt2 =
        (tplRef && tplRef.options && tplRef.options[idx]) ||
        (tplRef && tplRef.options ? tplRef.options[0] : null);
      if (!opt2) return;

      var existed = selected.get(templateKey);
      var seq =
        existed && typeof existed.__pickSeq === "number" && existed.__pickSeq > 0
          ? existed.__pickSeq
          : nextPickSeq();

      var next = {
        name: tplRef.name,
        sizeDays: tplRef.sizeDays,
        unit: tplRef.unit,
        cost: Number(opt2.cost || 0),
        price: Number(opt2.price || 0),
        desc: opt2.desc || "",
        qty: Math.max(1, Number(qtyValue) || 1),
        sourceTableName: opt2.sourceTableName || "",
        supplier: opt2.supplier || "",
        __optIndex: idx,
        __supplierManual: !!supplierManualFlag,
        __pickSeq: seq,
      };

      setSelected(templateKey, next);
    }

    // 需求1/3：主视觉供应商变更时，所有“未手动指定供应商且包含该供应商选项”的已选产品自动切换
    function syncAutoSuppliersFromMainVisual(allTemplatesMap) {
      var defSup = normalizeName(mainVisualDefaultSupplier || "");
      if (!defSup) return;

      selected.forEach(function (v, key) {
        if (!v) return;
        if (v.__supplierManual) return;

        var tpl = allTemplatesMap.get(key);
        if (!tpl) return;

        var idxStrict = findOptionIndexBySupplierStrict(tpl, defSup);
        if (idxStrict < 0) return; // 该产品不存在此供应商选项，不跟随

        applyPick(key, tpl, idxStrict, v.qty, false);
      });
    }

    // === 读表结构：大类索引 ===
    var metas = [];
    var metasParsed = [];
    var majorIndex = new Map(); // major -> { major, metas: [...], suppliers:Set() }

    // === 全量产品模板 ===
    var allTemplates = [];
    var allTemplatesMap = new Map(); // templateKey -> template

    function rebuildAllTemplatesMap() {
      allTemplatesMap = new Map();
      for (var i = 0; i < allTemplates.length; i++) {
        allTemplatesMap.set(allTemplates[i].key, allTemplates[i]);
      }
    }

    // === 产品大类下拉多选状态 ===
    var majorSelected = new Set();

    function getSelectedMajors() {
      var arr = [];
      majorSelected.forEach(function (m) {
        arr.push(m);
      });
      return arr;
    }

    function toggleMajorDropdown(open) {
      var isOpen = majorDropdownPanel.style.display === "block";
      var next = open != null ? !!open : !isOpen;
      majorDropdownPanel.style.display = next ? "block" : "none";

      var arrow = majorDropdownBtn.querySelector(".major-btn-arrow");
      if (arrow) arrow.textContent = next ? "▲" : "▼";
    }

    function updateMajorBtnText() {
      var textEl = majorDropdownBtn.querySelector(".major-btn-text");
      if (!textEl) return;

      var arr = getSelectedMajors();
      arr.sort(function (a, b) {
        return String(a).localeCompare(String(b), "zh");
      });

      if (arr.length === 0) {
        textEl.textContent = "请选择产品大类";
        return;
      }

      var preview = arr.slice(0, 4).join("、");
      textEl.textContent =
        "已选 " +
        arr.length +
        " 个：" +
        preview +
        (arr.length > 4 ? "..." : "");
    }

    function renderMajorDropdownOptions() {
      majorDropdownPanel.innerHTML = "";

      if (!majorIndex || majorIndex.size === 0) {
        majorDropdownBtn.disabled = true;
        updateMajorBtnText();
        return;
      }
      majorDropdownBtn.disabled = false;

      var majors = [];
      majorIndex.forEach(function (v) {
        majors.push(v.major);
      });

      // 这里仍用 UI 的排序规则，不影响导出科目归并逻辑
      majors.sort(function (a, b) {
        return String(a).localeCompare(String(b), "zh");
      });

      for (var i = 0; i < majors.length; i++) {
        var major = majors[i];
        var bucket = majorIndex.get(major);

        var supplierArr = [];
        if (bucket && bucket.suppliers)
          bucket.suppliers.forEach(function (x) {
            supplierArr.push(x);
          });

        supplierArr.sort(function (x, y) {
          return String(x).localeCompare(String(y), "zh");
        });

        var supText = "";
        if (supplierArr.length > 0) {
          var preview = supplierArr.slice(0, 3).join("、");
          supText =
            "供应商：" + preview + (supplierArr.length > 3 ? "等" : "");
        }

        var item = document.createElement("div");
        item.className = "major-item";

        var cb = document.createElement("input");
        cb.type = "checkbox";
        cb.checked = majorSelected.has(major);

        var box = document.createElement("div");
        box.innerHTML =
          '<div class="major-label">' +
          escapeHtml(major) +
          "</div>" +
          (supText
            ? '<div class="major-sup">' + escapeHtml(supText) + "</div>"
            : "");

        item.appendChild(cb);
        item.appendChild(box);

        (function (mj, cbEl) {
          item.addEventListener("click", function (e) {
            if (e && e.target && e.target.tagName === "INPUT") return;
            cbEl.checked = !cbEl.checked;
            cbEl.dispatchEvent(new Event("change"));
          });

          cbEl.addEventListener("change", function () {
            if (this.checked) majorSelected.add(mj);
            else majorSelected.delete(mj);

            updateMajorBtnText();
            renderProducts();
          });
        })(major, cb);

        majorDropdownPanel.appendChild(item);
      }

      updateMajorBtnText();
    }

    // === 视图状态 ===
    var productViewMode = "pool"; // pool | selected
    var productSearchQuery = "";

    function updateViewButtonsText() {
      if (!viewPoolBtn || !viewSelectedBtn) return;

      var tokens = splitTokens(productSearchQuery);

      // pool 的“范围”：搜索时全量；未搜索时按大类
      var poolScope = [];
      if (tokens.length > 0) {
        poolScope = allTemplates.slice();
      } else {
        var majors = getSelectedMajors();
        if (majors.length > 0) {
          var majorSet = new Set(majors);
          for (var i = 0; i < allTemplates.length; i++) {
            if (majorSet.has(allTemplates[i].major))
              poolScope.push(allTemplates[i]);
          }
        }
      }

      var poolCount = 0;
      for (var j = 0; j < poolScope.length; j++) {
        if (!selected.has(poolScope[j].key)) poolCount++;
      }

      viewPoolBtn.textContent = "待选库（" + poolCount + "）";
      viewSelectedBtn.textContent = "已选产品（" + selected.size + "）";

      if (productViewMode === "pool") {
        viewPoolBtn.classList.add("active");
        viewSelectedBtn.classList.remove("active");
      } else {
        viewSelectedBtn.classList.add("active");
        viewPoolBtn.classList.remove("active");
      }
    }

    function setProductViewMode(mode) {
      productViewMode = mode === "selected" ? "selected" : "pool";
      updateViewButtonsText();
      renderProducts();
    }

    function setProductSearchQuery(q) {
      productSearchQuery = String(q || "");
      renderProducts();
    }

    async function loadProductsFromTable(meta) {
      var table = await bitable.base.getTableById(meta.id);

      var fCode = (await getFieldByAnyName(table, ["产品编号"], false)).field;
      var fName = (await getFieldByAnyName(table, ["产品名称"], true)).field;
      var fSizeDays = (await getFieldByAnyName(table, ["尺寸/天数"], true)).field;
      var fUnit = (await getFieldByAnyName(table, ["计算单位"], true)).field;
      var fCost = (await getFieldByAnyName(table, ["成本单价"], true)).field;
      var fPrice = (await getFieldByAnyName(table, ["单价"], true)).field;
      var fDesc = (await getFieldByAnyName(table, ["产品描述"], false)).field;

      var parsed = parseTableName(meta.name);
      var major = parsed.major || meta.name || "";
      var supplier = parsed.supplier || "";

      var records = await getAllRecords(table, 200);
      var items = [];

      for (var i = 0; i < records.length; i++) {
        var r = records[i];
        var recordId = r.recordId || r.id;
        var fields = r.fields || {};

        items.push({
          sourceTableId: meta.id,
          sourceTableName: major,
          supplier: supplier,
          sourceTableNameRaw: meta.name || "",
          recordId: recordId,

          code: fCode ? (toPlainText(fields[fCode.id]) || "") : "",
          name: toPlainText(fields[fName.id]) || "未命名产品",
          sizeDays: toPlainText(fields[fSizeDays.id]) || "",
          unit: toPlainText(fields[fUnit.id]) || "",
          cost: toNumber(fields[fCost.id]),
          price: toNumber(fields[fPrice.id]),
          desc: fDesc ? (toPlainText(fields[fDesc.id]) || "") : "",

          qty: 1,
        });
      }

      return items;
    }

    async function reloadAllProducts() {
      clearMsg();
      productList.innerHTML =
        '<div class="hint">正在加载全部产品库（含所有大类）...</div>';

      var warnings = [];
      var allItems = [];

      var results = await runWithConcurrency(
        metasParsed,
        async function (meta) {
          var items = await loadProductsFromTable({
            id: meta.id,
            name: meta.name,
          });
          return { meta: meta, items: items };
        },
        4
      );

      for (var i = 0; i < results.length; i++) {
        var r = results[i];
        if (!r) continue;
        if (r.__error) {
          var metaName =
            metasParsed[i] && metasParsed[i].name ? metasParsed[i].name : "未知表";
          warnings.push(
            "表【" +
              metaName +
              "】加载失败：" +
              (r.__error && r.__error.message
                ? r.__error.message
                : String(r.__error))
          );
          continue;
        }
        if (r.items && r.items.length) allItems = allItems.concat(r.items);
      }

      // 构建“主视觉喷绘”价格簿（从物料搭建里取）
      mainVisualPriceBook = new Map();
      for (var x = 0; x < allItems.length; x++) {
        var it = allItems[x];
        if (normalizeName(it.name) !== "主视觉喷绘") continue;
        if (normalizeName(it.sourceTableName) !== "物料搭建") continue;

        var sup = it.supplier || "默认供应商";
        if (!mainVisualPriceBook.has(sup)) {
          mainVisualPriceBook.set(sup, {
            cost: Number(it.cost || 0),
            price: Number(it.price || 0),
            unit: it.unit || "平方",
            desc: it.desc || "",
            sourceTableName: it.sourceTableName || "物料搭建",
          });
        }
      }
      if (mainVisualPriceBook.size === 0) {
        mainVisualPriceBook = null;
      }

      updateMainVisualQtyUnit();

      allTemplates = buildProductTemplates(allItems);
      rebuildAllTemplatesMap();

      // 用模板重算“占位已选”，并执行默认供应商跟随逻辑（新规则：只要该产品包含主视觉供应商选项就跟随）
      selected.forEach(function (ex, key) {
        var tpl = allTemplatesMap.get(key);
        if (!tpl) return;

        var idx = 0;

        if (ex.__supplierManual) {
          if (typeof ex.__optIndex === "number") idx = ex.__optIndex;
          if (!tpl.options[idx])
            idx = findOptionIndexBySupplier(tpl, ex.supplier || "");
        } else {
          var followIdx = getFollowSupplierIndexForTemplate(tpl);
          if (followIdx != null) {
            idx = followIdx;
          } else if (ex.supplier) {
            var keepIdx = findOptionIndexBySupplierStrict(tpl, ex.supplier || "");
            if (keepIdx >= 0) idx = keepIdx;
            else idx = 0;
          } else if (
            typeof ex.__optIndex === "number" &&
            tpl.options[ex.__optIndex]
          ) {
            idx = ex.__optIndex;
          } else {
            idx = 0;
          }
        }

        // applyPick 会保留 __pickSeq（若已存在）
        applyPick(key, tpl, idx, ex.qty, !!ex.__supplierManual);
      });

      syncAutoSuppliersFromMainVisual(allTemplatesMap);

      renderProducts();

      if (warnings.length > 0)
        setMsg("部分表无法加载：\n" + warnings.join("\n"), "err");
      else setMsg("产品库加载完成（全量）。", "ok");
    }

    function ensureMainVisualPriceBook() {
      if (mainVisualPriceBook) return Promise.resolve(mainVisualPriceBook);
      return Promise.reject(
        new Error("未在【物料搭建】中找到“主视觉喷绘”的定价记录，无法匹配单价。")
      );
    }

    function renderMainVisualList() {
      if (!mvList) return;

      if (!mainVisualItems.length) {
        mvList.innerHTML = "尚未添加主视觉喷绘。";
        return;
      }

      var unitText = mvQtyUnit ? mvQtyUnit.textContent : "";
      if (!unitText) unitText = "平方";

      var html = "";
      for (var i = 0; i < mainVisualItems.length; i++) {
        var it = mainVisualItems[i];
        html +=
          '<div class="table-item" style="padding:6px 0; border:none;">' +
          '<div class="table-name">' +
          escapeHtml(it.supplier || "默认供应商") +
          "｜尺寸：" +
          escapeHtml(it.size) +
          "｜数量：" +
          escapeHtml(it.qty) +
          (unitText ? " " + escapeHtml(unitText) : "") +
          "</div>" +
          '<button data-mv-del="' +
          i +
          '" class="smallbtn" type="button">删除</button>' +
          "</div>";
      }
      mvList.innerHTML = html;

      var dels = mvList.querySelectorAll("button[data-mv-del]");
      for (var j = 0; j < dels.length; j++) {
        dels[j].addEventListener("click", function () {
          var idx = Number(this.getAttribute("data-mv-del"));
          if (isFinite(idx) && idx >= 0) {
            mainVisualItems.splice(idx, 1);
            renderMainVisualList();
          }
        });
      }
    }

    function refillMainVisualSuppliers() {
      if (!mvSupplier) return;
      mvSupplier.innerHTML = "";

      var bucket = majorIndex.get("物料搭建");
      var suppliers = [];
      if (bucket && bucket.suppliers)
        bucket.suppliers.forEach(function (x) {
          suppliers.push(x);
        });

      suppliers.sort(function (a, b) {
        return String(a).localeCompare(String(b), "zh");
      });

      if (suppliers.length === 0) suppliers = ["默认供应商"];

      for (var i = 0; i < suppliers.length; i++) {
        var o = document.createElement("option");
        o.value = suppliers[i];
        o.textContent = suppliers[i];
        mvSupplier.appendChild(o);
      }

      if (mainVisualDefaultSupplier) {
        for (var k = 0; k < mvSupplier.options.length; k++) {
          if (mvSupplier.options[k].value === mainVisualDefaultSupplier) {
            mvSupplier.value = mainVisualDefaultSupplier;
            break;
          }
        }
      }
      mainVisualDefaultSupplier = mvSupplier.value || "";
      updateMainVisualQtyUnit();
    }

    function computePoolScopeTemplates() {
      var tokens = splitTokens(productSearchQuery);

      // 搜索时直接用全量库
      if (tokens.length > 0) return allTemplates.slice();

      // 未搜索时按大类过滤
      var majors = getSelectedMajors();
      if (!majors || majors.length === 0) return [];

      var majorSet = new Set(majors);
      var out = [];
      for (var i = 0; i < allTemplates.length; i++) {
        if (majorSet.has(allTemplates[i].major)) out.push(allTemplates[i]);
      }
      return out;
    }

    function computeSearchedTemplates(baseTemplates) {
      var tokens = splitTokens(productSearchQuery);
      if (tokens.length === 0) return baseTemplates.slice();

      var scored = [];
      for (var j = 0; j < baseTemplates.length; j++) {
        var s = scoreTemplateByQuery(baseTemplates[j], tokens);
        if (s > 0) scored.push({ tpl: baseTemplates[j], score: s });
      }

      scored.sort(function (a, b) {
        if (a.score !== b.score) return b.score - a.score;
        var an = normalizeName(a.tpl.name);
        var bn = normalizeName(b.tpl.name);
        if (an !== bn) return an.localeCompare(bn, "zh");
        return normalizeName(a.tpl.sizeDays).localeCompare(
          normalizeName(b.tpl.sizeDays),
          "zh"
        );
      });

      var out = [];
      for (var k = 0; k < scored.length; k++) out.push(scored[k].tpl);
      return out;
    }

    function renderSelectedGrouped() {
      var tokens = splitTokens(productSearchQuery);

      var items = [];
      selected.forEach(function (picked, key) {
        if (String(key).indexOf("主视觉喷绘||") === 0) return;

        var tpl = allTemplatesMap.get(key) || null;
        var major =
          (tpl && tpl.major) ||
          (picked && picked.sourceTableName ? picked.sourceTableName : "") ||
          "其他";

        var name = (tpl && tpl.name) || (picked && picked.name) || "";
        var sizeDays =
          (tpl && tpl.sizeDays) || (picked && picked.sizeDays) || "";
        var unit = (tpl && tpl.unit) || (picked && picked.unit) || "";

        if (tokens.length > 0) {
          var score = 0;
          if (tpl) score = scoreTemplateByQuery(tpl, tokens);
          else {
            var fakeTpl = {
              name: name,
              sizeDays: sizeDays,
              unit: unit,
              major: major,
              options: [
                {
                  supplier: picked ? picked.supplier : "",
                  desc: picked ? picked.desc : "",
                },
              ],
            };
            score = scoreTemplateByQuery(fakeTpl, tokens);
          }
          if (score <= 0) return;
        }

        items.push({
          key: key,
          tpl: tpl,
          picked: picked,
          major: major,
          name: name,
          sizeDays: sizeDays,
          unit: unit,
        });
      });

      if (!items || items.length === 0) {
        productList.innerHTML =
          '<div class="hint">当前没有已选产品（或搜索无匹配）。</div>';
        return;
      }

      var groups = new Map();
      for (var i = 0; i < items.length; i++) {
        var mj = items[i].major || "其他";
        var arr = groups.get(mj);
        if (!arr) {
          arr = [];
          groups.set(mj, arr);
        }
        arr.push(items[i]);
      }

      var majors = [];
      groups.forEach(function (_, k) {
        majors.push(k);
      });
      majors.sort(function (a, b) {
        return String(a).localeCompare(String(b), "zh");
      });

      for (var gi = 0; gi < majors.length; gi++) {
        (function () {
          var mj = majors[gi];
          var arr = groups.get(mj) || [];

          // 需求2：同一大类内按“选择先后”排序
          arr.sort(function (a, b) {
            var as =
              a.picked && typeof a.picked.__pickSeq === "number"
                ? a.picked.__pickSeq
                : 0;
            var bs =
              b.picked && typeof b.picked.__pickSeq === "number"
                ? b.picked.__pickSeq
                : 0;

            if (as && bs && as !== bs) return as - bs;
            if (as && !bs) return -1;
            if (!as && bs) return 1;

            // 兜底：名称/尺寸
            var an = normalizeName(a.name);
            var bn = normalizeName(b.name);
            if (an !== bn) return an.localeCompare(bn, "zh");
            return normalizeName(a.sizeDays).localeCompare(
              normalizeName(b.sizeDays),
              "zh"
            );
          });

          var block = document.createElement("div");
          block.className = "group-block";
          block.style.setProperty("--gcolor", majorToColor(mj));

          var header = document.createElement("div");
          header.className = "group-header";
          header.innerHTML =
            '<div class="group-title">' +
            escapeHtml(mj) +
            "</div>" +
            '<div class="group-sub">已选 ' +
            arr.length +
            " 项</div>";

          var body = document.createElement("div");
          body.className = "group-body";

          for (var i = 0; i < arr.length; i++) {
            (function () {
              var item = arr[i];
              var key = item.key;
              var picked = item.picked || {};
              var tpl = item.tpl;

              var row = document.createElement("div");
              row.className = "prow";

              var cb = document.createElement("input");
              cb.type = "checkbox";
              cb.checked = true;

              var showCost = picked ? picked.cost : 0;
              var showPrice = picked ? picked.price : 0;
              var showDesc = picked ? picked.desc : "";

              var mid = document.createElement("div");
              var majorText = item.major
                ? '<span class="hl">' + escapeHtml(item.major) + "</span>　"
                : "";
              var missingTplHint = !tpl
                ? '<div class="pmeta"><span class="hl">提示</span>：该产品在当前产品库中未找到（仍保留在已选列表，可取消勾选移除）。</div>'
                : "";
              mid.innerHTML =
                '<div class="pname">' +
                escapeHtml(item.name || "未命名产品") +
                "</div>" +
                '<div class="pmeta">' +
                majorText +
                '<span class="hl">成本：¥' +
                (showCost || 0) +
                "</span>　" +
                '<span class="hl">单价：¥' +
                (showPrice || 0) +
                "</span></div>" +
                (showDesc
                  ? '<div class="pmeta"><span class="hl">描述</span>：' +
                    escapeHtml(showDesc) +
                    "</div>"
                  : "") +
                missingTplHint;

              var supplierSel = document.createElement("select");
              supplierSel.className = "select";

              if (!tpl || !tpl.options || tpl.options.length === 0) {
                supplierSel.disabled = true;
                var oo = document.createElement("option");
                oo.value = "0";
                oo.textContent =
                  picked && picked.supplier ? picked.supplier : "默认供应商";
                supplierSel.appendChild(oo);
              } else {
                supplierSel.disabled = false;
                for (var oi = 0; oi < tpl.options.length; oi++) {
                  var opt = tpl.options[oi];
                  var o = document.createElement("option");
                  o.value = String(oi);
                  o.textContent = displaySupplierName(opt);
                  supplierSel.appendChild(o);
                }

                // 新规则：如果该产品存在主视觉供应商选项，则默认显示主视觉供应商
                var followIdx = getFollowSupplierIndexForTemplate(tpl);
                var defaultIdx = followIdx != null ? followIdx : 0;

                supplierSel.value =
                  picked && typeof picked.__optIndex === "number"
                    ? String(picked.__optIndex)
                    : String(defaultIdx);
              }

              var qty = document.createElement("input");
              qty.className = "qty";
              qty.type = "number";
              qty.min = "1";
              qty.step = "1";
              qty.value = String(picked && picked.qty ? picked.qty : 1);

              var qtyWrap = document.createElement("div");
              qtyWrap.className = "qtywrap";
              qtyWrap.appendChild(qty);

              var unitSpan = document.createElement("span");
              unitSpan.className = "unit";
              unitSpan.textContent = item.unit ? String(item.unit) : "";
              qtyWrap.appendChild(unitSpan);

              cb.addEventListener("change", function () {
                if (!this.checked) {
                  deleteSelected(key);
                  renderProducts();
                } else {
                  this.checked = true;
                }
              });

              supplierSel.addEventListener("change", function () {
                if (!tpl) return;
                var it2 = selected.get(key);
                if (!it2) return;
                applyPick(key, tpl, supplierSel.value, qty.value, true);
                renderProducts();
              });

              qty.addEventListener("input", function () {
                var it3 = selected.get(key);
                if (!it3) return;
                it3.qty = Math.max(1, Number(qty.value) || 1);
                setSelected(key, it3);
              });

              row.appendChild(cb);
              row.appendChild(mid);
              row.appendChild(supplierSel);
              row.appendChild(qtyWrap);
              body.appendChild(row);
            })();
          }

          block.appendChild(header);
          block.appendChild(body);
          productList.appendChild(block);
        })();
      }
    }

    // ✅ 待选库：始终开放供应商/数量，且支持预设暂存（重渲染不丢）
    function renderPoolList() {
      var scope = computePoolScopeTemplates();
      var tokens = splitTokens(productSearchQuery);

      if ((!scope || scope.length === 0) && tokens.length === 0) {
        productList.innerHTML =
          '<div class="hint">请先在步骤1选择至少一个产品大类；或者直接在搜索框输入关键字（会从全部产品中搜索）。</div>';
        return;
      }

      var visible = computeSearchedTemplates(scope);

      // pool：隐藏已选
      var list = [];
      for (var i = 0; i < visible.length; i++) {
        var tpl = visible[i];
        if (selected.has(tpl.key)) continue;
        list.push(tpl);
      }

      if (!list || list.length === 0) {
        productList.innerHTML =
          '<div class="hint">当前待选库为空（可能都已选中，或搜索无匹配）。</div>';
        return;
      }

      for (var idx = 0; idx < list.length; idx++) {
        (function () {
          var tpl = list[idx];

          var row = document.createElement("div");
          row.className = "prow";

          var cb = document.createElement("input");
          cb.type = "checkbox";
          cb.checked = false;

          var showCost = tpl.options[0] ? tpl.options[0].cost : 0;
          var showPrice = tpl.options[0] ? tpl.options[0].price : 0;
          var showDesc = tpl.options[0] ? tpl.options[0].desc : "";

          var mid = document.createElement("div");
          var majorText = tpl.major
            ? '<span class="hl">' + escapeHtml(tpl.major) + "</span>　"
            : "";
          mid.innerHTML =
            '<div class="pname">' +
            escapeHtml(tpl.name) +
            "</div>" +
            '<div class="pmeta">' +
            majorText +
            '<span class="hl">成本：¥' +
            (showCost || 0) +
            "</span>　" +
            '<span class="hl">单价：¥' +
            (showPrice || 0) +
            "</span></div>" +
            (showDesc
              ? '<div class="pmeta"><span class="hl">描述</span>：' +
                escapeHtml(showDesc) +
                "</div>"
              : "");

          var supplierSel = document.createElement("select");
          supplierSel.className = "select";
          supplierSel.disabled = false;

          for (var oi = 0; oi < tpl.options.length; oi++) {
            var opt = tpl.options[oi];
            var o = document.createElement("option");
            o.value = String(oi);
            o.textContent = displaySupplierName(opt);
            supplierSel.appendChild(o);
          }

          var qty = document.createElement("input");
          qty.className = "qty";
          qty.type = "number";
          qty.min = "1";
          qty.step = "1";
          qty.disabled = false;

          var qtyWrap = document.createElement("div");
          qtyWrap.className = "qtywrap";
          qtyWrap.appendChild(qty);

          var unitSpan = document.createElement("span");
          unitSpan.className = "unit";
          unitSpan.textContent = tpl.unit ? String(tpl.unit) : "";
          qtyWrap.appendChild(unitSpan);

          // 初始化显示：优先用草稿，否则如果该产品存在主视觉供应商选项就显示主视觉供应商，否则默认第一个
          var followIdx0 = getFollowSupplierIndexForTemplate(tpl);
          var defaultIdx = followIdx0 != null ? followIdx0 : 0;

          var draft = poolDraft.get(tpl.key) || null;

          var idxToShow = defaultIdx;
          if (
            draft &&
            draft.supplierTouched &&
            typeof draft.supplierIndex === "number" &&
            isFinite(draft.supplierIndex)
          ) {
            idxToShow = Math.max(
              0,
              Math.min(tpl.options.length - 1, draft.supplierIndex)
            );
          }
          supplierSel.value = String(idxToShow);

          qty.value = String(draft && draft.qty ? draft.qty : 1);

          supplierSel.addEventListener("change", function () {
            // 手动选过供应商：标记 touched
            upsertPoolDraftSupplier(tpl.key, supplierSel.value);
          });

          qty.addEventListener("input", function () {
            upsertPoolDraftQty(tpl.key, qty.value);
          });

          cb.addEventListener("change", function () {
            if (!this.checked) return;

            var d = poolDraft.get(tpl.key) || null;

            var idxToUse = Number(supplierSel.value) || 0;
            var qtyToUse = qty.value;

            // 未手动选供应商时：如果该产品存在主视觉供应商选项，则强制跟随主视觉供应商
            if (!d || !d.supplierTouched) {
              var followIdx = getFollowSupplierIndexForTemplate(tpl);
              if (followIdx != null) {
                idxToUse = followIdx;
                supplierSel.value = String(idxToUse);
              }
            }

            applyPick(
              tpl.key,
              tpl,
              idxToUse,
              qtyToUse,
              !!(d && d.supplierTouched)
            );

            // 勾选后该产品从待选库隐藏；同时把它的“待选库预设”清掉（后续移除时也会重置）
            resetPoolDraftAndPersist(tpl.key);

            renderProducts();
          });

          row.appendChild(cb);
          row.appendChild(mid);
          row.appendChild(supplierSel);
          row.appendChild(qtyWrap);
          productList.appendChild(row);
        })();
      }
    }

    function renderProducts() {
      productList.innerHTML = "";
      updateViewButtonsText();

      if (!allTemplates || allTemplates.length === 0) {
        productList.innerHTML =
          '<div class="hint">产品库尚未加载完成，请稍等或点击“重新加载产品”。</div>';
        return;
      }

      if (productViewMode === "selected") {
        renderSelectedGrouped();
      } else {
        renderPoolList();
      }
    }

    // === 初始化读取表列表 ===
    var metas = [];
    try {
      metas = await getAllTables(bitable);
    } catch (e2) {
      setMsg(
        "读取数据表列表失败：" + (e2 && e2.message ? e2.message : String(e2)),
        "err"
      );
      return;
    }

    for (var mi = 0; mi < metas.length; mi++) {
      var m = metas[mi];
      var name = m && m.name ? m.name : "";
      var p = parseTableName(name);

      var mp = {
        id: m.id,
        name: name || m.id,
        major: p.major || name || m.id,
        supplier: p.supplier || "",
      };
      metasParsed.push(mp);

      var bucket = majorIndex.get(mp.major);
      if (!bucket) {
        bucket = { major: mp.major, metas: [], suppliers: new Set() };
        majorIndex.set(mp.major, bucket);
      }
      bucket.metas.push(mp);
      if (mp.supplier) bucket.suppliers.add(mp.supplier);
    }

    // === 事件绑定 ===
    majorDropdownBtn.addEventListener("click", function (e) {
      e.preventDefault();
      toggleMajorDropdown();
    });

    document.addEventListener("click", function (e) {
      var t = e && e.target ? e.target : null;
      if (!t) return;

      var insideBtn = majorDropdownBtn.contains(t);
      var insidePanel = majorDropdownPanel.contains(t);
      if (!insideBtn && !insidePanel) toggleMajorDropdown(false);
    });

    if (viewPoolBtn) {
      viewPoolBtn.addEventListener("click", function () {
        setProductViewMode("pool");
      });
    }
    if (viewSelectedBtn) {
      viewSelectedBtn.addEventListener("click", function () {
        setProductViewMode("selected");
      });
    }

    if (productSearchInput) {
      productSearchInput.addEventListener("input", function () {
        setProductSearchQuery(this.value || "");
      });
    }

    if (needMainVisualCb && mvBox) {
      needMainVisualCb.addEventListener("change", function () {
        mvBox.style.display = this.checked ? "block" : "none";
        if (this.checked) {
          refillMainVisualSuppliers();
          updateMainVisualQtyUnit();
          renderMainVisualList();
          syncAutoSuppliersFromMainVisual(allTemplatesMap);
          renderProducts();
        }
      });
    }

    if (mvSupplier) {
      mvSupplier.addEventListener("change", function () {
        mainVisualDefaultSupplier = mvSupplier.value || "";
        updateMainVisualQtyUnit();
        syncAutoSuppliersFromMainVisual(allTemplatesMap);
        renderProducts();
      });
    }

    if (mvAdd) {
      mvAdd.addEventListener("click", function () {
        clearMsg();

        var supplier = mvSupplier ? mvSupplier.value : "";
        var size = mvSize ? String(mvSize.value || "").trim() : "";
        var qty = mvQty ? Math.max(1, Number(mvQty.value) || 1) : 1;

        if (!size) {
          setMsg("请填写主视觉喷绘尺寸。", "err");
          return;
        }

        ensureMainVisualPriceBook()
          .then(function () {
            mainVisualItems.push({ supplier: supplier, size: size, qty: qty });
            if (mvSize) mvSize.value = "";
            if (mvQty) mvQty.value = "1";
            renderMainVisualList();
            setMsg("已添加一条主视觉喷绘。", "ok");
          })
          .catch(function (e) {
            setMsg(
              "主视觉喷绘配置失败：" + (e && e.message ? e.message : String(e)),
              "err"
            );
          });
      });
    }

    if (selectAllBtn) {
      selectAllBtn.addEventListener("click", function () {
        majorSelected.clear();
        majorIndex.forEach(function (bucket) {
          if (bucket && bucket.major) majorSelected.add(bucket.major);
        });
        renderMajorDropdownOptions();
        renderProducts();
      });
    }

    if (clearAllBtn) {
      clearAllBtn.addEventListener("click", function () {
        majorSelected.clear();
        renderMajorDropdownOptions();
        renderProducts();
      });
    }

    if (reloadBtn) {
      reloadBtn.addEventListener("click", function () {
        reloadAllProducts().catch(function (e) {
          setMsg(
            "刷新产品失败：" + (e && e.message ? e.message : String(e)),
            "err"
          );
        });
      });
    }

    if (selectAllVisibleProductsBtn) {
      selectAllVisibleProductsBtn.addEventListener("click", function () {
        clearMsg();

        if (!allTemplates || allTemplates.length === 0) return;

        if (productViewMode !== "pool") {
          setMsg("“全选当前列表”建议在【待选库】使用。", "err");
          return;
        }

        var scope = computePoolScopeTemplates();
        var visible = computeSearchedTemplates(scope);

        var added = 0;
        for (var i = 0; i < visible.length; i++) {
          var tpl = visible[i];
          if (selected.has(tpl.key)) continue;

          var idx = 0;
          var followIdx = getFollowSupplierIndexForTemplate(tpl);
          if (followIdx != null) idx = followIdx;

          applyPick(tpl.key, tpl, idx, 1, false);
          // 全选行为也重置预设
          resetPoolDraftAndPersist(tpl.key);

          added++;
        }

        renderProducts();
        setMsg("已批量选中 " + added + " 个产品。", "ok");
      });
    }

    if (clearAllSelectedProductsBtn) {
      clearAllSelectedProductsBtn.addEventListener("click", function () {
        clearMsg();
        if (selected.size === 0) {
          setMsg("当前没有已选产品。", "err");
          return;
        }
        var ok = window.confirm("确定清空全部已选产品吗？");
        if (!ok) return;
        clearSelected();
        renderProducts();
        setMsg("已清空全部已选产品。", "ok");
      });
    }

    exportBtn.addEventListener("click", function () {
      clearMsg();

      var itemsToExport = [];
      selected.forEach(function (v) {
        itemsToExport.push(v);
      });

      var quoteName = quoteNameInput ? quoteNameInput.value : "";
      if (!quoteName) quoteName = loadQuoteName() || "报价单";

      var needMv = needMainVisualCb && needMainVisualCb.checked;

      if (needMv && mainVisualItems.length > 0) {
        ensureMainVisualPriceBook()
          .then(function (book) {
            for (var i = 0; i < mainVisualItems.length; i++) {
              var mv = mainVisualItems[i];
              var sup = mv.supplier || "默认供应商";
              var pb = book.get(sup);

              if (!pb) {
                throw new Error(
                  "供应商【" +
                    sup +
                    "】缺少“主视觉喷绘”定价，请先在【物料搭建（" +
                    sup +
                    "）】补齐。"
                );
              }

              itemsToExport.push({
                name: "主视觉喷绘",
                sizeDays: mv.size,
                unit: pb.unit || "平方",
                cost: Number(pb.cost || 0),
                price: Number(pb.price || 0),
                desc: pb.desc || "",
                qty: Math.max(1, Number(mv.qty) || 1),
                sourceTableName: pb.sourceTableName || "物料搭建",
                supplier: sup,

                // ✅ 新需求：导出时主视觉喷绘要排在【物料搭建】大类最上方，这里记录添加顺序
                __mvSeq: i,
              });
            }

            doExport(itemsToExport, quoteName);
          })
          .catch(function (e) {
            setMsg(
              "导出失败：" + (e && e.message ? e.message : String(e)),
              "err"
            );
          });
        return;
      }

      doExport(itemsToExport, quoteName);

      function doExport(arr, qn) {
        if (!arr || arr.length === 0) {
          setMsg("请至少选择一个产品，或添加主视觉喷绘。", "err");
          return;
        }

        exportXlsx(arr, qn)
          .then(function () {
            setMsg(
              "已导出 XLSX（含公式/列宽/换行/行高/分组/小计/含税合计）。",
              "ok"
            );
          })
          .catch(function (e) {
            setMsg(
              "导出失败：" + (e && e.message ? e.message : String(e)),
              "err"
            );
          });
      }
    });

    // === 首次渲染 ===
    renderMajorDropdownOptions();
    refillMainVisualSuppliers();
    updateMainVisualQtyUnit();

    // 恢复已选 + 恢复待选预设
    var draftArr = restoreSelectedDraft();
    restorePoolDraft();

    if (draftArr && draftArr.length > 0) {
      setProductViewMode("selected");
    } else {
      setProductViewMode("pool");
    }

    // 启动即加载全量产品库
    reloadAllProducts().catch(function (e) {
      setMsg(
        "初始化加载产品失败：" + (e && e.message ? e.message : String(e)),
        "err"
      );
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", function () {
      main().catch(function (e) {
        setMsg("初始化异常：" + (e && e.message ? e.message : String(e)), "err");
      });
    });
  } else {
    main().catch(function (e) {
      setMsg("初始化异常：" + (e && e.message ? e.message : String(e)), "err");
    });
  }
})();
