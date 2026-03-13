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

  function getSelectedMajors() {
    var sel = document.getElementById("majorSelect");
    if (!sel) return [];
    var majors = [];
    for (var i = 0; i < sel.options.length; i++) {
      if (sel.options[i].selected) majors.push(sel.options[i].value);
    }
    return majors;
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

    // A 科目（组内序号），B 产品名称，C 尺寸/天数，D 数量，E 单位，F 成本单价，G 成本总价(公式)，H 单价，I 总价(公式)，J 产品描述
    var headers = [
      "科目",
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

    var colWidths = [8, 24, 24, 9, 9, 9, 9, 9, 9, 35];
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

    for (var hc = 8; hc <= 9; hc++) {
      ws.getCell(2, hc).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFFF00" },
      };
    }

    var typeOrder = [
      "物料搭建",
      "印刷制作",
      "线上小程序",
      "系统设计",
      "系统适配",
      "人员执行",
      "硬件设备",
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

      if (Object.prototype.hasOwnProperty.call(typeOrderMap, t)) return t;
      return "其他";
    }

    var selectedArr = [];
    if (selected && typeof selected.forEach === "function") {
      selected.forEach(function (v) {
        selectedArr.push(v);
      });
    } else if (Array.isArray(selected)) {
      selectedArr = selected.slice();
    }

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

    function addGroupSubtotal(startRn, endRn) {
      var subtotalRowNum = rowNum;

      ws.getRow(subtotalRowNum).height = 20;

      ws.mergeCells(subtotalRowNum, 1, subtotalRowNum, 6);
      ws.getCell(subtotalRowNum, 1).value = "小计";

      ws.getCell(subtotalRowNum, 7).value = {
        formula: "SUM(G" + startRn + ":G" + endRn + ")",
      };
      ws.getCell(subtotalRowNum, 9).value = {
        formula: "SUM(I" + startRn + ":I" + endRn + ")",
      };

      styleRowAllCols(subtotalRowNum, null, true);

      ws.getCell(subtotalRowNum, 7).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFFF00" },
      };
      ws.getCell(subtotalRowNum, 9).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFFF00" },
      };

      subtotalRows.push(subtotalRowNum);
      rowNum++;
    }

    function addGrandTotalWithTax() {
      ws.getRow(rowNum).height = 20;

      ws.mergeCells(rowNum, 1, rowNum, 6);
      ws.getCell(rowNum, 1).value = "含税合计（税价6%）";

      if (subtotalRows.length === 0) {
        ws.getCell(rowNum, 7).value = 0;
        ws.getCell(rowNum, 9).value = 0;
      } else {
        var gRefs = [];
        var iRefs = [];
        for (var i = 0; i < subtotalRows.length; i++) {
          gRefs.push("G" + subtotalRows[i]);
          iRefs.push("I" + subtotalRows[i]);
        }

        ws.getCell(rowNum, 7).value = {
          formula: "SUM(" + gRefs.join(",") + ")",
        };
        ws.getCell(rowNum, 9).value = {
          formula: "SUM(" + iRefs.join(",") + ")*1.06",
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
          addGroupSubtotal(groupStartRow, groupEndRow);
        }

        groupType = typeLabel;
        groupIdx++;
        groupSerial = 0;

        addGroupHeader(groupType);

        groupStartRow = rowNum;
        groupEndRow = rowNum - 1;
      }

      groupSerial++;

      var desc = normalizeNewlines(it.desc || "");
      if (it && it.supplier) {
        desc = (desc ? desc + "\n" : "") + "供应商：" + it.supplier;
      }

      var rn = rowNum;
      var row = ws.getRow(rn);

      row.getCell(1).value = groupSerial;
      row.getCell(2).value = it.name || "";
      row.getCell(3).value = it.sizeDays || "";
      row.getCell(4).value = Math.max(1, Number(it.qty) || 1);
      row.getCell(5).value = it.unit || "";
      row.getCell(6).value = Number(it.cost || 0);
      row.getCell(7).value = { formula: "F" + rn + "*D" + rn };
      row.getCell(8).value = Number(it.price || 0);
      row.getCell(9).value = { formula: "H" + rn + "*D" + rn };
      row.getCell(10).value = desc;

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

  function buildProductTemplates(allItems) {
    var map = new Map();

    for (var i = 0; i < allItems.length; i++) {
      var it = allItems[i];

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

      // 如果同一个 key 混入不同大类，这里保留第一个 major（通常不会发生）
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
      return normalizeName(a.sizeDays).localeCompare(normalizeName(b.sizeDays), "zh");
    });

    for (var j = 0; j < arr.length; j++) {
      arr[j].options.sort(function (x, y) {
        return normalizeName(x.supplier).localeCompare(normalizeName(y.supplier), "zh");
      });
    }

    return arr;
  }

  function lower(s) {
    return String(s || "").toLowerCase();
  }

  function splitTokens(q) {
    var s = normalizeName(q);
    if (!s) return [];
    // 支持空格/中文空格/逗号等分隔
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

  function isFollowMajor(major) {
    var m = normalizeName(major);
    return m === "物料搭建" || m === "印刷制作";
  }

  function scoreTemplateByQuery(tpl, tokens) {
    if (!tokens || tokens.length === 0) return 1; // 无搜索时都可显示，给个基础分
    var name = lower(tpl && tpl.name ? tpl.name : "");
    var sizeDays = lower(tpl && tpl.sizeDays ? tpl.sizeDays : "");
    var unit = lower(tpl && tpl.unit ? tpl.unit : "");
    var major = lower(tpl && tpl.major ? tpl.major : "");
    var desc = "";
    var suppliersText = "";

    // 轻量拼接：取前几个 option 的描述/供应商做搜索参考
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

  async function main() {
    clearMsg();

    var majorSelectEl = document.getElementById("majorSelect");
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
    var mvAdd = document.getElementById("addMainVisual");
    var mvList = document.getElementById("mainVisualList");

    if (!majorSelectEl || !productList || !exportBtn) {
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

    // selected: templateKey -> picked item（可导出）
    var selected = new Map();

    function persistSelectedDraft() {
      var arr = [];
      selected.forEach(function (v, key) {
        arr.push({
          key: key,
          qty: v && v.qty != null ? Number(v.qty) : 1,
          supplier: v && v.supplier ? String(v.supplier) : "",
          __optIndex: v && typeof v.__optIndex === "number" ? v.__optIndex : 0,
          __supplierManual: !!(v && v.__supplierManual),
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
        // 先塞一个“占位选择”，等模板加载后再按模板重算价格/描述等
        for (var i = 0; i < arr.length; i++) {
          var x = arr[i];
          if (!x || !x.key) continue;
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
            __placeholder: true,
          });
        }
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
    }

    function clearSelected() {
      selected.clear();
      persistSelectedDraft();
    }

    // === 主视觉喷绘 & 供应商跟随逻辑 ===
    var mainVisualItems = []; // { supplier, size, qty }
    var mainVisualPriceBook = null; // Map supplier -> { cost, price, unit, desc, sourceTableName }

    // 当前“主视觉喷绘供应商”选择（用于默认跟随）
    var mainVisualDefaultSupplier = "";

    function getDefaultSupplierForMajor(major) {
      if (!isFollowMajor(major)) return "";
      return normalizeName(mainVisualDefaultSupplier || "");
    }

    function findOptionIndexBySupplier(tplRef, supplierName) {
      if (!tplRef || !tplRef.options || tplRef.options.length === 0) return 0;

      var target = normalizeName(supplierName);
      if (!target) return 0;

      // 先按“展示名”匹配（空供应商 -> 默认供应商）
      for (var i = 0; i < tplRef.options.length; i++) {
        var opt = tplRef.options[i];
        var disp = normalizeName(displaySupplierName(opt));
        if (disp === target) return i;
      }

      // 再按“原始 supplier 字段”匹配
      for (var j = 0; j < tplRef.options.length; j++) {
        var opt2 = tplRef.options[j];
        var raw = normalizeName(opt2 && opt2.supplier ? opt2.supplier : "");
        if (raw === target) return j;
      }

      // 如果目标是“默认供应商”，尽量找 raw supplier 为空的那条
      if (target === "默认供应商") {
        for (var k = 0; k < tplRef.options.length; k++) {
          var opt3 = tplRef.options[k];
          if (!normalizeName(opt3 && opt3.supplier ? opt3.supplier : "")) return k;
        }
      }

      return 0;
    }

    var metas = [];
    var metasParsed = [];
    var majorIndex = new Map(); // major -> { major, metas: [{id,name,major,supplier}], suppliers:Set() }
    var currentTemplates = [];
    var currentTemplatesMap = new Map(); // templateKey -> template

    var productViewMode = "pool"; // pool | selected
    var productSearchQuery = "";

    function updateViewButtonsText() {
      if (!viewPoolBtn || !viewSelectedBtn) return;

      var poolCount = 0;
      var selCount = 0;

      for (var i = 0; i < currentTemplates.length; i++) {
        if (selected.has(currentTemplates[i].key)) selCount++;
        else poolCount++;
      }

      viewPoolBtn.textContent = "待选库（" + poolCount + "）";
      viewSelectedBtn.textContent = "已选产品（" + selCount + "）";

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
      renderProducts(currentTemplates);
    }

    function setProductSearchQuery(q) {
      productSearchQuery = String(q || "");
      renderProducts(currentTemplates);
    }

    function rebuildTemplatesMap() {
      currentTemplatesMap = new Map();
      for (var i = 0; i < currentTemplates.length; i++) {
        currentTemplatesMap.set(currentTemplates[i].key, currentTemplates[i]);
      }
    }

    function applyPick(templateKey, tplRef, supplierIndex, qtyValue, supplierManualFlag) {
      var idx = Math.max(0, Number(supplierIndex) || 0);
      var opt2 = (tplRef && tplRef.options && tplRef.options[idx]) || (tplRef && tplRef.options ? tplRef.options[0] : null);
      if (!opt2) return;

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
      };

      setSelected(templateKey, next);
    }

    function syncAutoSuppliersFromMainVisual() {
      var defSup = normalizeName(mainVisualDefaultSupplier || "");
      if (!defSup) return;

      // 仅对【物料搭建】【印刷制作】且未手动改过供应商的已选产品生效
      selected.forEach(function (v, key) {
        if (!v) return;
        if (v.__supplierManual) return;

        var tpl = currentTemplatesMap.get(key);
        if (!tpl) return;

        var major = tpl.major || v.sourceTableName || "";
        if (!isFollowMajor(major)) return;

        var idx = findOptionIndexBySupplier(tpl, defSup);
        // 保留数量
        applyPick(key, tpl, idx, v.qty, false);
      });
    }

    // 初始化读取表列表
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
        major: p.major || (name || m.id),
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

    var typeOrder = [
      "物料搭建",
      "印刷制作",
      "线上小程序",
      "系统设计",
      "系统适配",
      "人员执行",
      "硬件设备",
      "摄影摄像",
      "视觉设计",
      "项目管理",
      "其他",
    ];
    var typeOrderMap = {};
    for (var oi = 0; oi < typeOrder.length; oi++) typeOrderMap[typeOrder[oi]] = oi;

    function renderMajorSelectOptions() {
      majorSelectEl.innerHTML = "";

      if (!majorIndex || majorIndex.size === 0) {
        var opt0 = document.createElement("option");
        opt0.value = "";
        opt0.textContent = "当前 Base 没有数据表";
        majorSelectEl.appendChild(opt0);
        majorSelectEl.disabled = true;
        return;
      }
      majorSelectEl.disabled = false;

      var majors = [];
      majorIndex.forEach(function (v) {
        majors.push(v.major);
      });

      majors.sort(function (a, b) {
        var ai = typeOrderMap[a];
        var bi = typeOrderMap[b];
        if (ai == null && bi == null) return String(a).localeCompare(String(b), "zh");
        if (ai == null) return 1;
        if (bi == null) return -1;
        return ai - bi;
      });

      for (var i = 0; i < majors.length; i++) {
        var major = majors[i];
        var bucket = majorIndex.get(major);

        var supplierArr = [];
        bucket.suppliers.forEach(function (x) {
          supplierArr.push(x);
        });
        supplierArr.sort(function (x, y) {
          return String(x).localeCompare(String(y), "zh");
        });

        var suffix = "";
        if (supplierArr.length > 0) {
          var preview = supplierArr.slice(0, 3).join("、");
          suffix = "（供应商：" + preview + (supplierArr.length > 3 ? "等" : "") + "）";
        }

        var opt = document.createElement("option");
        opt.value = major;
        opt.textContent = major + suffix;
        majorSelectEl.appendChild(opt);
      }
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
          sourceTableName: major, // 用大类参与导出分组
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

    function computeVisibleTemplates(templates) {
      var arr = [];
      for (var i = 0; i < templates.length; i++) {
        var tpl = templates[i];
        var isSel = selected.has(tpl.key);
        if (productViewMode === "pool") {
          if (isSel) continue; // 已选从库中隐藏
        } else {
          if (!isSel) continue;
        }
        arr.push(tpl);
      }

      var tokens = splitTokens(productSearchQuery);
      if (tokens.length === 0) {
        return arr;
      }

      var scored = [];
      for (var j = 0; j < arr.length; j++) {
        var s = scoreTemplateByQuery(arr[j], tokens);
        if (s > 0) scored.push({ tpl: arr[j], score: s });
      }

      scored.sort(function (a, b) {
        if (a.score !== b.score) return b.score - a.score;
        var an = normalizeName(a.tpl.name);
        var bn = normalizeName(b.tpl.name);
        if (an !== bn) return an.localeCompare(bn, "zh");
        return normalizeName(a.tpl.sizeDays).localeCompare(normalizeName(b.tpl.sizeDays), "zh");
      });

      var out = [];
      for (var k = 0; k < scored.length; k++) out.push(scored[k].tpl);
      return out;
    }

    function renderProducts(templates) {
      productList.innerHTML = "";
      updateViewButtonsText();

      if (!templates || templates.length === 0) {
        productList.innerHTML = '<div class="hint">选中的产品大类没有可用产品（可能无记录）。</div>';
        return;
      }

      var visible = computeVisibleTemplates(templates);

      if (!visible || visible.length === 0) {
        var hint =
          productViewMode === "pool"
            ? "当前待选库为空（可能都已选中，或搜索无匹配）。"
            : "当前没有已选产品（或搜索无匹配）。";
        productList.innerHTML = '<div class="hint">' + escapeHtml(hint) + "</div>";
        return;
      }

      for (var i = 0; i < visible.length; i++) {
        (function () {
          var tpl = visible[i];
          var picked = selected.get(tpl.key);

          var row = document.createElement("div");
          row.className = "prow";

          var cb = document.createElement("input");
          cb.type = "checkbox";
          cb.checked = !!picked;

          var showCost = picked ? picked.cost : (tpl.options[0] ? tpl.options[0].cost : 0);
          var showPrice = picked ? picked.price : (tpl.options[0] ? tpl.options[0].price : 0);
          var showDesc = picked ? picked.desc : (tpl.options[0] ? tpl.options[0].desc : "");

          var mid = document.createElement("div");
          var majorText = tpl.major ? ('<span class="hl">' + escapeHtml(tpl.major) + "</span>　") : "";
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
          supplierSel.disabled = !picked;

          for (var oi = 0; oi < tpl.options.length; oi++) {
            var opt = tpl.options[oi];
            var o = document.createElement("option");
            o.value = String(oi);
            o.textContent = displaySupplierName(opt);
            supplierSel.appendChild(o);
          }

          // 默认供应商跟随：物料搭建/印刷制作 -> 跟随主视觉供应商（仅在未手动改动时生效）
          var defaultSup = getDefaultSupplierForMajor(tpl.major);
          var defaultIdx = defaultSup ? findOptionIndexBySupplier(tpl, defaultSup) : 0;

          supplierSel.value =
            picked && typeof picked.__optIndex === "number"
              ? String(picked.__optIndex)
              : String(defaultIdx);

          var qty = document.createElement("input");
          qty.className = "qty";
          qty.type = "number";
          qty.min = "1";
          qty.step = "1";
          qty.value = String(picked ? picked.qty : 1);
          qty.disabled = !picked;

          cb.addEventListener("change", function () {
            if (this.checked) {
              qty.disabled = false;
              supplierSel.disabled = false;

              // 选中时：如果该大类需要跟随主视觉，则按主视觉供应商做默认；否则用当前 select 值
              var idxToUse = Number(supplierSel.value) || 0;

              var defSup2 = getDefaultSupplierForMajor(tpl.major);
              if (defSup2) {
                idxToUse = findOptionIndexBySupplier(tpl, defSup2);
                supplierSel.value = String(idxToUse);
              }

              applyPick(tpl.key, tpl, idxToUse, qty.value, false);

              // 从待选库选中后，为避免“选中即消失”造成困惑，自动切到“已选产品”
              if (productViewMode === "pool") {
                setProductViewMode("selected");
              } else {
                renderProducts(currentTemplates);
              }
            } else {
              qty.disabled = true;
              supplierSel.disabled = true;
              deleteSelected(tpl.key);
              renderProducts(currentTemplates);
            }
          });

          supplierSel.addEventListener("change", function () {
            var it2 = selected.get(tpl.key);
            if (!it2) return;

            // 一旦手动改供应商，则不再跟随主视觉（直到取消勾选/重新选择）
            applyPick(tpl.key, tpl, supplierSel.value, qty.value, true);
            renderProducts(currentTemplates);
          });

          qty.addEventListener("input", function () {
            var it3 = selected.get(tpl.key);
            if (!it3) return;
            it3.qty = Math.max(1, Number(qty.value) || 1);
            setSelected(tpl.key, it3);
          });

          row.appendChild(cb);
          row.appendChild(mid);
          row.appendChild(supplierSel);
          row.appendChild(qty);
          productList.appendChild(row);
        })();
      }
    }

    function renderMainVisualList() {
      if (!mvList) return;

      if (!mainVisualItems.length) {
        mvList.innerHTML = "尚未添加主视觉喷绘。";
        return;
      }

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

      // 尝试保留之前的选择，否则用第一个
      if (mainVisualDefaultSupplier) {
        for (var k = 0; k < mvSupplier.options.length; k++) {
          if (mvSupplier.options[k].value === mainVisualDefaultSupplier) {
            mvSupplier.value = mainVisualDefaultSupplier;
            break;
          }
        }
      }
      mainVisualDefaultSupplier = mvSupplier.value || "";
    }

    async function ensureMainVisualPriceBook() {
      if (mainVisualPriceBook) return mainVisualPriceBook;

      mainVisualPriceBook = new Map();

      var bucket = majorIndex.get("物料搭建");
      if (!bucket || !bucket.metas || bucket.metas.length === 0) {
        throw new Error("未找到【物料搭建】产品库，无法匹配“主视觉喷绘”单价。请检查数据表命名。");
      }

      for (var i = 0; i < bucket.metas.length; i++) {
        var meta = bucket.metas[i];
        var items = await loadProductsFromTable({ id: meta.id, name: meta.name });

        for (var k = 0; k < items.length; k++) {
          if (normalizeName(items[k].name) === "主视觉喷绘") {
            var sup = items[k].supplier || "默认供应商";
            mainVisualPriceBook.set(sup, {
              cost: Number(items[k].cost || 0),
              price: Number(items[k].price || 0),
              unit: items[k].unit || "平方",
              desc: items[k].desc || "",
              sourceTableName: items[k].sourceTableName || "物料搭建",
            });
            break;
          }
        }
      }

      if (mainVisualPriceBook.size === 0) {
        throw new Error("【物料搭建】中未找到产品名称=“主视觉喷绘”的记录，无法匹配单价。");
      }

      return mainVisualPriceBook;
    }

    function rehydrateSelectionsFromTemplates() {
      // 使用模板重算占位选择（成本/单价/描述），并套用“默认跟随”逻辑
      currentTemplates.forEach(function (tpl) {
        var ex = selected.get(tpl.key);
        if (!ex) return;

        // 计算应选供应商 option
        var idx = 0;

        if (ex.__supplierManual) {
          // 手动选的：优先保持 optIndex，其次按 supplier 名称找
          if (typeof ex.__optIndex === "number") idx = ex.__optIndex;
          if (!tpl.options[idx]) idx = findOptionIndexBySupplier(tpl, ex.supplier || "");
        } else {
          // 自动跟随的：优先跟随主视觉供应商；若无，则尽量按原 supplier 保持
          var defSup = getDefaultSupplierForMajor(tpl.major);
          if (defSup) idx = findOptionIndexBySupplier(tpl, defSup);
          else if (ex.supplier) idx = findOptionIndexBySupplier(tpl, ex.supplier);
          else if (typeof ex.__optIndex === "number") idx = ex.__optIndex;
        }

        applyPick(tpl.key, tpl, idx, ex.qty, !!ex.__supplierManual);
      });
    }

    async function refreshProducts() {
      clearMsg();

      var majors = getSelectedMajors();
      if (!majors || majors.length === 0) {
        currentTemplates = [];
        rebuildTemplatesMap();
        productList.innerHTML = '<div class="hint">请先在步骤1选择至少一个产品大类</div>';
        updateViewButtonsText();
        return;
      }

      productList.innerHTML = '<div class="hint">加载中...</div>';

      var selectedMetas = [];
      for (var i = 0; i < majors.length; i++) {
        var bucket = majorIndex.get(majors[i]);
        if (bucket && bucket.metas) {
          for (var j = 0; j < bucket.metas.length; j++) selectedMetas.push(bucket.metas[j]);
        }
      }

      var allItems = [];
      var warnings = [];

      for (var k = 0; k < selectedMetas.length; k++) {
        try {
          var items = await loadProductsFromTable({
            id: selectedMetas[k].id,
            name: selectedMetas[k].name,
          });
          allItems = allItems.concat(items);
        } catch (e) {
          warnings.push(
            "表【" +
              selectedMetas[k].name +
              "】加载失败：" +
              (e && e.message ? e.message : String(e))
          );
        }
      }

      var templates = buildProductTemplates(allItems);

      // 清理已选：不在本次模板中的产品从已选里移除
      var alive = new Set();
      for (var ti = 0; ti < templates.length; ti++) alive.add(templates[ti].key);
      selected.forEach(function (v, key) {
        if (!alive.has(key)) selected.delete(key);
      });
      persistSelectedDraft();

      currentTemplates = templates;
      rebuildTemplatesMap();

      // 用模板“重算”占位选择，并套用自动跟随供应商逻辑
      rehydrateSelectionsFromTemplates();

      // 主视觉供应商变化后，推送到未手动改过的已选（物料搭建/印刷制作）
      syncAutoSuppliersFromMainVisual();

      renderProducts(currentTemplates);

      if (warnings.length > 0) setMsg("部分表无法加载：\n" + warnings.join("\n"), "err");
    }

    // === 事件绑定 ===

    majorSelectEl.addEventListener("change", function () {
      refreshProducts().catch(function (e) {
        setMsg("刷新产品失败：" + (e && e.message ? e.message : String(e)), "err");
      });
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
          renderMainVisualList();
          // 打开主视觉配置时，同步一次默认供应商到未手动的已选项
          syncAutoSuppliersFromMainVisual();
          renderProducts(currentTemplates);
        }
      });
    }

    if (mvSupplier) {
      mvSupplier.addEventListener("change", function () {
        mainVisualDefaultSupplier = mvSupplier.value || "";
        // 主视觉供应商变更 -> 推送到未手动改过的已选（物料搭建/印刷制作）
        syncAutoSuppliersFromMainVisual();
        renderProducts(currentTemplates);
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
            setMsg("主视觉喷绘配置失败：" + (e && e.message ? e.message : String(e)), "err");
          });
      });
    }

    if (selectAllBtn) {
      selectAllBtn.addEventListener("click", function () {
        for (var i = 0; i < majorSelectEl.options.length; i++) {
          majorSelectEl.options[i].selected = true;
        }
        refreshProducts().catch(function (e) {
          setMsg("刷新产品失败：" + (e && e.message ? e.message : String(e)), "err");
        });
      });
    }

    if (clearAllBtn) {
      clearAllBtn.addEventListener("click", function () {
        for (var i = 0; i < majorSelectEl.options.length; i++) {
          majorSelectEl.options[i].selected = false;
        }

        clearSelected();
        currentTemplates = [];
        rebuildTemplatesMap();
        productList.innerHTML = '<div class="hint">请先在步骤1选择至少一个产品大类</div>';
        updateViewButtonsText();
      });
    }

    if (reloadBtn) {
      reloadBtn.addEventListener("click", function () {
        refreshProducts().catch(function (e) {
          setMsg("刷新产品失败：" + (e && e.message ? e.message : String(e)), "err");
        });
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
              });
            }

            doExport(itemsToExport, quoteName);
          })
          .catch(function (e) {
            setMsg("导出失败：" + (e && e.message ? e.message : String(e)), "err");
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
            setMsg("已导出 XLSX（含公式/列宽/换行/行高/分组/小计/含税合计）。", "ok");
          })
          .catch(function (e) {
            setMsg("导出失败：" + (e && e.message ? e.message : String(e)), "err");
          });
      }
    });

    // === 首次渲染 ===
    renderMajorSelectOptions();

    // 还原“单次操作暂存”的已选
    var draftArr = restoreSelectedDraft();

    // 如果暂存里有内容：尝试自动勾选相关大类（按 sourceTableName 反推可能的大类）
    // 注意：暂存占位里没有 sourceTableName，这里只尽量根据 supplierManual 以外的信息无法推断；
    // 因此仅当用户本次已选择大类时才能完整重建。这里做“无害尝试”：如果之前本次操作已经选过大类，用户通常会保留。
    //（实际大类勾选仍以用户当前选择为准）
    if (draftArr && draftArr.length > 0) {
      // 不强行选大类（避免误选）；但会提示用户可直接切到“已选产品”查看（当大类已选时）
      setProductViewMode("selected");
    } else {
      setProductViewMode("pool");
    }

    // 主视觉默认供应商初始化（即使未勾选主视觉，也可用于默认跟随）
    refillMainVisualSuppliers();

    updateViewButtonsText();
    setMsg("初始化完成：请在步骤1选择产品大类。", "ok");
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
