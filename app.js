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
    if (
      bitable.base &&
      typeof bitable.base.getTableMetaList === "function"
    ) {
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
    var box = document.getElementById("tableList");
    if (!box) return [];
    var inputs = box.querySelectorAll(
      'input[type="checkbox"][data-major]'
    );
    var majors = [];
    for (var i = 0; i < inputs.length; i++) {
      if (inputs[i].checked) majors.push(inputs[i].getAttribute("data-major"));
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

        ws.getCell(rowNum, 7).value = { formula: "SUM(" + gRefs.join(",") + ")" };
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
     /*  if (it && it.supplier) {
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
      row.getCell(9).value = { formula: "H" + rn + "*D" + rn }; */
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
          options: [],
        };
        map.set(k, tpl);
      }

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

  async function main() {
    clearMsg();

    var tableListEl = document.getElementById("tableList");
    var productList = document.getElementById("productList");
    var exportBtn = document.getElementById("exportXlsx");
    var selectAllBtn = document.getElementById("selectAllTables");
    var clearAllBtn = document.getElementById("clearAllTables");
    var reloadBtn = document.getElementById("reloadProducts");
    var quoteNameInput = getQuoteNameInput();

    var needMainVisualCb = document.getElementById("needMainVisual");
    var mvBox = document.getElementById("mainVisualConfig");
    var mvSupplier = document.getElementById("mainVisualSupplier");
    var mvSize = document.getElementById("mainVisualSize");
    var mvQty = document.getElementById("mainVisualQty");
    var mvAdd = document.getElementById("addMainVisual");
    var mvList = document.getElementById("mainVisualList");

    if (!tableListEl || !productList || !exportBtn) {
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

    var selected = new Map(); // templateKey -> picked item
    var metas = [];
    var metasParsed = [];
    var majorIndex = new Map(); // major -> { major, metas: [{id,name,major,supplier}], suppliers:Set() }
    var currentTemplates = [];

    var mainVisualItems = []; // { supplier, size, qty }
    var mainVisualPriceBook = null; // Map supplier -> { cost, price, unit, desc, sourceTableName }

    try {
      metas = await getAllTables(bitable);
    } catch (e2) {
      setMsg("读取数据表列表失败：" + (e2 && e2.message ? e2.message : String(e2)), "err");
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

    function renderMajorCheckboxes() {
      tableListEl.innerHTML = "";

      if (!majorIndex || majorIndex.size === 0) {
        tableListEl.innerHTML = '<div class="hint">当前 Base 没有数据表</div>';
        return;
      }

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

        var item = document.createElement("div");
        item.className = "table-item";

        var cb = document.createElement("input");
        cb.type = "checkbox";
        cb.setAttribute("data-major", major);

        var name = document.createElement("div");
        name.className = "table-name";

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

        name.textContent = major + suffix;

        cb.addEventListener("change", function () {
          refreshProducts().catch(function (e) {
            setMsg("刷新产品失败：" + (e && e.message ? e.message : String(e)), "err");
          });
        });

        item.appendChild(cb);
        item.appendChild(name);
        tableListEl.appendChild(item);
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

    function renderProducts(templates) {
      productList.innerHTML = "";

      if (!templates || templates.length === 0) {
        productList.innerHTML = '<div class="hint">选中的产品大类没有可用产品（可能无记录）。</div>';
        return;
      }

      for (var i = 0; i < templates.length; i++) {
        (function () {
          var tpl = templates[i];
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
          mid.innerHTML =
            '<div class="pname">' +
            escapeHtml(tpl.name) +
            "</div>" +
            '<div class="pmeta"><span class="hl">成本：¥' +
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
            o.textContent = opt.supplier ? opt.supplier : "默认供应商";
            supplierSel.appendChild(o);
          }

          supplierSel.value =
            picked && typeof picked.__optIndex === "number"
              ? String(picked.__optIndex)
              : "0";

          var qty = document.createElement("input");
          qty.className = "qty";
          qty.type = "number";
          qty.min = "1";
          qty.step = "1";
          qty.value = String(picked ? picked.qty : 1);
          qty.disabled = !picked;

          function applyPick(templateKey, tplRef, supplierIndex, qtyValue) {
            var idx = Math.max(0, Number(supplierIndex) || 0);
            var opt2 = tplRef.options[idx] || tplRef.options[0];
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
            };

            selected.set(templateKey, next);
          }

          cb.addEventListener("change", function () {
            if (this.checked) {
              qty.disabled = false;
              supplierSel.disabled = false;
              applyPick(tpl.key, tpl, supplierSel.value, qty.value);
            } else {
              qty.disabled = true;
              supplierSel.disabled = true;
              selected.delete(tpl.key);
            }
            renderProducts(currentTemplates);
          });

          supplierSel.addEventListener("change", function () {
            if (!selected.get(tpl.key)) return;
            applyPick(tpl.key, tpl, supplierSel.value, qty.value);
            renderProducts(currentTemplates);
          });

          qty.addEventListener("input", function () {
            var it2 = selected.get(tpl.key);
            if (!it2) return;
            it2.qty = Math.max(1, Number(qty.value) || 1);
            selected.set(tpl.key, it2);
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

    async function refreshProducts() {
      clearMsg();

      var majors = getSelectedMajors();
      if (!majors || majors.length === 0) {
        currentTemplates = [];
        productList.innerHTML = '<div class="hint">请先在步骤1勾选至少一个产品大类</div>';
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

      var alive = new Set();
      for (var ti = 0; ti < templates.length; ti++) alive.add(templates[ti].key);
      selected.forEach(function (v, key) {
        if (!alive.has(key)) selected.delete(key);
      });

      currentTemplates = templates;
      renderProducts(currentTemplates);

      if (warnings.length > 0) setMsg("部分表无法加载：\n" + warnings.join("\n"), "err");
    }

    if (needMainVisualCb && mvBox) {
      needMainVisualCb.addEventListener("change", function () {
        mvBox.style.display = this.checked ? "block" : "none";
        if (this.checked) {
          refillMainVisualSuppliers();
          renderMainVisualList();
        }
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
        var inputs = tableListEl.querySelectorAll('input[type="checkbox"][data-major]');
        for (var i = 0; i < inputs.length; i++) inputs[i].checked = true;
        refreshProducts().catch(function (e) {
          setMsg("刷新产品失败：" + (e && e.message ? e.message : String(e)), "err");
        });
      });
    }

    if (clearAllBtn) {
      clearAllBtn.addEventListener("click", function () {
        var inputs = tableListEl.querySelectorAll('input[type="checkbox"][data-major]');
        for (var i = 0; i < inputs.length; i++) inputs[i].checked = false;

        selected.clear();
        currentTemplates = [];
        productList.innerHTML = '<div class="hint">请先在步骤1勾选至少一个产品大类</div>';
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
                  "供应商【" + sup + "】缺少“主视觉喷绘”定价，请先在【物料搭建（" + sup + "）】补齐。"
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

    renderMajorCheckboxes();
    setMsg("初始化完成：请在步骤1勾选产品大类。", "ok");
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
