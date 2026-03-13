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
        if (x && typeof x === "object" && typeof x.text === "string") parts.push(x.text);
        else if (x != null) parts.push(String(x));
      }
      return parts.join("");
    }

    if (typeof v === "object") {
      if (typeof v.text === "string") return v.text;

      if (v && (typeof v.value === "string" || typeof v.value === "number")) return String(v.value);
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

  function getSelectedTableIds() {
    var box = document.getElementById("tableList");
    if (!box) return [];
    var inputs = box.querySelectorAll('input[type="checkbox"][data-table-id]');
    var ids = [];
    for (var i = 0; i < inputs.length; i++) {
      if (inputs[i].checked) ids.push(inputs[i].getAttribute("data-table-id"));
    }
    return ids;
  }

  function buildKey(tableId, recordId) {
    return String(tableId) + ":" + String(recordId);
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
      if (/[\u2E80-\u9FFF\uF900-\uFAFF]/.test(ch)) units += 2; // CJK 及部首等
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
      throw new Error("ExcelJS 未加载：请检查 index.html 是否已引入 exceljs.min.js");
    }

    var fileTitle = safeFileName(quoteName || "报价单");

    var wb = new window.ExcelJS.Workbook();
    wb.creator = "Feishu Quote Plugin";
    wb.created = new Date();

    var ws = wb.addWorksheet("报价单", {
      properties: { defaultRowHeight: 20 },
    });

    // 固定列（需求1/2/3）：删除产品类型/产品编号，最左侧新增“科目”
    // A..J：科目、产品名称、尺寸/天数、数量、单位、成本单价、成本总价、单价、总价、产品描述
    var headers = [
      "科目", // A
      "产品名称", // B
      "尺寸/天数", // C
      "数量", // D
      "单位", // E
      "成本单价", // F
      "成本总价", // G (公式)
      "单价", // H
      "总价", // I (公式)
      "产品描述", // J
    ];

    // 列宽：A窄一点，B/C宽，J更宽，其它适中
    // A=8, B=24, C=24, J=35，其它=9
    var colWidths = [8, 24, 24, 9, 9, 9, 9, 9, 9, 35];
    for (var c = 1; c <= colWidths.length; c++) {
      ws.getColumn(c).width = colWidths[c - 1];
    }

    // 标题行（合并 A1:J1）
    ws.addRow([fileTitle]);
    ws.mergeCells(1, 1, 1, headers.length);

    var titleCell = ws.getCell(1, 1);
    titleCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFED7D31" }, // ARGB：FF + ED7D31
    };
    titleCell.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 16 };
    titleCell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    ws.getRow(1).height = 30;

    // 表头行（第2行）
    ws.addRow(headers);
    var headerRow = ws.getRow(2);
    headerRow.font = { bold: true };
    headerRow.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    headerRow.height = 20;

    // 需求6：填充 A2:I2（2到2I）底色为 FFFF00
    for (var hc = 1; hc <= 9; hc++) {
      ws.getCell(2, hc).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFFF00" }, // FFFF00
      };
    }

    // 导出排序：按产品类型分组并按指定顺序输出；未知类型归并为“其他”
    var typeOrder = [
      "物料搭建",
      "印刷制作",
      "线上",
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
      var t = normalizeName(rawType);
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

      var ai = Object.prototype.hasOwnProperty.call(typeOrderMap, at) ? typeOrderMap[at] : typeOrder.length;
      var bi = Object.prototype.hasOwnProperty.call(typeOrderMap, bt) ? typeOrderMap[bt] : typeOrderMap["其他"];
      if (ai !== bi) return ai - bi;

      // 同类型下：产品名称、编号排序，保证输出稳定且“列在一起”
      var an = normalizeName(a && a.name ? a.name : "");
      var bn = normalizeName(b && b.name ? b.name : "");
      if (an !== bn) return an.localeCompare(bn, "zh");

      var ac = normalizeName(a && a.code ? a.code : "");
      var bc = normalizeName(b && b.code ? b.code : "");
      return ac.localeCompare(bc, "zh");
    });

    // 数据行从第3行开始：按产品类型分组输出（需求4/5）
    var rowNum = 3;

    var groupType = null;
    var groupIdx = 0;
    var groupStartRow = 0; // 本组第一条产品数据行
    var groupEndRow = 0; // 本组最后一条产品数据行

    function styleRowAllCols(rn, fillArgb, bold) {
      for (var cc = 1; cc <= headers.length; cc++) {
        var cell = ws.getCell(rn, cc);
        if (fillArgb) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fillArgb } };
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
      // 需求4：本组最上方插入标题行，底色BFBFBF，合并B到J，A列写“一/二/三...”
      ws.getRow(rowNum).height = 20;

      ws.getCell(rowNum, 1).value = toCnIndex(groupIdx);
      ws.mergeCells(rowNum, 2, rowNum, headers.length); // B..J
      ws.getCell(rowNum, 2).value = typeLabel;

      styleRowAllCols(rowNum, "FFBFBFBF", true);
      rowNum++;
    }

    function addGroupSubtotal(startRn, endRn) {
      // 需求5：本组最下方插入小计行：合并A到F写“小计”，G填成本总价合计，I填总价合计
      ws.getRow(rowNum).height = 20;

      ws.mergeCells(rowNum, 1, rowNum, 6); // A..F
      ws.getCell(rowNum, 1).value = "小计";

      ws.getCell(rowNum, 7).value = { formula: "SUM(G" + startRn + ":G" + endRn + ")" };
      ws.getCell(rowNum, 9).value = { formula: "SUM(I" + startRn + ":I" + endRn + ")" };

      styleRowAllCols(rowNum, null, true);

      // G、I 黄底
      ws.getCell(rowNum, 7).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
      ws.getCell(rowNum, 9).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };

      rowNum++;
    }

    for (var si = 0; si < selectedArr.length; si++) {
      var it = selectedArr[si];
      var typeLabel = canonicalType(it.sourceTableName || "");

      // 进入新分组：先为上一个分组补小计，再插入新组标题行
      if (typeLabel !== groupType) {
        if (groupType != null) {
          addGroupSubtotal(groupStartRow, groupEndRow);
        }
        groupType = typeLabel;
        groupIdx++;
        addGroupHeader(groupType);

        groupStartRow = rowNum;
        groupEndRow = rowNum - 1;
      }

      var desc = normalizeNewlines(it.desc || "");
      var rn = rowNum;
      var row = ws.getRow(rn);

      // A: 科目（留空）
      row.getCell(1).value = "";

      // B: 产品名称
      row.getCell(2).value = it.name || "";

      // C: 尺寸/天数
      row.getCell(3).value = it.sizeDays || "";

      // D: 数量
      row.getCell(4).value = Math.max(1, Number(it.qty) || 1);

      // E: 单位
      row.getCell(5).value = it.unit || "";

      // F: 成本单价
      row.getCell(6).value = Number(it.cost || 0);

      // G: 成本总价 = F * D
      row.getCell(7).value = { formula: "F" + rn + "*D" + rn };

      // H: 单价
      row.getCell(8).value = Number(it.price || 0);

      // I: 总价 = H * D
      row.getCell(9).value = { formula: "H" + rn + "*D" + rn };

      // J: 产品描述
      row.getCell(10).value = desc;

      // 自动换行 + 行高按本行最大内容行数设置；同时设置居中（水平/垂直）
      var maxLines = 1;
      for (var ci = 1; ci <= headers.length; ci++) {
        var cell = ws.getCell(rn, ci);
        cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };

        var v = cell.value;
        if (v && typeof v === "object" && v.formula) continue; // 公式列按1行估算
        if (typeof v === "number") continue;

        var t = v == null ? "" : String(v);
        var w = ws.getColumn(ci).width || 9;
        maxLines = Math.max(maxLines, estimateLines(t, w));
      }

      row.height = rowHeightByLines(Math.min(3, maxLines));
      groupEndRow = rn;
      rowNum++;
    }

    // 最后一个分组补小计
    if (groupType != null && groupStartRow <= groupEndRow) {
      addGroupSubtotal(groupStartRow, groupEndRow);
    }

    // 所有单元格：细线框 + 居中（标题/表头/数据区全覆盖）
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

    var tableListEl = document.getElementById("tableList");
    var productList = document.getElementById("productList");
    var exportBtn = document.getElementById("exportXlsx");
    var selectAllBtn = document.getElementById("selectAllTables");
    var clearAllBtn = document.getElementById("clearAllTables");
    var reloadBtn = document.getElementById("reloadProducts");
    var quoteNameInput = getQuoteNameInput();

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

    var selected = new Map();
    var metas = [];

    try {
      metas = await getAllTables(bitable);
    } catch (e2) {
      setMsg("读取数据表列表失败：" + (e2 && e2.message ? e2.message : String(e2)), "err");
      return;
    }

    function renderTableCheckboxes() {
      tableListEl.innerHTML = "";
      if (!metas || metas.length === 0) {
        tableListEl.innerHTML = '<div class="hint">当前 Base 没有数据表</div>';
        return;
      }

      for (var i = 0; i < metas.length; i++) {
        var meta = metas[i];

        var item = document.createElement("div");
        item.className = "table-item";

        var cb = document.createElement("input");
        cb.type = "checkbox";
        cb.setAttribute("data-table-id", meta.id);

        var name = document.createElement("div");
        name.className = "table-name";
        name.textContent = meta.name;

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

      var fCode = (await getFieldByAnyName(table, ["产品编号"], true)).field;
      var fName = (await getFieldByAnyName(table, ["产品名称"], true)).field;
      var fSizeDays = (await getFieldByAnyName(table, ["尺寸/天数"], true)).field;
      var fUnit = (await getFieldByAnyName(table, ["计算单位"], true)).field;
      var fCost = (await getFieldByAnyName(table, ["成本单价"], true)).field;
      var fPrice = (await getFieldByAnyName(table, ["单价"], true)).field;
      var fDesc = (await getFieldByAnyName(table, ["产品描述"], false)).field;

      var records = await getAllRecords(table, 200);
      var items = [];

      for (var i = 0; i < records.length; i++) {
        var r = records[i];
        var recordId = r.recordId || r.id;
        var fields = r.fields || {};

        items.push({
          sourceTableId: meta.id,
          sourceTableName: meta.name,
          recordId: recordId,

          code: toPlainText(fields[fCode.id]) || "",
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

    function renderProducts(allItems) {
      productList.innerHTML = "";

      if (!allItems || allItems.length === 0) {
        productList.innerHTML = '<div class="hint">选中的数据表没有可用产品（可能无记录）。</div>';
        return;
      }

      for (var i = 0; i < allItems.length; i++) {
        var item = allItems[i];
        var key = buildKey(item.sourceTableId, item.recordId);
        var picked = selected.get(key);

        var row = document.createElement("div");
        row.className = "prow";

        var cb = document.createElement("input");
        cb.type = "checkbox";
        cb.checked = !!picked;

        var mid = document.createElement("div");
        mid.innerHTML =
          '<div class="pname">' +
          item.name +
          "</div>" +
          '<div class="pmeta">' +
          "产品类型：" +
          (item.sourceTableName || "—") +
          "　编号：" +
          (item.code || "—") +
          "　尺寸/天数：" +
          (item.sizeDays || "—") +
          "　单位：" +
          (item.unit || "—") +
          "　成本：¥" +
          (item.cost || 0) +
          "　单价：¥" +
          (item.price || 0) +
          "</div>" +
          (item.desc ? '<div class="pmeta">描述：' + item.desc + "</div>" : "");

        var qty = document.createElement("input");
        qty.className = "qty";
        qty.type = "number";
        qty.min = "1";
        qty.step = "1";
        qty.value = String(picked ? picked.qty : 1);
        qty.disabled = !picked;

        cb.addEventListener(
          "change",
          (function (k, it, qtyEl) {
            return function () {
              if (this.checked) {
                qtyEl.disabled = false;
                var next = Object.assign({}, it);
                next.qty = Math.max(1, Number(qtyEl.value) || 1);
                selected.set(k, next);
              } else {
                qtyEl.disabled = true;
                selected.delete(k);
              }
            };
          })(key, item, qty)
        );

        qty.addEventListener(
          "input",
          (function (k2, qtyEl2) {
            return function () {
              var it2 = selected.get(k2);
              if (!it2) return;
              it2.qty = Math.max(1, Number(qtyEl2.value) || 1);
              selected.set(k2, it2);
            };
          })(key, qty)
        );

        row.appendChild(cb);
        row.appendChild(mid);
        row.appendChild(qty);
        productList.appendChild(row);
      }
    }

    async function refreshProducts() {
      clearMsg();

      var ids = getSelectedTableIds();
      if (!ids || ids.length === 0) {
        productList.innerHTML = '<div class="hint">请先在步骤1勾选至少一个数据表</div>';
        return;
      }

      productList.innerHTML = '<div class="hint">加载中...</div>';

      var selectedMetas = [];
      for (var i = 0; i < metas.length; i++) {
        if (ids.indexOf(metas[i].id) >= 0) selectedMetas.push(metas[i]);
      }

      var allItems = [];
      var warnings = [];

      for (var j = 0; j < selectedMetas.length; j++) {
        try {
          var items = await loadProductsFromTable(selectedMetas[j]);
          allItems = allItems.concat(items);
        } catch (e) {
          warnings.push(
            "表【" +
              selectedMetas[j].name +
              "】加载失败：" +
              (e && e.message ? e.message : String(e))
          );
        }
      }

      renderProducts(allItems);

      if (warnings.length > 0) {
        setMsg("部分表无法加载：\n" + warnings.join("\n"), "err");
      }
    }

    if (selectAllBtn) {
      selectAllBtn.addEventListener("click", function () {
        var inputs = tableListEl.querySelectorAll('input[type="checkbox"][data-table-id]');
        for (var i = 0; i < inputs.length; i++) inputs[i].checked = true;
        refreshProducts().catch(function (e) {
          setMsg("刷新产品失败：" + (e && e.message ? e.message : String(e)), "err");
        });
      });
    }

    if (clearAllBtn) {
      clearAllBtn.addEventListener("click", function () {
        var inputs = tableListEl.querySelectorAll('input[type="checkbox"][data-table-id]');
        for (var i = 0; i < inputs.length; i++) inputs[i].checked = false;
        productList.innerHTML = '<div class="hint">请先在步骤1勾选至少一个数据表</div>';
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

      if (selected.size === 0) {
        setMsg("请至少勾选一个产品。", "err");
        return;
      }

      var quoteName = quoteNameInput ? quoteNameInput.value : "";
      if (!quoteName) quoteName = loadQuoteName() || "报价单";

      exportXlsx(selected, quoteName)
        .then(function () {
          setMsg("已导出 XLSX（含公式/列宽/换行/行高/标题行/分组与小计）。", "ok");
        })
        .catch(function (e) {
          setMsg("导出失败：" + (e && e.message ? e.message : String(e)), "err");
        });
    });

    renderTableCheckboxes();
    setMsg("初始化完成：请在步骤1勾选数据表。", "ok");
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
