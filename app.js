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
