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

      // 常见：自动编号/公式等字段返回 { value: "...", ... }
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

  // Base JS SDK: getRecords() -> { records, hasMore, pageToken, total }
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

  function csvEscape(s) {
    var x = String(s == null ? "" : s);
    if (/[",\r\n]/.test(x)) return '"' + x.replace(/"/g, '""') + '"';
    return x;
  }

  function downloadText(filename, text, mime) {
    var blob = new Blob([text], { type: mime || "text/plain;charset=utf-8" });
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

  async function main() {
    clearMsg();

    var tableListEl = document.getElementById("tableList");
    var productList = document.getElementById("productList");
    var exportBtn = document.getElementById("exportCsv");
    var selectAllBtn = document.getElementById("selectAllTables");
    var clearAllBtn = document.getElementById("clearAllTables");
    var reloadBtn = document.getElementById("reloadProducts");

    if (!tableListEl || !productList || !exportBtn) {
      setMsg("页面元素缺失：请确认 index.html 与 app.js 已正确上传。", "err");
      return;
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

    // key(tableId:recordId) -> item
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

      // 按你的新要求：不再使用“产品类型”字段；产品类型来自数据表名称
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
          sourceTableName: meta.name, // 输出“产品类型”
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
          '<div class="pname">' + item.name + "</div>" +
          '<div class="pmeta">' +
            "产品类型：" + (item.sourceTableName || "—") +
            "　编号：" + (item.code || "—") +
            "　尺寸/天数：" + (item.sizeDays || "—") +
            "　单位：" + (item.unit || "—") +
            "　成本：¥" + (item.cost || 0) +
            "　单价：¥" + (item.price || 0) +
          "</div>" +
          (item.desc ? '<div class="pmeta">描述：' + item.desc + "</div>" : "");

        var qty = document.createElement("input");
        qty.className = "qty";
        qty.type = "number";
        qty.min = "1";
        qty.step = "1";
        qty.value = String(picked ? picked.qty : 1);
        qty.disabled = !picked;

        cb.addEventListener("change", (function (k, it, qtyEl) {
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
        })(key, item, qty));

        qty.addEventListener("input", (function (k2, qtyEl2) {
          return function () {
            var it2 = selected.get(k2);
            if (!it2) return;
            it2.qty = Math.max(1, Number(qtyEl2.value) || 1);
            selected.set(k2, it2);
          };
        })(key, qty));

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
          warnings.push("表【" + selectedMetas[j].name + "】加载失败：" + (e && e.message ? e.message : String(e)));
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

      // 列顺序固定：
      // A 产品编号
      // B 产品类型(表名)
      // C 尺寸/天数
      // D 数量
      // E 单位
      // F 成本单价
      // G 成本总价(公式)=F*D
      // H 单价
      // I 总价(公式)=H*D
      // J 产品描述
      var lines = [];
      lines.push([
        "产品编号",
        "产品类型",
        "尺寸/天数",
        "数量",
        "单位",
        "成本单价",
        "成本总价",
        "单价",
        "总价",
        "产品描述",
      ]);

      var rowIndex = 2; // 第 1 行是表头，数据从第 2 行开始
      selected.forEach(function (it) {
        // 公式用 A1 引用：成本总价=F{row}*D{row}，总价=H{row}*D{row}
        var costTotalFormula = "=F" + rowIndex + "*D" + rowIndex;
        var totalFormula = "=H" + rowIndex + "*D" + rowIndex;

        lines.push([
          it.code || "",
          it.sourceTableName || "",
          it.sizeDays || "",
          String(it.qty || 1),
          it.unit || "",
          String(it.cost || 0),
          costTotalFormula,
          String(it.price || 0),
          totalFormula,
          (it.desc || "").replace(/\r?\n/g, " "),
        ]);

        rowIndex++;
      });

      var csv = lines
        .map(function (row) {
          return row.map(csvEscape).join(",");
        })
        .join("\r\n");

      var filename = "报价单.csv";
      downloadText(filename, csv, "text/csv;charset=utf-8");
      setMsg("已导出 CSV（含公式）：请在飞书在线表格中导入。", "ok");
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
