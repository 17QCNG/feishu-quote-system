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

      // 关键：很多字段（自动编号/公式等）会返回 { value: "1", ... }
      if (typeof v.value === "string" || typeof v.value === "number") return String(v.value);

      // 某些选择/人员对象可能有 name
      if (typeof v.name === "string") return v.name;

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

  function asRichText(text) {
    return [{ type: "text", text: text || "" }];
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

  // getRecords() 返回 { records, hasMore, pageToken, total }，这里拉全量（最多每页 200，循环分页）
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

  async function main() {
    clearMsg();

    var tableSelect = document.getElementById("tableSelect");
    var productList = document.getElementById("productList");
    var customerName = document.getElementById("customerName");
    var outputMode = document.getElementById("outputMode");
    var quoteTableSelect = document.getElementById("quoteTableSelect");
    var generateBtn = document.getElementById("generateQuote");

    if (!tableSelect || !productList || !customerName || !outputMode || !quoteTableSelect || !generateBtn) {
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

    // 选中产品：recordId -> item
    var selected = new Map();

    // 加载表列表（给产品表/目标表下拉框）
    var metas = [];
    try {
      metas = await getAllTables(bitable);
    } catch (e2) {
      setMsg("读取数据表列表失败：" + (e2 && e2.message ? e2.message : String(e2)), "err");
      return;
    }

    tableSelect.innerHTML = '<option value="">请选择产品库表...</option>';
    quoteTableSelect.innerHTML = '<option value="">请选择写入的目标表...</option>';

    for (var m = 0; m < metas.length; m++) {
      var opt = document.createElement("option");
      opt.value = metas[m].id;
      opt.textContent = metas[m].name;
      tableSelect.appendChild(opt);

      var opt2 = document.createElement("option");
      opt2.value = metas[m].id;
      opt2.textContent = metas[m].name;
      quoteTableSelect.appendChild(opt2);
    }

    // 切换输出方式时给个提示
    outputMode.addEventListener("change", function () {
      clearMsg();
      if (outputMode.value === "bitable") {
        setMsg("已切换为：写入多维表格。请在下方选择“写入的目标表”。", "ok");
      } else {
        setMsg("已切换为：导出 CSV。生成后可在飞书在线表格中导入。", "ok");
      }
      setTimeout(function () {
        clearMsg();
      }, 1500);
    });

    tableSelect.addEventListener("change", async function () {
      clearMsg();
      selected.clear();

      var tableId = tableSelect.value;
      if (!tableId) {
        productList.innerHTML = '<p class="hint">请先选择产品库表</p>';
        return;
      }

      productList.innerHTML = '<p class="hint">加载中...</p>';

      try {
        var table = await bitable.base.getTableById(tableId);

        // 你要求识别的 7 个字段（这里带少量兼容别名，避免你以后改名）
        var codeRes = await getFieldByAnyName(table, ["产品编号", "产品编码", "编号", "编码"], true);
        var nameRes = await getFieldByAnyName(table, ["产品名称", "名称"], true);
        var typeRes = await getFieldByAnyName(table, ["产品类型", "类型"], false);
        var unitRes = await getFieldByAnyName(table, ["计算单位", "单位"], false);
        var costRes = await getFieldByAnyName(table, ["成本单价", "成本", "成本价"], false);
        var priceRes = await getFieldByAnyName(table, ["单价", "售价", "报价"], true);
        var descRes = await getFieldByAnyName(table, ["产品描述", "描述", "备注"], false);

        var codeField = codeRes.field;
        var nameField = nameRes.field;
        var typeField = typeRes.field;
        var unitField = unitRes.field;
        var costField = costRes.field;
        var priceField = priceRes.field;
        var descField = descRes.field;

        var records = await getAllRecords(table, 200);

        if (!records || records.length === 0) {
          productList.innerHTML = '<p class="hint">该表没有记录，请先添加产品数据。</p>';
          return;
        }

        productList.innerHTML = "";

        for (var i = 0; i < records.length; i++) {
          var r = records[i];
          var recordId = r.recordId || r.id;
          var fields = r.fields || {};

          var item = {
            code: toPlainText(fields[codeField.id]) || "",
            name: toPlainText(fields[nameField.id]) || "未命名产品",
            type: typeField ? (toPlainText(fields[typeField.id]) || "") : "",
            unit: unitField ? (toPlainText(fields[unitField.id]) || "") : "",
            cost: costField ? toNumber(fields[costField.id]) : 0,
            price: toNumber(fields[priceField.id]),
            desc: descField ? (toPlainText(fields[descField.id]) || "") : "",
            qty: 1,
          };

          var row = document.createElement("div");
          row.className = "row";

          var cb = document.createElement("input");
          cb.type = "checkbox";

          var mid = document.createElement("div");
          mid.innerHTML =
            '<div class="pname">' + item.name + "</div>" +
            '<div class="pmeta">' +
              "编号：" + (item.code || "—") +
              (item.type ? "　类型：" + item.type : "") +
              (item.unit ? "　单位：" + item.unit : "") +
              "　单价：¥" + item.price +
              (costField ? "　成本：¥" + item.cost : "") +
            "</div>" +
            (item.desc ? '<div class="pmeta">描述：' + item.desc + "</div>" : "");

          var qty = document.createElement("input");
          qty.className = "qty";
          qty.type = "number";
          qty.min = "1";
          qty.step = "1";
          qty.value = "1";
          qty.disabled = true;

          cb.addEventListener("change", (function (rid, it, qtyEl) {
            return function () {
              if (this.checked) {
                qtyEl.disabled = false;
                var next = Object.assign({}, it);
                next.qty = Math.max(1, Number(qtyEl.value) || 1);
                selected.set(rid, next);
              } else {
                qtyEl.disabled = true;
                selected.delete(rid);
              }
            };
          })(recordId, item, qty));

          qty.addEventListener("input", (function (rid, qtyEl) {
            return function () {
              var it2 = selected.get(rid);
              if (!it2) return;
              it2.qty = Math.max(1, Number(qtyEl.value) || 1);
              selected.set(rid, it2);
            };
          })(recordId, qty));

          row.appendChild(cb);
          row.appendChild(mid);
          row.appendChild(qty);
          productList.appendChild(row);
        }
      } catch (e3) {
        setMsg("加载产品失败：" + (e3 && e3.message ? e3.message : String(e3)), "err");
        productList.innerHTML = '<p class="hint">加载失败，请检查字段名是否一致。</p>';
      }
    });

    generateBtn.addEventListener("click", async function () {
      clearMsg();

      var cname = (customerName.value || "").trim();
      if (!cname) {
        setMsg("请先填写客户名称。", "err");
        return;
      }
      if (selected.size === 0) {
        setMsg("请至少勾选一个产品。", "err");
        return;
      }

      // 先把明细整理出来（无论导出/写入都用同一份数据）
      var lines = [];
      lines.push([
        "客户名称",
        "产品编号",
        "产品名称",
        "产品类型",
        "计算单位",
        "成本单价",
        "单价",
        "产品描述",
        "数量",
        "小计",
      ]);

      var total = 0;
      var items = [];
      for (const it of selected.values()) {
        var subtotal = (Number(it.qty) || 0) * (Number(it.price) || 0);
        total += subtotal;

        items.push(it);
        lines.push([
          cname,
          it.code || "",
          it.name || "",
          it.type || "",
          it.unit || "",
          String(it.cost || 0),
          String(it.price || 0),
          (it.desc || "").replace(/\r?\n/g, " "),
          String(it.qty || 1),
          String(subtotal),
        ]);
      }
      lines.push(["合计", "", "", "", "", "", "", "", "", String(total)]);

      if (outputMode.value === "csv") {
        var csv = lines
          .map(function (row) {
            return row.map(csvEscape).join(",");
          })
          .join("\r\n");

        downloadText("报价单-" + (cname || "客户") + ".csv", csv, "text/csv;charset=utf-8");
        setMsg("已导出 CSV：可在飞书在线表格中导入生成报价单。", "ok");

        selected.clear();
        customerName.value = "";
        tableSelect.dispatchEvent(new Event("change"));
        return;
      }

      // 写入多维表格模式
      var targetTableId = quoteTableSelect.value;
      if (!targetTableId) {
        setMsg("请先选择要写入的目标表（输出方式=写入多维表格）。", "err");
        return;
      }

      try {
        var quoteTable = await bitable.base.getTableById(targetTableId);

        // 必需字段：客户名称、数量、产品名称、单价（产品编号也强烈建议）
        var fCustomer = (await getFieldByAnyName(quoteTable, ["客户名称"], true)).field;
        var fQty = (await getFieldByAnyName(quoteTable, ["数量"], true)).field;
        var fName = (await getFieldByAnyName(quoteTable, ["产品名称", "名称"], true)).field;
        var fPrice = (await getFieldByAnyName(quoteTable, ["单价", "售价", "报价"], true)).field;

        // 可选字段：产品编号/类型/单位/成本/描述
        var fCode = (await getFieldByAnyName(quoteTable, ["产品编号", "产品编码", "编号", "编码"], false)).field;
        var fType = (await getFieldByAnyName(quoteTable, ["产品类型", "类型"], false)).field;
        var fUnit = (await getFieldByAnyName(quoteTable, ["计算单位", "单位"], false)).field;
        var fCost = (await getFieldByAnyName(quoteTable, ["成本单价", "成本", "成本价"], false)).field;
        var fDesc = (await getFieldByAnyName(quoteTable, ["产品描述", "描述", "备注"], false)).field;

        var count = 0;
        for (var i = 0; i < items.length; i++) {
          var it3 = items[i];
          var fs = {};
          fs[fCustomer.id] = asRichText(cname);
          fs[fQty.id] = it3.qty;
          fs[fName.id] = asRichText(it3.name);
          fs[fPrice.id] = it3.price;

          if (fCode) fs[fCode.id] = asRichText(it3.code);
          if (fType) fs[fType.id] = asRichText(it3.type);
          if (fUnit) fs[fUnit.id] = asRichText(it3.unit);
          if (fCost) fs[fCost.id] = it3.cost;
          if (fDesc) fs[fDesc.id] = asRichText(it3.desc);

          await quoteTable.addRecord({ fields: fs });
          count++;
        }

        setMsg("已写入目标表：" + count + " 条报价明细。", "ok");

        selected.clear();
        customerName.value = "";
        tableSelect.dispatchEvent(new Event("change"));
      } catch (e4) {
        setMsg("写入失败：" + (e4 && e4.message ? e4.message : String(e4)), "err");
      }
    });

    setMsg("初始化完成：请选择产品库表开始操作。", "ok");
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
