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

  // Base JS SDK: getRecords() 返回 { records, hasMore, pageToken, total }
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

  async function main() {
    clearMsg();

    var tableSelect = document.getElementById("tableSelect");
    var productList = document.getElementById("productList");
    var customerName = document.getElementById("customerName");
    var generateBtn = document.getElementById("generateQuote");

    if (!tableSelect || !productList || !customerName || !generateBtn) {
      setMsg("页面元素缺失：请确认 index.html 与 app.js 已正确上传，并且没有重复 id。", "err");
      return;
    }

    var bitable;
    try {
      bitable = await waitForBitable(5000);
    } catch (e) {
      setMsg(
        "飞书 SDK 未加载：请在【飞书多维表格】里通过【扩展/自定义插件】打开本页面；如果你用 GitHub Pages 直接访问会失败（这是正常的）。",
        "err"
      );
      return;
    }

    // recordId -> item
    var selected = new Map();

    // 表列表
    var metas = [];
    try {
      metas = await getAllTables(bitable);
    } catch (e2) {
      setMsg("读取数据表列表失败：" + (e2 && e2.message ? e2.message : String(e2)), "err");
      return;
    }

    tableSelect.innerHTML = '<option value="">请选择产品库表...</option>';
    for (var m = 0; m < metas.length; m++) {
      var opt = document.createElement("option");
      opt.value = metas[m].id;
      opt.textContent = metas[m].name;
      tableSelect.appendChild(opt);
    }

    tableSelect.addEventListener("change", async function () {
      clearMsg();
      selected.clear();

      var tableId = tableSelect.value;
      if (!tableId) {
        productList.innerHTML = '<div class="hint">请先选择产品库表</div>';
        return;
      }

      productList.innerHTML = '<div class="hint">加载中...</div>';

      try {
        var table = await bitable.base.getTableById(tableId);

        // 识别你要求的 7 个字段（按你给的中文名；另带少量兼容）
        var fCode = (await getFieldByAnyName(table, ["产品编号", "产品编码", "编号", "编码"], true)).field;
        var fName = (await getFieldByAnyName(table, ["产品名称", "名称"], true)).field;
        var fType = (await getFieldByAnyName(table, ["产品类型", "类型"], false)).field;
        var fUnit = (await getFieldByAnyName(table, ["计算单位", "单位"], false)).field;
        var fCost = (await getFieldByAnyName(table, ["成本单价", "成本", "成本价"], false)).field;
        var fPrice = (await getFieldByAnyName(table, ["单价", "售价", "报价"], true)).field;
        var fDesc = (await getFieldByAnyName(table, ["产品描述", "描述", "备注"], false)).field;

        var records = await getAllRecords(table, 200);

        if (!records || records.length === 0) {
          productList.innerHTML = '<div class="hint">该表没有记录，请先添加产品数据。</div>';
          return;
        }

        productList.innerHTML = "";

        for (var i = 0; i < records.length; i++) {
          var r = records[i];
          var recordId = r.recordId || r.id;
          var fields = r.fields || {};

          var item = {
            code: toPlainText(fields[fCode.id]) || "",
            name: toPlainText(fields[fName.id]) || "未命名产品",
            type: fType ? (toPlainText(fields[fType.id]) || "") : "",
            unit: fUnit ? (toPlainText(fields[fUnit.id]) || "") : "",
            cost: fCost ? toNumber(fields[fCost.id]) : 0,
            price: toNumber(fields[fPrice.id]),
            desc: fDesc ? (toPlainText(fields[fDesc.id]) || "") : "",
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
              (fCost ? "　成本：¥" + item.cost : "") +
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
              var it = selected.get(rid);
              if (!it) return;
              it.qty = Math.max(1, Number(qtyEl.value) || 1);
              selected.set(rid, it);
            };
          })(recordId, qty));

          row.appendChild(cb);
          row.appendChild(mid);
          row.appendChild(qty);
          productList.appendChild(row);
        }
      } catch (e3) {
        setMsg("加载产品失败：" + (e3 && e3.message ? e3.message : String(e3)), "err");
        productList.innerHTML = '<div class="hint">加载失败，请检查字段名与表内是否一致。</div>';
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

      try {
        var quoteTable = await bitable.base.getTableByName("报价单");

        // 最基本必需字段（建议你的“报价单”表至少有这些）
        var customerField = await quoteTable.getFieldByName("客户名称");
        var qtyField = await quoteTable.getFieldByName("数量");
        var nameField = (await getFieldByAnyName(quoteTable, ["产品名称", "名称"], true)).field;
        var priceField = (await getFieldByAnyName(quoteTable, ["单价", "售价", "报价"], true)).field;

        // 可选：如果“报价单”表也有这些列，就自动写入（建议做成 文本/数字 字段）
        var codeField = (await getFieldByAnyName(quoteTable, ["产品编号", "产品编码"], false)).field;
        var costField = (await getFieldByAnyName(quoteTable, ["成本单价", "成本", "成本价"], false)).field;
        var descField = (await getFieldByAnyName(quoteTable, ["产品描述", "描述", "备注"], false)).field;

        var count = 0;
        for (const it of selected.values()) {
          var fs = {};
          fs[customerField.id] = asRichText(cname);
          fs[qtyField.id] = it.qty;
          fs[nameField.id] = asRichText(it.name);
          fs[priceField.id] = it.price;

          if (codeField) fs[codeField.id] = asRichText(it.code);
          if (costField) fs[costField.id] = it.cost;
          if (descField) fs[descField.id] = asRichText(it.desc);

          await quoteTable.addRecord({ fields: fs });
          count++;
        }

        setMsg("已生成报价单：写入 " + count + " 条明细到“报价单”表。", "ok");

        selected.clear();
        customerName.value = "";
        tableSelect.dispatchEvent(new Event("change"));
      } catch (e4) {
        setMsg("写入报价单失败：" + (e4 && e4.message ? e4.message : String(e4)), "err");
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
