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

  async function getAllTables(bitable) {
    // 优先：getTableMetaList -> [{id,name}]
    if (bitable.base && typeof bitable.base.getTableMetaList === "function") {
      var metaList = await bitable.base.getTableMetaList();
      return metaList || [];
    }

    // 兜底：getTableList + getName
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
      // 关键：必须在飞书多维表格里作为插件打开。浏览器直接打开 Pages 会失败（这是正常的）
      bitable = await waitForBitable(5000);
    } catch (e) {
      setMsg(
        "飞书 SDK 未加载：请在【飞书多维表格】里通过【扩展/自定义插件】打开本页面，不要直接在浏览器访问。",
        "err"
      );
      return;
    }

    // 选中产品：recordId -> {name, code, price, qty}
    var selected = new Map();

    // 加载表列表
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

        // 按字段名找字段（字段名必须一字不差）
        var nameField = await table.getFieldByName("产品名称");
        var codeField = await table.getFieldByName("产品编码");
        var priceField = await table.getFieldByName("单价");

        var records = await table.getRecords({ pageSize: 200 });

        if (!records || records.length === 0) {
          productList.innerHTML = '<div class="hint">该表没有记录，请先添加产品数据。</div>';
          return;
        }

        productList.innerHTML = "";

        for (var i = 0; i < records.length; i++) {
          var r = records[i];
          var recordId = r.recordId || r.id;
          var fields = r.fields || {};

          var rawName = fields[nameField.id];
          var rawCode = fields[codeField.id];
          var rawPrice = fields[priceField.id];

          var pname = toPlainText(rawName) || "未命名产品";
          var pcode = toPlainText(rawCode) || "";
          var price = toNumber(rawPrice);

          var row = document.createElement("div");
          row.className = "row";

          var cb = document.createElement("input");
          cb.type = "checkbox";

          var mid = document.createElement("div");
          mid.innerHTML =
            '<div class="pname">' +
            pname +
            "</div>" +
            '<div class="pmeta">编码：' +
            (pcode || "—") +
            "　单价：¥" +
            price +
            "</div>";

          var qty = document.createElement("input");
          qty.className = "qty";
          qty.type = "number";
          qty.min = "1";
          qty.step = "1";
          qty.value = "1";
          qty.disabled = true;

          cb.addEventListener("change", (function (rid, name, code, p, qtyEl) {
            return function () {
              if (this.checked) {
                qtyEl.disabled = false;
                selected.set(rid, { name: name, code: code, price: p, qty: Math.max(1, Number(qtyEl.value) || 1) });
              } else {
                qtyEl.disabled = true;
                selected.delete(rid);
              }
            };
          })(recordId, pname, pcode, price, qty));

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
        setMsg(
          "加载产品失败：" +
            (e3 && e3.message ? e3.message : String(e3)) +
            "。请确认产品表存在字段：产品名称/产品编码/单价",
          "err"
        );
        productList.innerHTML = '<div class="hint">加载失败，请检查字段名是否一致。</div>';
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

        var customerField = await quoteTable.getFieldByName("客户名称");
        var productField = await quoteTable.getFieldByName("产品名称");
        var qtyField = await quoteTable.getFieldByName("数量");
        var priceField = await quoteTable.getFieldByName("单价");

        var count = 0;
        selected.forEach(async function () {});

        // 逐条写入
        for (const it of selected.values()) {
          await quoteTable.addRecord({
            fields: {
              [customerField.id]: asRichText(cname),
              [productField.id]: asRichText(it.name),
              [qtyField.id]: it.qty,
              [priceField.id]: it.price,
            },
          });
          count++;
        }

        setMsg("已生成报价单：写入 " + count + " 条明细到“报价单”表。", "ok");

        // 清空
        selected.clear();
        customerName.value = "";
        tableSelect.dispatchEvent(new Event("change"));
      } catch (e4) {
        setMsg(
          "写入报价单失败：" +
            (e4 && e4.message ? e4.message : String(e4)) +
            "。请确认“报价单”表存在字段：客户名称/产品名称/数量/单价",
          "err"
        );
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

