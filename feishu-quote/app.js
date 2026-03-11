const { bitable } = window;

let selectedProducts = [];
let productTable = null;

async function init() {
  try {
    const tables = await bitable.base.getTableList();
    const tableSelect = document.getElementById('tableSelect');
    
    tables.forEach(table => {
      const option = document.createElement('option');
      option.value = table.id;
      option.textContent = table.name;
      tableSelect.appendChild(option);
    });
    
    tableSelect.addEventListener('change', async (e) => {
      const tableId = e.target.value;
      if (!tableId) return;
      
      productTable = await bitable.base.getTable(tableId);
      await loadProducts();
    });
    
    document.getElementById('generateQuote').addEventListener('click', generateQuote);
    
  } catch (error) {
    alert('初始化失败: ' + error);
  }
}

async function loadProducts() {
  const list = document.getElementById('productList');
  list.innerHTML = '<p style="color: #999;">加载中...</p>';
  
  try {
    const records = await productTable.getRecords({ pageSize: 100 });
    
    if (records.length === 0) {
      list.innerHTML = '<p style="color: #999;">暂无数据</p>';
      return;
    }
    
    list.innerHTML = '';
    records.forEach(record => {
      const div = document.createElement('div');
      div.style.cssText = 'padding: 10px; margin: 5px 0; background: #f5f5f5; border-radius: 4px; cursor: pointer;';
      div.innerHTML = `
        <strong>${record.fields['产品名称'] || '未命名'}</strong><br>
        <small>编码: ${record.fields['产品编码'] || 'N/A'} | 单价: ¥${record.fields['单价'] || 0}</small>
      `;
      
      div.onclick = () => {
        const isSelected = div.style.background === 'rgb(212, 232, 255)';
        div.style.background = isSelected ? '#f5f5f5' : '#d4e8ff';
        div.style.border = isSelected ? 'none' : '2px solid #3370ff';
        
        const idx = selectedProducts.findIndex(p => p.id === record.id);
        if (idx > -1) {
          selectedProducts.splice(idx, 1);
        } else {
          selectedProducts.push(record);
        }
      };
      
      list.appendChild(div);
    });
  } catch (error) {
    list.innerHTML = '<p style="color: red;">加载失败</p>';
  }
}

async function generateQuote() {
  const customerName = document.getElementById('customerName').value;
  
  if (!customerName) {
    showMessage('请输入客户名称', 'error');
    return;
  }
  
  if (selectedProducts.length === 0) {
    showMessage('请至少选择一个产品', 'error');
    return;
  }
  
  try {
    const quoteTable = await bitable.base.getTableByName('报价单');
    
    for (const product of selectedProducts) {
      await quoteTable.addRecord({
        fields: {
          '客户名称': customerName,
          '产品名称': product.fields['产品名称'],
          '数量': 1,
          '单价': product.fields['单价'] || 0
        }
      });
    }
    
    showMessage(`✅ 成功！共 ${selectedProducts.length} 个产品`, 'success');
    
    selectedProducts = [];
    document.querySelectorAll('#productList > div').forEach(div => {
      div.style.background = '#f5f5f5';
      div.style.border = 'none';
    });
    document.getElementById('customerName').value = '';
    
  } catch (error) {
    showMessage('失败: ' + error, 'error');
  }
}

function showMessage(text, type) {
  const div = document.getElementById('message');
  div.textContent = text;
  div.style.display = 'block';
  div.style.background = type === 'success' ? '#e8f8f2' : '#ffece8';
  div.style.color = type === 'success' ? '#00b42a' : '#f53f3f';
  
  setTimeout(() => {
    div.style.display = 'none';
  }, 3000);
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init);
} else {
  init();
}