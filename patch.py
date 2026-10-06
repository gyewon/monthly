# -*- coding: utf-8 -*-
import sys

def patch():
    html = open('index.html', encoding='utf-8').read()
    oldHtml = '<!-- Detailed Items Breakdown -->\n          <div class="table-responsive mb-3">'
    newHtml = '''<!-- Detailed Items Breakdown -->
          <div class="d-flex justify-content-between align-items-center mb-2">
            <h4 style="font-size: 1rem; margin: 0; color: var(--text-main);">상세 시뮬레이션 항목</h4>
            <button class="btn btn-outline-primary btn-sm" id="btn-add-sim-item">
              <i class="fa-solid fa-plus"></i> 시뮬레이션 전용 항목 추가
            </button>
          </div>
          <div class="table-responsive mb-3">'''
    if oldHtml in html:
        html = html.replace(oldHtml, newHtml)
        open('index.html', 'w', encoding='utf-8').write(html)
        print('Patched HTML')
    else:
        print('HTML patch target not found')

    js = open('app.js', encoding='utf-8').read()
    oldAllocs = '''    ...(appState.allocations?.dongwook || []).map(i => ({ ...i, pathStr: 'allocations.dongwook' }))\n  ];'''
    newAllocs = '''    ...(appState.allocations?.dongwook || []).map(i => ({ ...i, pathStr: 'allocations.dongwook' })),\n    ...(appState.simulationItems || []).map(i => ({ ...i, pathStr: 'simulationItems', isCustom: true }))\n  ];'''
    if oldAllocs in js:
        js = js.replace(oldAllocs, newAllocs)
        print('Patched Allocs')
    else:
        print('Allocs patch target not found')

    oldFilter = '''.filter(i => appState.simulationCategories.includes(i.category))'''
    newFilter = '''.filter(i => appState.simulationCategories.includes(i.category) || i.isCustom)'''
    if oldFilter in js:
        js = js.replace(oldFilter, newFilter)
        print('Patched Filter')

    oldRow = '''      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${item.name}</td>
        <td style="text-align:center;"><span class="text-muted" style="font-size:0.9em; font-weight:600;">${dayText}</span></td>
        <td><span class="badge ${getAllocCategoryBadgeClass(item.catLabel)}">${item.catLabel}</span></td>
        <td class="cell-amount"><div class="editable-cell" title="클릭하여 기납입액 수정" onclick="editInlineCell(this, '${item.pathStr}', '${item.id}', 'currentBalance', 'number')">${formatKRW(currentBal)}</div></td>
        <td style="text-align:center;"><div class="editable-cell text-muted" title="클릭하여 기준일자 입력 (예: 20240520 또는 2024.05.20)" style="font-size: 0.85rem;" onclick="editInlineCell(this, '${item.pathStr}', '${item.id}', 'currentBalanceDate', 'text')">${displayDate}</div></td>
        <td class="cell-amount">${formatKRW(amt)}</td>
        <td class="cell-amount" style="font-weight:700; color:var(--accent-primary);">${formatKRW(itemExpected)}</td>
      `;'''
    newRow = '''      const tr = document.createElement('tr');
      const nameHtml = item.isCustom ? `<div class="editable-cell" onclick="editInlineCell(this, '${item.pathStr}', '${item.id}', 'name', 'text')">${item.name}</div>` : item.name;
      const amtHtml = item.isCustom ? `<div class="editable-cell" onclick="editInlineCell(this, '${item.pathStr}', '${item.id}', 'amount', 'number')">${formatKRW(amt)}</div>` : formatKRW(amt);
      const catHtml = item.isCustom ? `<div class="editable-cell" onclick="editInlineCell(this, '${item.pathStr}', '${item.id}', 'category', 'text')"><span class="badge ${getAllocCategoryBadgeClass(item.catLabel)}">${item.catLabel}</span></div>` : `<span class="badge ${getAllocCategoryBadgeClass(item.catLabel)}">${item.catLabel}</span>`;
      const delHtml = item.isCustom ? ` <i class="fa-solid fa-times text-danger" style="cursor:pointer;" onclick="deleteSimItem('${item.id}')" title="삭제"></i>` : '';
      tr.innerHTML = `
        <td>${nameHtml}${delHtml}</td>
        <td style="text-align:center;"><span class="text-muted" style="font-size:0.9em; font-weight:600;">${dayText}</span></td>
        <td>${catHtml}</td>
        <td class="cell-amount"><div class="editable-cell" title="클릭하여 기납입액 수정" onclick="editInlineCell(this, '${item.pathStr}', '${item.id}', 'currentBalance', 'number')">${formatKRW(currentBal)}</div></td>
        <td style="text-align:center;"><div class="editable-cell text-muted" title="클릭하여 기준일자 입력 (예: 20240520 또는 2024.05.20)" style="font-size: 0.85rem;" onclick="editInlineCell(this, '${item.pathStr}', '${item.id}', 'currentBalanceDate', 'text')">${displayDate}</div></td>
        <td class="cell-amount">${amtHtml}</td>
        <td class="cell-amount" style="font-weight:700; color:var(--accent-primary);">${formatKRW(itemExpected)}</td>
      `;'''
    
    if oldRow in js:
        js = js.replace(oldRow, newRow)
        print('Patched Row')
    else:
        print('Row patch target not found')
    
    oldFuncEnd = '''  }

  function renderFixedExpenses() {'''
    newFuncEnd = '''  }

  window.deleteSimItem = function(id) {
    if(confirm('이 시뮬레이션 항목을 삭제하시겠습니까?')) {
      appState.simulationItems = (appState.simulationItems || []).filter(i => i.id !== id);
      saveState();
    }
  };

  document.addEventListener('click', (e) => {
    if (e.target.closest('#btn-add-sim-item')) {
      if (!appState.simulationItems) appState.simulationItems = [];
      const itemName = prompt('시뮬레이션 전용 항목명을 입력하세요 (예: 연말 보너스)');
      if (!itemName) return;
      const amountStr = prompt('월 평균 납입/발생 예상액을 숫자로 입력하세요 (예: 500000)');
      if (!amountStr) return;
      
      appState.simulationItems.push({
        id: Date.now().toString(),
        name: itemName,
        amount: parseInt(amountStr, 10) || 0,
        category: '기타',
        currentBalance: 0,
        currentBalanceDate: '',
        day: '-'
      });
      saveState();
    }
  });

  function renderFixedExpenses() {'''
    if oldFuncEnd in js:
        js = js.replace(oldFuncEnd, newFuncEnd)
        print('Patched Func End')
    
    open('app.js', 'w', encoding='utf-8').write(js)

patch()
