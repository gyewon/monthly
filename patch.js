const fs = require('fs');
let html = fs.readFileSync('index.html', 'utf8');
let oldHtml = '<!-- Detailed Items Breakdown -->\n          <div class=\"table-responsive mb-3\">';
let newHtml = '<!-- Detailed Items Breakdown -->\n          <div class=\"d-flex justify-content-between align-items-center mb-2\">\n            <h4 style=\"font-size: 1rem; margin: 0; color: var(--text-main);\">상세 시뮬레이션 항목</h4>\n            <button class=\"btn btn-outline-primary btn-sm\" id=\"btn-add-sim-item\">\n              <i class=\"fa-solid fa-plus\"></i> 시뮬레이션 항목 추가\n            </button>\n          </div>\n          <div class=\"table-responsive mb-3\">';
html = html.replace(oldHtml, newHtml);
fs.writeFileSync('index.html', html);

let js = fs.readFileSync('app.js', 'utf8');

// Patch 1: appState.simulationItems
let oldAllocs = '    ...(appState.allocations?.dongwook || []).map(i => ({ ...i, pathStr: \\'allocations.dongwook\\' }))\n  ];';
let newAllocs = '    ...(appState.allocations?.dongwook || []).map(i => ({ ...i, pathStr: \\'allocations.dongwook\\' })),\n    ...(appState.simulationItems || []).map(i => ({ ...i, pathStr: \\'simulationItems\\', isCustom: true }))\n  ];';
js = js.replace(oldAllocs, newAllocs);

// Patch 2: filter
let oldFilter = '.filter(i => appState.simulationCategories.includes(i.category))';
let newFilter = '.filter(i => appState.simulationCategories.includes(i.category) || i.isCustom)';
js = js.replace(oldFilter, newFilter);

// Patch 3: editable cells for custom items
let oldRow = '      const tr = document.createElement(\\'tr\\');\n      tr.innerHTML = \n        <td></td>\n        <td style=\"text-align:center;\"><span class=\"text-muted\" style=\"font-size:0.9em; font-weight:600;\"></span></td>\n        <td><span class=\"badge \"></span></td>\n        <td class=\"cell-amount\"><div class=\"editable-cell\" title=\"클릭하여 기납입액 수정\" onclick=\"editInlineCell(this, \\'\\', \\'\\', \\'currentBalance\\', \\'number\\')\"></div></td>\n        <td style=\"text-align:center;\"><div class=\"editable-cell text-muted\" title=\"클릭하여 기준일자 입력 (예: 20240520 또는 2024.05.20)\" style=\"font-size: 0.85rem;\" onclick=\"editInlineCell(this, \\'\\', \\'\\', \\'currentBalanceDate\\', \\'text\\')\"></div></td>\n        <td class=\"cell-amount\"></td>\n        <td class=\"cell-amount\" style=\"font-weight:700; color:var(--accent-primary);\"></td>\n      ;';

let newRow = '      const tr = document.createElement(\\'tr\\');\n      const nameHtml = item.isCustom ? <div class=\"editable-cell\" onclick=\"editInlineCell(this, \\'\\', \\'\\', \\'name\\', \\'text\\')\"></div> : item.name;\n      const amtHtml = item.isCustom ? <div class=\"editable-cell\" onclick=\"editInlineCell(this, \\'\\', \\'\\', \\'amount\\', \\'number\\')\"></div> : formatKRW(amt);\n      const catHtml = item.isCustom ? <div class=\"editable-cell\" onclick=\"editInlineCell(this, \\'\\', \\'\\', \\'category\\', \\'text\\')\"><span class=\"badge \"></span></div> : <span class=\"badge \"></span>;\n      const delHtml = item.isCustom ?  <i class=\"fa-solid fa-times text-danger\" style=\"cursor:pointer;\" onclick=\"deleteSimItem(\\'\\')\" title=\"삭제\"></i> : \\'\\';\n      tr.innerHTML = \n        <td></td>\n        <td style=\"text-align:center;\"><span class=\"text-muted\" style=\"font-size:0.9em; font-weight:600;\"></span></td>\n        <td></td>\n        <td class=\"cell-amount\"><div class=\"editable-cell\" title=\"클릭하여 기납입액 수정\" onclick=\"editInlineCell(this, \\'\\', \\'\\', \\'currentBalance\\', \\'number\\')\"></div></td>\n        <td style=\"text-align:center;\"><div class=\"editable-cell text-muted\" title=\"클릭하여 기준일자 입력 (예: 20240520 또는 2024.05.20)\" style=\"font-size: 0.85rem;\" onclick=\"editInlineCell(this, \\'\\', \\'\\', \\'currentBalanceDate\\', \\'text\\')\"></div></td>\n        <td class=\"cell-amount\"></td>\n        <td class=\"cell-amount\" style=\"font-weight:700; color:var(--accent-primary);\"></td>\n      ;';

js = js.replace(oldRow, newRow);

fs.writeFileSync('app.js', js);
console.log('Patched OK');
