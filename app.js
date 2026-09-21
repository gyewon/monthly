/**
 * Dongwook & Gyewon Monthly Budget Manager
 * Application Logic Engine JavaScript
 */

// Initial Excel Default Data
const DEFAULT_EXCEL_DATA = {
  currentMonth: '2026-03',
  incomes: [
    { id: 'inc_g1', person: 'gyewon', title: '기본급 (급여)', amount: 4700000, note: '계원 월급' },
    { id: 'inc_d1', person: 'dongwook', title: '기본급 (급여)', amount: 3750000, note: '동욱 월급' }
  ],
  allocations: {
    gyewon: [
      { id: 'ag1', name: '생활비', amount: 1024768, bank: '신한', schedule: '22일 자동이체', category: '생활비' },
      { id: 'ag2', name: '주택청약', amount: 100000, bank: '국민', schedule: '22일 자동이체', category: '저축/적금' },
      { id: 'ag3', name: '경조사', amount: 150000, bank: '신한', schedule: '22일/오빠자동이체', category: '비상금/경조사' },
      { id: 'ag4', name: '오빠 IRP', amount: 200000, bank: '신한', schedule: '22일/오빠자동이체', category: '투자/연금' },
      { id: 'ag5', name: '원이 IRP', amount: 200000, bank: '신한', schedule: '22일 자동이체', category: '투자/연금' },
      { id: 'ag6', name: '오빠 ISA', amount: 200000, bank: '신한', schedule: '', category: '투자/연금' },
      { id: 'ag7', name: '원이 ISA', amount: 400000, bank: '키움', schedule: '', category: '투자/연금' },
      { id: 'ag8', name: '케이뱅크', amount: 300000, bank: '코인', schedule: '22일 자동이체', category: '투자/연금' },
      { id: 'ag9', name: '대출금', amount: 1875232, bank: '고정지출', schedule: '22일 자동이체', category: '대출금' },
      { id: 'ag10', name: '자율적금', amount: 100000, bank: '', schedule: '22일/오빠자동이체', category: '저축/적금' },
      { id: 'ag11', name: 'IRP', amount: 50000, bank: '신한투자증권', schedule: '15일회사자동이체(별도)', category: '투자/연금' }
    ],
    dongwook: [
      { id: 'ad1', name: '생활비', amount: 1750000, bank: '', schedule: '16일/자동이체', category: '생활비' },
      { id: 'ad2', name: '적금', amount: 1000000, bank: '어머니', schedule: '16일/자동이체', category: '저축/적금' },
      { id: 'ad3', name: '미래적금', amount: 500000, bank: '신한', schedule: '16일/자동이체', category: '저축/적금' },
      { id: 'ad4', name: '오빠 ISA', amount: 300000, bank: '신한', schedule: '16일/자동이체', category: '투자/연금' },
      { id: 'ad5', name: '에어비앤비', amount: 100000, bank: '저축', schedule: '', category: '저축/적금' },
      { id: 'ad6', name: '제로스토어', amount: 100000, bank: '저축', schedule: '', category: '저축/적금' }
    ]
  },
  fixedExpenses: [
    { id: 'fe1', name: '각자용돈', amount: 760000, day: '2일', method: '현금', category: '용돈', note: '토스고정지출 자동이체', actualMarch: 760000, isPaid: true },
    { id: 'fe2', name: '곗돈', amount: 240000, day: '2일/26일', method: '현금', category: '곗돈', note: '토스고정지출 자동이체 (10k / 230k)', actualMarch: 240000, isPaid: true },
    { id: 'fe3', name: '네이버멤버십', amount: 2450, day: '7일', method: '하나카드', category: '구독', note: '새마을금고 / 토스고정지출 자동이체', actualMarch: 2450, isPaid: true },
    { id: 'fe4', name: '힘이보험료(실비)', amount: 50000, day: '10일', method: '현금', category: '보험', note: '엄마가내주고적금넣음 / 토스고정지출 자동이체', actualMarch: 50000, isPaid: true },
    { id: 'fe5', name: '동욱보험비', amount: 177931, day: '10일', method: '오빠카드(삼성)', category: '보험', note: '토스고정지출 자동이체', actualMarch: 177931, isPaid: true },
    { id: 'fe6', name: '쿠팡멤버십', amount: 7890, day: '10일', method: '오빠카드', category: '구독', note: '토스고정지출 자동이체', actualMarch: 7890, isPaid: true },
    { id: 'fe7', name: '동욱폰요금', amount: 32600, day: '26일', method: '오빠카드(삼성)', category: '핸드폰요금', note: '', actualMarch: 32600, isPaid: false },
    { id: 'fe8', name: '원이보험비(DB)', amount: 138340, day: '16일', method: '하나카드', category: '보험', note: '보험아줌마 수기 결제', actualMarch: 138340, isPaid: true },
    { id: 'fe9', name: '원이보험비(삼성화재)', amount: 55000, day: '수기', method: '우리카드', category: '보험', note: '', actualMarch: 55000, isPaid: false },
    { id: 'fe10', name: '관리비', amount: 160000, day: '수기', method: '카드', category: '주거', note: '가스비, 수도, 전기', actualMarch: 154437, isPaid: false },
    { id: 'fe11', name: '계원폰요금', amount: 60000, day: '수기', method: '카드', category: '핸드폰요금', note: '8월까지 확인 필요 (첫 사용시 2.5만원 할인)', actualMarch: 29950, isPaid: false },
    { id: 'fe12', name: '하이패스', amount: 50000, day: '수기', method: '신한카드', category: '교통', note: '', actualMarch: 50000, isPaid: false },
    { id: 'fe13', name: '베인', amount: 1990, day: '수기', method: '카드', category: '구독', note: '', actualMarch: 1990, isPaid: false }
  ],
  monthlyHistory: [
    { month: '2025-12', totalIncome: 8450000, totalExpenses: 4100000, fixedExpenses: 1774768, remaining: 4350000, savings: 2850000 },
    { month: '2026-01', totalIncome: 8450000, totalExpenses: 4050000, fixedExpenses: 1750000, remaining: 4400000, savings: 2850000 },
    { month: '2026-02', totalIncome: 8450000, totalExpenses: 4080000, fixedExpenses: 1765000, remaining: 4370000, savings: 2850000 },
    { month: '2026-03', totalIncome: 8450000, totalExpenses: 4088506, fixedExpenses: 1698448, remaining: 4361494, savings: 2850000 }
  ]
};

// Application State
let appState = loadState();

// Chart Instances
let trendChartInstance = null;
let categoryChartInstance = null;
let paymentChartInstance = null;

// DOM Loaded Initialization
document.addEventListener('DOMContentLoaded', () => {
  initApp();
});

function initApp() {
  setupEventListeners();
  renderAll();
}

// State Management & LocalStorage
function loadState() {
  const saved = localStorage.getItem('dongwook_gyewon_budget_app_v2');
  if (saved) {
    try {
      return JSON.parse(saved);
    } catch (e) {
      console.error('Failed to parse state:', e);
    }
  }
  return JSON.parse(JSON.stringify(DEFAULT_EXCEL_DATA));
}

function saveState() {
  localStorage.setItem('dongwook_gyewon_budget_app_v2', JSON.stringify(appState));
  renderAll();
}

// Event Listeners Registration
function setupEventListeners() {
  // Navigation Tabs
  document.querySelectorAll('.nav-item').forEach(btn => {
    btn.addEventListener('click', (e) => {
      document.querySelectorAll('.nav-item').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-page').forEach(p => p.classList.remove('active'));

      const targetTab = btn.getAttribute('data-tab');
      btn.classList.add('active');
      document.getElementById(`tab-${targetTab}`).classList.add('active');

      // Update Page Title
      const titles = {
        'dashboard': '재정 대시보드',
        'income-alloc': '수입 & 배분 관리',
        'fixed-expenses': '고정지출 상세 관리',
        'history': '월별 기록 & 추이'
      };
      document.getElementById('page-title').innerText = titles[targetTab] || '월별 재정 관리';

      // Re-render charts when entering dashboard tab
      if (targetTab === 'dashboard') {
        renderCharts();
      }
    });
  });

  // Global Month Selector
  const monthPicker = document.getElementById('global-month-picker');
  if (monthPicker) {
    monthPicker.value = appState.currentMonth || '2026-03';
    monthPicker.addEventListener('change', (e) => {
      appState.currentMonth = e.target.value;
      saveState();
      showToast(`${e.target.value} 월로 변경되었습니다.`);
    });
  }

  // Theme Toggle
  const themeToggle = document.getElementById('theme-toggle');
  themeToggle.addEventListener('change', (e) => {
    document.documentElement.setAttribute('data-theme', e.target.checked ? 'dark' : 'light');
  });

  // Action Buttons
  document.getElementById('btn-copy-summary').addEventListener('click', copySummaryToClipboard);
  document.getElementById('btn-export-json').addEventListener('click', exportBackupJSON);
  
  const importTrigger = document.getElementById('btn-import-trigger');
  const importInput = document.getElementById('import-json-input');
  importTrigger.addEventListener('click', () => importInput.click());
  importInput.addEventListener('change', importBackupJSON);

  // Income Addition Modal Trigger
  document.getElementById('btn-add-income-modal').addEventListener('click', () => openAddIncomeModal());

  // Allocations Addition Triggers
  document.getElementById('btn-add-alloc-gyewon').addEventListener('click', () => openAddAllocModal('gyewon'));
  document.getElementById('btn-add-alloc-dongwook').addEventListener('click', () => openAddAllocModal('dongwook'));

  // Fixed Expense Addition Trigger
  document.getElementById('btn-add-fixed-expense').addEventListener('click', () => openAddFixedExpenseModal());

  // Filters for Fixed Expenses
  document.getElementById('filter-method').addEventListener('change', renderFixedExpensesTable);
  document.getElementById('filter-category').addEventListener('change', renderFixedExpensesTable);
  document.getElementById('filter-paid').addEventListener('change', renderFixedExpensesTable);
  document.getElementById('search-fixed-expense').addEventListener('input', renderFixedExpensesTable);

  // Modal Close
  document.getElementById('modal-close-btn').addEventListener('click', closeModal);
  document.getElementById('modal-cancel-btn').addEventListener('click', closeModal);
}

// Master Render Function
function renderAll() {
  const calcs = calculateTotals();

  // Render KPI Banner
  document.getElementById('kpi-total-income').innerText = formatKRW(calcs.totalIncome);
  document.getElementById('kpi-total-expenses').innerText = formatKRW(calcs.totalExpenses);
  
  const remainElem = document.getElementById('kpi-remaining-balance');
  remainElem.innerText = formatKRW(calcs.remainingBalance);
  
  if (calcs.remainingBalance < 0) {
    remainElem.classList.add('text-danger');
  } else {
    remainElem.classList.remove('text-danger');
  }

  document.getElementById('kpi-remain-gyewon').innerText = formatKRW(calcs.remainGyewon);
  document.getElementById('kpi-remain-dongwook').innerText = formatKRW(calcs.remainDongwook);

  document.getElementById('kpi-income-gyewon').innerText = formatCompactKRW(calcs.incomeGyewon);
  document.getElementById('kpi-income-dongwook').innerText = formatCompactKRW(calcs.incomeDongwook);
  document.getElementById('kpi-fixed-total').innerText = formatKRW(calcs.fixedExpenseTotal);
  document.getElementById('kpi-total-savings').innerText = formatKRW(calcs.totalSavings);
  document.getElementById('kpi-savings-rate').innerText = `${calcs.savingsRate}%`;

  // Render Tables
  renderIncomeTables();
  renderAllocationTables();
  renderFixedExpensesTable();
  renderTimeline();
  renderPaymentMethodSummary();
  renderHistoryTable();

  // Render Charts
  renderCharts();
}

// Calculation Helper
function calculateTotals() {
  // Incomes
  let incomeGyewon = 0;
  let incomeDongwook = 0;

  appState.incomes.forEach(inc => {
    if (inc.person === 'gyewon') incomeGyewon += Number(inc.amount) || 0;
    else incomeDongwook += Number(inc.amount) || 0;
  });

  const totalIncome = incomeGyewon + incomeDongwook;

  // Allocations
  let allocGyewonSum = appState.allocations.gyewon.reduce((acc, curr) => acc + (Number(curr.amount) || 0), 0);
  let allocDongwookSum = appState.allocations.dongwook.reduce((acc, curr) => acc + (Number(curr.amount) || 0), 0);
  const totalAllocations = allocGyewonSum + allocDongwookSum;

  const remainGyewon = incomeGyewon - allocGyewonSum;
  const remainDongwook = incomeDongwook - allocDongwookSum;

  // Fixed Expenses
  const fixedBaseTotal = appState.fixedExpenses.reduce((acc, curr) => acc + (Number(curr.amount) || 0), 0);
  const fixedActualTotal = appState.fixedExpenses.reduce((acc, curr) => acc + (Number(curr.actualMarch) || 0), 0);

  // Total Expenses
  const totalExpenses = totalAllocations;

  // Savings & Investments (Categories: 저축/적금, 투자/연금)
  let totalSavings = 0;
  [...appState.allocations.gyewon, ...appState.allocations.dongwook].forEach(item => {
    if (item.category === '저축/적금' || item.category === '투자/연금') {
      totalSavings += Number(item.amount) || 0;
    }
  });

  const remainingBalance = totalIncome - totalExpenses;
  const savingsRate = totalIncome > 0 ? ((totalSavings / totalIncome) * 100).toFixed(1) : 0;

  return {
    incomeGyewon,
    incomeDongwook,
    totalIncome,
    allocGyewonSum,
    allocDongwookSum,
    remainGyewon,
    remainDongwook,
    totalAllocations,
    fixedExpenseTotal: fixedBaseTotal,
    fixedExpenseActualTotal: fixedActualTotal,
    totalExpenses,
    totalSavings,
    remainingBalance,
    savingsRate
  };
}

// Render Incomes
function renderIncomeTables() {
  const gyewonTable = document.querySelector('#table-income-gyewon tbody');
  const dongwookTable = document.querySelector('#table-income-dongwook tbody');

  gyewonTable.innerHTML = '';
  dongwookTable.innerHTML = '';

  let sumG = 0, sumD = 0;

  appState.incomes.forEach(inc => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><div class="editable-cell" title="클릭하여 수정" onclick="editInlineCell(this, 'incomes', '${inc.id}', 'title')">${inc.title}</div></td>
      <td><div class="editable-cell cell-amount text-accent" title="클릭하여 월급 수정" onclick="editInlineCell(this, 'incomes', '${inc.id}', 'amount', 'number')">${formatKRW(inc.amount)}</div></td>
      <td><div class="editable-cell" title="클릭하여 수정" onclick="editInlineCell(this, 'incomes', '${inc.id}', 'note')">${inc.note || '-'}</div></td>
      <td style="white-space:nowrap; text-align:right;">
        <button class="btn btn-outline-primary btn-sm" title="수입/월급 수정" onclick="openEditIncomeModal('${inc.id}')"><i class="fa-solid fa-pen-to-square"></i> 수정</button>
        <button class="btn btn-outline-danger btn-sm" title="삭제" onclick="deleteItem('incomes', '${inc.id}')"><i class="fa-solid fa-trash"></i></button>
      </td>
    `;

    if (inc.person === 'gyewon') {
      sumG += Number(inc.amount) || 0;
      gyewonTable.appendChild(tr);
    } else {
      sumD += Number(inc.amount) || 0;
      dongwookTable.appendChild(tr);
    }
  });

  document.getElementById('gyewon-income-sum').innerText = formatKRW(sumG);
  document.getElementById('dongwook-income-sum').innerText = formatKRW(sumD);
}

// Render Allocations Tables
function renderAllocationTables() {
  const calcs = calculateTotals();

  // Update Gyewon Summary Bar
  document.getElementById('g-summary-income').innerText = formatKRW(calcs.incomeGyewon);
  document.getElementById('g-summary-alloc').innerText = formatKRW(calcs.allocGyewonSum);
  const gRemain = calcs.incomeGyewon - calcs.allocGyewonSum;
  const gRemainElem = document.getElementById('g-summary-remain');
  gRemainElem.innerText = formatKRW(gRemain);
  if (gRemain < 0) gRemainElem.classList.add('negative'); else gRemainElem.classList.remove('negative');

  // Update Dongwook Summary Bar
  document.getElementById('d-summary-income').innerText = formatKRW(calcs.incomeDongwook);
  document.getElementById('d-summary-alloc').innerText = formatKRW(calcs.allocDongwookSum);
  const dRemain = calcs.incomeDongwook - calcs.allocDongwookSum;
  const dRemainElem = document.getElementById('d-summary-remain');
  dRemainElem.innerText = formatKRW(dRemain);
  if (dRemain < 0) dRemainElem.classList.add('negative'); else dRemainElem.classList.remove('negative');

  // Populate Gyewon Allocation Rows
  const tbodyG = document.querySelector('#table-alloc-gyewon tbody');
  tbodyG.innerHTML = '';
  appState.allocations.gyewon.forEach(item => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><div class="editable-cell" onclick="editInlineCell(this, 'allocations.gyewon', '${item.id}', 'name')">${item.name}</div></td>
      <td><div class="editable-cell cell-amount" onclick="editInlineCell(this, 'allocations.gyewon', '${item.id}', 'amount', 'number')">${formatKRW(item.amount)}</div></td>
      <td><div class="editable-cell" onclick="editInlineCell(this, 'allocations.gyewon', '${item.id}', 'bank')">${item.bank || '-'}</div></td>
      <td><div class="editable-cell" onclick="editInlineCell(this, 'allocations.gyewon', '${item.id}', 'schedule')">${item.schedule || '-'}</div></td>
      <td><div class="editable-cell" onclick="editInlineCell(this, 'allocations.gyewon', '${item.id}', 'category')">${item.category || '-'}</div></td>
      <td><button class="btn btn-outline-danger btn-sm" onclick="deleteItem('allocations.gyewon', '${item.id}')"><i class="fa-solid fa-xmark"></i></button></td>
    `;
    tbodyG.appendChild(tr);
  });

  // Populate Dongwook Allocation Rows
  const tbodyD = document.querySelector('#table-alloc-dongwook tbody');
  tbodyD.innerHTML = '';
  appState.allocations.dongwook.forEach(item => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><div class="editable-cell" onclick="editInlineCell(this, 'allocations.dongwook', '${item.id}', 'name')">${item.name}</div></td>
      <td><div class="editable-cell cell-amount" onclick="editInlineCell(this, 'allocations.dongwook', '${item.id}', 'amount', 'number')">${formatKRW(item.amount)}</div></td>
      <td><div class="editable-cell" onclick="editInlineCell(this, 'allocations.dongwook', '${item.id}', 'bank')">${item.bank || '-'}</div></td>
      <td><div class="editable-cell" onclick="editInlineCell(this, 'allocations.dongwook', '${item.id}', 'schedule')">${item.schedule || '-'}</div></td>
      <td><div class="editable-cell" onclick="editInlineCell(this, 'allocations.dongwook', '${item.id}', 'category')">${item.category || '-'}</div></td>
      <td><button class="btn btn-outline-danger btn-sm" onclick="deleteItem('allocations.dongwook', '${item.id}')"><i class="fa-solid fa-xmark"></i></button></td>
    `;
    tbodyD.appendChild(tr);
  });
}

// Render Fixed Expenses Table
function renderFixedExpensesTable() {
  const tbody = document.querySelector('#table-fixed-expenses tbody');
  tbody.innerHTML = '';

  const filterMethod = document.getElementById('filter-method').value;
  const filterCategory = document.getElementById('filter-category').value;
  const filterPaid = document.getElementById('filter-paid').value;
  const searchKeyword = document.getElementById('search-fixed-expense').value.toLowerCase().trim();

  let filtered = appState.fixedExpenses.filter(item => {
    if (filterMethod !== 'ALL') {
      if (filterMethod === '오빠카드' && !item.method.includes('오빠카드')) return false;
      else if (filterMethod !== '오빠카드' && item.method !== filterMethod) return false;
    }
    if (filterCategory !== 'ALL' && item.category !== filterCategory) return false;
    if (filterPaid === 'PAID' && !item.isPaid) return false;
    if (filterPaid === 'UNPAID' && item.isPaid) return false;
    if (searchKeyword) {
      const matchName = item.name.toLowerCase().includes(searchKeyword);
      const matchNote = (item.note || '').toLowerCase().includes(searchKeyword);
      if (!matchName && !matchNote) return false;
    }
    return true;
  });

  let baseSum = 0;
  let actualSum = 0;
  let paidCount = 0;

  filtered.forEach(item => {
    baseSum += Number(item.amount) || 0;
    actualSum += Number(item.actualMarch) || 0;
    if (item.isPaid) paidCount++;

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="text-align:center;">
        <input type="checkbox" ${item.isPaid ? 'checked' : ''} onchange="toggleFixedPaid('${item.id}', this.checked)" style="transform: scale(1.2); cursor:pointer;">
      </td>
      <td><div class="editable-cell" onclick="editInlineCell(this, 'fixedExpenses', '${item.id}', 'name')">${item.name}</div></td>
      <td><div class="editable-cell cell-amount" onclick="editInlineCell(this, 'fixedExpenses', '${item.id}', 'amount', 'number')">${formatKRW(item.amount)}</div></td>
      <td><div class="editable-cell cell-amount text-accent" onclick="editInlineCell(this, 'fixedExpenses', '${item.id}', 'actualMarch', 'number')">${formatKRW(item.actualMarch)}</div></td>
      <td><div class="editable-cell" onclick="editInlineCell(this, 'fixedExpenses', '${item.id}', 'day')">${item.day || '-'}</div></td>
      <td><div class="editable-cell" onclick="editInlineCell(this, 'fixedExpenses', '${item.id}', 'method')">${item.method || '-'}</div></td>
      <td><div class="editable-cell" onclick="editInlineCell(this, 'fixedExpenses', '${item.id}', 'category')">${item.category || '-'}</div></td>
      <td><div class="editable-cell" onclick="editInlineCell(this, 'fixedExpenses', '${item.id}', 'note')">${item.note || '-'}</div></td>
      <td><button class="btn btn-outline-danger btn-sm" onclick="deleteItem('fixedExpenses', '${item.id}')"><i class="fa-solid fa-trash"></i></button></td>
    `;
    tbody.appendChild(tr);
  });

  document.getElementById('fixed-stat-base').innerText = formatKRW(baseSum);
  document.getElementById('fixed-stat-actual').innerText = formatKRW(actualSum);
  const totalCount = filtered.length;
  const pct = totalCount > 0 ? Math.round((paidCount / totalCount) * 100) : 0;
  document.getElementById('fixed-stat-completion').innerText = `${paidCount} / ${totalCount} (${pct}%)`;
}

// Render Timeline / Schedule
function renderTimeline() {
  const container = document.getElementById('schedule-timeline');
  container.innerHTML = '';

  const groups = {};
  appState.fixedExpenses.forEach(item => {
    const dayKey = item.day || '수기/미지정';
    if (!groups[dayKey]) groups[dayKey] = [];
    groups[dayKey].push(item);
  });

  // Sort keys
  const sortedKeys = Object.keys(groups).sort((a, b) => {
    const numA = parseInt(a) || 99;
    const numB = parseInt(b) || 99;
    return numA - numB;
  });

  sortedKeys.forEach(day => {
    const items = groups[day];
    const totalDayAmount = items.reduce((acc, c) => acc + (Number(c.actualMarch || c.amount) || 0), 0);

    const card = document.createElement('div');
    card.className = 'timeline-card';
    card.innerHTML = `
      <div class="timeline-card-header">
        <span class="timeline-date"><i class="fa-regular fa-clock"></i> ${day} 이체</span>
        <span class="timeline-count">${items.length}건 / ${formatKRW(totalDayAmount)}</span>
      </div>
      <ul class="timeline-item-list">
        ${items.map(i => `
          <li>
            <span>${i.name} (${i.method})</span>
            <strong>${formatKRW(i.actualMarch || i.amount)}</strong>
          </li>
        `).join('')}
      </ul>
    `;
    container.appendChild(card);
  });
}

// Render Payment Method Summary Cards
function renderPaymentMethodSummary() {
  const container = document.getElementById('payment-method-summary');
  container.innerHTML = '';

  const pmMap = {};
  appState.fixedExpenses.forEach(item => {
    const method = item.method || '기타';
    pmMap[method] = (pmMap[method] || 0) + (Number(item.actualMarch || item.amount) || 0);
  });

  Object.entries(pmMap).forEach(([pm, amt]) => {
    const card = document.createElement('div');
    card.className = 'pm-card';
    card.innerHTML = `
      <div class="pm-name">${pm}</div>
      <div class="pm-amount">${formatKRW(amt)}</div>
    `;
    container.appendChild(card);
  });
}

// Render Monthly History Table
function renderHistoryTable() {
  const tbody = document.querySelector('#table-monthly-history tbody');
  tbody.innerHTML = '';

  (appState.monthlyHistory || []).forEach((h, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><strong>${h.month}</strong></td>
      <td class="text-accent">${formatKRW(h.totalIncome)}</td>
      <td>${formatKRW(h.totalExpenses)}</td>
      <td>${formatKRW(h.fixedExpenses)}</td>
      <td>${formatKRW(h.savings)}</td>
      <td><span class="badge-balance">${formatKRW(h.remaining)}</span></td>
      <td>
        <button class="btn btn-outline-danger btn-sm" onclick="deleteHistoryRow(${idx})">
          <i class="fa-solid fa-trash"></i>
        </button>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

// Direct Inline Cell Editor
function editInlineCell(element, pathStr, itemId, fieldName, type = 'text') {
  if (element.querySelector('input')) return; // Already editing

  let currVal = getFieldValue(pathStr, itemId, fieldName);
  const originalVal = currVal;

  const input = document.createElement('input');
  input.type = type === 'number' ? 'number' : 'text';
  input.className = 'cell-input';
  input.value = currVal;

  element.innerHTML = '';
  element.appendChild(input);
  input.focus();

  const finishEdit = () => {
    let newVal = input.value.trim();
    if (type === 'number') newVal = Number(newVal) || 0;
    
    setFieldValue(pathStr, itemId, fieldName, newVal);
    saveState();
  };

  input.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      finishEdit();
    } else if (e.key === 'Escape') {
      element.innerText = type === 'number' ? formatKRW(originalVal) : (originalVal || '-');
    }
  });

  input.addEventListener('blur', finishEdit);
}

// Helpers for nested state access
function getFieldValue(pathStr, itemId, fieldName) {
  let list = resolvePath(pathStr);
  let target = list.find(i => i.id === itemId);
  return target ? target[fieldName] : '';
}

function setFieldValue(pathStr, itemId, fieldName, value) {
  let list = resolvePath(pathStr);
  let target = list.find(i => i.id === itemId);
  if (target) {
    target[fieldName] = value;
  }
}

function resolvePath(pathStr) {
  const parts = pathStr.split('.');
  let curr = appState;
  parts.forEach(p => {
    curr = curr[p];
  });
  return curr;
}

function deleteItem(pathStr, itemId) {
  if (confirm('이 항목을 삭제하시겠습니까?')) {
    let list = resolvePath(pathStr);
    const idx = list.findIndex(i => i.id === itemId);
    if (idx !== -1) {
      list.splice(idx, 1);
      saveState();
      showToast('항목이 삭제되었습니다.');
    }
  }
}

function toggleFixedPaid(itemId, isPaid) {
  const item = appState.fixedExpenses.find(i => i.id === itemId);
  if (item) {
    item.isPaid = isPaid;
    saveState();
  }
}

// Modal Handlers
let modalCurrentCallback = null;

function closeModal() {
  document.getElementById('item-modal').classList.remove('active');
}

function openAddIncomeModal() {
  const body = document.getElementById('modal-body');
  document.getElementById('modal-title').innerText = '수입 항목 추가';

  body.innerHTML = `
    <div class="form-group">
      <label>수입 대상자</label>
      <select id="modal-inc-person" class="form-select">
        <option value="gyewon">계원</option>
        <option value="dongwook">동욱</option>
      </select>
    </div>
    <div class="form-group">
      <label>수입 항목명 (예: 기본급, 보너스, 부수입)</label>
      <input type="text" id="modal-inc-title" class="form-control" placeholder="항목명 입력">
    </div>
    <div class="form-group">
      <label>금액 (원)</label>
      <input type="number" id="modal-inc-amount" class="form-control" placeholder="0">
    </div>
    <div class="form-group">
      <label>비고</label>
      <input type="text" id="modal-inc-note" class="form-control" placeholder="메모">
    </div>
  `;

  document.getElementById('item-modal').classList.add('active');

  document.getElementById('modal-save-btn').onclick = () => {
    const person = document.getElementById('modal-inc-person').value;
    const title = document.getElementById('modal-inc-title').value.trim();
    const amount = Number(document.getElementById('modal-inc-amount').value) || 0;
    const note = document.getElementById('modal-inc-note').value.trim();

    if (!title) { alert('수입 항목명을 입력하세요.'); return; }

    appState.incomes.push({
      id: 'inc_' + Date.now(),
      person,
      title,
      amount,
      note
    });

    closeModal();
    saveState();
    showToast('새 수입 항목이 추가되었습니다.');
  };
}

function openEditIncomeModal(incId) {
  const inc = appState.incomes.find(i => i.id === incId);
  if (!inc) return;

  const body = document.getElementById('modal-body');
  const personName = inc.person === 'gyewon' ? '계원' : '동욱';
  document.getElementById('modal-title').innerText = `${personName} 수입/월급 수정`;

  body.innerHTML = `
    <div class="form-group">
      <label>수입 대상자</label>
      <select id="modal-inc-person" class="form-select">
        <option value="gyewon" ${inc.person === 'gyewon' ? 'selected' : ''}>계원</option>
        <option value="dongwook" ${inc.person === 'dongwook' ? 'selected' : ''}>동욱</option>
      </select>
    </div>
    <div class="form-group">
      <label>수입 항목명 (예: 기본급, 보너스, 부수입)</label>
      <input type="text" id="modal-inc-title" class="form-control" value="${inc.title}">
    </div>
    <div class="form-group">
      <label>수입 / 월급 금액 (원) - 월급 인상/변동 시 수정</label>
      <input type="number" id="modal-inc-amount" class="form-control" value="${inc.amount}">
    </div>
    <div class="form-group">
      <label>비고</label>
      <input type="text" id="modal-inc-note" class="form-control" value="${inc.note || ''}">
    </div>
  `;

  document.getElementById('item-modal').classList.add('active');

  document.getElementById('modal-save-btn').onclick = () => {
    inc.person = document.getElementById('modal-inc-person').value;
    inc.title = document.getElementById('modal-inc-title').value.trim();
    inc.amount = Number(document.getElementById('modal-inc-amount').value) || 0;
    inc.note = document.getElementById('modal-inc-note').value.trim();

    closeModal();
    saveState();
    showToast('수입/월급 금액이 수정되었습니다.');
  };
}

function openAddAllocModal(personKey) {
  const body = document.getElementById('modal-body');
  const personName = personKey === 'gyewon' ? '계원' : '동욱';
  document.getElementById('modal-title').innerText = `${personName} 배분 항목 추가`;

  body.innerHTML = `
    <div class="form-group">
      <label>배분 항목명</label>
      <input type="text" id="modal-alloc-name" class="form-control" placeholder="예: 적금, 생활비, ISA">
    </div>
    <div class="form-group">
      <label>금액 (원)</label>
      <input type="number" id="modal-alloc-amount" class="form-control" placeholder="0">
    </div>
    <div class="form-group">
      <label>은행 / 기관명</label>
      <input type="text" id="modal-alloc-bank" class="form-control" placeholder="예: 신한, 국민, 키움">
    </div>
    <div class="form-group">
      <label>이체일 / 스케줄</label>
      <input type="text" id="modal-alloc-schedule" class="form-control" placeholder="예: 22일 자동이체">
    </div>
    <div class="form-group">
      <label>카테고리</label>
      <select id="modal-alloc-category" class="form-select">
        <option value="생활비">생활비</option>
        <option value="저축/적금">저축/적금</option>
        <option value="투자/연금">투자/연금</option>
        <option value="대출금">대출금</option>
        <option value="비상금/경조사">비상금/경조사</option>
      </select>
    </div>
  `;

  document.getElementById('item-modal').classList.add('active');

  document.getElementById('modal-save-btn').onclick = () => {
    const name = document.getElementById('modal-alloc-name').value.trim();
    const amount = Number(document.getElementById('modal-alloc-amount').value) || 0;
    const bank = document.getElementById('modal-alloc-bank').value.trim();
    const schedule = document.getElementById('modal-alloc-schedule').value.trim();
    const category = document.getElementById('modal-alloc-category').value;

    if (!name) { alert('항목명을 입력하세요.'); return; }

    appState.allocations[personKey].push({
      id: 'a_' + Date.now(),
      name,
      amount,
      bank,
      schedule,
      category
    });

    closeModal();
    saveState();
    showToast('새 배분 항목이 추가되었습니다.');
  };
}

function openAddFixedExpenseModal() {
  const body = document.getElementById('modal-body');
  document.getElementById('modal-title').innerText = '고정지출 항목 추가';

  body.innerHTML = `
    <div class="form-group">
      <label>고정지출 항목명</label>
      <input type="text" id="modal-fe-name" class="form-control" placeholder="예: 관리비, OTT구독">
    </div>
    <div class="form-group">
      <label>기준 금액 (원)</label>
      <input type="number" id="modal-fe-amount" class="form-control" placeholder="0">
    </div>
    <div class="form-group">
      <label>실제 지출 금액 (원)</label>
      <input type="number" id="modal-fe-actual" class="form-control" placeholder="0">
    </div>
    <div class="form-group">
      <label>이체일자</label>
      <input type="text" id="modal-fe-day" class="form-control" placeholder="예: 10일, 26일, 수기">
    </div>
    <div class="form-group">
      <label>결제수단</label>
      <input type="text" id="modal-fe-method" class="form-control" placeholder="예: 하나카드, 현금, 오빠카드">
    </div>
    <div class="form-group">
      <label>카테고리</label>
      <select id="modal-fe-category" class="form-select">
        <option value="보험">보험</option>
        <option value="구독">구독</option>
        <option value="용돈">용돈</option>
        <option value="핸드폰요금">핸드폰요금</option>
        <option value="주거">주거(관리비)</option>
        <option value="곗돈">곗돈</option>
        <option value="교통">교통</option>
      </select>
    </div>
    <div class="form-group">
      <label>비고</label>
      <input type="text" id="modal-fe-note" class="form-control" placeholder="메모">
    </div>
  `;

  document.getElementById('item-modal').classList.add('active');

  document.getElementById('modal-save-btn').onclick = () => {
    const name = document.getElementById('modal-fe-name').value.trim();
    const amount = Number(document.getElementById('modal-fe-amount').value) || 0;
    const actualMarch = Number(document.getElementById('modal-fe-actual').value) || amount;
    const day = document.getElementById('modal-fe-day').value.trim();
    const method = document.getElementById('modal-fe-method').value.trim();
    const category = document.getElementById('modal-fe-category').value;
    const note = document.getElementById('modal-fe-note').value.trim();

    if (!name) { alert('항목명을 입력하세요.'); return; }

    appState.fixedExpenses.push({
      id: 'fe_' + Date.now(),
      name,
      amount,
      actualMarch,
      day,
      method,
      category,
      note,
      isPaid: false
    });

    closeModal();
    saveState();
    showToast('고정지출 항목이 추가되었습니다.');
  };
}

// Chart.js Rendering Engine
function renderCharts() {
  const calcs = calculateTotals();

  // 1. Monthly Trend Chart
  const trendCtx = document.getElementById('monthlyTrendChart');
  if (trendCtx) {
    const history = appState.monthlyHistory || [];
    const labels = history.map(h => h.month);
    const incomeData = history.map(h => h.totalIncome);
    const expenseData = history.map(h => h.totalExpenses);
    const fixedData = history.map(h => h.fixedExpenses);
    const balanceData = history.map(h => h.remaining);

    if (trendChartInstance) trendChartInstance.destroy();

    trendChartInstance = new Chart(trendCtx, {
      type: 'line',
      data: {
        labels: labels,
        datasets: [
          {
            label: '총 수입',
            data: incomeData,
            borderColor: '#06b6d4',
            backgroundColor: 'rgba(6, 182, 212, 0.1)',
            borderWidth: 3,
            tension: 0.3,
            fill: true
          },
          {
            label: '총 배분/지출',
            data: expenseData,
            borderColor: '#f59e0b',
            backgroundColor: 'rgba(245, 158, 11, 0.1)',
            borderWidth: 3,
            tension: 0.3
          },
          {
            label: '고정지출',
            data: fixedData,
            borderColor: '#ef4444',
            backgroundColor: 'rgba(239, 68, 68, 0.1)',
            borderWidth: 2,
            borderDash: [5, 5],
            tension: 0.3
          },
          {
            label: '남은 잔액',
            data: balanceData,
            borderColor: '#10b981',
            backgroundColor: 'rgba(16, 185, 129, 0.2)',
            borderWidth: 3,
            tension: 0.3
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.dataset.label}: ${formatKRW(ctx.raw)}`
            }
          }
        },
        scales: {
          x: { grid: { color: 'rgba(255,255,255,0.05)' } },
          y: {
            grid: { color: 'rgba(255,255,255,0.05)' },
            ticks: {
              callback: (val) => `${(val / 10000).toLocaleString()}만원`
            }
          }
        }
      }
    });
  }

  // 2. Category Pie Chart
  const pieCtx = document.getElementById('categoryPieChart');
  if (pieCtx) {
    const catMap = {};
    [...appState.allocations.gyewon, ...appState.allocations.dongwook].forEach(item => {
      const cat = item.category || '기타';
      catMap[cat] = (catMap[cat] || 0) + Number(item.amount || 0);
    });

    if (categoryChartInstance) categoryChartInstance.destroy();

    categoryChartInstance = new Chart(pieCtx, {
      type: 'doughnut',
      data: {
        labels: Object.keys(catMap),
        datasets: [{
          data: Object.values(catMap),
          backgroundColor: ['#6366f1', '#a855f7', '#3b82f6', '#10b981', '#f59e0b', '#ec4899', '#06b6d4'],
          borderWidth: 0
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'right', labels: { color: '#94a3b8', font: { size: 11 } } },
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.label}: ${formatKRW(ctx.raw)}`
            }
          }
        },
        cutout: '70%'
      }
    });
  }

  // 3. Payment Method Chart
  const pmCtx = document.getElementById('paymentMethodChart');
  if (pmCtx) {
    const pmMap = {};
    appState.fixedExpenses.forEach(item => {
      const pm = item.method || '기타';
      pmMap[pm] = (pmMap[pm] || 0) + Number(item.actualMarch || item.amount || 0);
    });

    if (paymentChartInstance) paymentChartInstance.destroy();

    paymentChartInstance = new Chart(pmCtx, {
      type: 'bar',
      data: {
        labels: Object.keys(pmMap),
        datasets: [{
          label: '고정지출 액',
          data: Object.values(pmMap),
          backgroundColor: '#3b82f6',
          borderRadius: 6
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: { legend: { display: false } },
        scales: {
          x: { grid: { display: false } },
          y: { grid: { color: 'rgba(255,255,255,0.05)' } }
        }
      }
    });
  }
}

// Snapshot Monthly Records
function saveMonthlySnapshot() {
  const calcs = calculateTotals();
  const month = appState.currentMonth || '2026-03';

  const existingIdx = appState.monthlyHistory.findIndex(h => h.month === month);
  const record = {
    month: month,
    totalIncome: calcs.totalIncome,
    totalExpenses: calcs.totalExpenses,
    fixedExpenses: calcs.fixedExpenseActualTotal,
    remaining: calcs.remainingBalance,
    savings: calcs.totalSavings
  };

  if (existingIdx !== -1) {
    appState.monthlyHistory[existingIdx] = record;
  } else {
    appState.monthlyHistory.push(record);
  }

  // Sort history chronologically
  appState.monthlyHistory.sort((a, b) => a.month.localeCompare(b.month));

  saveState();
  showToast(`${month} 월 기록 마감이 저장되었습니다! 추이 차트에서 확인하세요.`);
}

function deleteHistoryRow(index) {
  if (confirm('선택한 월별 기록을 삭제하시겠습니까?')) {
    appState.monthlyHistory.splice(index, 1);
    saveState();
    showToast('월별 기록이 삭제되었습니다.');
  }
}

// Utility Functions
function formatKRW(num) {
  return (Number(num) || 0).toLocaleString('ko-KR') + ' 원';
}

function formatCompactKRW(num) {
  const val = Number(num) || 0;
  if (val >= 10000) {
    return Math.floor(val / 10000).toLocaleString() + '만';
  }
  return val.toLocaleString() + '원';
}

function showToast(msg, type = 'success') {
  const container = document.getElementById('toast-container');
  const toast = document.createElement('div');
  toast.className = `toast toast-${type}`;
  toast.innerHTML = `<i class="fa-solid fa-circle-check"></i> <span>${msg}</span>`;
  container.appendChild(toast);
  setTimeout(() => toast.remove(), 3000);
}

function copySummaryToClipboard() {
  const calcs = calculateTotals();
  const summaryText = `[동욱 & 계원 재정 요약 - ${appState.currentMonth}]
• 총 수입: ${formatKRW(calcs.totalIncome)} (계원 ${formatCompactKRW(calcs.incomeGyewon)} / 동욱 ${formatCompactKRW(calcs.incomeDongwook)})
• 총 지출/배분: ${formatKRW(calcs.totalExpenses)}
• 고정지출: ${formatKRW(calcs.fixedExpenseActualTotal)}
• 저축/적금/투자: ${formatKRW(calcs.totalSavings)} (저축률 ${calcs.savingsRate}%)
• 남은 잔액 (여유자금): ${formatKRW(calcs.remainingBalance)}`;

  navigator.clipboard.writeText(summaryText).then(() => {
    showToast('클립보드에 요약 정보가 복사되었습니다!');
  });
}

function exportBackupJSON() {
  const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(appState, null, 2));
  const downloadAnchor = document.createElement('a');
  downloadAnchor.setAttribute("href", dataStr);
  downloadAnchor.setAttribute("download", `dongwook_gyewon_budget_${appState.currentMonth}.json`);
  document.body.appendChild(downloadAnchor);
  downloadAnchor.click();
  downloadAnchor.remove();
  showToast('백업 파일이 다운로드되었습니다.');
}

function importBackupJSON(e) {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function(evt) {
    try {
      const data = JSON.parse(evt.target.result);
      if (data.allocations && data.fixedExpenses) {
        appState = data;
        saveState();
        showToast('데이터가 성공적으로 복원되었습니다.');
      } else {
        alert('유효하지 않은 백업 파일 형식입니다.');
      }
    } catch (err) {
      alert('파일 읽기 오류: ' + err.message);
    }
  };
  reader.readAsText(file);
}
