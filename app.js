/**
 * Dongwook & Gyewon Monthly Budget Manager
 * Application Logic Engine JavaScript
 */

// Initial Excel Default Data
const DEFAULT_EXCEL_DATA = {
  currentMonth: '2026-03',
  categories: {
    recipients: ['계원', '동욱', '공통 / 가구'],
    incomeItems: ['기본급 (급여)', '보너스 / 상여금', '인센티브', '부수입 / 알바', '투자수익 / 배당', '기타 수입']
  },
  incomes: [
    { id: 'inc_g1', person: '계원', title: '기본급 (급여)', amount: 4700000, day: '27일', note: '계원 월급날 (27일)' },
    { id: 'inc_d1', person: '동욱', title: '기본급 (급여)', amount: 3750000, day: '16일', note: '동욱 월급날 (16일)' }
  ],
  allocations: {
    gyewon: [
      { id: 'ag1', name: '생활비', amount: 1024768, bank: '신한', day: '22일', method: '자동이체', schedule: '22일 자동이체', category: '생활비' },
      { id: 'ag2', name: '주택청약', amount: 100000, bank: '국민', day: '22일', method: '자동이체', schedule: '22일 자동이체', category: '저축/적금' },
      { id: 'ag3', name: '경조사', amount: 150000, bank: '신한', day: '22일', method: '오빠자동이체', schedule: '22일/오빠자동이체', category: '비상금/경조사' },
      { id: 'ag4', name: '오빠 IRP', amount: 200000, bank: '신한', day: '22일', method: '오빠자동이체', schedule: '22일/오빠자동이체', category: '투자/연금' },
      { id: 'ag5', name: '원이 IRP', amount: 200000, bank: '신한', day: '22일', method: '자동이체', schedule: '22일 자동이체', category: '투자/연금' },
      { id: 'ag6', name: '오빠 ISA', amount: 200000, bank: '신한', day: '-', method: '-', schedule: '-', category: '투자/연금' },
      { id: 'ag7', name: '원이 ISA', amount: 400000, bank: '키움', day: '-', method: '-', schedule: '-', category: '투자/연금' },
      { id: 'ag8', name: '케이뱅크', amount: 300000, bank: '코인', day: '22일', method: '자동이체', schedule: '22일 자동이체', category: '투자/연금' },
      { id: 'ag9', name: '대출금', amount: 1875232, bank: '고정지출', day: '22일', method: '자동이체', schedule: '22일 자동이체', category: '대출금' },
      { id: 'ag10', name: '자율적금', amount: 100000, bank: '', day: '22일', method: '오빠자동이체', schedule: '22일/오빠자동이체', category: '저축/적금' },
      { id: 'ag11', name: 'IRP', amount: 50000, bank: '신한투자증권', day: '15일', method: '회사자동이체(별도)', schedule: '15일회사자동이체(별도)', category: '투자/연금' }
    ],
    dongwook: [
      { id: 'ad1', name: '생활비', amount: 1750000, bank: '', day: '16일', method: '자동이체', schedule: '16일/자동이체', category: '생활비' },
      { id: 'ad2', name: '적금', amount: 1000000, bank: '어머니', day: '16일', method: '자동이체', schedule: '16일/자동이체', category: '저축/적금' },
      { id: 'ad3', name: '미래적금', amount: 500000, bank: '신한', day: '16일', method: '자동이체', schedule: '16일/자동이체', category: '저축/적금' },
      { id: 'ad4', name: '오빠 ISA', amount: 300000, bank: '신한', day: '16일', method: '자동이체', schedule: '16일/자동이체', category: '투자/연금' },
      { id: 'ad5', name: '에어비앤비', amount: 100000, bank: '저축', day: '-', method: '-', schedule: '-', category: '저축/적금' },
      { id: 'ad6', name: '제로스토어', amount: 100000, bank: '저축', day: '-', method: '-', schedule: '-', category: '저축/적금' }
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

let currentIncomeRecipientFilter = 'ALL';

// Calculation Helper
function calculateTotals() {
  // Incomes
  let incomeGyewon = 0;
  let incomeDongwook = 0;
  let incomeCommon = 0;

  (appState.incomes || []).forEach(inc => {
    const amt = Number(inc.amount) || 0;
    const p = (inc.person || '').toLowerCase();
    if (p === 'gyewon' || p === '계원') incomeGyewon += amt;
    else if (p === 'dongwook' || p === '동욱') incomeDongwook += amt;
    else incomeCommon += amt;
  });

  const totalIncome = incomeGyewon + incomeDongwook + incomeCommon;

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
    incomeCommon,
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

// Render Unified Incomes Table
function renderIncomeTables() {
  const tbody = document.querySelector('#table-income-unified tbody');
  if (!tbody) return;
  tbody.innerHTML = '';

  const calcs = calculateTotals();

  const gSumElem = document.getElementById('gyewon-income-sum');
  const dSumElem = document.getElementById('dongwook-income-sum');
  const cSumElem = document.getElementById('common-income-sum');
  const uTotalElem = document.getElementById('unified-income-total');

  if (gSumElem) gSumElem.innerText = formatKRW(calcs.incomeGyewon);
  if (dSumElem) dSumElem.innerText = formatKRW(calcs.incomeDongwook);
  if (cSumElem) cSumElem.innerText = formatKRW(calcs.incomeCommon);
  if (uTotalElem) uTotalElem.innerText = formatKRW(calcs.totalIncome);

  // Render Recipient Filter Pills
  renderIncomeFilterPills();

  // Filter items
  const filtered = (appState.incomes || []).filter(inc => {
    if (currentIncomeRecipientFilter === 'ALL') return true;
    const p = (inc.person || '').toLowerCase();
    if (currentIncomeRecipientFilter === 'gyewon' || currentIncomeRecipientFilter === '계원') {
      return p === 'gyewon' || p === '계원';
    }
    if (currentIncomeRecipientFilter === 'dongwook' || currentIncomeRecipientFilter === '동욱') {
      return p === 'dongwook' || p === '동욱';
    }
    return inc.person === currentIncomeRecipientFilter;
  });

  filtered.forEach(inc => {
    let badgeClass = 'badge-info';
    let personLabel = inc.person || '기타';
    if (inc.person === 'gyewon' || inc.person === '계원') { badgeClass = 'badge-gyewon'; personLabel = '계원'; }
    else if (inc.person === 'dongwook' || inc.person === '동욱') { badgeClass = 'badge-dongwook'; personLabel = '동욱'; }
    else if (inc.person === 'common' || inc.person === '공통 / 가구') { badgeClass = 'badge-info'; personLabel = '공통/가구'; }

    const categoryName = inc.category || inc.title || '기본급 (급여)';

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><span class="badge ${badgeClass}">${personLabel}</span></td>
      <td><div class="editable-cell" title="클릭하여 카테고리 빠른 변경" onclick="editInlineCategory(this, '${inc.id}', event)"><span class="badge badge-secondary" style="font-weight:600; cursor:pointer;">${categoryName}</span></div></td>
      <td><div class="editable-cell" title="클릭하여 수정" onclick="editInlineCell(this, 'incomes', '${inc.id}', 'title')">${inc.title}</div></td>
      <td><div class="editable-cell cell-amount text-accent" title="클릭하여 월급/금액 수정" onclick="editInlineCell(this, 'incomes', '${inc.id}', 'amount', 'number')">${formatKRW(inc.amount)}</div></td>
      <td><div class="editable-cell text-success font-weight-bold" title="클릭하여 입금일 수정" onclick="editInlineCell(this, 'incomes', '${inc.id}', 'day')"><i class="fa-regular fa-calendar-check"></i> ${inc.day || '25일'}</div></td>
      <td><div class="editable-cell" title="클릭하여 수정" onclick="editInlineCell(this, 'incomes', '${inc.id}', 'note')">${inc.note || '-'}</div></td>
      <td style="white-space:nowrap; text-align:right;">
        <button class="btn btn-outline-primary btn-sm" title="수입/월급 수정" onclick="openEditIncomeModal('${inc.id}')"><i class="fa-solid fa-pen-to-square"></i></button>
        <button class="btn btn-outline-danger btn-sm" title="삭제" onclick="deleteItem('incomes', '${inc.id}')"><i class="fa-solid fa-trash"></i></button>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

function renderIncomeFilterPills() {
  const containers = [
    document.getElementById('income-recipient-pills'),
    document.getElementById('alloc-recipient-pills')
  ].filter(Boolean);

  if (containers.length === 0) return;

  // Get recipient categories
  const defaultRecipients = ['계원', '동욱', '공통 / 가구'];
  const userRecipients = (appState.categories && appState.categories.recipients) ? appState.categories.recipients : defaultRecipients;
  const allRecipients = Array.from(new Set([...defaultRecipients, ...userRecipients]));

  containers.forEach(container => {
    container.innerHTML = '';

    const allBtn = document.createElement('button');
    allBtn.className = `btn btn-sm ${currentIncomeRecipientFilter === 'ALL' ? 'btn-primary' : 'btn-glass'}`;
    allBtn.innerText = '전체 보기';
    allBtn.onclick = () => { 
      currentIncomeRecipientFilter = 'ALL'; 
      renderIncomeTables(); 
      renderAllocationTables(); 
    };
    container.appendChild(allBtn);

    allRecipients.forEach(r => {
      const btn = document.createElement('button');
      const isSelected = currentIncomeRecipientFilter === r;
      btn.className = `btn btn-sm ${isSelected ? 'btn-primary' : 'btn-glass'}`;
      btn.innerText = r;
      btn.onclick = () => { 
        currentIncomeRecipientFilter = r; 
        renderIncomeTables(); 
        renderAllocationTables(); 
      };
      container.appendChild(btn);
    });
  });
}

// Supabase Integration Credentials
const SUPABASE_URL = 'https://bdnqlcrpytkwuaonhgmm.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImJkbnFsY3JweXRrd3Vhb25oZ21tIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODk5MjYyMDEsImV4cCI6MjEwNTUwMjIwMX0.QLt4HgVyCoAR2oy1R5O3SdNe7N3XtVxP7rjabS3j_ew';

let supabaseClient = null;
if (window.supabase) {
  try {
    supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
  } catch (e) {
    console.error('Supabase Client Init Error:', e);
  }
}

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
  syncFromSupabase();
}

// State Normalization Helper
function normalizeAllocationItem(item) {
  if (!item) return item;
  if (!item.day && !item.method && item.schedule) {
    const str = String(item.schedule).trim();
    const match = str.match(/^(\d+일)\s*[\/]?\s*(.*)$/);
    if (match) {
      item.day = match[1];
      item.method = match[2] || '-';
    } else {
      const parts = str.split('/');
      if (parts.length > 1) {
        item.day = parts[0].trim();
        item.method = parts.slice(1).join('/').trim() || '-';
      } else if (str.includes('일')) {
        item.day = str;
        item.method = '-';
      } else {
        item.day = '-';
        item.method = str || '-';
      }
    }
  }
  if (!item.day) item.day = '-';
  if (!item.method) item.method = '-';
  item.schedule = `${item.day === '-' ? '' : item.day} ${item.method === '-' ? '' : item.method}`.trim() || '-';
  return item;
}

function normalizeState(state) {
  if (state && state.allocations) {
    if (Array.isArray(state.allocations.gyewon)) {
      state.allocations.gyewon.forEach(normalizeAllocationItem);
    }
    if (Array.isArray(state.allocations.dongwook)) {
      state.allocations.dongwook.forEach(normalizeAllocationItem);
    }
  }
  return state;
}

// State Management & LocalStorage & Supabase
function loadState() {
  let state = null;
  const saved = localStorage.getItem('dongwook_gyewon_budget_app_v2');
  if (saved) {
    try {
      state = JSON.parse(saved);
    } catch (e) {
      console.error('Failed to parse state:', e);
    }
  }
  if (!state) {
    state = JSON.parse(JSON.stringify(DEFAULT_EXCEL_DATA));
  }
  return normalizeState(state);
}

function saveState() {
  localStorage.setItem('dongwook_gyewon_budget_app_v2', JSON.stringify(appState));
  renderAll();
  syncToSupabase();
}

// Supabase Status Indicator Update
function updateSupabaseBadge(isConnected, message = '') {
  const badge = document.getElementById('supabase-status-badge');
  if (!badge) return;

  if (isConnected) {
    badge.className = 'badge badge-gyewon';
    badge.innerHTML = '<i class="fa-solid fa-cloud-check"></i> Supabase 연동 완료';
  } else {
    badge.className = 'badge badge-info';
    badge.innerHTML = `<i class="fa-solid fa-cloud"></i> 로컬 저장소 사용 (${message || '테이블 생성 필요'})`;
  }
}

// Sync from Supabase Cloud
async function syncFromSupabase() {
  if (!supabaseClient) {
    updateSupabaseBadge(false, 'SDK 미로드');
    return;
  }

  try {
    const { data, error } = await supabaseClient
      .from('monthly_budget_data')
      .select('data')
      .eq('id', 'main')
      .single();

    if (error) {
      if (error.code === 'PGRST116') {
        // No row exists yet, push current local appState to Supabase
        await syncToSupabase();
      } else {
        console.warn('Supabase fetch notice:', error.message);
        updateSupabaseBadge(false, '테이블 준비 중');
      }
    } else if (data && data.data) {
      appState = normalizeState(data.data);
      localStorage.setItem('dongwook_gyewon_budget_app_v2', JSON.stringify(appState));
      renderAll();
      updateSupabaseBadge(true);
      showToast('Supabase 클라우드에서 데이터를 연동했습니다!');
    }
  } catch (err) {
    console.error('Supabase sync error:', err);
    updateSupabaseBadge(false, '연결 실패');
  }
}

// Sync to Supabase Cloud
async function syncToSupabase() {
  if (!supabaseClient) return;

  try {
    const { error } = await supabaseClient
      .from('monthly_budget_data')
      .upsert({
        id: 'main',
        data: appState,
        updated_at: new Date().toISOString()
      });

    if (error) {
      console.warn('Supabase save notice:', error.message);
      updateSupabaseBadge(false, '테이블 생성 필요');
    } else {
      updateSupabaseBadge(true);
    }
  } catch (err) {
    console.error('Supabase save error:', err);
    updateSupabaseBadge(false, '저장 오류');
  }
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

  // Category Management & Income Addition Modal Triggers
  const btnCatManage = document.getElementById('btn-manage-categories');
  if (btnCatManage) {
    btnCatManage.addEventListener('click', () => openCategoryManagerModal());
  }
  const btnAllocCatManage = document.getElementById('btn-manage-alloc-categories');
  if (btnAllocCatManage) {
    btnAllocCatManage.addEventListener('click', () => openCategoryManagerModal());
  }
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

  // Generate Income Breakdown by Category
  const incomeByCategory = {};
  (appState.incomes || []).forEach(inc => {
    const cat = inc.title || '기타수입';
    const amt = Number(inc.amount) || 0;
    incomeByCategory[cat] = (incomeByCategory[cat] || 0) + amt;
  });

  const breakdownHTML = Object.entries(incomeByCategory)
    .sort((a, b) => b[1] - a[1]) // Sort descending by amount
    .map(([cat, amt]) => `<span class="badge badge-info">${cat} ${formatCompactKRW(amt)}</span>`)
    .join(' ');
  
  const breakdownElem = document.getElementById('kpi-income-breakdown');
  if (breakdownElem) {
    breakdownElem.innerHTML = breakdownHTML;
  }
  document.getElementById('kpi-fixed-total').innerText = formatKRW(calcs.fixedExpenseTotal);
  document.getElementById('kpi-total-savings').innerText = formatKRW(calcs.totalSavings);
  document.getElementById('kpi-savings-rate').innerText = `${calcs.savingsRate}%`;

  // Render Tables
  renderIncomeTables();
  renderAllocationTables();
  renderFixedExpensesTable();
  renderTimeline();
  renderPaymentMethodSummary();
  renderCategoryAllocationSummary();
  renderHistoryTable();

  // Render Charts
  renderCharts();
}

// Render Allocations Tables
function renderAllocationTables() {
  const calcs = calculateTotals();

  // Update Gyewon Summary Bar
  const gIncElem = document.getElementById('g-summary-income');
  const gAllocElem = document.getElementById('g-summary-alloc');
  const gRemainElem = document.getElementById('g-summary-remain');
  if (gIncElem) gIncElem.innerText = formatKRW(calcs.incomeGyewon);
  if (gAllocElem) gAllocElem.innerText = formatKRW(calcs.allocGyewonSum);
  if (gRemainElem) {
    gRemainElem.innerText = formatKRW(calcs.remainGyewon);
    if (calcs.remainGyewon < 0) gRemainElem.classList.add('negative');
    else gRemainElem.classList.remove('negative');
  }

  // Update Dongwook Summary Bar
  const dIncElem = document.getElementById('d-summary-income');
  const dAllocElem = document.getElementById('d-summary-alloc');
  const dRemainElem = document.getElementById('d-summary-remain');
  if (dIncElem) dIncElem.innerText = formatKRW(calcs.incomeDongwook);
  if (dAllocElem) dAllocElem.innerText = formatKRW(calcs.allocDongwookSum);
  if (dRemainElem) {
    dRemainElem.innerText = formatKRW(calcs.remainDongwook);
    if (calcs.remainDongwook < 0) dRemainElem.classList.add('negative');
    else dRemainElem.classList.remove('negative');
  }

  // Target Filter Visibility Control
  const gyewonCard = document.getElementById('alloc-gyewon-card');
  const dongwookCard = document.getElementById('alloc-dongwook-card');
  const allocContainer = document.getElementById('allocations-container');

  const filter = (currentIncomeRecipientFilter || 'ALL').trim();
  const filterLower = filter.toLowerCase();

  if (gyewonCard && dongwookCard) {
    if (filter === 'ALL' || filterLower === 'all') {
      gyewonCard.style.display = '';
      dongwookCard.style.display = '';
      if (allocContainer) allocContainer.style.gridTemplateColumns = '';
    } else if (filter === '계원' || filterLower === 'gyewon') {
      gyewonCard.style.display = '';
      dongwookCard.style.display = 'none';
      if (allocContainer) allocContainer.style.gridTemplateColumns = '1fr';
    } else if (filter === '동욱' || filterLower === 'dongwook') {
      gyewonCard.style.display = 'none';
      dongwookCard.style.display = '';
      if (allocContainer) allocContainer.style.gridTemplateColumns = '1fr';
    } else {
      gyewonCard.style.display = '';
      dongwookCard.style.display = '';
      if (allocContainer) allocContainer.style.gridTemplateColumns = '';
    }
  }

function getAllocCategoryBadgeClass(category) {
  switch(category) {
    case '생활비': return 'badge-info';
    case '저축/적금': return 'badge-success';
    case '투자/연금': return 'badge-primary';
    case '대출금': return 'badge-danger';
    case '비상금/경조사': return 'badge-warning';
    default: return 'badge-secondary';
  }
}

  // Populate Gyewon Allocation Rows
  const tbodyG = document.querySelector('#table-alloc-gyewon tbody');
  if (tbodyG) {
    tbodyG.innerHTML = '';
    (appState.allocations.gyewon || []).forEach(item => {
      const catBadge = getAllocCategoryBadgeClass(item.category);
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><div class="editable-cell" onclick="editInlineCell(this, 'allocations.gyewon', '${item.id}', 'name')">${item.name}</div></td>
        <td><div class="editable-cell cell-amount" onclick="editInlineCell(this, 'allocations.gyewon', '${item.id}', 'amount', 'number')">${formatKRW(item.amount)}</div></td>
        <td><div class="editable-cell" onclick="editInlineCell(this, 'allocations.gyewon', '${item.id}', 'bank')">${item.bank || '-'}</div></td>
        <td><div class="editable-cell text-primary font-weight-bold" onclick="editInlineCell(this, 'allocations.gyewon', '${item.id}', 'day')">${item.day || '-'}</div></td>
        <td><div class="editable-cell" onclick="editInlineCell(this, 'allocations.gyewon', '${item.id}', 'method')">${item.method || '-'}</div></td>
        <td><div class="editable-cell" title="클릭하여 카테고리 빠른 변경" onclick="editInlineAllocCategory(this, 'allocations.gyewon', '${item.id}', event)"><span class="badge ${catBadge}" style="font-weight:600; cursor:pointer;">${item.category || '기타'}</span></div></td>
        <td><button class="btn btn-outline-danger btn-sm" onclick="deleteItem('allocations.gyewon', '${item.id}')"><i class="fa-solid fa-xmark"></i></button></td>
      `;
      tbodyG.appendChild(tr);
    });
  }

  // Populate Dongwook Allocation Rows
  const tbodyD = document.querySelector('#table-alloc-dongwook tbody');
  if (tbodyD) {
    tbodyD.innerHTML = '';
    (appState.allocations.dongwook || []).forEach(item => {
      const catBadge = getAllocCategoryBadgeClass(item.category);
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td><div class="editable-cell" onclick="editInlineCell(this, 'allocations.dongwook', '${item.id}', 'name')">${item.name}</div></td>
        <td><div class="editable-cell cell-amount" onclick="editInlineCell(this, 'allocations.dongwook', '${item.id}', 'amount', 'number')">${formatKRW(item.amount)}</div></td>
        <td><div class="editable-cell" onclick="editInlineCell(this, 'allocations.dongwook', '${item.id}', 'bank')">${item.bank || '-'}</div></td>
        <td><div class="editable-cell text-primary font-weight-bold" onclick="editInlineCell(this, 'allocations.dongwook', '${item.id}', 'day')">${item.day || '-'}</div></td>
        <td><div class="editable-cell" onclick="editInlineCell(this, 'allocations.dongwook', '${item.id}', 'method')">${item.method || '-'}</div></td>
        <td><div class="editable-cell" title="클릭하여 카테고리 빠른 변경" onclick="editInlineAllocCategory(this, 'allocations.dongwook', '${item.id}', event)"><span class="badge ${catBadge}" style="font-weight:600; cursor:pointer;">${item.category || '기타'}</span></div></td>
        <td><button class="btn btn-outline-danger btn-sm" onclick="deleteItem('allocations.dongwook', '${item.id}')"><i class="fa-solid fa-xmark"></i></button></td>
      `;
      tbodyD.appendChild(tr);
    });
  }
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

// Render Master Timeline / Deposit & Expense Schedule
function renderTimeline() {
  const container = document.getElementById('schedule-timeline');
  container.innerHTML = '';

  const groups = {};

  // 1. Incomes (입금/월급)
  (appState.incomes || []).forEach(inc => {
    let rawDay = inc.day || '미지정';
    let dayKey = rawDay.includes('일') ? rawDay : (rawDay + '일');
    if (!groups[dayKey]) groups[dayKey] = [];
    groups[dayKey].push({
      type: 'INCOME',
      name: `${inc.person === 'gyewon' ? '계원' : '동욱'} ${inc.title}`,
      method: '입금',
      amount: Number(inc.amount) || 0
    });
  });

  // 2. Allocations (Gyewon)
  (appState.allocations.gyewon || []).forEach(item => {
    let dayKey = '수기/기타';
    if (item.schedule && item.schedule.includes('일')) {
      const match = item.schedule.match(/(\d+일)/);
      if (match) dayKey = match[1];
    }
    if (!groups[dayKey]) groups[dayKey] = [];
    groups[dayKey].push({
      type: 'ALLOCATION',
      name: `계원 ${item.name}`,
      method: item.bank || '자동이체',
      amount: Number(item.amount) || 0
    });
  });

  // 3. Allocations (Dongwook)
  (appState.allocations.dongwook || []).forEach(item => {
    let dayKey = '수기/기타';
    if (item.schedule && item.schedule.includes('일')) {
      const match = item.schedule.match(/(\d+일)/);
      if (match) dayKey = match[1];
    }
    if (!groups[dayKey]) groups[dayKey] = [];
    groups[dayKey].push({
      type: 'ALLOCATION',
      name: `동욱 ${item.name}`,
      method: item.bank || '자동이체',
      amount: Number(item.amount) || 0
    });
  });

  // 4. Fixed Expenses (고정지출 이체)
  (appState.fixedExpenses || []).forEach(item => {
    let rawDay = item.day || '수기/기타';
    let dayKey = (rawDay.includes('일') || rawDay === '수기' || rawDay.includes('수기')) ? rawDay : (rawDay ? rawDay + '일' : '수기/기타');
    if (!groups[dayKey]) groups[dayKey] = [];
    groups[dayKey].push({
      type: 'FIXED',
      name: item.name,
      method: item.method || '카드',
      amount: Number(item.actualMarch || item.amount) || 0
    });
  });

  // Sort keys
  const sortedKeys = Object.keys(groups).sort((a, b) => {
    const numA = parseInt(a) || 99;
    const numB = parseInt(b) || 99;
    return numA - numB;
  });

  sortedKeys.forEach(day => {
    const items = groups[day];
    const totalIncomeDay = items.filter(i => i.type === 'INCOME').reduce((acc, c) => acc + c.amount, 0);
    const totalOutDay = items.filter(i => i.type !== 'INCOME').reduce((acc, c) => acc + c.amount, 0);

    const card = document.createElement('div');
    card.className = 'timeline-card';
    card.innerHTML = `
      <div class="timeline-card-header">
        <span class="timeline-date"><i class="fa-regular fa-calendar-check"></i> ${day} 일정</span>
        <span class="timeline-count">
          ${totalIncomeDay > 0 ? `<span class="badge badge-gyewon">+${formatKRW(totalIncomeDay)}</span> ` : ''}
          ${totalOutDay > 0 ? `<span class="badge badge-info">-${formatKRW(totalOutDay)}</span>` : ''}
        </span>
      </div>
      <ul class="timeline-item-list">
        ${items.map(i => {
          let badgeClass = 'badge-dongwook';
          let badgeText = '지출';
          if (i.type === 'INCOME') { badgeClass = 'badge-gyewon'; badgeText = '입금'; }
          else if (i.type === 'ALLOCATION') { badgeClass = 'badge-info'; badgeText = '배분'; }

          return `
            <li>
              <span><span class="badge ${badgeClass}">${badgeText}</span> ${i.name} (${i.method})</span>
              <strong class="${i.type === 'INCOME' ? 'text-success' : ''}">${formatKRW(i.amount)}</strong>
            </li>
          `;
        }).join('')}
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

// Render Category Allocation Summary Cards on Dashboard
function renderCategoryAllocationSummary() {
  const container = document.getElementById('category-allocation-summary');
  if (!container) return;
  container.innerHTML = '';

  const catMap = {};
  let totalAlloc = 0;
  const allAllocations = [...(appState.allocations?.gyewon || []), ...(appState.allocations?.dongwook || [])];
  
  allAllocations.forEach(item => {
    const cat = item.category || '기타';
    const amt = Number(item.amount || 0);
    catMap[cat] = (catMap[cat] || 0) + amt;
    totalAlloc += amt;
  });

  const sortedCats = Object.entries(catMap).sort((a, b) => b[1] - a[1]);

  sortedCats.forEach(([cat, amt]) => {
    const badgeClass = getAllocCategoryBadgeClass(cat);
    const pct = totalAlloc > 0 ? ((amt / totalAlloc) * 100).toFixed(1) : 0;
    
    const card = document.createElement('div');
    card.className = 'pm-card';
    card.innerHTML = `
      <div class="pm-name"><span class="badge ${badgeClass}">${cat}</span></div>
      <div class="pm-amount" style="margin-top:4px;">${formatKRW(amt)}</div>
      <div style="font-size:11px; color:var(--text-muted); margin-top:2px;">${pct}% (${formatCompactKRW(amt)})</div>
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

function formatDayInput(val) {
  if (!val) return '';
  val = String(val).trim();
  // Pure numbers like "25" -> "25일"
  if (/^\d+$/.test(val)) {
    return val + '일';
  }
  // Slash numbers like "2/26" -> "2일/26일"
  if (/^\d+\/\d+$/.test(val)) {
    const parts = val.split('/');
    return `${parts[0]}일/${parts[1]}일`;
  }
  return val;
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
    if (type === 'number') {
      newVal = Number(newVal) || 0;
    } else if (fieldName === 'day' || fieldName === 'schedule') {
      newVal = formatDayInput(newVal);
    }
    
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
    if (pathStr.startsWith('allocations') && (fieldName === 'day' || fieldName === 'method')) {
      const d = target.day && target.day !== '-' ? target.day : '';
      const m = target.method && target.method !== '-' ? target.method : '';
      target.schedule = `${d} ${m}`.trim() || '-';
    }
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

// Helper to attach modal formatters
function attachIncomeModalFormatters() {
  const amountInput = document.getElementById('modal-inc-amount');
  const dayInput = document.getElementById('modal-inc-day');
  const catSelect = document.getElementById('modal-inc-cat-select');
  const titleInput = document.getElementById('modal-inc-title');

  // 1. Real-time Thousands Comma Formatting
  if (amountInput) {
    amountInput.addEventListener('input', (e) => {
      const raw = e.target.value.replace(/[^\d]/g, '');
      e.target.value = raw ? Number(raw).toLocaleString('ko-KR') : '';
    });
  }

  // 2. Automatic '일' Appending on Day Input
  if (dayInput) {
    const autoDay = (e) => {
      let val = e.target.value.trim();
      if (!val) return;
      if (/^\d+$/.test(val)) {
        e.target.value = val + '일';
      } else if (/^\d+\/\d+$/.test(val)) {
        const parts = val.split('/');
        e.target.value = `${parts[0]}일/${parts[1]}일`;
      }
    };
    dayInput.addEventListener('blur', autoDay);
    dayInput.addEventListener('change', autoDay);
  }

  // 3. Category Dropdown updates Title Input
  if (catSelect && titleInput) {
    catSelect.addEventListener('change', (e) => {
      if (e.target.value && e.target.value !== 'CUSTOM') {
        titleInput.value = e.target.value;
      } else if (e.target.value === 'CUSTOM') {
        titleInput.value = '';
        titleInput.focus();
      }
    });
  }
}

function openCategoryManagerModal() {
  const body = document.getElementById('modal-body');
  document.getElementById('modal-title').innerText = '카테고리 설정 (수입 & 배분 통합 관리)';

  if (!appState.categories) {
    appState.categories = {
      recipients: ['계원', '동욱', '공통 / 가구'],
      incomeItems: ['기본급 (급여)', '보너스 / 상여금', '인센티브', '부수입 / 알바', '투자수익 / 배당', '기타 수입'],
      allocCategories: ['생활비', '저축/적금', '투자/연금', '대출금', '비상금/경조사']
    };
  }
  if (!appState.categories.allocCategories) {
    appState.categories.allocCategories = ['생활비', '저축/적금', '투자/연금', '대출금', '비상금/경조사'];
  }

  const renderCatLists = () => {
    body.innerHTML = `
      <div class="card-box mb-3" style="padding: 16px; background: rgba(0,0,0,0.2);">
        <h4 style="font-size:14px; font-weight:700; margin-bottom:10px;"><i class="fa-solid fa-user-tag"></i> 수입 대상자 카테고리</h4>
        <div class="flex-gap-2 mb-3" style="flex-wrap:wrap;" id="recipient-cat-tags">
          ${appState.categories.recipients.map((r, idx) => `
            <span class="badge badge-info" style="padding: 6px 12px; font-size:12px; display:inline-flex; align-items:center; gap:8px;">
              ${r} 
              <i class="fa-solid fa-pen-to-square" style="cursor:pointer;" title="이름 수정" onclick="editRecipientCategory(${idx})"></i>
              <i class="fa-solid fa-xmark" style="cursor:pointer;" title="삭제" onclick="deleteRecipientCategory(${idx})"></i>
            </span>
          `).join('')}
        </div>
        <div class="flex-gap-2">
          <input type="text" id="new-recipient-input" class="form-control form-control-sm" placeholder="새 대상자 이름 (예: 부부 비상금, 사업체)">
          <button class="btn btn-primary btn-sm" onclick="addRecipientCategory()"><i class="fa-solid fa-plus"></i> 추가</button>
        </div>
      </div>

      <div class="card-box mb-3" style="padding: 16px; background: rgba(0,0,0,0.2);">
        <h4 style="font-size:14px; font-weight:700; margin-bottom:10px;"><i class="fa-solid fa-list-check"></i> 수입 항목 카테고리</h4>
        <div class="flex-gap-2 mb-3" style="flex-wrap:wrap;" id="item-cat-tags">
          ${appState.categories.incomeItems.map((item, idx) => `
            <span class="badge badge-gyewon" style="padding: 6px 12px; font-size:12px; display:inline-flex; align-items:center; gap:8px;">
              ${item} 
              <i class="fa-solid fa-pen-to-square" style="cursor:pointer;" title="이름 수정" onclick="editIncomeItemCategory(${idx})"></i>
              <i class="fa-solid fa-xmark" style="cursor:pointer;" title="삭제" onclick="deleteIncomeItemCategory(${idx})"></i>
            </span>
          `).join('')}
        </div>
        <div class="flex-gap-2">
          <input type="text" id="new-item-cat-input" class="form-control form-control-sm" placeholder="새 수입 항목명 (예: 주식배당, 임대수익)">
          <button class="btn btn-primary btn-sm" onclick="addIncomeItemCategory()"><i class="fa-solid fa-plus"></i> 추가</button>
        </div>
      </div>

      <div class="card-box" style="padding: 16px; background: rgba(0,0,0,0.2);">
        <h4 style="font-size:14px; font-weight:700; margin-bottom:10px;"><i class="fa-solid fa-layer-group"></i> 월급 배분 카테고리</h4>
        <div class="flex-gap-2 mb-3" style="flex-wrap:wrap;" id="alloc-cat-tags">
          ${(appState.categories.allocCategories || []).map((cat, idx) => `
            <span class="badge ${getAllocCategoryBadgeClass(cat)}" style="padding: 6px 12px; font-size:12px; display:inline-flex; align-items:center; gap:8px;">
              ${cat} 
              <i class="fa-solid fa-pen-to-square" style="cursor:pointer;" title="이름 수정" onclick="editAllocCategory(${idx})"></i>
              <i class="fa-solid fa-xmark" style="cursor:pointer;" title="삭제" onclick="deleteAllocCategory(${idx})"></i>
            </span>
          `).join('')}
        </div>
        <div class="flex-gap-2">
          <input type="text" id="new-alloc-cat-input" class="form-control form-control-sm" placeholder="새 배분 카테고리명 (예: 통신비, 교육비, 여행적금)">
          <button class="btn btn-primary btn-sm" onclick="addAllocCategory()"><i class="fa-solid fa-plus"></i> 추가</button>
        </div>
      </div>
    `;
  };

  renderCatLists();
  document.getElementById('item-modal').classList.add('active');

  document.getElementById('modal-save-btn').onclick = () => {
    closeModal();
    saveState();
    renderIncomeTables();
    renderAllocationTables();
    showToast('카테고리 설정이 저장되었습니다.');
  };
}

window.editRecipientCategory = function(idx) {
  const oldVal = appState.categories.recipients[idx];
  const newVal = prompt('수입 대상자 카테고리 이름을 수정하세요:', oldVal);
  if (newVal !== null && newVal.trim() !== '' && newVal.trim() !== oldVal) {
    const trimmed = newVal.trim();
    appState.categories.recipients[idx] = trimmed;
    (appState.incomes || []).forEach(inc => {
      if (inc.person === oldVal) inc.person = trimmed;
    });
    if (currentIncomeRecipientFilter === oldVal) {
      currentIncomeRecipientFilter = trimmed;
    }
    saveState();
    openCategoryManagerModal();
    renderIncomeTables();
    showToast(`수입 대상 카테고리가 '${trimmed}'(으)로 수정되었습니다.`);
  }
};

window.editIncomeItemCategory = function(idx) {
  const oldVal = appState.categories.incomeItems[idx];
  const newVal = prompt('수입 항목 카테고리 이름을 수정하세요:', oldVal);
  if (newVal !== null && newVal.trim() !== '' && newVal.trim() !== oldVal) {
    const trimmed = newVal.trim();
    appState.categories.incomeItems[idx] = trimmed;
    (appState.incomes || []).forEach(inc => {
      if (inc.category === oldVal) inc.category = trimmed;
    });
    saveState();
    openCategoryManagerModal();
    renderIncomeTables();
    showToast(`수입 항목 카테고리가 '${trimmed}'(으)로 수정되었습니다.`);
  }
};

window.addRecipientCategory = function() {
  const input = document.getElementById('new-recipient-input');
  if (!input) return;
  const val = input.value.trim();
  if (val) {
    if (!appState.categories.recipients.includes(val)) {
      appState.categories.recipients.push(val);
      saveState();
      openCategoryManagerModal();
      renderIncomeTables();
    }
  }
};

window.deleteRecipientCategory = function(idx) {
  if (appState.categories.recipients.length <= 1) {
    alert('최소 1개 이상의 대상 카테고리가 필요합니다.');
    return;
  }
  const removed = appState.categories.recipients.splice(idx, 1);
  saveState();
  openCategoryManagerModal();
  renderIncomeTables();
};

window.addIncomeItemCategory = function() {
  const input = document.getElementById('new-item-cat-input');
  if (!input) return;
  const val = input.value.trim();
  if (val) {
    if (!appState.categories.incomeItems.includes(val)) {
      appState.categories.incomeItems.push(val);
      saveState();
      openCategoryManagerModal();
      renderIncomeTables();
    }
  }
};

window.deleteIncomeItemCategory = function(idx) {
  saveState();
  openCategoryManagerModal();
  renderIncomeTables();
};

window.editInlineCategory = function(element, incId, evt) {
  if (evt) evt.stopPropagation();
  if (element.querySelector('select')) return;

  const inc = (appState.incomes || []).find(i => i.id === incId);
  if (!inc) return;

  const itemCatList = (appState.categories && appState.categories.incomeItems) ? appState.categories.incomeItems : ['기본급 (급여)', '보너스 / 상여금', '인센티브', '부수입 / 알바', '투자수익 / 배당', '기타 수입'];

  const select = document.createElement('select');
  select.className = 'form-select form-select-sm';
  select.style.width = '100%';
  select.style.minWidth = '110px';

  itemCatList.forEach(cat => {
    const opt = document.createElement('option');
    opt.value = cat;
    opt.innerText = cat;
    if ((inc.category || inc.title) === cat) opt.selected = true;
    select.appendChild(opt);
  });
  const customOpt = document.createElement('option');
  customOpt.value = 'NEW_CUSTOM';
  customOpt.innerText = '+ 새 카테고리 추가...';
  select.appendChild(customOpt);

  element.innerHTML = '';
  element.appendChild(select);

  let committed = false;
  const commitChange = (val) => {
    if (committed) return;
    committed = true;
    if (val === 'NEW_CUSTOM') {
      setTimeout(() => {
        const newCat = prompt('새 수입 항목 카테고리명을 입력하세요:');
        if (newCat && newCat.trim()) {
          const trimmed = newCat.trim();
          if (!appState.categories.incomeItems.includes(trimmed)) {
            appState.categories.incomeItems.push(trimmed);
          }
          inc.category = trimmed;
        }
        saveState();
        renderIncomeTables();
      }, 50);
    } else {
      if (val) inc.category = val;
      saveState();
      renderIncomeTables();
    }
  };

  select.onchange = (e) => commitChange(e.target.value);
  select.onblur = (e) => {
    if (select.value !== 'NEW_CUSTOM') {
      commitChange(select.value);
    }
  };

  setTimeout(() => select.focus(), 20);
};

window.editInlineAllocCategory = function(element, pathStr, itemId, evt) {
  if (evt) evt.stopPropagation();
  if (element.querySelector('select')) return;

  const list = resolvePath(pathStr);
  const item = list.find(i => i.id === itemId);
  if (!item) return;

  if (!appState.categories) appState.categories = {};
  if (!appState.categories.allocCategories) {
    appState.categories.allocCategories = ['생활비', '저축/적금', '투자/연금', '대출금', '비상금/경조사'];
  }
  const catList = appState.categories.allocCategories;

  const select = document.createElement('select');
  select.className = 'form-select form-select-sm';
  select.style.width = '100%';
  select.style.minWidth = '110px';

  catList.forEach(cat => {
    const opt = document.createElement('option');
    opt.value = cat;
    opt.innerText = cat;
    if (item.category === cat) opt.selected = true;
    select.appendChild(opt);
  });
  const customOpt = document.createElement('option');
  customOpt.value = 'NEW_CUSTOM';
  customOpt.innerText = '+ 새 카테고리 추가...';
  select.appendChild(customOpt);

  element.innerHTML = '';
  element.appendChild(select);

  let committed = false;
  const commitChange = (val) => {
    if (committed) return;
    committed = true;
    if (val === 'NEW_CUSTOM') {
      setTimeout(() => {
        const newCat = prompt('새 배분 카테고리명을 입력하세요:');
        if (newCat && newCat.trim()) {
          const trimmed = newCat.trim();
          if (!appState.categories.allocCategories.includes(trimmed)) {
            appState.categories.allocCategories.push(trimmed);
          }
          item.category = trimmed;
        }
        saveState();
        renderAllocationTables();
      }, 50);
    } else {
      if (val) item.category = val;
      saveState();
      renderAllocationTables();
    }
  };

  select.onchange = (e) => commitChange(e.target.value);
  select.onblur = (e) => {
    if (select.value !== 'NEW_CUSTOM') {
      commitChange(select.value);
    }
  };

  setTimeout(() => select.focus(), 20);
};

function openAddIncomeModal() {
  const body = document.getElementById('modal-body');
  document.getElementById('modal-title').innerText = '수입 항목 추가';

  const recipientList = (appState.categories && appState.categories.recipients) ? appState.categories.recipients : ['계원', '동욱', '공통 / 가구'];
  const recipientOptions = recipientList.map(r => `<option value="${r}">${r}</option>`).join('');

  const itemCatList = (appState.categories && appState.categories.incomeItems) ? appState.categories.incomeItems : ['기본급 (급여)', '보너스 / 상여금', '인센티브', '부수입 / 알바', '투자수익 / 배당', '기타 수입'];
  const itemCatOptions = itemCatList.map(cat => `<option value="${cat}">${cat}</option>`).join('');

  body.innerHTML = `
    <div class="form-group">
      <label>수입 대상자 (카테고리)</label>
      <select id="modal-inc-person" class="form-select">
        ${recipientOptions}
      </select>
    </div>
    <div class="form-group">
      <label>수입 항목 카테고리 선택</label>
      <select id="modal-inc-cat-select" class="form-select mb-2">
        <option value="">-- 카테고리 선택 --</option>
        ${itemCatOptions}
        <option value="CUSTOM">직접 입력...</option>
      </select>
      <input type="text" id="modal-inc-title" class="form-control" placeholder="수입 항목명 (예: 기본급(급여), 성과급)">
    </div>
    <div class="form-group">
      <label>수입 / 월급 금액 (원)</label>
      <input type="text" id="modal-inc-amount" class="form-control" placeholder="0 (숫자 입력 시 , 자동 적용)">
    </div>
    <div class="form-group">
      <label>입금일자 / 월급날 (예: 25일, 27일, 10일)</label>
      <input type="text" id="modal-inc-day" class="form-control" placeholder="숫자만 써도 '일' 자동 추가 (예: 25, 28)">
    </div>
    <div class="form-group">
      <label>비고</label>
      <input type="text" id="modal-inc-note" class="form-control" placeholder="메모">
    </div>
  `;

  document.getElementById('item-modal').classList.add('active');
  attachIncomeModalFormatters();

  document.getElementById('modal-save-btn').onclick = () => {
    const person = document.getElementById('modal-inc-person').value;
    const title = document.getElementById('modal-inc-title').value.trim();
    const rawAmt = document.getElementById('modal-inc-amount').value.replace(/,/g, '');
    const amount = Number(rawAmt) || 0;
    const dayInputVal = document.getElementById('modal-inc-day').value.trim();
    const day = formatDayInput(dayInputVal || '25일');
    const note = document.getElementById('modal-inc-note').value.trim();

    if (!title) { alert('수입 항목명을 입력하세요.'); return; }

    appState.incomes.push({
      id: 'inc_' + Date.now(),
      person,
      title,
      amount,
      day,
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
  document.getElementById('modal-title').innerText = `${inc.person || '수입'} 수정`;

  const recipientList = (appState.categories && appState.categories.recipients) ? appState.categories.recipients : ['계원', '동욱', '공통 / 가구'];
  const recipientOptions = recipientList.map(r => `<option value="${r}" ${inc.person === r ? 'selected' : ''}>${r}</option>`).join('');

  const itemCatList = (appState.categories && appState.categories.incomeItems) ? appState.categories.incomeItems : ['기본급 (급여)', '보너스 / 상여금', '인센티브', '부수입 / 알바', '투자수익 / 배당', '기타 수입'];
  const itemCatOptions = itemCatList.map(cat => `<option value="${cat}" ${inc.title.includes(cat) ? 'selected' : ''}>${cat}</option>`).join('');

  const initialAmountFormatted = inc.amount ? Number(inc.amount).toLocaleString('ko-KR') : '';

  body.innerHTML = `
    <div class="form-group">
      <label>수입 대상자 (카테고리)</label>
      <select id="modal-inc-person" class="form-select">
        ${recipientOptions}
      </select>
    </div>
    <div class="form-group">
      <label>수입 항목 카테고리 선택</label>
      <select id="modal-inc-cat-select" class="form-select mb-2">
        <option value="">-- 카테고리 선택 --</option>
        ${itemCatOptions}
        <option value="CUSTOM">직접 입력...</option>
      </select>
      <input type="text" id="modal-inc-title" class="form-control" value="${inc.title}" placeholder="수입 항목명 입력">
    </div>
    <div class="form-group">
      <label>수입 / 월급 금액 (원) - 숫자 입력 시 , 자동 생성</label>
      <input type="text" id="modal-inc-amount" class="form-control" value="${initialAmountFormatted}" placeholder="0">
    </div>
    <div class="form-group">
      <label>입금일자 / 월급날 (숫자 입력 시 '일' 자동 추가)</label>
      <input type="text" id="modal-inc-day" class="form-control" value="${inc.day || '25일'}" placeholder="예: 25, 28">
    </div>
    <div class="form-group">
      <label>비고</label>
      <input type="text" id="modal-inc-note" class="form-control" value="${inc.note || ''}" placeholder="메모">
    </div>
  `;

  document.getElementById('item-modal').classList.add('active');
  attachIncomeModalFormatters();

  document.getElementById('modal-save-btn').onclick = () => {
    inc.person = document.getElementById('modal-inc-person').value;
    inc.title = document.getElementById('modal-inc-title').value.trim();
    const rawAmt = document.getElementById('modal-inc-amount').value.replace(/,/g, '');
    inc.amount = Number(rawAmt) || 0;
    const dayInputVal = document.getElementById('modal-inc-day').value.trim();
    inc.day = formatDayInput(dayInputVal || '25일');
    inc.note = document.getElementById('modal-inc-note').value.trim();

    closeModal();
    saveState();
    showToast('수입/월급 정보가 성공적으로 수정되었습니다.');
  };
}

function openAddAllocModal(personKey) {
  const body = document.getElementById('modal-body');
  const personName = personKey === 'gyewon' ? '계원' : '동욱';
  document.getElementById('modal-title').innerText = `${personName} 배분 항목 추가`;

  const allocCatList = (appState.categories && appState.categories.allocCategories) ? appState.categories.allocCategories : ['생활비', '저축/적금', '투자/연금', '대출금', '비상금/경조사'];
  const allocCatOptions = allocCatList.map(cat => `<option value="${cat}">${cat}</option>`).join('');

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
      <label>이체일자</label>
      <input type="text" id="modal-alloc-day" class="form-control" placeholder="예: 22일, 16일">
    </div>
    <div class="form-group">
      <label>이체방식</label>
      <input type="text" id="modal-alloc-method" class="form-control" placeholder="예: 자동이체, 오빠자동이체">
    </div>
    <div class="form-group">
      <label>카테고리</label>
      <select id="modal-alloc-category" class="form-select">
        ${allocCatOptions}
      </select>
    </div>
  `;

  document.getElementById('item-modal').classList.add('active');

  document.getElementById('modal-save-btn').onclick = () => {
    const name = document.getElementById('modal-alloc-name').value.trim();
    const amount = Number(document.getElementById('modal-alloc-amount').value) || 0;
    const bank = document.getElementById('modal-alloc-bank').value.trim();
    let day = document.getElementById('modal-alloc-day').value.trim();
    let method = document.getElementById('modal-alloc-method').value.trim();
    const category = document.getElementById('modal-alloc-category').value;

    if (!name) { alert('항목명을 입력하세요.'); return; }
    if (day) day = formatDayInput(day);

    const dayVal = day || '-';
    const methodVal = method || '-';
    const schedule = `${dayVal === '-' ? '' : dayVal} ${methodVal === '-' ? '' : methodVal}`.trim() || '-';

    appState.allocations[personKey].push({
      id: 'a_' + Date.now(),
      name,
      amount,
      bank,
      day: dayVal,
      method: methodVal,
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
    [...(appState.allocations?.gyewon || []), ...(appState.allocations?.dongwook || [])].forEach(item => {
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
