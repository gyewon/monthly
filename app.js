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
      { id: 'ag9', name: '대출/이자', amount: 1875232, bank: '고정지출', day: '22일', method: '자동이체', schedule: '22일 자동이체', category: '대출/이자' },
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
    { id: 'fe2_sagol', name: '곗돈 - 사골', amount: 20000, day: '2일', method: '현금', category: '곗돈', note: '토스고정지출 자동이체', actualMarch: 20000, isPaid: true, isAutoTransfer: true },
    { id: 'fe2_yeojin_dogyeong', name: '곗돈 - 여진도경', amount: 20000, day: '26일', method: '현금', category: '곗돈', note: '토스고정지출 자동이체', actualMarch: 20000, isPaid: true, isAutoTransfer: true },
    { id: 'fe2_jiyoung_soyeon', name: '곗돈 - 지영소연', amount: 30000, day: '26일', method: '현금', category: '곗돈', note: '토스고정지출 자동이체', actualMarch: 30000, isPaid: true, isAutoTransfer: true },
    { id: 'fe2_newgrina', name: '곗돈 - 뉴그리나', amount: 20000, day: '26일', method: '현금', category: '곗돈', note: '토스고정지출 자동이체', actualMarch: 20000, isPaid: true, isAutoTransfer: true },
    { id: 'fe2_family', name: '곗돈 - 가족', amount: 100000, day: '26일', method: '현금', category: '곗돈', note: '토스고정지출 자동이체', actualMarch: 100000, isPaid: true, isAutoTransfer: true },
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
  ],
  currentAssets: 0
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
  const fixedActualTotal = fixedBaseTotal; // Unified with amount as requested

  // Total Expenses
  const totalExpenses = totalAllocations;

  // Savings & Investments
  const catSavings = (appState.categories && appState.categories.allocCategories) ? appState.categories.allocCategories[1] : '저축/적금';
  const catInvest = (appState.categories && appState.categories.allocCategories) ? appState.categories.allocCategories[2] : '투자/연금';
  const catEmergency = (appState.categories && appState.categories.allocCategories) ? appState.categories.allocCategories[4] : '비상금/경조사';

  let totalSavings = 0;
  let pureSavings = 0;
  let pureInvestment = 0;
  let pureEmergency = 0;
  [...appState.allocations.gyewon, ...appState.allocations.dongwook].forEach(item => {
    if (item.category === catSavings) {
      pureSavings += Number(item.amount) || 0;
      totalSavings += Number(item.amount) || 0;
    } else if (item.category === catInvest) {
      pureInvestment += Number(item.amount) || 0;
      totalSavings += Number(item.amount) || 0;
    } else if (item.category === catEmergency) {
      pureEmergency += Number(item.amount) || 0;
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
    pureSavings,
    pureInvestment,
    pureEmergency,
    savingsRate,
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



  const filtered = appState.incomes || [];
  let totalIncomeAmount = 0;

  filtered.forEach(inc => {
    totalIncomeAmount += Number(inc.amount) || 0;
    let badgeClass = 'badge-info';
    let personLabel = inc.person || '기타';
    if (inc.person === 'gyewon' || inc.person === '계원') { badgeClass = 'badge-gyewon'; personLabel = '계원'; }
    else if (inc.person === 'dongwook' || inc.person === '동욱') { badgeClass = 'badge-dongwook'; personLabel = '동욱'; }
    else if (inc.person === 'common' || inc.person === '공통 / 가구') { badgeClass = 'badge-info'; personLabel = '공통/가구'; }

    const categoryName = inc.category || inc.title || '기본급 (급여)';

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="text-align:center;">
        <input type="checkbox" ${inc.isPaid ? 'checked' : ''} onchange="toggleIncomePaid('${inc.id}', this.checked)" style="transform: scale(1.2); cursor:pointer;">
      </td>
      <td><span class="badge ${badgeClass}">${personLabel}</span></td>
      <td><div class="editable-cell" title="클릭하여 카테고리 빠른 변경" onclick="editInlineCategory(this, '${inc.id}', event)"><span class="badge ${getCategoryBadgeClass(categoryName)}" style="font-weight:600; cursor:pointer;">${categoryName}</span></div></td>
      <td><div class="editable-cell cell-amount text-accent" title="클릭하여 월급/금액 수정" onclick="editInlineCell(this, 'incomes', '${inc.id}', 'amount', 'number')">${formatKRW(inc.amount)}</div></td>
      <td><div class="editable-cell text-success font-weight-bold" title="클릭하여 입금일 수정" onclick="editInlineCell(this, 'incomes', '${inc.id}', 'day')">${formatDayDisplay(inc.day)}</div></td>
      <td><div class="editable-cell" title="클릭하여 수정" onclick="editInlineCell(this, 'incomes', '${inc.id}', 'note')">${inc.note || '-'}</div></td>
      <td style="white-space:nowrap; text-align:right;">
        <button class="btn btn-outline-primary btn-sm" title="수입/월급 수정" onclick="openEditIncomeModal('${inc.id}')"><i class="fa-solid fa-pen-to-square"></i></button>
        <button class="btn btn-outline-danger btn-sm" title="삭제" onclick="deleteItem('incomes', '${inc.id}')"><i class="fa-solid fa-trash"></i></button>
      </td>
    `;
    tbody.appendChild(tr);
  });

  // Render Total Footer
  let tfoot = document.querySelector('#table-income-unified tfoot');
  if (!tfoot) {
    tfoot = document.createElement('tfoot');
    document.getElementById('table-income-unified').appendChild(tfoot);
  }
  tfoot.innerHTML = `
    <tr style="background: rgba(255, 255, 255, 0.05); font-weight: bold; border-top: 2px solid rgba(255,255,255,0.1);">
      <td colspan="3" style="text-align: center; color: var(--text-color);">합계</td>
      <td class="text-accent" style="font-size: 1.1em;">${formatKRW(totalIncomeAmount)}</td>
      <td colspan="3"></td>
    </tr>
  `;
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
let fixedCategoryChartInstance = null;

// DOM Loaded Initialization
document.addEventListener('DOMContentLoaded', () => {
  initApp();
});

function initApp() {
  setupEventListeners();
  
  const savedTab = localStorage.getItem('activeTab') || 'dashboard';
  const targetBtn = document.querySelector(`.nav-item[data-tab="${savedTab}"]`);
  if (targetBtn) {
    targetBtn.click();
  }

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
  else item.day = item.day.toString().replace(/일$/, '').trim();
  if (!item.method) item.method = '-';
  item.schedule = `${item.day === '-' ? '' : item.day + '일'} ${item.method === '-' ? '' : item.method}`.trim() || '-';
  return item;
}

function normalizeState(state) {
  if (!state) state = {};
  if (!state.categories) state.categories = {};
  if (!state.allocations) state.allocations = { gyewon: [], dongwook: [] };
  if (!state.allocations.gyewon) state.allocations.gyewon = [];
  if (!state.allocations.dongwook) state.allocations.dongwook = [];
  if (!state.fixedExpenses) state.fixedExpenses = [];
  if (!state.monthlyHistory) state.monthlyHistory = [];
  if (!state.incomes) state.incomes = [];
  if (!state.categories.allocCategories) {
    state.categories.allocCategories = ['생활비', '저축/적금', '투자/연금', '대출/이자', '비상금/경조사'];
  }
  if (!state.simulationCategories || state.simulationCategories.length === 0) {
    state.simulationCategories = [
      state.categories.allocCategories[1], 
      state.categories.allocCategories[2], 
      state.categories.allocCategories[4]
    ].filter(Boolean);
  }

  // Split the former combined 곗돈 row into the current transfer schedule.
  if (state && Array.isArray(state.fixedExpenses)) {
    const legacyIndex = state.fixedExpenses.findIndex(item => item && item.name === '곗돈');
    if (legacyIndex !== -1) {
      const legacy = state.fixedExpenses[legacyIndex];
      const common = {
        method: legacy.method || '현금',
        category: '곗돈',
        note: '토스고정지출 자동이체',
        isPaid: Boolean(legacy.isPaid),
        isAutoTransfer: true
      };
      const splitGyeExpenses = [
        { id: 'fe2_sagol', name: '곗돈 - 사골', amount: 20000, day: '2일', actualMarch: 20000, ...common },
        { id: 'fe2_yeojin_dogyeong', name: '곗돈 - 여진도경', amount: 20000, day: '26일', actualMarch: 20000, ...common },
        { id: 'fe2_jiyoung_soyeon', name: '곗돈 - 지영소연', amount: 30000, day: '26일', actualMarch: 30000, ...common },
        { id: 'fe2_newgrina', name: '곗돈 - 뉴그리나', amount: 20000, day: '26일', actualMarch: 20000, ...common },
        { id: 'fe2_family', name: '곗돈 - 가족', amount: 100000, day: '26일', actualMarch: 100000, ...common }
      ];
      state.fixedExpenses.splice(legacyIndex, 1, ...splitGyeExpenses);
    }
  }

  if (state && state.allocations) {
    if (Array.isArray(state.allocations.gyewon)) {
      state.allocations.gyewon.forEach(normalizeAllocationItem);
    }
    if (Array.isArray(state.allocations.dongwook)) {
      state.allocations.dongwook.forEach(normalizeAllocationItem);
    }
  }
  // Migration for 대출금 -> 대출/이자
  if (state && state.categories && state.categories.allocCategories) {
    state.categories.allocCategories = state.categories.allocCategories.map(c => c === '대출금' ? '대출/이자' : c);
  }
  if (state && state.allocations) {
    if (state.allocations.gyewon) {
      state.allocations.gyewon.forEach(item => { if (item.category === '대출금') item.category = '대출/이자'; });
    }
    if (state.allocations.dongwook) {
      state.allocations.dongwook.forEach(item => { if (item.category === '대출금') item.category = '대출/이자'; });
    }
  }

  // Sync mismatched categories caused by earlier edits
  if (state && state.categories && state.categories.allocCategories && state.allocations) {
    const cats = state.categories.allocCategories;
    ['gyewon', 'dongwook'].forEach(p => {
      if (state.allocations[p]) {
        state.allocations[p].forEach(item => {
          if (item.category && !cats.includes(item.category)) {
            const match = cats.find(c => item.category.includes(c) || c.includes(item.category));
            if (match) item.category = match;
          }
        });
      }
    });
  }

  // Migration for bank 계원토스 -> 토스고정비
  if (state && Array.isArray(state.fixedExpenses)) {
    state.fixedExpenses.forEach(item => {
      if (item.bank === '계원토스') item.bank = '토스고정비';
    });
  }

  if (state.currentAssets === undefined) {
    state.currentAssets = 0;
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
        data: appState
      });

    if (error) {
      console.warn('Supabase save notice:', error.message);
      updateSupabaseBadge(false, '저장 실패');
    } else {
      updateSupabaseBadge(true);
    }
  } catch (err) {
    console.error('Supabase save error:', err);
    updateSupabaseBadge(false, '연결 오류');
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

      localStorage.setItem('activeTab', targetTab);

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
      if (targetTab === 'history') {
        renderSavingsProjection();
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
    btnCatManage.addEventListener('click', () => openCategoryManagerModal('income'));
  }
  const btnAllocCatManage = document.getElementById('btn-manage-alloc-categories');
  if (btnAllocCatManage) {
    btnAllocCatManage.addEventListener('click', () => openCategoryManagerModal('alloc'));
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
  updateMethodFilterDropdown();

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
  
  remainElem.classList.remove('text-danger', 'zero');
  if (calcs.remainingBalance < 0) {
    remainElem.classList.add('text-danger');
  } else if (calcs.remainingBalance === 0) {
    remainElem.classList.add('zero');
  }

  document.getElementById('kpi-remain-gyewon').innerText = formatKRW(calcs.remainGyewon);
  document.getElementById('kpi-remain-dongwook').innerText = formatKRW(calcs.remainDongwook);

  // Generate Income Breakdown by Category
  const incomeByCategory = {};
  (appState.incomes || []).forEach(inc => {
    const cat = inc.category || inc.title || '기본급 (급여)';
    const amt = Number(inc.amount) || 0;
    incomeByCategory[cat] = (incomeByCategory[cat] || 0) + amt;
  });

  const breakdownHTML = Object.entries(incomeByCategory)
    .sort((a, b) => b[1] - a[1]) // Sort descending by amount
    .map(([cat, amt]) => `<span class="badge ${getCategoryBadgeClass(cat)}">${cat} ${formatCompactKRW(amt)}</span>`)
    .join(' ');
  
  const breakdownElem = document.getElementById('kpi-income-breakdown');
  if (breakdownElem) {
    breakdownElem.innerHTML = breakdownHTML;
  }
  const fixedElem = document.getElementById('kpi-fixed-total');
  if (fixedElem) {
    fixedElem.innerText = formatKRW(calcs.fixedExpenseActualTotal) + '원';
  }
  const variableTotal = calcs.totalExpenses - calcs.fixedExpenseActualTotal;
  const varElem = document.getElementById('kpi-variable-total');
  if (varElem) {
    varElem.innerText = formatKRW(variableTotal) + '원';
  }
  
  const pSavingsElem = document.getElementById('kpi-pure-savings');
  const pInvestElem = document.getElementById('kpi-pure-investment');
  const pEmergencyElem = document.getElementById('kpi-pure-emergency');
  const pTotalElem = document.getElementById('kpi-total-savings');
  const pLabelSavings = document.getElementById('kpi-label-savings');
  const pLabelInvest = document.getElementById('kpi-label-invest');
  const pLabelEmergency = document.getElementById('kpi-label-emergency');

  if (pSavingsElem) pSavingsElem.innerText = formatKRW(calcs.pureSavings);
  if (pInvestElem) pInvestElem.innerText = formatKRW(calcs.pureInvestment);
  if (pEmergencyElem) pEmergencyElem.innerText = formatKRW(calcs.pureEmergency);
  if (pTotalElem) pTotalElem.innerText = formatKRW(calcs.totalSavings);

  if (pLabelSavings) pLabelSavings.innerText = (appState.categories && appState.categories.allocCategories) ? appState.categories.allocCategories[1] : '저축';
  if (pLabelInvest) pLabelInvest.innerText = (appState.categories && appState.categories.allocCategories) ? appState.categories.allocCategories[2] : '투자';
  if (pLabelEmergency) pLabelEmergency.innerText = (appState.categories && appState.categories.allocCategories) ? appState.categories.allocCategories[4] : '비상금';
  
  document.getElementById('kpi-savings-rate').innerHTML = `<span style="color: var(--accent-danger);">${calcs.savingsRate}</span>%`;

  // Render Tables
  renderIncomeTables();
  renderAllocationTables();
  renderFixedExpensesTable();
  renderTimeline();
  renderPaymentMethodSummary();
  renderFixedCategorySummary();
  renderCategoryAllocationSummary();

  // Render Charts
  renderCharts();
  renderSavingsProjection();
}

function getCategoryBadgeClass(category) {
  if (!category) return 'badge-secondary';
  if (category.includes('급여') || category.includes('기본급') || category.includes('월급')) return 'badge-salary';
  if (category.includes('보너스') || category.includes('상여')) return 'badge-warning';
  if (category.includes('인센티브')) return 'badge-success';
  if (category.includes('부수입') || category.includes('알바')) return 'badge-info';
  if (category.includes('투자') || category.includes('연금') || category.includes('배당')) return 'badge-primary';
  if (category.includes('기타수입') || category.includes('기타 수입')) return 'badge-total';
  
  if (category.includes('생활비')) return 'badge-info';
  if (category.includes('저축') || category.includes('적금')) return 'badge-success';
  if (category.includes('주거') || category.includes('관리비')) return 'badge-gyewon';
  if (category.includes('구독')) return 'badge-dongwook';
  if (category.includes('잔액')) return 'badge-remain';
  if (category.includes('대출')) return 'badge-danger';
  if (category.includes('비상금') || category.includes('경조사')) return 'badge-warning';
  return 'badge-secondary';
}

function getPaymentMethodHTML(method) {
  if (!method || method === '-') return '<span class="text-muted">-</span>';
  
  let company = method;
  if (appState.categories?.paymentMethodCompanies?.[method]) {
    company = appState.categories.paymentMethodCompanies[method];
  }

  let logoUrl = '';
  let icon = 'fa-solid fa-credit-card';
  let color = '#94a3b8'; // gray
  let bg = '#fff';

  if (company.includes('현금')) { 
    icon = 'fa-solid fa-money-bill-wave'; color = '#10b981'; 
  } else if (company.includes('하나')) { 
    logoUrl = 'https://www.google.com/s2/favicons?domain=hanabank.com&sz=64'; 
  } else if (company.includes('우리')) { 
    logoUrl = 'https://www.google.com/s2/favicons?domain=wooricard.com&sz=64'; 
  } else if (company.includes('신한')) { 
    logoUrl = 'https://www.google.com/s2/favicons?domain=shinhancard.com&sz=64'; 
  } else if (company.includes('삼성')) { 
    logoUrl = 'https://www.google.com/s2/favicons?domain=samsung.com&sz=64'; 
  } else if (company.includes('토스')) { 
    logoUrl = 'https://www.google.com/s2/favicons?domain=toss.im&sz=64'; 
  } else if (company.includes('국민') || company.includes('KB')) {
    logoUrl = 'https://www.google.com/s2/favicons?domain=kbstar.com&sz=64';
  } else if (company.includes('현대')) {
    logoUrl = 'https://www.google.com/s2/favicons?domain=hyundaicard.com&sz=64';
  } else if (company.includes('롯데')) {
    logoUrl = 'https://www.google.com/s2/favicons?domain=lottecard.co.kr&sz=64';
  } else if (company.includes('농협') || company.includes('NH')) {
    logoUrl = 'https://www.google.com/s2/favicons?domain=nonghyup.com&sz=64';
  } else if (company.includes('카카오')) {
    logoUrl = 'https://www.google.com/s2/favicons?domain=kakaobank.com&sz=64';
    bg = '#fee500';
  } else if (company.includes('자동이체') || company.includes('계좌')) { 
    icon = 'fa-solid fa-building-columns'; color = '#c084fc'; 
  }
  
  const iconElement = logoUrl 
    ? `<img src="${logoUrl}" alt="${method}" style="width: 18px; height: 18px; border-radius: 50%; object-fit: cover; background: ${bg}; vertical-align: middle; box-shadow: 0 0 2px rgba(255,255,255,0.2);">` 
    : `<i class="${icon}" style="color:${color}; font-size:16px; width:18px; text-align:center;"></i>`;

  return `<span style="display:inline-flex; align-items:center; gap:6px; font-weight:600; font-size:14.5px; color:#f8fafc; cursor:pointer; padding:5px 8px; border-radius:6px; background-color:rgba(255,255,255,0.05); border:1px solid rgba(255,255,255,0.1);">${iconElement} ${method}</span>`;
}

function getFixedCategoryHTML(category) {
  if (!category || category === '-') return '<span class="text-muted">-</span>';
  let emoji = '\ud83c\udff7\ufe0f';
  
  if (category.includes('\uc6a9\ub3c8')) emoji = '\ud83d\udc5b';
  else if (category.includes('\uacb0\ub3c8') || category.includes('\uacb0')) emoji = '\ud83d\udc65';
  else if (category.includes('\uad6c\ub3c5')) emoji = '\u25b6\ufe0f';
  else if (category.includes('\ubcf4\ud5d8')) emoji = '\ud83d\udee1\ufe0f';
  else if (category.includes('\ud578\ub4dc\ud3f0') || category.includes('\ud1b5\uc2e0')) emoji = '\ud83d\udcf1';
  else if (category.includes('\uc8fc\uac70') || category.includes('\uad00\ub9ac\ube44')) emoji = '\ud83c\udfe0';
  else if (category.includes('\uad50\ud1b5')) emoji = '\ud83d\ude8c';
  else if (category.includes('현금') || category.includes('\ud604\uae08')) emoji = '💰';

  return `<span style="display:inline-flex; align-items:center; gap:6px; font-weight:600; font-size:13px; color:#f8fafc; background-color:rgba(255,255,255,0.08); padding:5px 12px; border-radius:12px; cursor:pointer;"><span style="font-size:17px;">${emoji}</span> ${category}</span>`;
}
window.getAllocCategoryBadgeClass = getCategoryBadgeClass; // For backward compatibility with any other calls

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
    gRemainElem.classList.remove('negative', 'zero');
    if (calcs.remainGyewon < 0) gRemainElem.classList.add('negative');
    else if (calcs.remainGyewon === 0) gRemainElem.classList.add('zero');
  }

  // Update Dongwook Summary Bar
  const dIncElem = document.getElementById('d-summary-income');
  const dAllocElem = document.getElementById('d-summary-alloc');
  const dRemainElem = document.getElementById('d-summary-remain');
  if (dIncElem) dIncElem.innerText = formatKRW(calcs.incomeDongwook);
  if (dAllocElem) dAllocElem.innerText = formatKRW(calcs.allocDongwookSum);
  if (dRemainElem) {
    dRemainElem.innerText = formatKRW(calcs.remainDongwook);
    dRemainElem.classList.remove('negative', 'zero');
    if (calcs.remainDongwook < 0) dRemainElem.classList.add('negative');
    else if (calcs.remainDongwook === 0) dRemainElem.classList.add('zero');
  }

  // Target Filter Visibility Control (Removed filter hiding per user request)
  const gyewonCard = document.getElementById('alloc-gyewon-card');
  const dongwookCard = document.getElementById('alloc-dongwook-card');
  const allocContainer = document.getElementById('allocations-container');

  if (gyewonCard && dongwookCard) {
    gyewonCard.style.display = '';
    dongwookCard.style.display = '';
    if (allocContainer) allocContainer.style.gridTemplateColumns = '';
  }


  // Populate Gyewon Allocation Rows
  const tbodyG = document.querySelector('#table-alloc-gyewon tbody');
  if (tbodyG) {
    tbodyG.innerHTML = '';
    (appState.allocations.gyewon || []).forEach(item => {
      const catBadge = getAllocCategoryBadgeClass(item.category);
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td style="text-align:center;">
          <input type="checkbox" ${item.isPaid ? 'checked' : ''} onchange="toggleAllocPaid('gyewon', '${item.id}', this.checked)" style="transform: scale(1.2); cursor:pointer;">
        </td>
        <td><div class="editable-cell" onclick="editInlineCell(this, 'allocations.gyewon', '${item.id}', 'name')">${item.name}</div></td>
        <td><div class="editable-cell cell-amount" onclick="editInlineCell(this, 'allocations.gyewon', '${item.id}', 'amount', 'number')">${formatKRW(item.amount)}</div></td>
        <td><div class="editable-cell" onclick="editInlineCell(this, 'allocations.gyewon', '${item.id}', 'bank')">${item.bank || '-'}</div></td>
        <td><div class="editable-cell" onclick="editInlineCell(this, 'allocations.gyewon', '${item.id}', 'deposit')">${item.deposit || '-'}</div></td>
        <td><div class="editable-cell text-primary font-weight-bold" onclick="editInlineCell(this, 'allocations.gyewon', '${item.id}', 'day')">${formatDayDisplay(item.day)}</div></td>
        <td style="text-align:center;">
          <input type="checkbox" class="form-check-input" style="width:1.2rem; height:1.2rem; cursor:pointer;" 
                 ${(item.isAutoTransfer === true || (item.isAutoTransfer === undefined && item.method && item.method.includes('자동'))) ? 'checked' : ''} 
                 onclick="toggleAllocAutoTransfer('allocations.gyewon', '${item.id}')">
        </td>
        <td><div class="editable-cell" title="클릭하여 카테고리 빠른 변경" onclick="editInlineAllocCategory(this, 'allocations.gyewon', '${item.id}', event)"><span class="badge ${catBadge}" style="font-weight:600; cursor:pointer;">${item.category || '기타'}</span></div></td>
        <td><div class="editable-cell" onclick="editInlineCell(this, 'allocations.gyewon', '${item.id}', 'note')">${item.note || '-'}</div></td>
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
        <td style="text-align:center;">
          <input type="checkbox" ${item.isPaid ? 'checked' : ''} onchange="toggleAllocPaid('dongwook', '${item.id}', this.checked)" style="transform: scale(1.2); cursor:pointer;">
        </td>
        <td><div class="editable-cell" onclick="editInlineCell(this, 'allocations.dongwook', '${item.id}', 'name')">${item.name}</div></td>
        <td><div class="editable-cell cell-amount" onclick="editInlineCell(this, 'allocations.dongwook', '${item.id}', 'amount', 'number')">${formatKRW(item.amount)}</div></td>
        <td><div class="editable-cell" onclick="editInlineCell(this, 'allocations.dongwook', '${item.id}', 'bank')">${item.bank || '-'}</div></td>
        <td><div class="editable-cell" onclick="editInlineCell(this, 'allocations.dongwook', '${item.id}', 'deposit')">${item.deposit || '-'}</div></td>
        <td><div class="editable-cell text-primary font-weight-bold" onclick="editInlineCell(this, 'allocations.dongwook', '${item.id}', 'day')">${formatDayDisplay(item.day)}</div></td>
        <td style="text-align:center;">
          <input type="checkbox" class="form-check-input" style="width:1.2rem; height:1.2rem; cursor:pointer;" 
                 ${(item.isAutoTransfer === true || (item.isAutoTransfer === undefined && item.method && item.method.includes('자동'))) ? 'checked' : ''} 
                 onclick="toggleAllocAutoTransfer('allocations.dongwook', '${item.id}')">
        </td>
        <td><div class="editable-cell" title="클릭하여 카테고리 빠른 변경" onclick="editInlineAllocCategory(this, 'allocations.dongwook', '${item.id}', event)"><span class="badge ${catBadge}" style="font-weight:600; cursor:pointer;">${item.category || '기타'}</span></div></td>
        <td><div class="editable-cell" onclick="editInlineCell(this, 'allocations.dongwook', '${item.id}', 'note')">${item.note || '-'}</div></td>
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
  let paidSum = 0;
  let catData = {};
  let methodData = {};

  filtered.forEach(item => {
    const amt = Number(item.amount) || 0;
    baseSum += amt;
    actualSum += Number(item.actualMarch) || 0;
    if (item.isPaid) {
      paidCount++;
      paidSum += amt;
      const c = item.category || '기타';
      
      // Use actual method name as requested
      let m = item.method || '기타';
      m = m.trim();
      let majorCat = getPMType(m);
      
      catData[c] = (catData[c] || 0) + amt;
      methodData[majorCat] = (methodData[majorCat] || 0) + amt;
    }

    const methodHTML = getPaymentMethodHTML(item.method);
    const fCatHTML = getFixedCategoryHTML(item.category);
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="text-align:center;">
        <input type="checkbox" ${item.isPaid ? 'checked' : ''} onchange="toggleFixedPaid('${item.id}', this.checked)" style="transform: scale(1.2); cursor:pointer;">
      </td>
      <td><div class="editable-cell" onclick="editInlineCell(this, 'fixedExpenses', '${item.id}', 'name')">${item.name}</div></td>
      <td><div class="editable-cell cell-amount" onclick="editInlineCell(this, 'fixedExpenses', '${item.id}', 'amount', 'number')">${formatKRW(item.amount)}</div></td>
      <td><div class="editable-cell" onclick="editInlineCell(this, 'fixedExpenses', '${item.id}', 'bank')">${item.bank || '-'}</div></td>
      <td><div class="editable-cell" onclick="editInlineCell(this, 'fixedExpenses', '${item.id}', 'deposit')">${item.deposit || '-'}</div></td>
      <td><div class="editable-cell text-primary font-weight-bold" onclick="editInlineCell(this, 'fixedExpenses', '${item.id}', 'day')">${formatDayDisplay(item.day)}</div></td>
      <td style="text-align:center;">
        <input type="checkbox" class="form-check-input" style="width:1.2rem; height:1.2rem; cursor:pointer;" 
               ${(item.isAutoTransfer === true || (item.isAutoTransfer === undefined && item.note && item.note.includes('자동'))) ? 'checked' : ''} 
               onclick="toggleFixedAutoTransfer('${item.id}')">
      </td>
      <td onclick="editFixedExpenseMethod(this, '${item.id}')" style="cursor:pointer; vertical-align:middle;">${methodHTML}</td>
      <td onclick="editFixedExpenseCategory(this, '${item.id}')" style="cursor:pointer; vertical-align:middle;">${fCatHTML}</td>
      <td><div class="editable-cell" onclick="editInlineCell(this, 'fixedExpenses', '${item.id}', 'note')">${item.note || '-'}</div></td>
      <td><button class="btn btn-outline-danger btn-sm" onclick="deleteItem('fixedExpenses', '${item.id}')"><i class="fa-solid fa-trash"></i></button></td>
    `;
    tbody.appendChild(tr);
  });

  document.getElementById('fixed-stat-base').innerText = formatKRW(paidSum);

  // Calculate 생활비 budget from allocations
  const allAllocs = [...(appState.allocations?.gyewon || []), ...(appState.allocations?.dongwook || [])];
  const livingBudget = allAllocs
    .filter(item => (item.category || '').includes('생활비'))
    .reduce((sum, item) => sum + (Number(item.amount) || 0), 0);

  const FLEXIBLE_LIVING_BUDGET = appState.flexibleLivingBudget !== undefined ? Number(appState.flexibleLivingBudget) : 1000000;
  const flexBudgetElem = document.getElementById('fixed-stat-flexible-budget');
  if (flexBudgetElem) flexBudgetElem.innerText = formatKRW(FLEXIBLE_LIVING_BUDGET);
  const fixedExpenseBudget = livingBudget - FLEXIBLE_LIVING_BUDGET;
  const fixedRemaining = fixedExpenseBudget - paidSum;

  const livingBudgetElem = document.getElementById('fixed-stat-living-budget');
  const fixedBudgetElem = document.getElementById('fixed-stat-fixed-budget');
  const livingRemainElem = document.getElementById('fixed-stat-remaining');
  const dashLivingRemainElem = document.getElementById('dash-fixed-stat-remaining');
  
  if (livingBudgetElem) livingBudgetElem.innerText = formatKRW(livingBudget);
  if (fixedBudgetElem) fixedBudgetElem.innerText = formatKRW(fixedExpenseBudget);
  
  if (livingRemainElem) {
    livingRemainElem.innerText = formatKRW(fixedRemaining);
    livingRemainElem.className = `value ${fixedRemaining < 0 ? 'text-danger' : 'text-success'}`;
  }
  
  if (dashLivingRemainElem) {
    dashLivingRemainElem.innerText = formatKRW(fixedRemaining);
    dashLivingRemainElem.style.color = fixedRemaining < 0 ? 'var(--danger-color)' : 'var(--success-color)';
  }

  const totalCount = filtered.length;
  const pct = totalCount > 0 ? Math.round((paidCount / totalCount) * 100) : 0;
  const completionElem = document.getElementById('fixed-stat-completion');
  if (completionElem) {
    completionElem.innerHTML = `${paidCount} / ${totalCount} (<strong style="font-weight:900;">${pct}%</strong>)`;
  }
  
  const dateReqs = {};
  filtered.forEach(item => {
    if (!item.isPaid) return; // [표시/완료] 체크된 항목만 집계
    let day = item.day || '수기/미지정';
    if (!day.includes('일') && day !== '수기' && day !== '수기/미지정') day += '일';
    if (!dateReqs[day]) dateReqs[day] = { paid: 0 };
    dateReqs[day].paid += (Number(item.amount) || 0);
  });

  const reqContainer = document.getElementById('fixed-date-requirements');
  if (reqContainer) {
    reqContainer.innerHTML = '';
    const sortedDays = Object.keys(dateReqs).sort((a, b) => (parseInt(a) || 99) - (parseInt(b) || 99));
    
    if (sortedDays.length === 0) {
      reqContainer.innerHTML = '<div class="text-muted" style="text-align:center; margin-top:40px;">완료된 결제 내역이 없습니다.</div>';
    } else {
      let runningBalance = fixedExpenseBudget; // 전체 예산에서 시작
      sortedDays.forEach(d => {
        const stats = dateReqs[d];
        runningBalance -= stats.paid; // 완료된 금액 차감
        
        const row = document.createElement('div');
        row.style.display = 'flex';
        row.style.justifyContent = 'space-between';
        row.style.alignItems = 'center';
        row.style.padding = '8px 0';
        row.style.borderBottom = '1px solid rgba(255,255,255,0.05)';
        row.innerHTML = `
          <span style="font-weight:600;">${d}</span> 
          <div style="text-align:right; line-height:1.2;">
            <div style="font-size:0.75rem; color:var(--text-muted); margin-bottom:2px;">
              실제 출금: ${formatKRW(stats.paid)}
            </div>
            <strong class="${runningBalance < 0 ? 'text-danger' : 'text-primary'}">남은 예산: ${formatKRW(runningBalance)}</strong>
          </div>
        `;
        reqContainer.appendChild(row);
      });
    }
  }

  drawFixedCharts(catData, methodData, fixedRemaining);
}

let fixedCategoryChartInst = null;
let fixedMethodChartInst = null;

function drawFixedCharts(catData, methodData, fixedRemaining) {
  const catCtx = document.getElementById('fixedTabCategoryChart');
  const methodCtx = document.getElementById('fixedTabMethodChart');
  if (!catCtx || !methodCtx) return;

  if (fixedCategoryChartInst) fixedCategoryChartInst.destroy();
  if (fixedMethodChartInst) fixedMethodChartInst.destroy();

  const generateChartLabels = (chart) => {
    const data = chart.data;
    if (data.labels.length && data.datasets.length) {
      const dataset = data.datasets[0];
      const total = dataset.data.reduce((a, b) => a + b, 0);
      return data.labels.map((label, i) => {
        const value = dataset.data[i];
        const pct = total > 0 ? Math.round((value / total) * 100) : 0;
        let formattedVal = value.toLocaleString();
        const isHidden = !chart.getDataVisibility(i);
        
        const map = {'0':'𝟬','1':'𝟭','2':'𝟮','3':'𝟯','4':'𝟰','5':'𝟱','6':'𝟲','7':'𝟳','8':'𝟴','9':'𝟵','%':'%'};
        const boldPct = `${pct}%`.split('').map(c => map[c] || c).join('');
        
        return {
          text: `${label} ${boldPct} (${formattedVal}원)`,
          fillStyle: dataset.backgroundColor[i],
          hidden: isHidden,
          index: i,
          fontColor: isHidden ? '#94a3b8' : '#f8fafc',
          textDecoration: isHidden ? 'line-through' : '',
          lineWidth: 0
        };
      });
    }
    return [];
  };

  const pieOptions = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { 
        position: 'right', 
        labels: { 
          boxWidth: 12, 
          font: { size: 14, weight: 'normal' },
          color: '#f8fafc',
          generateLabels: generateChartLabels
        } 
      },
      tooltip: {
        callbacks: {
          label: (ctx) => ` ${ctx.label}: ${formatKRW(ctx.raw)}`
        }
      }
    }
  };

  const catBaseColors = [
    '#3b82f6', '#10b981', '#f59e0b', '#ef4444', '#8b5cf6', 
    '#06b6d4', '#f43f5e', '#14b8a6', '#6366f1', '#f97316',
    '#ec4899', '#84cc16'
  ];

  const methodBaseColors = [
    '#8b5cf6', '#ec4899', '#14b8a6', '#f97316', '#0ea5e9',
    '#84cc16', '#ef4444', '#10b981', '#eab308', '#6366f1'
  ];

  // Sort data descending by amount
  const sortedCatLabels = Object.keys(catData).sort((a, b) => catData[b] - catData[a]);
  const sortedCatValues = sortedCatLabels.map(k => catData[k]);
  
  if (fixedRemaining > 0) {
    sortedCatLabels.push('미사용금액');
    sortedCatValues.push(fixedRemaining);
  }

  const catBgColors = sortedCatLabels.map((label, i) => label === '미사용금액' ? '#475569' : catBaseColors[i % catBaseColors.length]);

  fixedCategoryChartInst = new Chart(catCtx, {
    type: 'doughnut',
    data: {
      labels: sortedCatLabels,
      datasets: [{
        data: sortedCatValues,
        backgroundColor: catBgColors,
        borderWidth: 0
      }]
    },
    options: pieOptions
  });

  // Sort data descending by amount
  const sortedMethodLabels = Object.keys(methodData).sort((a, b) => methodData[b] - methodData[a]);
  const sortedMethodValues = sortedMethodLabels.map(k => methodData[k]);

  if (fixedRemaining > 0) {
    sortedMethodLabels.push('미사용금액');
    sortedMethodValues.push(fixedRemaining);
  }

  const methodBgColors = sortedMethodLabels.map((label, i) => label === '미사용금액' ? '#475569' : methodBaseColors[i % methodBaseColors.length]);

  fixedMethodChartInst = new Chart(methodCtx, {
    type: 'doughnut',
    data: {
      labels: sortedMethodLabels,
      datasets: [{
        data: sortedMethodValues,
        backgroundColor: methodBgColors,
        borderWidth: 0
      }]
    },
    options: pieOptions
  });
}

// Render Master Timeline / Deposit & Expense Schedule
function renderTimeline() {
  const container = document.getElementById('schedule-timeline');
  container.innerHTML = '';

  const groups = {};

  // 1. Incomes (입금/월급)
  (appState.incomes || []).forEach(inc => {
    if (!inc.isPaid) return;
    let rawDay = inc.day || '미지정';
    let dayKey = rawDay.includes('일') ? rawDay : (rawDay + '일');
    if (!groups[dayKey]) groups[dayKey] = [];
    
    // Fix: person string is already Korean in some datasets, handle both cases
    let personName = inc.person;
    if (personName === 'gyewon') personName = '계원';
    if (personName === 'dongwook') personName = '동욱';

    groups[dayKey].push({
      type: 'INCOME',
      name: `${personName} ${inc.title}`,
      method: '급여/입금',
      amount: Number(inc.amount) || 0,
      category: inc.title,
      isPaid: !!inc.isPaid
    });
  });

  // 2. Allocations (Gyewon)
  (appState.allocations?.gyewon || []).forEach(item => {
    if (!item.isPaid) return;
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
      amount: Number(item.amount) || 0,
      category: item.category,
      isPaid: !!item.isPaid
    });
  });

  // 3. Allocations (Dongwook)
  (appState.allocations?.dongwook || []).forEach(item => {
    if (!item.isPaid) return;
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
      amount: Number(item.amount) || 0,
      category: item.category,
      isPaid: !!item.isPaid
    });
  });

  // 4. Fixed Expenses (고정지출 이체)
  (appState.fixedExpenses || []).forEach(item => {
    if (!item.isPaid) return;
    let rawDay = item.day || '수기/기타';
    let dayKey = (rawDay.includes('일') || rawDay === '수기' || rawDay.includes('수기')) ? rawDay : (rawDay ? rawDay + '일' : '수기/기타');
    if (!groups[dayKey]) groups[dayKey] = [];
    groups[dayKey].push({
      type: 'FIXED',
      name: item.name,
      method: item.method || '결제수단',
      amount: Number(item.amount) || 0,
      category: item.category,
      isPaid: !!item.isPaid
    });
  });

  // Sort keys
  const sortedKeys = Object.keys(groups).sort((a, b) => {
    const numA = parseInt(a) || 99;
    const numB = parseInt(b) || 99;
    return numA - numB;
  });

  if (sortedKeys.length === 0) {
    container.innerHTML = `
      <div style="text-align:center; padding: 40px 20px; color: #64748b;">
        <i class="fa-regular fa-calendar-xmark fa-3x" style="margin-bottom:15px; opacity:0.5;"></i>
        <p style="font-size:15px; margin:0;">표시할 일정이 없습니다.</p>
        <p style="font-size:13px; margin-top:5px; opacity:0.8;">각 탭에서 <strong>[표시]</strong>를 체크한 항목만 이곳에 나타납니다.</p>
      </div>
    `;
    return;
  }

  sortedKeys.forEach(day => {
    const items = groups[day];

    items.sort((a, b) => {
      if (a.type === 'INCOME' && b.type !== 'INCOME') return -1;
      if (a.type !== 'INCOME' && b.type === 'INCOME') return 1;
      const catA = a.category || '';
      const catB = b.category || '';
      return catA.localeCompare(catB, 'ko-KR');
    });

    const totalIncomeDay = items.filter(i => i.type === 'INCOME').reduce((acc, c) => acc + c.amount, 0);
    const totalOutDay = items.filter(i => i.type !== 'INCOME').reduce((acc, c) => acc + c.amount, 0);

    const card = document.createElement('div');
    card.className = 'timeline-card';
    card.innerHTML = `
      <div class="timeline-card-header">
        <span class="timeline-date">${day} 일정</span>
        <span class="timeline-count">
          ${totalIncomeDay > 0 ? `<span class="badge badge-salary">+${formatKRW(totalIncomeDay)}</span> ` : ''}
          ${totalOutDay > 0 ? `<span class="badge badge-remain">-${formatKRW(totalOutDay)}</span>` : ''}
        </span>
      </div>
      <ul class="timeline-item-list">
        ${items.map(i => {
          let badgeClass = 'badge-secondary';
          let badgeText = '지출';
          
          if (i.type === 'INCOME') {
            badgeClass = 'badge-salary';
            badgeText = '수입';
          } else if (i.type === 'FIXED') {
            badgeClass = 'badge-danger';
            badgeText = '고정비';
          } else if (i.type === 'ALLOCATION') {
            badgeClass = getCategoryBadgeClass(i.category);
            if ((i.category || '').includes('저축') || (i.category || '').includes('적금')) badgeText = '저축';
            else if ((i.category || '').includes('투자') || (i.category || '').includes('연금')) badgeText = '투자';
            else if ((i.category || '').includes('생활비')) badgeText = '생활비';
            else if ((i.category || '').includes('비상금') || (i.category || '').includes('경조사')) badgeText = '비상금';
            else badgeText = '지출';
          }

          return `
            <li>
              <span><span class="badge ${badgeClass}">${badgeText}</span> ${i.name} <small class="text-muted">(${i.method})</small></span>
              <strong class="${i.type === 'INCOME' ? 'text-success' : 'text-danger'}">${i.type === 'INCOME' ? '+' : '-'}${formatKRW(i.amount)}</strong>
            </li>
          `;
        }).join('')}
      </ul>
    `;
    container.appendChild(card);
  });
}

// Helper to get PM Type
function getPMType(pm) {
  if (!pm) return '미분류';
  const pmStr = String(pm);
  
  if (!appState.categories.paymentMethodTypes) {
    appState.categories.paymentMethodTypes = {};
  }
  if (appState.categories.paymentMethodTypes[pmStr]) {
    const type = appState.categories.paymentMethodTypes[pmStr];
    if (type === 'card') return '카드';
    if (type === 'cash') return '현금/이체';
    return type;
  }
  return pmStr.includes('카드') ? '카드' : '현금/이체';
}

// Render Payment Method Summary Cards
function renderPaymentMethodSummary() {
  const container = document.getElementById('payment-method-summary');
  container.innerHTML = '';

  const pmMap = {};
  const groupMap = {};
  let totalFixed = 0;

  appState.fixedExpenses.forEach(item => {
    let method = item.method || '기타';
    method = method.trim();
    const amt = Number(item.amount) || 0;
    
    pmMap[method] = (pmMap[method] || 0) + amt;
    totalFixed += amt;
    
    const major = getPMType(method);
    groupMap[major] = (groupMap[major] || 0) + amt;
  });

  if (totalFixed > 0) {
    const totalWrap = document.createElement('div');
    totalWrap.style.display = 'flex';
    totalWrap.style.gap = '10px';
    totalWrap.style.flexWrap = 'wrap';
    totalWrap.style.marginBottom = '10px';
    totalWrap.style.width = '100%';

    let html = `
      <div class="pm-card pm-card-total" style="border: 1px solid var(--accent-total); background-color: rgba(139, 92, 246, 0.05); flex: 1; min-width: 150px;">
        <div class="pm-name" style="margin-bottom:8px;"><span class="badge badge-total" style="font-weight: bold; padding: 5px 10px; font-size:13px;">총 결제액</span></div>
        <div class="pm-amount" style="color: var(--text-main); font-weight: 700;">${formatKRW(totalFixed)}</div>
      </div>
    `;

    const colors = ['#10B981', '#3B82F6', '#F59E0B', '#EF4444', '#8B5CF6', '#EC4899'];
    let cIdx = 0;

    Object.entries(groupMap).sort((a,b)=>b[1]-a[1]).forEach(([major, amt]) => {
      const color = colors[cIdx % colors.length];
      html += `
        <div class="pm-card pm-card-total" style="border: 1px solid ${color}; background-color: ${color}10; flex: 1; min-width: 150px;">
          <div class="pm-name" style="margin-bottom:8px;"><span class="badge" style="background-color: ${color}; color:#fff; font-weight: bold; padding: 5px 10px; font-size:13px;"><i class="fa-solid fa-layer-group"></i> ${major}</span></div>
          <div class="pm-amount" style="color: var(--text-main); font-weight: 700;">${formatKRW(amt)}</div>
        </div>
      `;
      cIdx++;
    });

    totalWrap.innerHTML = html;
    container.appendChild(totalWrap);
  }

  const sortedPm = Object.entries(pmMap).sort((a, b) => b[1] - a[1]);
  sortedPm.forEach(([pm, amt]) => {
    const pct = totalFixed > 0 ? ((amt / totalFixed) * 100).toFixed(1) : 0;
    const card = document.createElement('div');
    card.className = 'pm-card';
    card.innerHTML = `
      <div class="pm-name" style="margin-bottom:8px;">${getPaymentMethodHTML(pm)}</div>
      <div class="pm-amount" style="color:var(--text-main);">${formatKRW(amt)}</div>
      <div style="font-size:11px; color:var(--text-muted); margin-top:4px;">총 고정비 대비 ${pct}%</div>
    `;
    container.appendChild(card);
  });
}

function renderFixedCategorySummary() {
  const container = document.getElementById('fixed-category-summary');
  if (!container) return;
  container.innerHTML = '';

  const fcMap = {};
  let totalFixed = 0;
  appState.fixedExpenses.forEach(item => {
    const cat = item.category || '기타';
    const amt = Number(item.amount) || 0;
    fcMap[cat] = (fcMap[cat] || 0) + amt;
    totalFixed += amt;
  });

  if (totalFixed > 0) {
    const totalCard = document.createElement('div');
    totalCard.className = 'pm-card pm-card-total';
    totalCard.style.border = '1px solid var(--accent-total)';
    totalCard.style.backgroundColor = 'rgba(139, 92, 246, 0.05)';
    totalCard.innerHTML = `
      <div class="pm-name"><span class="badge badge-total" style="font-weight: bold;">총 합계</span></div>
      <div class="pm-amount" style="margin-top:4px; color: var(--text-main); font-weight: 700;">${formatKRW(totalFixed)}</div>
    `;
    container.appendChild(totalCard);
  }

  const sortedFc = Object.entries(fcMap).sort((a, b) => b[1] - a[1]);
  sortedFc.forEach(([cat, amt]) => {
    const pct = totalFixed > 0 ? ((amt / totalFixed) * 100).toFixed(1) : 0;
    const card = document.createElement('div');
    card.className = 'pm-card';
    card.innerHTML = `
      <div class="pm-name">${cat}</div>
      <div class="pm-amount" style="color:var(--text-main);">${formatKRW(amt)}</div>
      <div style="font-size:11px; color:var(--text-muted); margin-top:2px;">총 고정비 대비 ${pct}%</div>
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

  const calcs = calculateTotals();
  const baseIncome = calcs.totalIncome > 0 ? calcs.totalIncome : 1;

  // Add Total Sum Card (Top)
  if (totalAlloc > 0) {
    const allocPct = ((totalAlloc / baseIncome) * 100).toFixed(1);
    const totalCard = document.createElement('div');
    totalCard.className = 'pm-card';
    totalCard.style.border = '1px solid var(--accent-total)';
    totalCard.style.backgroundColor = 'rgba(139, 92, 246, 0.05)';
    totalCard.innerHTML = `
      <div class="pm-name"><span class="badge badge-total" style="font-weight: bold;">총 사용 합계</span></div>
      <div class="pm-amount" style="margin-top:4px; color: var(--text-main); font-weight: 700;">${formatKRW(totalAlloc)}</div>
      <div style="font-size:11px; color:var(--text-muted); margin-top:2px;">총 수입 대비 ${allocPct}%</div>
    `;
    container.appendChild(totalCard);
  }

  // Add Remaining Balance Card (Top)
  const remainPct = ((calcs.remainingBalance / baseIncome) * 100).toFixed(1);
  const remainCard = document.createElement('div');
  remainCard.className = 'pm-card';
  remainCard.style.border = '1px solid var(--accent-remain)';
  remainCard.style.backgroundColor = 'rgba(14, 165, 233, 0.05)';
  remainCard.innerHTML = `
    <div class="pm-name"><span class="badge badge-remain" style="font-weight: bold;">남은 잔액</span></div>
    <div class="pm-amount" style="margin-top:4px; color: ${calcs.remainingBalance < 0 ? 'var(--accent-danger)' : 'var(--accent-remain)'}; font-weight: 700;">${formatKRW(calcs.remainingBalance)}</div>
    <div style="font-size:11px; color:var(--text-muted); margin-top:2px;">총 수입 대비 ${remainPct}%</div>
  `;
  container.appendChild(remainCard);

  const sortedCats = Object.entries(catMap).sort((a, b) => b[1] - a[1]);

  sortedCats.forEach(([cat, amt]) => {
    const badgeClass = getAllocCategoryBadgeClass(cat);
    const pct = ((amt / baseIncome) * 100).toFixed(1);
    
    const card = document.createElement('div');
    card.className = 'pm-card';
    card.innerHTML = `
      <div class="pm-name"><span class="badge ${badgeClass}">${cat}</span></div>
      <div class="pm-amount" style="margin-top:4px;">${formatKRW(amt)}</div>
      <div style="font-size:11px; color:var(--text-muted); margin-top:2px;">총 수입 대비 ${pct}%</div>
    `;
    container.appendChild(card);
  });
}

function formatDayDisplay(val) {
  if (!val || val === '-') return '-';
  const s = String(val).trim();
  if (s === '수기') return s;
  return s.replace(/일+$/, '') + '일';
}

function formatDayInput(val) {
  if (!val) return '';
  return String(val).trim().replace(/일+$/, '');
}

// Direct Inline Cell Editor
function editInlineCell(element, pathStr, itemId, fieldName, type = 'text') {
  if (element.querySelector('input')) return; // Already editing

  let currVal = getFieldValue(pathStr, itemId, fieldName);
  const originalVal = currVal;

  const input = document.createElement('input');
  input.type = 'text'; // Always use text to allow commas
  input.className = 'cell-input';

  if (type === 'number') {
    // Format initial value with commas
    input.value = Number(currVal) ? Number(currVal).toLocaleString('ko-KR') : (currVal === 0 ? '0' : '');
    // Format dynamically as user types
    input.addEventListener('input', (e) => {
      const el = e.target;
      const originalValue = el.value;
      const selectionStart = el.selectionStart || 0;

      // Count raw characters before cursor
      const rawBeforeCursorLength = originalValue.slice(0, selectionStart).replace(/[^0-9-]/g, '').length;

      let raw = originalValue.replace(/[^0-9-]/g, '');
      let newValue = '';
      if (raw) {
        if (raw === '-') {
          newValue = '-';
        } else {
          newValue = Number(raw).toLocaleString('ko-KR');
        }
      }

      el.value = newValue;

      // Restore cursor position
      let newCursorPos = 0;
      let rawCount = 0;
      for (let i = 0; i < newValue.length; i++) {
        if (rawCount === rawBeforeCursorLength) {
          newCursorPos = i;
          break;
        }
        if (/[0-9-]/.test(newValue[i])) {
          rawCount++;
        }
      }
      if (rawCount === rawBeforeCursorLength) {
        newCursorPos = newValue.length;
      }

      el.setSelectionRange(newCursorPos, newCursorPos);
    });
  } else {
    input.value = currVal;
  }

  element.innerHTML = '';
  element.appendChild(input);
  input.focus();

  const finishEdit = () => {
    let newVal = input.value.trim();
    if (type === 'number') {
      newVal = Number(newVal.replace(/[^0-9-]/g, '')) || 0;
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

window.toggleAllocAutoTransfer = function(pathStr, itemId) {
  let list = resolvePath(pathStr);
  if (!list) return;
  let item = list.find(i => i.id === itemId);
  if (item) {
    let currentlyAuto = false;
    if (item.isAutoTransfer !== undefined) {
      currentlyAuto = item.isAutoTransfer;
    } else if (item.method) {
      currentlyAuto = item.method.includes('자동');
    }

    const actionText = currentlyAuto ? '해제' : '설정';
    if (!confirm(`해당 항목의 자동이체를 ${actionText}하시겠습니까?`)) {
      renderAll(); // Revert visual state
      return;
    }

    item.isAutoTransfer = !currentlyAuto;
    if (!item.isAutoTransfer && item.method && item.method.includes('자동')) {
      item.method = '';
    } else if (item.isAutoTransfer) {
      item.method = '자동이체';
    }
    saveState();
    renderAll();
  }
};

window.toggleFixedAutoTransfer = function(itemId) {
  let item = appState.fixedExpenses.find(i => i.id === itemId);
  if (item) {
    let currentlyAuto = false;
    if (item.isAutoTransfer !== undefined) {
      currentlyAuto = item.isAutoTransfer;
    } else if (item.note) {
      currentlyAuto = item.note.includes('자동');
    }

    const actionText = currentlyAuto ? '해제' : '설정';
    if (!confirm(`해당 고정지출의 자동이체를 ${actionText}하시겠습니까?`)) {
      renderAll(); // Revert visual state
      return;
    }

    item.isAutoTransfer = !currentlyAuto;
    saveState();
    renderAll();
  }
};

window.sortAllocations = function(person, field) {
  if (!appState.allocations[person]) return;
  
  if (!window.sortState) window.sortState = {};
  if (!window.sortState[person]) window.sortState[person] = { field: null, asc: true };
  
  const state = window.sortState[person];
  if (state.field === field) {
    state.asc = !state.asc; // Toggle sort direction
  } else {
    state.field = field;
    state.asc = true;
  }
  
  appState.allocations[person].sort((a, b) => {
    let valA = a[field] || '';
    let valB = b[field] || '';
    if (valA < valB) return state.asc ? -1 : 1;
    if (valA > valB) return state.asc ? 1 : -1;
    return 0;
  });
  
  saveState();
  renderAllocationTables();
};

window.sortFixedExpenses = function(field) {
  if (!appState.fixedExpenses) return;
  
  if (!window.sortState) window.sortState = {};
  if (!window.sortState.fixedExpenses) window.sortState.fixedExpenses = { field: null, asc: true };
  
  const state = window.sortState.fixedExpenses;
  if (state.field === field) {
    state.asc = !state.asc; // Toggle sort direction
  } else {
    state.field = field;
    state.asc = true;
  }
  
  appState.fixedExpenses.sort((a, b) => {
    let valA = a[field] || '';
    let valB = b[field] || '';

    if (field === 'day') {
      let numA = parseInt(valA);
      let numB = parseInt(valB);
      if (isNaN(numA)) numA = 999;
      if (isNaN(numB)) numB = 999;
      
      if (numA !== numB) {
        return state.asc ? (numA - numB) : (numB - numA);
      }
    }
    
    if (valA < valB) return state.asc ? -1 : 1;
    if (valA > valB) return state.asc ? 1 : -1;
    return 0;
  });
  
  saveState();
  renderFixedExpensesTable();
};

function toggleFixedPaid(itemId, isPaid) {
  const item = appState.fixedExpenses.find(i => i.id === itemId);
  if (item) {
    item.isPaid = isPaid;
    saveState();
    renderTimeline(); // Update calendar
  }
}

function toggleAllocPaid(person, itemId, isPaid) {
  const list = appState.allocations[person];
  if (list) {
    const item = list.find(i => i.id === itemId);
    if (item) {
      item.isPaid = isPaid;
      saveState();
      renderTimeline(); // Update calendar
    }
  }
}

window.editFlexibleBudget = function() {
  const currentVal = appState.flexibleLivingBudget !== undefined ? Number(appState.flexibleLivingBudget) : 1000000;
  const newVal = prompt("유동 생활비(목표 예산)를 입력하세요 (숫자만 입력):", currentVal);
  if (newVal !== null) {
    const num = Number(newVal.replace(/[^0-9-]/g, ''));
    if (!isNaN(num)) {
      appState.flexibleLivingBudget = num;
      saveState();
      renderAll();
    } else {
      alert("올바른 숫자를 입력해주세요.");
    }
  }
};
function toggleIncomePaid(itemId, isPaid) {
  const item = appState.incomes.find(i => i.id === itemId);
  if (item) {
    item.isPaid = isPaid;
    saveState();
    renderTimeline(); // Update calendar
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
      const el = e.target;
      const originalValue = el.value;
      const selectionStart = el.selectionStart || 0;

      // Count raw characters before cursor
      const rawBeforeCursorLength = originalValue.slice(0, selectionStart).replace(/[^0-9-]/g, '').length;

      let raw = originalValue.replace(/[^0-9-]/g, '');
      let newValue = '';
      if (raw) {
        if (raw === '-') {
          newValue = '-';
        } else {
          newValue = Number(raw).toLocaleString('ko-KR');
        }
      }

      el.value = newValue;

      // Restore cursor position
      let newCursorPos = 0;
      let rawCount = 0;
      for (let i = 0; i < newValue.length; i++) {
        if (rawCount === rawBeforeCursorLength) {
          newCursorPos = i;
          break;
        }
        if (/[0-9-]/.test(newValue[i])) {
          rawCount++;
        }
      }
      if (rawCount === rawBeforeCursorLength) {
        newCursorPos = newValue.length;
      }

      el.setSelectionRange(newCursorPos, newCursorPos);
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

window.currentCategoryModalMode = 'all';

function openCategoryManagerModal(mode) {
  if (mode) {
    window.currentCategoryModalMode = mode;
  }
  const currentMode = window.currentCategoryModalMode;

  const body = document.getElementById('modal-body');
  
  let title = '카테고리 설정 (수입 & 배분 통합 관리)';
  if (currentMode === 'income') title = '카테고리 설정 (수입 통합 관리)';
  else if (currentMode === 'alloc') title = '카테고리 설정 (월급 배분 내역)';
  
  document.getElementById('modal-title').innerText = title;

  if (!appState.categories) {
    appState.categories = {
      recipients: ['계원', '동욱', '공통 / 가구'],
      incomeItems: ['기본급 (급여)', '보너스 / 상여금', '인센티브', '부수입 / 알바', '투자수익 / 배당', '기타 수입'],
      allocCategories: ['생활비', '저축/적금', '투자/연금', '대출/이자', '비상금/경조사'],
      paymentMethods: ['현금', '하나카드', '오빠카드(삼성)', '오빠카드', '우리카드', '신한카드', '카드'],
      fixedCategories: ['용돈', '곳돈', '구독', '보험', '핸드폰요금', '주거', '교통']
    };
  }
  if (!appState.categories.allocCategories) {
    appState.categories.allocCategories = ['생활비', '저축/적금', '투자/연금', '대출/이자', '비상금/경조사'];
  }
  if (!appState.categories.paymentMethods) {
    appState.categories.paymentMethods = ['현금', '하나카드', '오빠카드(삼성)', '오빠카드', '우리카드', '신한카드', '카드'];
  }
  if (!appState.categories.fixedCategories) {
    appState.categories.fixedCategories = ['용돈', '곳돈', '구독', '보험', '핸드폰요금', '주거', '교통'];
  }

  const renderCatLists = () => {
    let html = '';

    if (currentMode === 'all' || currentMode === 'income') {
      html += `
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

        <div class="card-box ${currentMode === 'income' ? '' : 'mb-3'}" style="padding: 16px; background: rgba(0,0,0,0.2);">
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
      `;
    }

    if (currentMode === 'all' || currentMode === 'alloc') {
      html += `
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
    }

    body.innerHTML = html;
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



window.addAllocCategory = function() {
  const input = document.getElementById('new-alloc-cat-input');
  if (!input) return;
  const val = input.value.trim();
  if (val) {
    if (!appState.categories.allocCategories.includes(val)) {
      appState.categories.allocCategories.push(val);
      saveState();
      openCategoryManagerModal();
      renderAllocationTables();
    }
  }
};

window.editAllocCategory = function(idx) {
  const oldVal = appState.categories.allocCategories[idx];
  const newVal = prompt('새 배분 카테고리명을 입력하세요:', oldVal);
  if (newVal && newVal.trim() !== '') {
    const trimmed = newVal.trim();
    appState.categories.allocCategories[idx] = trimmed;
    ['gyewon', 'dongwook'].forEach(person => {
      (appState.allocations[person] || []).forEach(item => {
        if (item.category === oldVal) item.category = trimmed;
      });
    });
    saveState();
    openCategoryManagerModal();
    renderAllocationTables();
  }
};

window.deleteAllocCategory = function(idx) {
  if (appState.categories.allocCategories.length <= 1) {
    alert('최소 1개 이상의 배분 항목이 필요합니다.');
    return;
  }
  appState.categories.allocCategories.splice(idx, 1);
  saveState();
  openCategoryManagerModal();
  renderAllocationTables();
};

// Payment Method CRUD
window.addPaymentMethod = function() {
  const input = document.getElementById('new-payment-method-input');
  if (!input) return;
  const val = input.value.trim();
  if (val) {
    if (!appState.categories.paymentMethods.includes(val)) {
      appState.categories.paymentMethods.push(val);
      saveState();
      openCategoryManagerModal();
      renderFixedExpensesTable();
      updateMethodFilterDropdown();
    }
  }
};

window.editPaymentMethod = function(idx) {
  const oldVal = appState.categories.paymentMethods[idx];
  const newVal = prompt('결제 수단 이름을 수정하세요:', oldVal);
  if (newVal && newVal.trim() !== '' && newVal.trim() !== oldVal) {
    const trimmed = newVal.trim();
    appState.categories.paymentMethods[idx] = trimmed;
    // Update all fixed expenses using this method
    (appState.fixedExpenses || []).forEach(item => {
      if (item.method === oldVal) item.method = trimmed;
    });
    saveState();
    openCategoryManagerModal();
    renderFixedExpensesTable();
    updateMethodFilterDropdown();
  }
};

window.deletePaymentMethod = function(idx) {
  if ((appState.categories.paymentMethods || []).length <= 1) {
    alert('최소 1개 이상의 결제 수단이 필요합니다.');
    return;
  }
  appState.categories.paymentMethods.splice(idx, 1);
  saveState();
  openCategoryManagerModal();
  renderFixedExpensesTable();
  updateMethodFilterDropdown();
};

function updateMethodFilterDropdown() {
  const sel = document.getElementById('filter-method');
  if (!sel) return;
  const current = sel.value;
  const methods = appState.categories?.paymentMethods || [];
  sel.innerHTML = `<option value="ALL">전체</option>${methods.map(m => `<option value="${m}" ${current === m ? 'selected' : ''}>${m}</option>`).join('')}`;
}

window.editFixedExpenseMethod = function(element, itemId) {
  const item = (appState.fixedExpenses || []).find(i => i.id === itemId);
  if (!item) return;
  const methods = appState.categories?.paymentMethods || [];
  openCategoryPickerModal({
    title: '\uacb0\uc81c \uc218\ub2e8 \uc120\ud0dd',
    currentCategory: item.method,
    categoryList: methods,
    type: 'paymentMethod',
    onSelect: (selected) => {
      item.method = selected;
      saveState();
      renderFixedExpensesTable();
      closeModal();
    }
  });
};

window.openPaymentMethodManagerModal = function() {
  document.getElementById('modal-title').innerText = '결제수단 및 카테고리 관리';
  const body = document.getElementById('modal-body');

  if (!appState.categories) appState.categories = {};
  if (!appState.categories.paymentMethods) {
    appState.categories.paymentMethods = ['현금', '하나카드', '오빠카드(삼성)', '오빠카드', '우리카드', '신한카드', '카드'];
  }
  if (!appState.categories.fixedCategories) {
    appState.categories.fixedCategories = ['용돈', '곳돈', '구독', '보험', '핸드폰요금', '주거', '교통'];
  }

  // Use temporary state to prevent auto-saving
  let tempPMs = [...appState.categories.paymentMethods];
  let tempFCats = [...appState.categories.fixedCategories];
  let tempExpenses = JSON.parse(JSON.stringify(appState.fixedExpenses || []));

  window.draggedPMIdx = null;
  window.draggedFCatIdx = null;

  window.onDragStartPM = (e, idx) => { draggedPMIdx = idx; e.dataTransfer.effectAllowed = 'move'; e.target.style.opacity = '0.5'; };
  window.onDragOverPM = (e) => { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; };
  window.onDragEndPM = (e) => { e.target.style.opacity = '1'; draggedPMIdx = null; draggedFCatIdx = null; };
  
  window.onDropPM = (e, targetIdx) => {
    e.preventDefault();
    e.stopPropagation();
    if (draggedPMIdx !== null && draggedPMIdx !== targetIdx) {
      const draggedVal = tempPMs[draggedPMIdx];
      const targetVal = tempPMs[targetIdx];
      const newType = getPMType(targetVal);
      
      tempPMs.splice(draggedPMIdx, 1);
      const newTargetIdx = tempPMs.indexOf(targetVal);
      tempPMs.splice(newTargetIdx, 0, draggedVal);
      
      if (!appState.categories.paymentMethodTypes) appState.categories.paymentMethodTypes = {};
      appState.categories.paymentMethodTypes[draggedVal] = newType;
      renderAll();
    }
  };

  window.onDragOverPMGroup = (e) => { e.preventDefault(); e.currentTarget.style.borderColor = 'var(--accent-primary)'; };
  window.onDragLeavePMGroup = (e) => { e.currentTarget.style.borderColor = 'transparent'; };
  window.onDropPMGroup = (e, groupName) => {
    e.preventDefault();
    e.currentTarget.style.borderColor = 'transparent';
    if (draggedPMIdx !== null) {
      const draggedVal = tempPMs[draggedPMIdx];
      if (getPMType(draggedVal) !== groupName) {
        if (!appState.categories.paymentMethodTypes) appState.categories.paymentMethodTypes = {};
        appState.categories.paymentMethodTypes[draggedVal] = groupName;
        renderAll();
      }
    }
  };

  window.onDragStartFCat = (e, idx) => { draggedFCatIdx = idx; e.dataTransfer.effectAllowed = 'move'; e.target.style.opacity = '0.5'; };
  window.onDropFCat = (e, targetIdx) => {
    e.preventDefault();
    e.stopPropagation();
    if (draggedFCatIdx !== null && draggedFCatIdx !== targetIdx) {
      const item = tempFCats.splice(draggedFCatIdx, 1)[0];
      tempFCats.splice(targetIdx, 0, item);
      renderAll();
    }
  };

  const renderAll = () => {
    const groupedPMs = {};
    tempPMs.forEach(pm => {
      const type = getPMType(pm);
      if (!groupedPMs[type]) groupedPMs[type] = [];
      groupedPMs[type].push(pm);
    });

    let pmGroupsHTML = '';
    Object.keys(groupedPMs).forEach(group => {
      pmGroupsHTML += `
        <h5 style="font-size:12px; font-weight:700; color: var(--accent-primary); margin-bottom:10px;">
          <i class="fa-solid fa-layer-group"></i> ${group}
          <i class="fa-solid fa-pen-to-square text-muted" style="cursor:pointer; margin-left: 8px;" title="대분류 이름 수정" onclick="editPMGroupInModal('${group}')"></i>
        </h5>
        <div class="flex-gap-2 mb-3" style="flex-wrap:wrap; min-height: 30px; padding: 5px; border: 1px dashed transparent;" 
             ondragover="onDragOverPMGroup(event)" 
             ondragleave="onDragLeavePMGroup(event)"
             ondrop="onDropPMGroup(event, '${group}')">
          ${groupedPMs[group].map(pm => {
            const idx = tempPMs.indexOf(pm);
            return `
            <span draggable="true" 
                  ondragstart="onDragStartPM(event, ${idx})" 
                  ondragover="onDragOverPM(event)" 
                  ondrop="onDropPM(event, ${idx})" 
                  ondragend="onDragEndPM(event)"
                  style="display:inline-flex; align-items:center; gap:8px; cursor:grab; padding:4px 8px; border-radius:4px; transition: background 0.2s;">
              ${getPaymentMethodHTML(pm)}
              <i class="fa-solid fa-pen-to-square text-muted" style="cursor:pointer;" title="대분류/소분류 수정" onclick="editPMInModal(${idx})"></i>
              <i class="fa-solid fa-xmark text-muted" style="cursor:pointer;" onclick="deletePMInModal(${idx})"></i>
            </span>
          `;}).join('')}
          <div style="display:inline-flex; align-items:center; margin-left: 8px;">
            <input type="text" id="new-pm-group-input-${group}" class="form-control form-control-sm" style="width:160px;" placeholder="${group.includes('카드') ? '별명, 카드사 (예: 밥,신한)' : '결제수단명 입력'}" onkeydown="if(event.key==='Enter') addPMToGroup('${group}')">
            <button class="btn btn-outline-primary btn-sm" style="margin-left:4px;" onclick="addPMToGroup('${group}')"><i class="fa-solid fa-plus"></i></button>
          </div>
        </div>
      `;
    });

    body.innerHTML = `
      <div class="card-box mb-3" style="padding: 16px; background: rgba(0,0,0,0.2);">
        <h4 style="font-size:14px; font-weight:700; margin-bottom:10px;"><i class="fa-solid fa-credit-card"></i> 결제 수단</h4>
        <p class="text-muted mb-3" style="font-size:12px;">원하시는 대분류(예: 주거래카드)에 소분류(결제수단명)를 등록하세요. '카드'의 경우 카드사 정보(로고)도 등록할 수 있습니다.</p>
        
        ${pmGroupsHTML}

        <div class="flex-gap-2" style="align-items:center; margin-top:10px; padding-top:10px; border-top:1px solid rgba(255,255,255,0.1);">
          <input type="text" id="new-pm-modal-type" class="form-control form-control-sm" style="width:130px;" placeholder="새 대분류명 입력">
          <input type="text" id="new-pm-modal-input" class="form-control form-control-sm" placeholder="새 결제수단명 입력">
          <button class="btn btn-secondary btn-sm" onclick="addPaymentMethodFromModal()"><i class="fa-solid fa-plus"></i> 대분류 추가</button>
        </div>
      </div>

      <div class="card-box" style="padding: 16px; background: rgba(0,0,0,0.2);">
        <h4 style="font-size:14px; font-weight:700; margin-bottom:10px;"><i class="fa-solid fa-tag"></i> 고정지출 카테고리</h4>
        <p class="text-muted mb-3" style="font-size:12px;">고정지출 항목의 카테고리를 관리합니다.</p>
        <div class="flex-gap-2 mb-3" style="flex-wrap:wrap;">
          ${tempFCats.map((cat, idx) => `
            <span draggable="true"
                  ondragstart="onDragStartFCat(event, ${idx})"
                  ondragover="onDragOverPM(event)"
                  ondrop="onDropFCat(event, ${idx})"
                  ondragend="onDragEndPM(event)"
                  style="display:inline-flex; align-items:center; gap:8px; cursor:grab; padding:4px 8px; border-radius:4px;">
              ${getFixedCategoryHTML(cat)}
              <i class="fa-solid fa-pen-to-square text-muted" style="cursor:pointer;" onclick="editFCatInModal(${idx})"></i>
              <i class="fa-solid fa-xmark text-muted" style="cursor:pointer;" onclick="deleteFCatInModal(${idx})"></i>
            </span>
          `).join('')}
        </div>
        <div class="flex-gap-2">
          <input type="text" id="new-fcat-modal-input" class="form-control form-control-sm" placeholder="새 카테고리 (예: 통신비, 교육, 운동)">
          <button class="btn btn-primary btn-sm" onclick="addFCatFromModal()"><i class="fa-solid fa-plus"></i> 추가</button>
        </div>
      </div>
    `;
  };

  window.addPaymentMethodFromModal = function() {
    const input = document.getElementById('new-pm-modal-input');
    const typeInput = document.getElementById('new-pm-modal-type');
    if (!input) return;
    const rawVal = input.value.trim();
    const typeVal = typeInput && typeInput.value.trim() ? typeInput.value.trim() : '미분류';

    let val = rawVal;
    let company = '';
    if (rawVal.includes(',')) {
      const parts = rawVal.split(',');
      val = parts[0].trim();
      company = parts.slice(1).join(',').trim();
    }

    if (val && !tempPMs.includes(val)) {
      tempPMs.push(val);
      if (!appState.categories.paymentMethodTypes) appState.categories.paymentMethodTypes = {};
      appState.categories.paymentMethodTypes[val] = typeVal;

      if (company) {
        if (!appState.categories.paymentMethodCompanies) appState.categories.paymentMethodCompanies = {};
        appState.categories.paymentMethodCompanies[val] = company;
      }

      renderAll();
    }
  };

  window.addPMToGroup = function(groupName) {
    const input = document.getElementById(`new-pm-group-input-${groupName}`);
    if (!input) return;
    const rawVal = input.value.trim();
    if (!rawVal) return;

    let val = rawVal;
    let company = '';
    if (rawVal.includes(',')) {
      const parts = rawVal.split(',');
      val = parts[0].trim();
      company = parts.slice(1).join(',').trim();
    }

    if (!val) return;

    if (!tempPMs.includes(val)) {
      tempPMs.push(val);
      if (!appState.categories.paymentMethodTypes) appState.categories.paymentMethodTypes = {};
      appState.categories.paymentMethodTypes[val] = groupName;

      if (company) {
        if (!appState.categories.paymentMethodCompanies) appState.categories.paymentMethodCompanies = {};
        appState.categories.paymentMethodCompanies[val] = company;
      }
      renderAll();
    }
  };

  window.addFCatFromModal = function() {
    const input = document.getElementById('new-fcat-modal-input');
    if (!input) return;
    const val = input.value.trim();
    if (val && !tempFCats.includes(val)) {
      tempFCats.push(val); renderAll();
    }
  };

  window.editPMGroupInModal = function(oldGroupName) {
    const newGroupName = prompt('대분류 이름을 수정하세요:', oldGroupName);
    if (!newGroupName || newGroupName.trim() === '' || newGroupName.trim() === oldGroupName) return;
    const trimmed = newGroupName.trim();
    
    if (!appState.categories.paymentMethodTypes) appState.categories.paymentMethodTypes = {};
    
    tempPMs.forEach(pm => {
      if (getPMType(pm) === oldGroupName) {
        appState.categories.paymentMethodTypes[pm] = trimmed;
      }
    });
    renderAll();
  };

  window.editPMInModal = function(idx) {
    const oldVal = tempPMs[idx];
    const oldType = getPMType(oldVal);

    const oldCompany = (appState.categories.paymentMethodCompanies && appState.categories.paymentMethodCompanies[oldVal]) || '';
    
    const isCard = oldType.includes('카드') || oldVal.includes('카드');
    let promptMsg = '결제수단 별명을 수정하세요 (예: 데이트용 카드):';
    let defaultText = oldVal;
    
    if (isCard) {
      promptMsg = '결제수단 별명과 카드사를 쉼표(,)로 구분해서 적어주세요. (예: 밥값카드, 신한)';
      defaultText = oldCompany ? `${oldVal}, ${oldCompany}` : oldVal;
    }

    const newVal = prompt(promptMsg, defaultText);
    if (!newVal || newVal.trim() === '') return;

    let trimmed = newVal.trim();
    let trimmedCompany = null;

    if (isCard && trimmed.includes(',')) {
      const parts = trimmed.split(',');
      trimmed = parts[0].trim();
      trimmedCompany = parts.slice(1).join(',').trim();
    } else if (isCard && !trimmed.includes(',')) {
      trimmedCompany = newVal.includes(',') ? '' : oldCompany;
    }

    if (!trimmed) return;

    tempPMs[idx] = trimmed;
    tempExpenses.forEach(item => { if (item.method === oldVal) item.method = trimmed; });
    
    if (!appState.categories.paymentMethodTypes) appState.categories.paymentMethodTypes = {};
    if (trimmed !== oldVal) delete appState.categories.paymentMethodTypes[oldVal];
    appState.categories.paymentMethodTypes[trimmed] = oldType; // Keep the same company/group
    
    if (!appState.categories.paymentMethodCompanies) appState.categories.paymentMethodCompanies = {};
    if (trimmed !== oldVal) delete appState.categories.paymentMethodCompanies[oldVal];
    if (trimmedCompany) {
      appState.categories.paymentMethodCompanies[trimmed] = trimmedCompany;
    }

    renderAll();
  };
  window.deletePMInModal = function(idx) {
    if (tempPMs.length <= 1) { alert('최소 1개 이상 필요'); return; }
    const oldVal = tempPMs[idx];
    const inUse = tempExpenses.some(item => item.method === oldVal);
    if (inUse) {
      alert('이 결제수단이 지정된 지출 내역이 있어 삭제할 수 없습니다.');
      return;
    }
    tempPMs.splice(idx, 1);
    renderAll();
  };
  window.editFCatInModal = function(idx) {
    const oldVal = tempFCats[idx];
    const newVal = prompt('카테고리 이름을 수정하세요:', oldVal);
    if (newVal && newVal.trim() !== '' && newVal.trim() !== oldVal) {
      const trimmed = newVal.trim();
      tempFCats[idx] = trimmed;
      tempExpenses.forEach(item => { if (item.category === oldVal) item.category = trimmed; });
      renderAll();
    }
  };
  window.deleteFCatInModal = function(idx) {
    if (tempFCats.length <= 1) { alert('최소 1개 이상 필요'); return; }
    const oldVal = tempFCats[idx];
    const inUse = tempExpenses.some(item => item.category === oldVal);
    if (inUse) {
      alert('이 카테고리가 지정된 지출 내역이 있어 삭제할 수 없습니다.');
      return;
    }
    tempFCats.splice(idx, 1);
    renderAll();
  };

  renderAll();

  // Save changes only when modal save button is clicked
  document.getElementById('modal-save-btn').onclick = () => {
    appState.categories.paymentMethods = tempPMs;
    appState.categories.fixedCategories = tempFCats;
    appState.fixedExpenses = tempExpenses;
    saveState();
    renderFixedExpensesTable();
    updateMethodFilterDropdown();
    closeModal();
  };
  
  document.getElementById('item-modal').classList.add('active');
};

window.editFixedExpenseCategory = function(element, itemId) {
  const item = (appState.fixedExpenses || []).find(i => i.id === itemId);
  if (!item) return;
  const cats = appState.categories?.fixedCategories || [];
  openCategoryPickerModal({
    title: '\uace0\uc815\uc9c0\ucd9c \uce74\ud14c\uace0\ub9ac \uc120\ud0dd',
    currentCategory: item.category,
    categoryList: cats,
    type: 'fixedCategory',
    onSelect: (selected) => {
      item.category = selected;
      saveState();
      renderFixedExpensesTable();
      closeModal();
    }
  });
};

// Mobile & Web Touch-Friendly Category Picker Modal
function openCategoryPickerModal({ title, currentCategory, categoryList, onSelect, onAddNew, type }) {
  const body = document.getElementById('modal-body');
  document.getElementById('modal-title').innerText = title || '카테고리 선택';

  const categoryCards = (categoryList || []).map(cat => {
    const isSelected = cat === currentCategory;
    let badgeHTML = '';
    if (type === 'paymentMethod') {
      badgeHTML = getPaymentMethodHTML(cat);
    } else if (type === 'fixedCategory') {
      badgeHTML = getFixedCategoryHTML(cat);
    } else {
      const badgeClass = typeof getAllocCategoryBadgeClass === 'function' ? getAllocCategoryBadgeClass(cat) : 'badge-info';
      badgeHTML = `<span class="badge ${badgeClass}" style="font-size: 13px; padding: 6px 12px;">${cat}</span>`;
    }

    return `
      <button class="btn btn-outline-light category-pick-btn ${isSelected ? 'active-cat' : ''}" 
              style="padding: 12px 16px; font-size: 14px; font-weight: 600; border-radius: 10px; text-align: left; display: flex; align-items: center; justify-content: space-between; width: 100%; margin-bottom: 8px; cursor: pointer; background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.1);"
              onclick="selectPickedCategory('${cat}')">
        ${badgeHTML}
        ${isSelected ? '<i class="fa-solid fa-circle-check text-success" style="font-size: 18px;"></i>' : '<i class="fa-solid fa-chevron-right text-muted" style="font-size: 12px;"></i>'}
      </button>
    `;
  }).join('');

  body.innerHTML = `
    <p class="text-muted mb-3" style="font-size: 13px;">변경할 카테고리를 아래 목록에서 터치/클릭하세요.</p>
    <div class="category-picker-list mb-3" style="max-height: 280px; overflow-y: auto; padding-right: 4px;">
      ${categoryCards}
    </div>
    <div class="d-flex flex-gap-2">
      <button class="btn btn-primary btn-sm w-100" style="padding: 10px;" onclick="promptNewCategoryInPicker()">
        <i class="fa-solid fa-plus"></i> 새 카테고리 직접 입력
      </button>
      <button class="btn btn-secondary btn-sm" style="padding: 10px; white-space: nowrap;" onclick="openCategoryManagerModal()">
        <i class="fa-solid fa-gear"></i> 카테고리 관리
      </button>
    </div>
  `;

  window._activeCategoryPickerOnSelect = onSelect;
  window._activeCategoryPickerOnAddNew = onAddNew;

  document.getElementById('item-modal').classList.add('active');
}

window.selectPickedCategory = function(cat) {
  closeModal();
  if (window._activeCategoryPickerOnSelect) {
    window._activeCategoryPickerOnSelect(cat);
  }
};

window.promptNewCategoryInPicker = function() {
  const newCat = prompt('새 카테고리명을 입력하세요:');
  if (newCat && newCat.trim()) {
    const trimmed = newCat.trim();
    closeModal();
    if (window._activeCategoryPickerOnAddNew) {
      window._activeCategoryPickerOnAddNew(trimmed);
    } else if (window._activeCategoryPickerOnSelect) {
      window._activeCategoryPickerOnSelect(trimmed);
    }
  }
};

window.editInlineCategory = function(element, incId, evt) {
  if (evt) evt.stopPropagation();

  const inc = (appState.incomes || []).find(i => i.id === incId);
  if (!inc) return;

  const itemCatList = (appState.categories && appState.categories.incomeItems) ? appState.categories.incomeItems : ['기본급 (급여)', '보너스 / 상여금', '인센티브', '부수입 / 알바', '투자수익 / 배당', '기타 수입'];

  openCategoryPickerModal({
    title: `${inc.title || '수입 항목'} 카테고리 변경`,
    currentCategory: inc.category || inc.title,
    categoryList: itemCatList,
    onSelect: (selectedCat) => {
      inc.category = selectedCat;
      saveState();
      renderIncomeTables();
      renderCharts();
    },
    onAddNew: (newCat) => {
      if (!appState.categories.incomeItems.includes(newCat)) {
        appState.categories.incomeItems.push(newCat);
      }
      inc.category = newCat;
      saveState();
      renderIncomeTables();
      renderCharts();
    }
  });
};

window.editInlineAllocCategory = function(element, pathStr, itemId, evt) {
  if (evt) evt.stopPropagation();

  const list = resolvePath(pathStr);
  const item = list.find(i => i.id === itemId);
  if (!item) return;

  if (!appState.categories) appState.categories = {};
  if (!appState.categories.allocCategories) {
    appState.categories.allocCategories = ['생활비', '저축/적금', '투자/연금', '대출/이자', '비상금/경조사'];
  }
  const catList = appState.categories.allocCategories;

  openCategoryPickerModal({
    title: `${item.name || '배분 항목'} 카테고리 변경`,
    currentCategory: item.category,
    categoryList: catList,
    onSelect: (selectedCat) => {
      item.category = selectedCat;
      saveState();
      renderAllocationTables();
      renderCategoryAllocationSummary();
      renderCharts();
    },
    onAddNew: (newCat) => {
      if (!appState.categories.allocCategories.includes(newCat)) {
        appState.categories.allocCategories.push(newCat);
      }
      item.category = newCat;
      saveState();
      renderAllocationTables();
      renderCategoryAllocationSummary();
      renderCharts();
    }
  });
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
      <select id="modal-inc-cat-select" class="form-select">
        <option value="">-- 카테고리 선택 --</option>
        ${itemCatOptions}
      </select>
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
    const title = document.getElementById('modal-inc-cat-select').value;
    if (!title) { alert('수입 항목 카테고리를 선택하세요.'); return; }
    
    const rawAmt = document.getElementById('modal-inc-amount').value.replace(/,/g, '');
    const amount = Number(rawAmt) || 0;
    const dayInputVal = document.getElementById('modal-inc-day').value.trim();
    const day = formatDayInput(dayInputVal || '25일');
    const note = document.getElementById('modal-inc-note').value.trim();



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
      <select id="modal-inc-cat-select" class="form-select">
        <option value="">-- 카테고리 선택 --</option>
        ${itemCatOptions}
      </select>
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
    const title = document.getElementById('modal-inc-cat-select').value;
    if (!title) { alert('수입 항목 카테고리를 선택하세요.'); return; }
    inc.title = title;
    inc.category = title;
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

  const allocCatList = (appState.categories && appState.categories.allocCategories) ? appState.categories.allocCategories : ['생활비', '저축/적금', '투자/연금', '대출/이자', '비상금/경조사'];
  const allocCatOptions = allocCatList.map(cat => `<option value="${cat}">${cat}</option>`).join('');

  body.innerHTML = `
    <div class="form-group">
      <label>배분 항목명</label>
      <input type="text" id="modal-alloc-name" class="form-control" placeholder="예: 적금, 생활비, ISA">
    </div>
    <div class="form-group">
      <label>금액 (원)</label>
      <input type="text" id="modal-alloc-amount" class="form-control" placeholder="0 (숫자 입력 시 , 자동 적용)" oninput="this.value = this.value.replace(/[^0-9]/g, '').replace(/\B(?=(\d{3})+(?!\d))/g, ',')">
    </div>
    <div class="form-group">
      <label>출금</label>
      <input type="text" id="modal-alloc-bank" class="form-control" placeholder="예: 신한, 국민">
    </div>
    <div class="form-group">
      <label>입금</label>
      <input type="text" id="modal-alloc-deposit" class="form-control" placeholder="예: 국민, 키움">
    </div>
    <div class="form-group">
      <label>이체일자</label>
      <input type="text" id="modal-alloc-day" class="form-control" placeholder="예: 22일, 16일">
    </div>
    <div class="form-group" style="display: flex; align-items: center; gap: 8px;">
      <input type="checkbox" id="modal-alloc-isAutoTransfer" class="form-check-input" style="width:1.2rem; height:1.2rem;">
      <label for="modal-alloc-isAutoTransfer" style="margin-bottom:0;">자동이체 여부</label>
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
    const rawAmt = document.getElementById('modal-alloc-amount').value.replace(/,/g, '');
    const amount = Number(rawAmt) || 0;
    const bank = document.getElementById('modal-alloc-bank').value.trim();
    const deposit = document.getElementById('modal-alloc-deposit').value.trim();
    let day = document.getElementById('modal-alloc-day').value.trim();
    const isAutoTransfer = document.getElementById('modal-alloc-isAutoTransfer').checked;
    const category = document.getElementById('modal-alloc-category').value;

    if (!name) { alert('항목명을 입력하세요.'); return; }
    if (day) day = formatDayInput(day);

    const dayVal = day || '-';
    const methodVal = isAutoTransfer ? '자동이체' : '';
    const schedule = `${dayVal === '-' ? '' : dayVal} ${methodVal}`.trim() || '-';

    appState.allocations[personKey].push({
      id: 'a_' + Date.now(),
      name,
      amount,
      bank,
      deposit,
      day: dayVal,
      method: methodVal,
      isAutoTransfer,
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
      <label>금액 (원)</label>
      <input type="text" id="modal-fe-amount" class="form-control" placeholder="0 (숫자 입력 시 , 자동 적용)" oninput="this.value = this.value.replace(/[^0-9]/g, '').replace(/\B(?=(\d{3})+(?!\d))/g, ',')">
    </div>
    <div class="form-group">
      <label>출금처</label>
      <input type="text" id="modal-fe-bank" class="form-control" placeholder="예: 토스, 신한은행">
    </div>
    <div class="form-group">
      <label>입금처</label>
      <input type="text" id="modal-fe-deposit" class="form-control" placeholder="예: 국민은행, 관리사무소">
    </div>
    <div class="form-group">
      <label>이체일자</label>
      <input type="text" id="modal-fe-day" class="form-control" placeholder="예: 10일, 26일, 수기">
    </div>
    <div class="form-group form-check mt-2 mb-3">
      <input type="checkbox" class="form-check-input" id="modal-fe-autotransfer" checked>
      <label class="form-check-label" for="modal-fe-autotransfer">자동이체 여부</label>
    </div>
    <div class="form-group">
      <label>결제수단</label>
      <select id="modal-fe-method" class="form-select">
        ${(appState.categories?.paymentMethods || []).map(m => `<option value="${m}">${m}</option>`).join('')}
      </select>
    </div>
    <div class="form-group">
      <label>카테고리</label>
      <select id="modal-fe-category" class="form-select">
        ${(appState.categories?.fixedCategories || []).map(c => `<option value="${c}">${c}</option>`).join('')}
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
    const rawAmt = document.getElementById('modal-fe-amount').value.replace(/,/g, '');
    const amount = Number(rawAmt) || 0;
    const actualMarch = amount; // Merged with amount
    const bank = document.getElementById('modal-fe-bank').value.trim();
    const deposit = document.getElementById('modal-fe-deposit').value.trim();
    const day = document.getElementById('modal-fe-day').value.trim();
    const method = document.getElementById('modal-fe-method').value;
    const category = document.getElementById('modal-fe-category').value;
    const note = document.getElementById('modal-fe-note').value.trim();
    const isAutoTransfer = document.getElementById('modal-fe-autotransfer').checked;

    if (!name) { alert('항목명을 입력하세요.'); return; }

    appState.fixedExpenses.push({
      id: 'fe_' + Date.now(),
      name,
      amount,
      actualMarch,
      bank,
      deposit,
      day,
      method,
      category,
      note,
      isPaid: false,
      isAutoTransfer
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

    const calcs = calculateTotals();
    if (calcs.remainingBalance > 0) {
      catMap['남은 잔액'] = calcs.remainingBalance;
    }

    const sortedEntries = Object.entries(catMap).sort((a, b) => b[1] - a[1]);
    
    const totalAmount = sortedEntries.reduce((sum, e) => sum + e[1], 0);
    const originalCats = sortedEntries.map(e => e[0]);
    const labels = sortedEntries.map(e => {
      const percentage = totalAmount > 0 ? Math.round((e[1] / totalAmount) * 100) : 0;
      return `${e[0]} ${percentage}% (${formatKRW(e[1])})`;
    });
    const dataValues = sortedEntries.map(e => e[1]);
    
    const bgColors = originalCats.map((cat, index) => {
      const badgeClass = getAllocCategoryBadgeClass(cat);
      switch(badgeClass) {
        case 'badge-info': return '#06b6d4';
        case 'badge-success': return '#10b981';
        case 'badge-primary': return '#6366f1';
        case 'badge-warning': return '#f59e0b';
        case 'badge-danger': return '#ef4444';
        case 'badge-salary': return '#ec4899';
        case 'badge-total': return '#8b5cf6';
        case 'badge-remain': return '#84cc16';
        case 'badge-gyewon': return '#fb7185';
        case 'badge-dongwook': return '#38bdf8';
        case 'badge-secondary': 
        default: 
          // fallback palette for duplicates
          const fallbackColors = ['#94a3b8', '#cbd5e1', '#64748b', '#475569', '#334155'];
          return fallbackColors[index % fallbackColors.length];
      }
    });

    if (categoryChartInstance) categoryChartInstance.destroy();

    categoryChartInstance = new Chart(pieCtx, {
      type: 'doughnut',
      data: {
        labels: labels,
        datasets: [{
          data: dataValues,
          backgroundColor: bgColors,
          borderWidth: 0
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        layout: { padding: { top: 15 } },
        plugins: {
          legend: { 
            position: 'right', 
            labels: { 
              font: { size: 11 },
              generateLabels: function(chart) {
                const data = chart.data;
                if (data.labels.length && data.datasets.length) {
                  return data.labels.map((label, i) => {
                    const meta = chart.getDatasetMeta(0);
                    const style = meta.controller.getStyle(i);
                    return {
                      text: label,
                      fillStyle: style.backgroundColor,
                      strokeStyle: style.borderColor,
                      lineWidth: style.borderWidth,
                      hidden: isNaN(data.datasets[0].data[i]) || meta.data[i].hidden,
                      index: i,
                      fontColor: '#ffffff',
                      color: '#ffffff'
                    };
                  });
                }
                return [];
              }
            } 
          },
          tooltip: {
            callbacks: {
              label: (ctx) => `${ctx.label}`
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
      let pm = item.method || '기타';
      pm = pm.trim();
      pmMap[pm] = (pmMap[pm] || 0) + Number(item.amount || 0);
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

  // 4. Fixed Category Chart
  const fcCtx = document.getElementById('fixedCategoryChart');
  if (fcCtx) {
    const fcMap = {};
    appState.fixedExpenses.forEach(item => {
      const cat = item.category || '기타';
      fcMap[cat] = (fcMap[cat] || 0) + Number(item.amount || 0);
    });

    if (fixedCategoryChartInstance) fixedCategoryChartInstance.destroy();

    const fcLabels = Object.keys(fcMap);
    const fcData = Object.values(fcMap);
    const fcColors = fcLabels.map(cat => {
      if (cat.includes('용돈')) return '#f472b6'; // pink
      if (cat.includes('곗돈') || cat.includes('곗')) return '#fbbf24'; // amber
      if (cat.includes('구독')) return '#38bdf8'; // sky blue
      if (cat.includes('보험')) return '#818cf8'; // indigo
      if (cat.includes('핸드폰') || cat.includes('통신')) return '#34d399'; // emerald
      if (cat.includes('주거') || cat.includes('관리비')) return '#fb7185'; // rose
      if (cat.includes('교통')) return '#a3e635'; // lime
      return '#94a3b8'; // gray
    });

    fixedCategoryChartInstance = new Chart(fcCtx, {
      type: 'doughnut',
      data: {
        labels: fcLabels,
        datasets: [{
          data: fcData,
          backgroundColor: fcColors,
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
              label: function(context) {
                const val = context.raw;
                const total = context.dataset.data.reduce((a,b)=>a+b,0);
                const pct = total > 0 ? ((val/total)*100).toFixed(1) : 0;
                return ` ${context.label}: ${formatKRW(val)}원 (${pct}%)`;
              }
            }
          }
        },
        cutout: '70%'
      }
    });
  }
}

// Savings & Investment Projection
let projectionChartInstance = null;

function renderSavingsProjection() {
  const calcs = calculateTotals();
  const allAllocs = [
    ...(appState.allocations?.gyewon || []).map(i => ({ ...i, pathStr: 'allocations.gyewon' })), 
    ...(appState.allocations?.dongwook || []).map(i => ({ ...i, pathStr: 'allocations.dongwook' }))
  ];

  const allocCats = appState.categories?.allocCategories || ['생활비', '저축/적금', '투자/연금', '대출/이자', '비상금/경조사'];
  
  if (!appState.simulationCategories || appState.simulationCategories.length === 0) {
    appState.simulationCategories = [allocCats[1], allocCats[2], allocCats[4]].filter(Boolean);
  }

  const filterContainer = document.getElementById('simulation-category-filters');
  if (filterContainer) {
    filterContainer.innerHTML = '';
    allocCats.forEach((cat) => {
      if (!cat) return;
      const isChecked = appState.simulationCategories.includes(cat);
      const div = document.createElement('div');
      div.className = 'form-check form-check-inline';
      div.innerHTML = `
        <input class="form-check-input" type="checkbox" id="sim-cat-${cat}" value="${cat}" ${isChecked ? 'checked' : ''}>
        <label class="form-check-label" for="sim-cat-${cat}" style="color:var(--text-color); cursor:pointer; margin-left: 5px;">
          <span class="badge ${getAllocCategoryBadgeClass(cat)}">${cat}</span>
        </label>
      `;
      const cb = div.querySelector('input');
      cb.addEventListener('change', (e) => {
        if (e.target.checked) {
          appState.simulationCategories.push(cat);
        } else {
          appState.simulationCategories = appState.simulationCategories.filter(c => c !== cat);
        }
        saveState(); // this will trigger renderAll -> renderSavingsProjection
      });
      filterContainer.appendChild(div);
    });
  }

  // Gather items by category
  const allItems = allAllocs
    .filter(i => appState.simulationCategories.includes(i.category))
    .map(i => ({ ...i, catLabel: i.category }));

  const monthlySavings = allItems.filter(i => i.category === allocCats[1]).reduce((sum, i) => sum + (Number(i.amount) || 0), 0);
  const monthlyInvest = allItems.filter(i => i.category === allocCats[2]).reduce((sum, i) => sum + (Number(i.amount) || 0), 0);
  const monthlyEmergency = allItems.filter(i => i.category === allocCats[4]).reduce((sum, i) => sum + (Number(i.amount) || 0), 0);
  const monthlyTotal = allItems.reduce((sum, i) => sum + (Number(i.amount) || 0), 0);

  // Update KPI
  const simSavingsEl = document.getElementById('sim-monthly-savings');
  const simInvestEl = document.getElementById('sim-monthly-invest');
  const simEmergencyEl = document.getElementById('sim-monthly-emergency');
  const simTotalEl = document.getElementById('sim-monthly-total');
  if (simSavingsEl) simSavingsEl.innerText = formatKRW(monthlySavings);
  if (simInvestEl) simInvestEl.innerText = formatKRW(monthlyInvest);
  if (simEmergencyEl) simEmergencyEl.innerText = formatKRW(monthlyEmergency);
  if (simTotalEl) simTotalEl.innerText = formatKRW(monthlyTotal);

  // Render detail table
  const detailTbody = document.querySelector('#table-savings-detail tbody');
  if (detailTbody) {
    detailTbody.innerHTML = '';

    const targetDateInput = document.getElementById('simulation-target-date');
    let targetDateStr = targetDateInput ? targetDateInput.value : '';
    const baseMonthStr = appState.currentMonth || '2026-03';
    const [baseY, baseM] = baseMonthStr.split('-').map(Number);
    
    if (!targetDateStr) {
      const nextY = baseY + Math.floor((baseM + 12 - 1) / 12);
      const nextM = ((baseM + 12 - 1) % 12) + 1;
      targetDateStr = `${nextY}-${String(nextM).padStart(2, '0')}`;
      if (targetDateInput) targetDateInput.value = targetDateStr;
    }
    
    const [targetY, targetM] = targetDateStr.split('-').map(Number);
    let monthsDiff = (targetY - baseY) * 12 + (targetM - baseM);
    if (monthsDiff < 0) monthsDiff = 0;
    
    const simTableLabel = document.getElementById('sim-table-target-label');
    if (simTableLabel) simTableLabel.innerText = `${targetY}년 ${targetM}월`;

    const simTargetLabel = document.getElementById('sim-target-date-label');
    if (simTargetLabel) simTargetLabel.innerText = `${targetY}년 ${targetM}월`;

    let currentAssets = 0;
    allItems.forEach(item => {
      const amt = Number(item.amount) || 0;
      const currentBal = Number(item.currentBalance) || 0;
      currentAssets += currentBal;

      const dateText = item.currentBalanceDate || '-';
      const dayText = item.day && item.day !== '-' ? item.day : '-';
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>${item.name}</td>
        <td style="text-align:center;"><span class="text-muted" style="font-size:0.9em; font-weight:600;">${dayText}</span></td>
        <td><span class="badge ${getAllocCategoryBadgeClass(item.catLabel)}">${item.catLabel}</span></td>
        <td class="cell-amount"><div class="editable-cell" title="클릭하여 기납입액 수정" onclick="editInlineCell(this, '${item.pathStr}', '${item.id}', 'currentBalance', 'number')">${formatKRW(currentBal)}</div></td>
        <td style="text-align:center;"><div class="editable-cell text-muted" title="클릭하여 기준일자 입력 (예: 2024.05)" style="font-size: 0.85rem;" onclick="editInlineCell(this, '${item.pathStr}', '${item.id}', 'currentBalanceDate', 'text')">${dateText}</div></td>
        <td class="cell-amount">${formatKRW(amt)}</td>
        <td class="cell-amount" style="font-weight:700; color:var(--accent-primary);">${formatKRW(amt * monthsDiff + currentBal)}</td>
      `;
      detailTbody.appendChild(tr);
    });

    // Add total row
    if (allItems.length > 0) {
      const totalTr = document.createElement('tr');
      totalTr.style.borderTop = '2px solid var(--border-color)';
      totalTr.style.fontWeight = '700';
      totalTr.innerHTML = `
        <td colspan="3" style="text-align:right;">합계</td>
        <td class="cell-amount">${formatKRW(currentAssets)}</td>
        <td></td>
        <td class="cell-amount">${formatKRW(monthlyTotal)}</td>
        <td class="cell-amount" style="color:var(--accent-primary);">${formatKRW(monthlyTotal * monthsDiff + currentAssets)}</td>
      `;
      detailTbody.appendChild(totalTr);
    }
  }

  const currentAssets = allItems.reduce((sum, item) => sum + (Number(item.currentBalance) || 0), 0);
  
  // Re-calculate target date here again for KPI and Chart
  const targetDateInputOut = document.getElementById('simulation-target-date');
  let targetDateStrOut = targetDateInputOut ? targetDateInputOut.value : '';
  const baseMonthStrOut = appState.currentMonth || '2026-03';
  const [baseYOut, baseMOut] = baseMonthStrOut.split('-').map(Number);
  
  if (!targetDateStrOut) {
    const nextY = baseYOut + Math.floor((baseMOut + 12 - 1) / 12);
    const nextM = ((baseMOut + 12 - 1) % 12) + 1;
    targetDateStrOut = `${nextY}-${String(nextM).padStart(2, '0')}`;
  }
  const [targetYOut, targetMOut] = targetDateStrOut.split('-').map(Number);
  let monthsDiffOut = (targetYOut - baseYOut) * 12 + (targetMOut - baseMOut);
  if (monthsDiffOut < 0) monthsDiffOut = 0;

  const yearlyTotalExpected = (monthlyTotal * monthsDiffOut) + currentAssets;
  const yearlyTotalEl = document.getElementById('sim-yearly-total-assets');
  if (yearlyTotalEl) {
    yearlyTotalEl.innerText = formatKRW(yearlyTotalExpected) + ' 원';
  }

  // Render Projection Chart
  const projCtx = document.getElementById('projectionChart');
  if (projCtx) {
    if (projectionChartInstance) projectionChartInstance.destroy();

    const labels = [];
    const savingsData = [];
    const investData = [];
    const emergencyData = [];
    const totalData = [];

    const maxChartMonths = Math.min(Math.max(monthsDiffOut, 1), 120); 
    for (let m = 1; m <= maxChartMonths; m++) {
      labels.push(`${m}개월`);
      savingsData.push(monthlySavings * m);
      investData.push(monthlyInvest * m);
      emergencyData.push(monthlyEmergency * m);
      totalData.push(currentAssets + (monthlyTotal * m));
    }

    projectionChartInstance = new Chart(projCtx, {
      type: 'line',
      data: {
        labels: labels,
        datasets: [
          {
            label: allocCats[1],
            data: savingsData,
            borderColor: '#10b981',
            backgroundColor: 'rgba(16, 185, 129, 0.1)',
            fill: true,
            tension: 0.3,
            pointRadius: 4,
            pointHoverRadius: 6
          },
          {
            label: allocCats[2],
            data: investData,
            borderColor: '#6366f1',
            backgroundColor: 'rgba(99, 102, 241, 0.1)',
            fill: true,
            tension: 0.3,
            pointRadius: 4,
            pointHoverRadius: 6
          },
          {
            label: allocCats[4],
            data: emergencyData,
            borderColor: '#f59e0b',
            backgroundColor: 'rgba(245, 158, 11, 0.1)',
            fill: true,
            tension: 0.3,
            pointRadius: 4,
            pointHoverRadius: 6
          },
          {
            label: '예상 총 자산 (현재자산 포함)',
            data: totalData,
            borderColor: '#ec4899',
            backgroundColor: 'rgba(236, 72, 153, 0.05)',
            borderDash: [5, 5],
            fill: false,
            tension: 0.3,
            pointRadius: 4,
            pointHoverRadius: 6,
            borderWidth: 2
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: { intersect: false, mode: 'index' },
        plugins: {
          legend: { position: 'top', labels: { color: '#94a3b8', font: { size: 11 } } },
          tooltip: {
            callbacks: {
              label: (ctx) => ` ${ctx.dataset.label}: ${formatKRW(ctx.raw)}`
            }
          }
        },
        scales: {
          x: {
            ticks: { color: '#64748b' },
            grid: { color: 'rgba(148, 163, 184, 0.1)' }
          },
          y: {
            ticks: {
              color: '#64748b',
              callback: function(value) {
                if (value >= 10000) return Math.floor(value / 10000).toLocaleString() + '만';
                return value.toLocaleString() + '원';
              }
            },
            grid: { color: 'rgba(148, 163, 184, 0.1)' }
          }
        }
      }
    });
  }
}

// Snapshot Monthly Records


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
