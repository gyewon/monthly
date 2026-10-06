# -*- coding: utf-8 -*-
js = open('app.js', encoding='utf-8').read()
newFuncs = '''
window.deleteSimItem = function(id) {
  if(confirm('이 시뮬레이션 항목을 삭제하시겠습니까?')) {
    appState.simulationItems = (appState.simulationItems || []).filter(i => i.id !== id);
    saveState(); // this re-renders
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
'''
if 'deleteSimItem' not in js:
    open('app.js', 'a', encoding='utf-8').write('\n' + newFuncs)
    print('Appended functions')
