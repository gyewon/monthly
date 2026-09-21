import codecs

file_path = 'app.js'
with codecs.open(file_path, 'r', encoding='euc-kr') as f:
    content = f.read()

target = """  document.getElementById('kpi-income-gyewon').innerText = formatCompactKRW(calcs.incomeGyewon);
  document.getElementById('kpi-income-dongwook').innerText = formatCompactKRW(calcs.incomeDongwook);"""

replacement = """  // Generate Income Breakdown by Category
  const incomeByCategory = {};
  (appState.incomes || []).forEach(inc => {
    const cat = inc.title || '기타수입';
    const amt = Number(inc.amount) || 0;
    incomeByCategory[cat] = (incomeByCategory[cat] || 0) + amt;
  });

  const breakdownHTML = Object.entries(incomeByCategory)
    .sort((a, b) => b[1] - a[1])
    .map(([cat, amt]) => `<span class="badge badge-info">${cat} ${formatCompactKRW(amt)}</span>`)
    .join(' ');
  
  const breakdownElem = document.getElementById('kpi-income-breakdown');
  if (breakdownElem) {
    breakdownElem.innerHTML = breakdownHTML;
  }"""

if target in content:
    content = content.replace(target, replacement)
    with codecs.open(file_path, 'w', encoding='euc-kr') as f:
        f.write(content)
    print("Success")
else:
    print("Target not found")
