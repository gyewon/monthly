const SUPABASE_URL = 'https://bdnqlcrpytkwuaonhgmm.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImJkbnFsY3JweXRrd3Vhb25oZ21tIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODk5MjYyMDEsImV4cCI6MjEwNTUwMjIwMX0.QLt4HgVyCoAR2oy1R5O3SdNe7N3XtVxP7rjabS3j_ew';

async function run() {
  const fetchRes = await fetch(`${SUPABASE_URL}/rest/v1/monthly_budget_data?id=eq.main`, {
    headers: {
      'apikey': SUPABASE_ANON_KEY,
      'Authorization': `Bearer ${SUPABASE_ANON_KEY}`
    }
  });
  
  const fetchedData = await fetchRes.json();
  const appState = fetchedData[0].data;
  
  let changed = false;
  
  // 1. Update categories
  if (appState.categories && appState.categories.incomeItems) {
    appState.categories.incomeItems = appState.categories.incomeItems.map(item => {
      if (item === '기본급 (급여)') {
        changed = true;
        return '급여';
      }
      return item;
    });
  }
  
  // 2. Update incomes
  if (appState.incomes && Array.isArray(appState.incomes)) {
    appState.incomes.forEach(inc => {
      if (inc.category === '기본급 (급여)') {
        inc.category = '급여';
        changed = true;
      }
      if (inc.title === '기본급 (급여)') {
        inc.title = '급여';
        changed = true;
      }
    });
  }
  
  if (!changed) {
    console.log("No '기본급 (급여)' found to update.");
  }
  
  const patchRes = await fetch(`${SUPABASE_URL}/rest/v1/monthly_budget_data?id=eq.main`, {
    method: 'PATCH',
    headers: {
      'apikey': SUPABASE_ANON_KEY,
      'Authorization': `Bearer ${SUPABASE_ANON_KEY}`,
      'Content-Type': 'application/json',
      'Prefer': 'return=minimal'
    },
    body: JSON.stringify({
      data: appState,
      updated_at: new Date().toISOString()
    })
  });
  
  if (!patchRes.ok) {
    console.error("Failed to patch", await patchRes.text());
  } else {
    console.log("Successfully renamed '기본급 (급여)' to '급여' in Supabase.");
  }
}

run();
