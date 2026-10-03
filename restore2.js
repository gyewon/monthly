const fs = require('fs');
let code = fs.readFileSync('app.js', 'utf8');
const DEFAULT_EXCEL_DATA_str = code.substring(code.indexOf('const DEFAULT_EXCEL_DATA = {'), code.indexOf('let supabaseClient'));

const SUPABASE_URL = 'https://bdnqlcrpytkwuaonhgmm.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImJkbnFsY3JweXRrd3Vhb25oZ21tIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODk5MjYyMDEsImV4cCI6MjEwNTUwMjIwMX0.QLt4HgVyCoAR2oy1R5O3SdNe7N3XtVxP7rjabS3j_ew';

const fetchStr = `
${DEFAULT_EXCEL_DATA_str}

fetch('${SUPABASE_URL}/rest/v1/monthly_budget_data?id=eq.main', {
  method: 'PATCH',
  headers: {
    'apikey': '${SUPABASE_ANON_KEY}',
    'Authorization': 'Bearer ${SUPABASE_ANON_KEY}',
    'Content-Type': 'application/json',
    'Prefer': 'return=minimal'
  },
  body: JSON.stringify({
    data: DEFAULT_EXCEL_DATA,
    updated_at: new Date().toISOString()
  })
}).then(r => {
  if (r.ok) console.log('Successfully restored DEFAULT_EXCEL_DATA to Supabase.');
  else console.error('Failed to restore');
});
`;

eval(fetchStr);
