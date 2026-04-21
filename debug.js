const h = require('fs').readFileSync('C:/Users/El-Wattaneya/Desktop/basset/index.html','utf8');
// Check what's actually in the dashboard section
const idx = h.indexOf('xl:grid-cols-6');
if (idx > -1) {
  console.log('FOUND grid-cols-6 at:', idx);
  console.log(h.substring(idx-30, idx+300));
} else {
  console.log('NOT FOUND grid-cols-6');
  // Try to find the stat cards
  const idx2 = h.indexOf('statTotalSales');
  console.log('statTotalSales at:', idx2);
  if(idx2>-1) console.log(h.substring(idx2-200, idx2+100));
}
