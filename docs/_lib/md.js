const fs = require('fs');
const B = '☐';
function Doc() {
  const L = [];
  const api = {
    B,
    p: s => { L.push(s); return api; },
    tbl: (h, r) => { L.push('| ' + h.join(' | ') + ' |');
      L.push('|' + h.map(() => '---').join('|') + '|');
      r.forEach(x => L.push('| ' + x.map(c => String(c).replace(/\n/g, '<br>')).join(' | ') + ' |'));
      L.push(''); return api; },
    save: f => { fs.writeFileSync(f, L.join('\n')); console.log('md written', f, L.length, 'lines'); },
  };
  return api;
}
module.exports = { Doc, B };
