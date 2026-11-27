'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { execFileSync } = require('child_process');

const ROOT = __dirname;
const HTML = path.join(ROOT, 'index.html');
const extractedPath = path.join(require('os').tmpdir(), 'schedule-to-form1-dead-species-script.js');

function extractInlineScript(html) {
  const re = /<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g;
  let m, last = '';
  while ((m = re.exec(html))) {
    const tag = m[0].slice(0, m[0].indexOf('>'));
    if (/\ssrc\s*=/.test(tag)) continue;
    last = m[1];
  }
  if (!last) throw new Error('No inline script found in index.html');
  return last;
}

function makeEl(id, extras) {
  return Object.assign({
    id,
    value: '',
    classList: { add() {}, remove() {}, contains() { return false; } },
    style: {},
    files: [],
    innerHTML: '',
    onclick: null,
    addEventListener() {},
    click() {},
    querySelector() { return null; },
    querySelectorAll() { return [] }
  }, extras || {});
}

function loadApi() {
  const html = fs.readFileSync(HTML, 'utf8');
  const script = extractInlineScript(html);
  fs.writeFileSync(extractedPath, script);
  execFileSync(process.execPath, ['--check', extractedPath], { stdio: 'pipe' });

  const els = {
    drop: makeEl('drop'),
    file: makeEl('file'),
    photoPathsBox: makeEl('photoPathsBox'),
    photoPaths: makeEl('photoPaths'),
    photoRemarksOut: makeEl('photoRemarksOut'),
    similarDefectOut: makeEl('similarDefectOut'),
    out: makeEl('out'),
    project: makeEl('project'),
    extraGroups: makeEl('extraGroups')
  };
  const radios = [
    Object.assign(makeEl('deadU'), { name: 'dead', value: 'Unidentified', checked: true }),
    Object.assign(makeEl('deadD'), { name: 'dead', value: 'Dead tree', checked: false }),
    Object.assign(makeEl('modeC'), { name: 'mode', value: 'construction', checked: true }),
    Object.assign(makeEl('modeP'), { name: 'mode', value: 'private', checked: false })
  ];
  const document = {
    getElementById(id) { return els[id] || makeEl(id); },
    querySelectorAll(sel) {
      if (String(sel).includes('name=dead') || String(sel).includes('name=mode')) return radios;
      return [];
    },
    querySelector() { return radios[0]; }
  };
  const sandbox = {
    document, console, setTimeout, clearTimeout,
    Number, String, Array, Map, Set, Object, Math, JSON,
    parseInt, parseFloat, isNaN, Infinity, Date, RegExp, Error, Uint8Array, TextEncoder,
    Blob: function () {}, FileReader: function () {}, XLSX: {}
  };
  sandbox.window = sandbox;
  sandbox.globalThis = sandbox;
  sandbox.self = sandbox;
  vm.createContext(sandbox);
  vm.runInContext(script + `
    globalThis.__API = {
      isDeadSpecies,
      isNonSpeciesRow,
      isItalicSci,
      countIdentifiedSpecies,
      applyDeadLabel,
      section33IntroText,
      section5IntroText,
      speciesBreakdown,
      speciesTableHtml,
      xTableRows,
      wTable,
      sciHtml,
      sciSpan,
      qtyLabel,
      formBGroups,
      formARows,
      formPasteSheets,
      formSpeciesText,
      buildSection4Body,
      buildGroupIntro,
      collectNewfoundOrigin,
      isNewfoundTree,
      isTprpGroupTree,
      isTprpIndividualTree,
      classifyKind,
      isTG,
      normTG
    };
  `, sandbox);
  return { api: sandbox.__API, els };
}

function assert(cond, msg) {
  if (!cond) throw new Error(msg);
}
function deepEqual(a, b, msg) {
  const sa = JSON.stringify(a);
  const sb = JSON.stringify(b);
  if (sa !== sb) throw new Error((msg || 'deepEqual failed') + '\n  actual:   ' + sa + '\n  expected: ' + sb);
}

function tree(no, sci, chi, extra) {
  return Object.assign({ no, sci, chi, remarks: '', rec: '', triage: '', zone: '', source: '', dbh: '', ht: '', spread: '' }, extra || {});
}

function parseSpeciesTable(html) {
  const rows = [];
  const re = /<tr>([\s\S]*?)<\/tr>/g;
  let m, i = 0;
  while ((m = re.exec(html))) {
    const cells = [...m[1].matchAll(/<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/g)].map(x => x[1]);
    if (!i) { i++; continue; } // header
    const num = String(cells[0] || '').replace(/<[^>]+>/g, '').trim();
    const sciRaw = cells[1] || '';
    const chi = String(cells[2] || '').replace(/<[^>]+>/g, '').trim();
    const qty = String(cells[3] || '').replace(/<[^>]+>/g, '').trim();
    const sci = sciRaw.replace(/<[^>]+>/g, '').trim();
    const italic = /<i>/.test(sciRaw);
    rows.push({ num, sci, chi, qty, italic, isTotal: num === 'Total' || sci === 'Total' || /^Total/.test(num) });
  }
  return rows;
}

function splitFormAB(active) {
  const formA = active.filter(t => /yellow|orange|red|black/i.test(t.triage) || (t.rec && !/^(n\.?a\.?|-|no mitigation)/i.test(t.rec)));
  const formB = active.filter(t => !formA.includes(t));
  return { formA, formB };
}

const { api } = loadApi();
let passed = 0;
function test(name, fn) {
  fn();
  passed++;
  console.log('ok  ' + name);
}

const ficus = (no, extra) => tree(no, 'Ficus microcarpa', '細葉榕', extra);
const aleurites = (no, extra) => tree(no, 'Aleurites moluccana', '石栗', extra);
const deadTree = (no, extra) => tree(no, 'Dead tree', '死樹', extra);
const unidentified = (no, extra) => tree(no, 'Unidentified', 'Unidentified', extra);

test('3.3 / Section 5 species count excludes Dead tree and Unidentified', () => {
  const active = [
    ficus('T1'), ficus('T2'), aleurites('T3'),
    deadTree('T4'), unidentified('T5'), deadTree('T6')
  ];
  const n = api.countIdentifiedSpecies(active);
  assert(n === 2, 'expected 2 identified species, got ' + n);
  const intro33 = api.section33IntroText(n, 3);
  assert(intro33.startsWith('A total of 2 species were recorded'), intro33);
  assert(intro33.includes('3 dead trees with unknown species were observed'), intro33);
  assert(!/A total of 3 species/.test(intro33) && !/A total of 4 species/.test(intro33), intro33);
  const intro5 = api.section5IntroText(6, n, 'Ficus microcarpa', '細葉榕 (2)', 3);
  assert(intro5.startsWith('A total of 6 trees and 2 species were recorded'), intro5);
  assert(intro5.includes('3 dead trees with unknown species were observed'), intro5);
});

test('species table has no numbered Dead tree / Unidentified row', () => {
  const active = [ficus('T1'), aleurites('T2'), deadTree('T3'), unidentified('T4')];
  const rows = api.speciesBreakdown(active);
  const html = api.speciesTableHtml(rows, active.length);
  const parsed = parseSpeciesTable(html);
  const data = parsed.filter(r => !r.isTotal);
  const numberedDead = data.filter(r => r.num !== '' && /dead tree|unidentified/i.test(r.sci + ' ' + r.chi));
  assert(numberedDead.length === 0, 'numbered dead rows: ' + JSON.stringify(numberedDead));
  const identified = data.filter(r => r.num !== '');
  assert(identified.length === 2, 'identified rows: ' + JSON.stringify(identified));
  assert(identified[0].num === '1' && identified[1].num === '2', JSON.stringify(identified));
  assert(identified.every(r => /Ficus|Aleurites/.test(r.sci)), JSON.stringify(identified));
});

test('unnumbered quantity line is present and table Total equals all trees', () => {
  const active = [ficus('T1'), ficus('T2'), deadTree('T3'), deadTree('T4'), unidentified('T5')];
  const rows = api.speciesBreakdown(active);
  assert(rows.filter(r => !r.nonSpecies).length === 1, JSON.stringify(rows));
  const deadLines = rows.filter(r => r.nonSpecies);
  assert(deadLines.length === 2, 'expected two unnumbered lines: ' + JSON.stringify(deadLines));
  assert(deadLines.every(r => api.isNonSpeciesRow(r)), JSON.stringify(deadLines));
  const sum = rows.reduce((n, r) => n + r.n, 0);
  assert(sum === 5, 'qty sum ' + sum);
  const html = api.speciesTableHtml(rows, active.length);
  const parsed = parseSpeciesTable(html);
  const total = parsed.find(r => r.isTotal);
  assert(total && total.qty === '5', JSON.stringify(total));
  const unnumbered = parsed.filter(r => !r.isTotal && r.num === '');
  assert(unnumbered.length === 2, JSON.stringify(unnumbered));
  assert(unnumbered.some(r => r.sci === 'Dead tree' && r.chi === '死樹' && r.qty === '2'), JSON.stringify(unnumbered));
  assert(unnumbered.some(r => r.sci === 'Unidentified' && r.qty === '1'), JSON.stringify(unnumbered));
  const xrows = api.xTableRows('', rows, active.length);
  const xData = xrows.filter(r => r[1] && r[1].v !== 'Scientific Name' && r[0].v !== 'Total');
  const xDead = xData.filter(r => r[0].v === '');
  const xLive = xData.filter(r => r[0].v !== '');
  assert(xLive.length === 1 && xLive[0][0].v === 1 && xLive[0][1].v === 'Ficus microcarpa', JSON.stringify(xLive));
  assert(xDead.length === 2, JSON.stringify(xDead));
  const xTotal = xrows.find(r => r[0] && r[0].v === 'Total');
  assert(xTotal && xTotal[3].v === 5, JSON.stringify(xTotal));
});

test('per-TG breakdown uses the same unnumbered dead rule', () => {
  const tg = [
    tree('TG726-1', 'Ficus microcarpa', '細葉榕', { zone: 'TG726' }),
    tree('TG726-2', 'Dead tree', '死樹', { zone: 'TG726' }),
    tree('TG726-3', 'Dead tree', '死樹', { zone: 'TG726' })
  ];
  const title = 'TG726 (' + tg.length + ' nos.)';
  assert(title === 'TG726 (3 nos.)', title);
  const rows = api.speciesBreakdown(tg);
  assert(rows.filter(r => !r.nonSpecies).length === 1, JSON.stringify(rows));
  assert(rows.find(r => r.sci === 'Ficus microcarpa').n === 1, JSON.stringify(rows));
  const dead = rows.find(r => r.nonSpecies);
  assert(dead && dead.sci === 'Dead tree' && dead.n === 2, JSON.stringify(rows));
  const html = api.speciesTableHtml(rows, tg.length);
  const parsed = parseSpeciesTable(html);
  assert(parsed.find(r => r.isTotal).qty === '3', JSON.stringify(parsed));
  assert(parsed.filter(r => !r.isTotal && r.num !== '').length === 1, JSON.stringify(parsed));
  assert(parsed.some(r => r.num === '' && r.sci === 'Dead tree' && r.qty === '2'), JSON.stringify(parsed));
});

test('Form 1 B / Sheet2: A+B totals unchanged; dead not counted as a species', () => {
  const active = [
    ficus('T1', { triage: 'Yellow', rec: 'Crown cleaning', ht: '8' }),
    aleurites('T2', { triage: '', rec: '', ht: '6' }),
    deadTree('T3', { triage: '', rec: '', ht: '5' }),
    unidentified('T4', { triage: '', rec: '', ht: '4' }),
    ficus('T5', { triage: '', rec: '', ht: '7' })
  ];
  const { formA, formB } = splitFormAB(active);
  assert(formA.length + formB.length === active.length, 'A+B ' + formA.length + '+' + formB.length);
  assert(formA.length === 1 && formB.length === 4, 'A=' + formA.length + ' B=' + formB.length);
  const groups = api.formBGroups(formB);
  const qty = groups.reduce((n, g) => n + g.n, 0);
  assert(qty === formB.length, 'Sheet2 qty ' + qty + ' vs B ' + formB.length);
  const identified = groups.filter(g => !g.nonSpecies && !api.isNonSpeciesRow(g));
  const dead = groups.filter(g => g.nonSpecies || api.isNonSpeciesRow(g));
  assert(identified.length === 2, 'form B species: ' + JSON.stringify(groups));
  assert(dead.length === 2, 'form B dead lines: ' + JSON.stringify(dead));
  assert(identified.every(g => /Ficus|Aleurites/.test(g.sci)), JSON.stringify(identified));
  assert(dead.every(g => /Dead tree|Unidentified/.test(g.sci)), JSON.stringify(dead));
  assert(groups[groups.length - 2].nonSpecies && groups[groups.length - 1].nonSpecies, 'dead lines should be last: ' + JSON.stringify(groups));
  const sheets = api.formPasteSheets({ formA, formB });
  assert(sheets[0].name === 'Sheet1' && sheets[1].name === 'Sheet2', sheets.map(s => s.name).join(','));
  const sheet2 = sheets[1].xml;
  assert(/<dimension ref="A1:E4"\/>/.test(sheet2), 'Sheet2 should be headerless A1 paste with 4 data rows, 5 cols: ' + sheet2.slice(0, 400));
  assert(!/Scientific Name|Tree Species/.test(sheet2), 'Sheet2 must stay headerless');
  const pasteQty = groups.reduce((n, g) => n + g.n, 0);
  assert(pasteQty === 4, pasteQty);
  assert(formA.length + pasteQty === 5, 'A+B tree total');
  assert(/Dead tree/.test(sheet2) && /Unidentified/.test(sheet2), sheet2);
});

test('identified species still italic; dead/unidentified names not italic', () => {
  const rows = api.speciesBreakdown([ficus('T1'), deadTree('T2'), unidentified('T3')]);
  const html = api.speciesTableHtml(rows, 3);
  assert(/<i>Ficus microcarpa<\/i>/.test(html), html);
  assert(!/<i>Dead tree<\/i>/.test(html), html);
  assert(!/<i>Unidentified<\/i>/.test(html), html);
  const parsed = parseSpeciesTable(html);
  const live = parsed.find(r => r.sci === 'Ficus microcarpa');
  const dead = parsed.find(r => r.sci === 'Dead tree');
  const unk = parsed.find(r => r.sci === 'Unidentified');
  assert(live && live.italic, JSON.stringify(live));
  assert(dead && !dead.italic, JSON.stringify(dead));
  assert(unk && !unk.italic, JSON.stringify(unk));
  assert(api.isItalicSci('Ficus microcarpa', '細葉榕'));
  assert(!api.isItalicSci('Dead tree', '死樹'));
  assert(!api.isItalicSci('Unidentified', 'Unidentified'));
  assert(!api.isItalicSci('枯樹', ''));
  const xrows = api.xTableRows('', rows, 3);
  const xLive = xrows.find(r => r[1] && r[1].v === 'Ficus microcarpa');
  const xDead = xrows.find(r => r[1] && r[1].v === 'Dead tree');
  assert(xLive && xLive[1].s === 5, JSON.stringify(xLive));
  assert(xDead && xDead[1].s === 4, JSON.stringify(xDead));
  const word = api.wTable(rows, 3);
  assert(/Ficus microcarpa/.test(word) && /<w:i\/>/.test(word), word.slice(0, 400));
  const deadChunk = word.split('Dead tree')[0].slice(-200);
  // the run for Dead tree should not include italic
  assert( /<w:t xml:space="preserve">Dead tree<\/w:t>/.test(word), word);
  const deadRun = word.match(/<w:r>(?:<w:rPr>[\s\S]*?<\/w:rPr>)?<w:t xml:space="preserve">Dead tree<\/w:t><\/w:r>/);
  assert(deadRun, 'missing Dead tree run');
  assert(!/<w:i\/>/.test(deadRun[0]), deadRun[0]);
  const liveRun = word.match(/<w:r>(?:<w:rPr>[\s\S]*?<\/w:rPr>)?<w:t xml:space="preserve">Ficus microcarpa<\/w:t><\/w:r>/);
  assert(liveRun && /<w:i\/>/.test(liveRun[0]), liveRun && liveRun[0]);
  const spanLive = api.sciSpan('Ficus microcarpa', '細葉榕');
  const spanDead = api.sciSpan('Dead tree', '死樹');
  assert(/class="sci"/.test(spanLive), spanLive);
  assert(!/class="sci"/.test(spanDead) && spanDead === 'Dead tree', spanDead);
});

test('Section 4 / groupIntro / Table 1 tree totals unchanged', () => {
  const active = [
    tree('T1', 'Ficus microcarpa', '細葉榕', { zone: 'IND', source: 'TPRP' }),
    tree('TG726-1', 'Aleurites moluccana', '石栗', { zone: 'TG726', source: 'TPRP' }),
    tree('N1', 'Dead tree', '死樹', { zone: 'N.A.', source: 'New' }),
    tree('N2', 'Ficus microcarpa', '細葉榕', { zone: 'N.A.', source: 'New' }),
    tree('N3', 'Unidentified', 'Unidentified', { zone: 'N.A.', source: 'New' })
  ];
  const indN = active.filter(t => api.isTprpIndividualTree(t)).length;
  const grpN = active.filter(t => api.isTprpGroupTree(t)).length;
  const newN = active.filter(t => api.isNewfoundTree(t)).length;
  const deadN = active.filter(t => api.isDeadSpecies(t.sci, t.chi)).length;
  const aliveN = active.length - deadN;
  assert(active.length === 5, 'total trees');
  assert(newN === 3, 'newly found should include dead: ' + newN);
  assert(deadN === 2 && aliveN === 3, 'alive/dead ' + aliveN + '/' + deadN);
  const s4 = api.buildSection4Body({
    projectPhrase: '', total: active.length, indN, grpN, newN, aliveN, deadN, removedN: 0, protectedChecks: []
  });
  assert(s4.includes('total of 5 nos. of surveyed tree'), s4);
  assert(s4.includes('3 alive trees and 2 dead trees'), s4);
  assert(s4.includes('3 (Alive) + 2 (Dead Trees) = 5.'), s4);
  const intro = api.buildGroupIntro(['TG726'], ['TG726'], active.length, indN, grpN, newN);
  assert(intro.includes('total of 5 nos. of surveyed tree'), intro);
  assert(intro.includes('3 nos. of trees newly found in the survey'), intro);
  const t1 = api.collectNewfoundOrigin(active);
  assert(t1.newN === 3, 'Table 1 newN ' + t1.newN);
  assert(!t1.rows.some(r => api.isDeadSpecies(r.sci, r.chi)), JSON.stringify(t1.rows));
  assert(t1.placeholderTrees === 2, 'placeholders ' + t1.placeholderTrees);
  const identifiedNew = t1.nativeTrees + t1.exoticTrees;
  assert(identifiedNew === 1, 'Table 1 identified trees ' + identifiedNew);
});

test('radio Dead tree → Unidentified remaps the label but still excludes it from species', () => {
  const raw = [
    { no: 'T1', sci: 'Dead tree', chi: '死樹' },
    { no: 'T2', sci: '枯樹', chi: '' },
    { no: 'T3', sci: '', chi: '' },
    { no: 'T4', sci: 'Ficus microcarpa', chi: '細葉榕' }
  ];
  const asUnid = raw.map(t => Object.assign({ remarks: '' }, t, api.applyDeadLabel(t.sci, t.chi, 'Unidentified')));
  assert(asUnid[0].sci === 'Unidentified' && asUnid[0].chi === 'Unidentified', JSON.stringify(asUnid[0]));
  assert(asUnid[1].sci === 'Unidentified', JSON.stringify(asUnid[1]));
  assert(asUnid[2].sci === 'Unidentified', JSON.stringify(asUnid[2]));
  assert(asUnid[3].sci === 'Ficus microcarpa', JSON.stringify(asUnid[3]));
  assert(api.countIdentifiedSpecies(asUnid) === 1, 'count after remap');
  const rows = api.speciesBreakdown(asUnid);
  assert(rows.length === 2, JSON.stringify(rows));
  assert(rows[0].sci === 'Ficus microcarpa' && !rows[0].nonSpecies, JSON.stringify(rows[0]));
  assert(rows[1].sci === 'Unidentified' && rows[1].nonSpecies && rows[1].n === 3, JSON.stringify(rows[1]));
  const intro = api.section33IntroText(api.countIdentifiedSpecies(asUnid), 3);
  assert(intro.startsWith('A total of 1 species were recorded'), intro);

  const asDead = raw.map(t => Object.assign({ remarks: '' }, t, api.applyDeadLabel(t.sci, t.chi, 'Dead tree')));
  assert(asDead[0].sci === 'Dead tree' && asDead[0].chi === '死樹', JSON.stringify(asDead[0]));
  assert(asDead[2].sci === 'Dead tree', JSON.stringify(asDead[2]));
  assert(api.countIdentifiedSpecies(asDead) === 1);
  const rowsD = api.speciesBreakdown(asDead);
  assert(rowsD[1].sci === 'Dead tree' && rowsD[1].chi === '死樹' && rowsD[1].nonSpecies && rowsD[1].n === 3, JSON.stringify(rowsD[1]));
});

test('a file with only identified species is unchanged (no extra blank dead line)', () => {
  const active = [ficus('T1'), aleurites('T2'), ficus('T3')];
  const rows = api.speciesBreakdown(active);
  assert(rows.length === 2, JSON.stringify(rows));
  assert(rows.every(r => !r.nonSpecies && !api.isNonSpeciesRow(r)), JSON.stringify(rows));
  const html = api.speciesTableHtml(rows, active.length);
  const parsed = parseSpeciesTable(html);
  const data = parsed.filter(r => !r.isTotal);
  assert(data.length === 2, JSON.stringify(data));
  assert(data.every(r => r.num !== ''), 'unexpected blank line: ' + JSON.stringify(data));
  assert(parsed.find(r => r.isTotal).qty === '3', JSON.stringify(parsed));
  assert(!/Dead tree|Unidentified/.test(html), html);
  const xrows = api.xTableRows('', rows, active.length);
  const xData = xrows.filter(r => r[1] && r[1].v !== 'Scientific Name' && r[0].v !== 'Total');
  assert(xData.length === 2 && xData.every(r => r[0].v !== ''), JSON.stringify(xData));
  assert(!xData.some(r => /Dead tree|Unidentified/i.test(String(r[1].v))), JSON.stringify(xData));
  const intro33 = api.section33IntroText(api.countIdentifiedSpecies(active), 0);
  assert(intro33.startsWith('A total of 2 species were recorded'), intro33);
  assert(!intro33.includes('dead trees with unknown species'), intro33);
  const { formA, formB } = splitFormAB(active);
  const groups = api.formBGroups(formB);
  assert(groups.every(g => !g.nonSpecies), JSON.stringify(groups));
  assert(formA.length + formB.length === 3);
});

test('死樹 / 枯樹 labels are excluded from species even without remapping', () => {
  const active = [
    ficus('T1'),
    tree('T2', '死樹', '死樹'),
    tree('T3', '枯樹', '枯樹')
  ];
  assert(api.countIdentifiedSpecies(active) === 1);
  const rows = api.speciesBreakdown(active);
  assert(rows[0].sci === 'Ficus microcarpa');
  assert(rows.slice(1).every(r => r.nonSpecies), JSON.stringify(rows));
});

console.log('passed ' + passed + ' tests');
