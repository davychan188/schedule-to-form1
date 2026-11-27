'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { execFileSync } = require('child_process');

const ROOT = __dirname;
const HTML = path.join(ROOT, 'index.html');
const extractedPath = path.join(require('os').tmpdir(), 'schedule-to-form1-spelling-script.js');

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
    querySelectorAll() { return []; }
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
    photoSpellingOut: makeEl('photoSpellingOut'),
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
    parseInt, parseFloat, isNaN, Infinity, Date, RegExp, Error, Uint8Array,
    Blob: function () {}, FileReader: function () {}, XLSX: {}
  };
  sandbox.window = sandbox;
  sandbox.globalThis = sandbox;
  sandbox.self = sandbox;
  vm.createContext(sandbox);
  vm.runInContext(script + `
    globalThis.__API = {
      editDistance,
      analyzePhotoPaths,
      collectPhotoRemarkSpellings,
      findPhotoRemarkTypos,
      suggestPhotoSpellCorrections,
      isExactOfficialRemark,
      isPhotoSpellKnownWord,
      photoSpellingTitle,
      photoSpellingWarnHtml,
      photoSpellingSheet,
      photoRemarksHtml,
      withPhotoRemarks,
      checkSheets,
      renderPhotoRemarks,
      lastReportData: () => lastReportData,
      setLastReportData(v) { lastReportData = v; }
    };
  `, sandbox);
  return { api: sandbox.__API, els, sandbox };
}

function assert(cond, msg) {
  if (!cond) throw new Error(msg);
}
function deepEqual(a, b, msg) {
  const sa = JSON.stringify(a);
  const sb = JSON.stringify(b);
  if (sa !== sb) throw new Error((msg || 'deepEqual failed') + '\n  actual:   ' + sa + '\n  expected: ' + sb);
}

const { api, els } = loadApi();
let passed = 0;
function test(name, fn) {
  fn();
  passed++;
  console.log('ok  ' + name);
}

test('node --check already ran on the extracted inline script', () => {
  assert(fs.existsSync(extractedPath), 'extracted script missing');
});

test('Lage cavity on trunk base flags Lage and suggests Large', () => {
  const flags = api.findPhotoRemarkTypos('Lage cavity on trunk base');
  assert(flags.length === 1, 'flags ' + JSON.stringify(flags));
  assert(flags[0].typo === 'Lage', 'typo: ' + flags[0].typo);
  assert(/\bLarge\b/.test(flags[0].suggestion), 'suggestion: ' + flags[0].suggestion);
  const a = api.analyzePhotoPaths('T3_Lage cavity on trunk base.JPG');
  const rows = api.collectPhotoRemarkSpellings(a);
  assert(rows.length === 1, JSON.stringify(rows));
  assert(rows[0].item === 'Lage cavity on trunk base', rows[0].item);
  assert(rows[0].typo === 'Lage' && /\bLarge\b/.test(rows[0].suggestion), JSON.stringify(rows[0]));
  assert(rows[0].nos[0] === 'T3' && rows[0].photos === 1, JSON.stringify(rows[0]));
  assert(a.rows[0].item === 'Lage cavity on trunk base', 'must keep original wording: ' + a.rows[0].item);
});

test('correct Large cavity on trunk base and exact official remarks are not flagged', () => {
  deepEqual(api.findPhotoRemarkTypos('Large cavity on trunk base'), []);
  deepEqual(api.findPhotoRemarkTypos('Wound on trunk'), []);
  deepEqual(api.findPhotoRemarkTypos('Cavity on trunk'), []);
  deepEqual(api.findPhotoRemarkTypos('Leaning'), []);
  deepEqual(api.findPhotoRemarkTypos('Co-dominant branches'), []);
  deepEqual(api.findPhotoRemarkTypos('Slightly leaning'), []);
  const a = api.analyzePhotoPaths([
    'T1_Large cavity on trunk base.JPG',
    'T2_Wound on trunk.JPG',
    'T3_Cavity on trunk.JPG'
  ].join('\n'));
  deepEqual(api.collectPhotoRemarkSpellings(a), []);
});

test('simple plurals and official wording are not typos', () => {
  const items = [
    'Dead branches',
    'Co-dominant trunks',
    'Abnormal bark cracks on trunk',
    'Wounds on trunk',
    'Cavity on roots',
    'cavities on trunk',
    'Crack on root',
    'Fungal fruiting bodies on trunk base'
  ];
  for (const item of items) {
    const flags = api.findPhotoRemarkTypos(item);
    assert(flags.length === 0, item + ' -> ' + JSON.stringify(flags));
  }
});

test('trailing (1)/(2), quoting, and skipped labels are not false positives', () => {
  const text = [
    '"R:\\\\Photos\\\\T3_Wound on trunk (1).JPG"',
    '"R:\\\\Photos\\\\T3_Wound on trunk (2).JPG"',
    'T4_Tag.jpg',
    'T5_Whole view.JPG',
    'T6_Wholeview.jpg',
    'T7_Trunk base.jpg',
    'T8_Crown.JPG',
    'T9_Lage cavity on trunk base (1).JPG',
    'T9_Lage cavity on trunk base (2).JPG'
  ].join('\n');
  const a = api.analyzePhotoPaths(text);
  const items = a.rows.map(r => r.item).sort();
  assert(items.includes('Wound on trunk'), JSON.stringify(items));
  assert(items.includes('Lage cavity on trunk base'), JSON.stringify(items));
  assert(!items.includes('Tag') && !items.includes('Whole view') && !items.includes('Trunk base') && !items.includes('Crown'), JSON.stringify(items));
  const rows = api.collectPhotoRemarkSpellings(a);
  assert(rows.length === 1, JSON.stringify(rows));
  assert(rows[0].typo === 'Lage' && rows[0].item === 'Lage cavity on trunk base', JSON.stringify(rows[0]));
  assert(rows[0].photos === 2 && rows[0].nos.length === 1 && rows[0].nos[0] === 'T9', JSON.stringify(rows[0]));
  const official = a.rows.find(r => r.item === 'Wound on trunk');
  assert(official && official.photos === 2, JSON.stringify(official));
});

test('photo-only paths with no schedule still warn', () => {
  els.out.innerHTML = '';
  api.setLastReportData(null);
  els.photoPaths.value = [
    '"C:\\\\Pics\\\\T12_Lage cavity on trunk base.JPG"',
    'T13_Wound on trunk.JPG'
  ].join('\n');
  api.renderPhotoRemarks();
  const html = els.photoSpellingOut.innerHTML;
  assert(html.includes('photoSpellingBox'), 'missing warning box: ' + html);
  assert(html.includes('Lage'), html);
  assert(html.includes('Large'), html);
  assert(html.includes('Lage cavity on trunk base'), html);
  assert(html.includes('T12'), html);
  assert(!html.includes('最接近的表內 description'), 'forbidden column');
  assert(!els.out.innerHTML, 'should not require or write a schedule report');
  const rows = api.collectPhotoRemarkSpellings(api.analyzePhotoPaths(els.photoPaths.value));
  const title = api.photoSpellingTitle(rows);
  assert(title === '請核對：1 項相片 Remarks 英文串字', title);
});

test('no warning box or excel sheet when there are no typo hits', () => {
  const a = api.analyzePhotoPaths('T1_Wound on trunk.JPG\nT2_Large cavity on trunk base.JPG');
  const rows = api.collectPhotoRemarkSpellings(a);
  deepEqual(rows, []);
  assert(api.photoSpellingWarnHtml(rows) === '', 'html should be empty');
  assert(api.photoSpellingSheet(rows) == null, 'sheet should be omitted');
  els.photoPaths.value = 'T1_Wound on trunk.JPG';
  api.setLastReportData(null);
  api.renderPhotoRemarks();
  assert(!els.photoSpellingOut.innerHTML, 'box should be omitted: ' + els.photoSpellingOut.innerHTML);
});

test('Chinese, tree numbers, extensions and unparsed lines are not spelling targets', () => {
  const a = api.analyzePhotoPaths([
    'T3_死樹.JPG',
    'README.md',
    'just a note',
    'AT-0001_Tag.jpg',
    'T4_Wound on trunk.JPG'
  ].join('\n'));
  assert(a.unparsed.length >= 3, JSON.stringify(a.unparsed));
  const rows = api.collectPhotoRemarkSpellings(a);
  deepEqual(rows, []);
  const flags = api.findPhotoRemarkTypos('死樹 T3 JPG');
  deepEqual(flags, []);
});

test('unknown but plausible long words are not flagged; scientific names are not flagged', () => {
  deepEqual(api.findPhotoRemarkTypos('Photosynthesis on trunk'), []);
  deepEqual(api.findPhotoRemarkTypos('Ficus microcarpa wound on trunk'), []);
  deepEqual(api.findPhotoRemarkTypos('Sign of Phauda flammans infestation'), []);
});

test('common location words close to lexicon words are not false positives', () => {
  for (const item of ['Wound on left side', 'Crack from trunk', 'Conflict with wall', 'Roof conflict']) {
    const flags = api.findPhotoRemarkTypos(item);
    assert(flags.length === 0, item + ' -> ' + JSON.stringify(flags));
  }
});

test('original remark wording is kept; suggestions are not applied to the A-Z table', () => {
  const a = api.analyzePhotoPaths('T8_Lage cavity on trunk base.JPG\nT8_Wound on trunk.JPG');
  deepEqual(a.rows.map(r => r.item), ['Lage cavity on trunk base', 'Wound on trunk']);
  const rows = api.collectPhotoRemarkSpellings(a);
  assert(rows[0].item === 'Lage cavity on trunk base', rows[0].item);
  assert(a.rows.find(r => r.item === 'Lage cavity on trunk base').item !== 'Large cavity on trunk base');
});

test('editDistance 1 insertion catches Lage vs Large', () => {
  assert(api.editDistance('lage', 'large') === 1, String(api.editDistance('lage', 'large')));
  assert(api.editDistance('large', 'large') === 0);
});

test('withPhotoRemarks and checkSheets include typo sheet only when there are hits', () => {
  els.photoPaths.value = 'T3_Lage cavity on trunk base.JPG';
  const rep = api.withPhotoRemarks({ project: 'Test' });
  assert((rep.photoRemarkSpellings || []).length === 1, JSON.stringify(rep.photoRemarkSpellings));
  const sheets = api.checkSheets(rep);
  const names = sheets.map(s => s && s.name);
  assert(names.includes('相片Remarks串字'), JSON.stringify(names));
  assert(names.includes('相片Remarks用字'), JSON.stringify(names));
  assert(!JSON.stringify(sheets).includes('最接近的表內 description'), 'forbidden wording in sheets');

  els.photoPaths.value = 'T3_Wound on trunk.JPG';
  const clean = api.withPhotoRemarks({ project: 'Test' });
  deepEqual(clean.photoRemarkSpellings, []);
  const cleanSheets = api.checkSheets(clean);
  const cleanNames = cleanSheets.map(s => s && s.name);
  assert(!cleanNames.includes('相片Remarks串字'), JSON.stringify(cleanNames));
});

test('title and warning html stay in the remarks-warning family', () => {
  const rows = [{ item: 'Lage cavity on trunk base', typo: 'Lage', suggestion: 'Large', nos: ['T3'], photos: 2 }];
  const html = api.photoSpellingWarnHtml(rows);
  assert(html.includes('class="warn noprint"'), html);
  assert(html.includes('請核對：1 項相片 Remarks 英文串字'), html);
  assert(html.includes('不會自動更改'), html);
  assert(html.includes('疑似錯字'), html);
  assert(html.includes('建議寫法'), html);
});

console.log('\n' + passed + ' tests passed');
