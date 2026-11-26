'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { execFileSync } = require('child_process');

const ROOT = __dirname;
const HTML = path.join(ROOT, 'index.html');
const extractedPath = path.join(require('os').tmpdir(), 'schedule-to-form1-script.js');

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
  const listeners = {};
  return Object.assign({
    id,
    value: '',
    classList: { add() {}, remove() {}, contains() { return false; } },
    style: {},
    files: [],
    innerHTML: '',
    onclick: null,
    ondragover: null,
    ondragleave: null,
    ondrop: null,
    onchange: null,
    checked: false,
    addEventListener(type, fn) {
      (listeners[type] || (listeners[type] = [])).push(fn);
    },
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
    document,
    console,
    setTimeout,
    clearTimeout,
    Number,
    String,
    Array,
    Map,
    Set,
    Object,
    Math,
    JSON,
    parseInt,
    parseFloat,
    isNaN,
    Infinity,
    Date,
    RegExp,
    Error,
    Uint8Array,
    Blob: function () {},
    FileReader: function () {},
    XLSX: {}
  };
  sandbox.window = sandbox;
  sandbox.globalThis = sandbox;
  sandbox.self = sandbox;
  vm.createContext(sandbox);
  vm.runInContext(script + `
    globalThis.__API = {
      stripPhotoPathLine,
      filenameFromPhotoPath,
      parsePhotoFilename,
      isStandardPhotoViewType,
      analyzePhotoPaths,
      collectRemarkMismatches,
      compareRemarkWording,
      splitRemarks,
      sortTreeNos,
      uniqueTreeNos,
      photoRemarksTitle,
      bestDefect
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

const { api, els } = loadApi();
let passed = 0;
function test(name, fn) {
  fn();
  passed++;
  console.log('ok  ' + name);
}

test('quoted Windows paths yield tree no and remark', () => {
  const line = '"R:\\Photos\\T3_Wound on trunk (1).JPG"';
  const stripped = api.stripPhotoPathLine(line);
  assert(stripped === 'R:\\Photos\\T3_Wound on trunk (1).JPG', 'unquote failed: ' + stripped);
  const fn = api.filenameFromPhotoPath(stripped);
  assert(fn === 'T3_Wound on trunk (1).JPG', 'filename: ' + fn);
  const p = api.parsePhotoFilename(fn);
  assert(p && p.treeNo === 'T3', 'treeNo: ' + (p && p.treeNo));
  assert(p.rest === 'Wound on trunk', 'rest: ' + p.rest);
});

test('TG-T and T and T- numbers are kept as written', () => {
  const a = api.parsePhotoFilename('TG-T0175_Co-dominant branches.jpg');
  assert(a && a.treeNo === 'TG-T0175' && a.rest === 'Co-dominant branches', JSON.stringify(a));
  const b = api.parsePhotoFilename('TG-T0001_Tag.jpg');
  assert(b && b.treeNo === 'TG-T0001' && b.rest === 'Tag', JSON.stringify(b));
  const c = api.parsePhotoFilename('T-2225_Leaning.jpg');
  assert(c && c.treeNo === 'T-2225' && c.rest === 'Leaning', JSON.stringify(c));
  const d = api.parsePhotoFilename('T28A_Wilted.JPG');
  assert(d && d.treeNo === 'T28A' && d.rest === 'Wilted', JSON.stringify(d));
  const e = api.parsePhotoFilename('T117a_Wound.JPG');
  assert(e && e.treeNo === 'T117a', JSON.stringify(e));
  const f = api.parsePhotoFilename('T28_Wound.JPG');
  assert(f && f.treeNo === 'T28', JSON.stringify(f));
});

test('trailing (1)/(2) stripped for wording grouping; T148 (1)/(2) same tree', () => {
  const a = api.parsePhotoFilename('T3_Wound on trunk (1).JPG');
  const b = api.parsePhotoFilename('T3_Wound on trunk (2).JPG');
  assert(a.rest === 'Wound on trunk' && b.rest === 'Wound on trunk', a.rest + ' / ' + b.rest);
  const c = api.parsePhotoFilename('T148 (1)_Wound.JPG');
  const d = api.parsePhotoFilename('T148 (2)_Wound.JPG');
  assert(c.treeNo === 'T148' && d.treeNo === 'T148', c.treeNo + ' / ' + d.treeNo);
  const analysis = api.analyzePhotoPaths([
    '"C:\\\\Pics\\\\T3_Wound on trunk (1).JPG"',
    '"C:\\\\Pics\\\\T3_Wound on trunk (2).JPG"',
    'T148 (1)_Wound.JPG',
    'T148 (2)_Wound.JPG'
  ].join('\n'));
  const woundTrunk = analysis.rows.find(r => r.item === 'Wound on trunk');
  const wound = analysis.rows.find(r => r.item === 'Wound');
  assert(woundTrunk && woundTrunk.nos.length === 1 && woundTrunk.nos[0] === 'T3' && woundTrunk.photos === 2, JSON.stringify(woundTrunk));
  assert(wound && wound.nos.length === 1 && wound.nos[0] === 'T148' && wound.photos === 2, JSON.stringify(wound));
});

test('Tag / Whole view / Trunk base / Crown (and Wholeview) skipped', () => {
  const paths = [
    'T1_Tag.jpg',
    'T2_Whole view.JPG',
    'T3_Wholeview.jpg',
    'T4_Trunk base.jpg',
    'T5_Trunkbase.JPG',
    'T6_Crown.JPG',
    'T7_Wound on trunk.JPG'
  ].join('\n');
  const a = api.analyzePhotoPaths(paths);
  assert(a.photoCount === 7, 'photoCount ' + a.photoCount);
  assert(a.rows.length === 1 && a.rows[0].item === 'Wound on trunk', JSON.stringify(a.rows));
  assert(a.skippedView === 6, 'skippedView ' + a.skippedView);
});

test('filename with only tree number contributes no remark wording', () => {
  const a = api.analyzePhotoPaths('T527.JPG\nT28.JPG\n/home/user/photos/T6.JPG');
  assert(a.photoCount === 3, 'photoCount ' + a.photoCount);
  assert(a.rows.length === 0, 'rows ' + JSON.stringify(a.rows));
  assert(a.noRemark === 3, 'noRemark ' + a.noRemark);
  assert(a.unparsed.length === 0, 'unparsed ' + JSON.stringify(a.unparsed));
});

test('comma-split remarks via splitRemarks', () => {
  const a = api.analyzePhotoPaths('T3_Wound on trunk, Leaning, N.A..JPG');
  const items = a.rows.map(r => r.item);
  assert(items.includes('Wound on trunk') && items.includes('Leaning'), JSON.stringify(items));
  assert(!items.some(x => /n\.?a/i.test(x)), 'N.A. should be excluded: ' + JSON.stringify(items));
  assert(a.rows.every(r => r.nos[0] === 'T3' && r.photos === 1), JSON.stringify(a.rows));
});

test('English A-Z sort with case tie-break (leaning before Leaning)', () => {
  const items = ['Leaning', 'leaning', 'Wound on trunk', 'Bending', 'asymmetric crown'];
  const sorted = items.slice().sort(api.compareRemarkWording);
  deepEqual(sorted, ['asymmetric crown', 'Bending', 'leaning', 'Leaning', 'Wound on trunk']);
  const a = api.analyzePhotoPaths([
    'T2_Leaning.JPG',
    'T1_leaning.JPG',
    'T3_Wound on trunk.JPG',
    'T4_Bending.JPG',
    'T5_asymmetric crown.JPG'
  ].join('\n'));
  deepEqual(a.rows.map(r => r.item), ['asymmetric crown', 'Bending', 'leaning', 'Leaning', 'Wound on trunk']);
});

test('analysis works with paths only (no trees/Excel)', () => {
  els.photoPaths.value = '';
  const a = api.analyzePhotoPaths([
    '\uFEFF"R:\\Photos\\TG-T0175_Co-dominant branches.jpg",',
    '',
    '/unix/photos/T28A_Wilted.JPG',
    'T28_Dead branch.JPG'
  ].join('\n'));
  assert(a.rows.length === 3, JSON.stringify(a.rows));
  deepEqual(a.rows.map(r => r.item), ['Co-dominant branches', 'Dead branch', 'Wilted']);
  deepEqual(a.rows.map(r => r.nos), [['TG-T0175'], ['T28'], ['T28A']]);
  assert(!els.out.innerHTML, 'should not require writing a schedule report');
});

test('unparsed lines are noted and not guessed', () => {
  const a = api.analyzePhotoPaths([
    'T3_Wound.JPG',
    'README.md',
    'just a note',
    'AT-0001_Tag.jpg',
    'C:\\\\folder\\\\'
  ].join('\n'));
  assert(a.rows.length === 1 && a.rows[0].item === 'Wound', JSON.stringify(a.rows));
  assert(a.unparsed.length === 4, 'unparsed ' + JSON.stringify(a.unparsed));
  assert(a.unparsed.includes('README.md'), 'missing README.md');
  assert(a.unparsed.includes('just a note'), 'missing note');
  assert(a.unparsed.includes('AT-0001_Tag.jpg'), 'AT-0001 should not be guessed');
});

test('schedule remarks table is unchanged when both sources exist', () => {
  const trees = [
    { no: 'T1', remarks: 'Leaning, Wound on trunk' },
    { no: 'T2', remarks: 'leaning' }
  ];
  const scheduleOnly = api.collectRemarkMismatches(trees);
  const photo = api.analyzePhotoPaths('T9_Bending.JPG\nT1_Cavity on trunk.JPG');
  const scheduleAfter = api.collectRemarkMismatches(trees);
  deepEqual(scheduleAfter, scheduleOnly);
  assert(scheduleOnly.map(r => r.item).join('|') === 'leaning|Leaning|Wound on trunk', scheduleOnly.map(r => r.item).join('|'));
  assert(photo.rows.map(r => r.item).join('|') === 'Bending|Cavity on trunk', photo.rows.map(r => r.item).join('|'));
  const scheduleItems = new Set(scheduleOnly.map(r => r.item + '@' + r.nos.join(',')));
  for (const r of photo.rows) {
    assert(!scheduleItems.has(r.item + '@' + r.nos.join(',')), 'photo wording should not be merged into schedule: ' + r.item);
  }
});

test('T28 and T28A stay distinct; photo counts do not replace tree counts', () => {
  const a = api.analyzePhotoPaths([
    'T28_Wound.JPG',
    'T28A_Wound.JPG',
    'T28_Wound (2).JPG'
  ].join('\n'));
  assert(a.rows.length === 1, JSON.stringify(a.rows));
  assert(a.rows[0].nos.length === 2, '棵數 should be unique trees: ' + a.rows[0].nos);
  assert(a.rows[0].photos === 3, '張數 ' + a.rows[0].photos);
  deepEqual(a.rows[0].nos, ['T28', 'T28A']);
});

test('blank lines, trailing commas, mixed slashes and BOM are tolerated', () => {
  const text = '\uFEFF"R:\\A\\B\\T1_Crack.JPG",\r\n\r\n/unix/C/T2_Crack.jpg,\n';
  const a = api.analyzePhotoPaths(text);
  assert(a.unparsed.length === 0, JSON.stringify(a.unparsed));
  assert(a.rows.length === 1 && a.rows[0].item === 'Crack' && a.rows[0].photos === 2, JSON.stringify(a.rows));
});

test('title style matches schedule remarks wording', () => {
  const a = api.analyzePhotoPaths('T1_Leaning.JPG\nT2_not a standard phrase.JPG');
  const title = api.photoRemarksTitle(a.rows);
  assert(/^全部相片 Remarks（2 項用字，其中 1 項不是標準 description）$/.test(title), title);
  const exact = api.analyzePhotoPaths('T1_Leaning.JPG');
  const t2 = api.photoRemarksTitle(exact.rows);
  assert(t2 === '全部相片 Remarks（1 項用字）', t2);
});

console.log('\n' + passed + ' tests passed');
