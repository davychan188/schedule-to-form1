
'use strict';

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { execFileSync } = require('child_process');

const ROOT = __dirname;
const HTML = path.join(ROOT, 'index.html');
const extractedPath = path.join(require('os').tmpdir(), 'schedule-to-form1-similar-script.js');

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
      areSimilarDefects,
      collectSimilarDefects,
      collectSimilarDefectsFromPhotoAnalysis,
      collectRemarkMismatches,
      collectPostureConflicts,
      analyzePhotoPaths,
      splitRemarks,
      sortTreeNos,
      similarDefectTitle,
      similarDefectWarnHtml,
      similarDefectSheet,
      formatSimilarWordings,
      uniqueRemarkItems,
      speciesBreakdown,
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
function wordingsOf(row) {
  if (!row) return [];
  return (row.items || []).slice().sort((a, b) => String(a).localeCompare(String(b), 'en'));
}

const { api } = loadApi();
let passed = 0;
function test(name, fn) {
  fn();
  passed++;
  console.log('ok  ' + name);
}

function scheduleTree(no, remarks, extra) {
  return Object.assign({ no, sci: 'Ficus microcarpa', chi: '細葉榕', remarks }, extra || {});
}

test('leaning + severely leaning (user example) warns', () => {
  const rows = api.collectSimilarDefects([
    scheduleTree('T1', 'Leaning, Severely leaning')
  ]);
  assert(rows.length === 1 && rows[0].no === 'T1', JSON.stringify(rows));
  deepEqual(wordingsOf(rows[0]), ['Leaning', 'Severely leaning']);
  assert(rows[0].wordings.includes('Leaning') && rows[0].wordings.includes('Severely leaning'), rows[0].wordings);
});

test('leaning + slightly leaning warns', () => {
  const rows = api.collectSimilarDefects([
    scheduleTree('T2', 'Leaning, Slightly leaning')
  ]);
  assert(rows.length === 1 && rows[0].no === 'T2', JSON.stringify(rows));
  deepEqual(wordingsOf(rows[0]), ['Leaning', 'Slightly leaning']);
});

test('sparse foliage + severely sparse foliage warns', () => {
  const rows = api.collectSimilarDefects([
    scheduleTree('T3', 'Sparse foliage, Severely sparse foliage')
  ]);
  assert(rows.length === 1, JSON.stringify(rows));
  deepEqual(wordingsOf(rows[0]), ['Severely sparse foliage', 'Sparse foliage']);
});

test('Dead branch + Dead branches + Minor dead branch warn', () => {
  const rows = api.collectSimilarDefects([
    scheduleTree('T4', 'Dead branch, Dead branches, Minor dead branch')
  ]);
  assert(rows.length === 1, JSON.stringify(rows));
  deepEqual(wordingsOf(rows[0]), ['Dead branch', 'Dead branches', 'Minor dead branch']);
});

test('Wound on trunk vs Wound on branch does NOT warn', () => {
  const rows = api.collectSimilarDefects([
    scheduleTree('T5', 'Wound on trunk, Wound on branch')
  ]);
  assert(rows.length === 0, JSON.stringify(rows));
  assert(!api.areSimilarDefects('Wound on trunk', 'Wound on branch'), 'pair should not be similar');
});

test('Co-dominant branches vs Co-dominant trunks does NOT warn', () => {
  const rows = api.collectSimilarDefects([
    scheduleTree('T6', 'Co-dominant branches, Co-dominant trunks')
  ]);
  assert(rows.length === 0, JSON.stringify(rows));
});

test('Restricted root growth + Trunk in conflict with hoarding does NOT warn', () => {
  const rows = api.collectSimilarDefects([
    scheduleTree('T7', 'Restricted root growth, Trunk in conflict with hoarding')
  ]);
  assert(rows.length === 0, JSON.stringify(rows));
});

test('bending + leaning does NOT appear in this new warning (still only posture check)', () => {
  const trees = [scheduleTree('T8', 'Bending, Leaning')];
  const similar = api.collectSimilarDefects(trees);
  const posture = api.collectPostureConflicts(trees);
  assert(similar.length === 0, 'similar should be empty: ' + JSON.stringify(similar));
  assert(posture.length === 1 && posture[0].no === 'T8', JSON.stringify(posture));
  assert((posture[0].terms || []).join(',') === 'Bending,Leaning', JSON.stringify(posture[0].terms));
});

test('photo-path only (no Excel) still warns for the same tree', () => {
  const analysis = api.analyzePhotoPaths([
    'T9_Leaning.JPG',
    'T9_Severely leaning.JPG',
    'T10_Tag.JPG',
    'T10_Wound on trunk.JPG'
  ].join('\n'));
  const rows = api.collectSimilarDefectsFromPhotoAnalysis(analysis);
  assert(rows.length === 1 && rows[0].no === 'T9', JSON.stringify(rows));
  deepEqual(wordingsOf(rows[0]), ['Leaning', 'Severely leaning']);
});

test('original wording preserved; no auto-rename', () => {
  const unofficial = 'severely leaning';
  const rows = api.collectSimilarDefects([
    scheduleTree('T11', 'Leaning, ' + unofficial)
  ]);
  assert(rows.length === 1, JSON.stringify(rows));
  assert(rows[0].items.includes(unofficial), 'unofficial wording lost: ' + JSON.stringify(rows[0].items));
  assert(!rows[0].items.some(x => x !== 'Leaning' && x !== unofficial), 'renamed: ' + JSON.stringify(rows[0].items));
  assert(rows[0].wordings.indexOf(unofficial) !== -1, rows[0].wordings);
  const official = api.bestDefect(unofficial);
  assert(official.type !== 'exact', 'severely leaning should stay unofficial: ' + JSON.stringify(official));
});

test('report totals unchanged: remarks stay separate and are not merged', () => {
  const trees = [
    scheduleTree('T12', 'Leaning, Severely leaning, Wound on trunk'),
    scheduleTree('T13', 'Bending, Leaning')
  ];
  const remarks = api.collectRemarkMismatches(trees);
  const similar = api.collectSimilarDefects(trees);
  const remarksAfter = api.collectRemarkMismatches(trees);
  deepEqual(remarksAfter, remarks);
  const leaning = remarks.find(r => r.item === 'Leaning');
  const severe = remarks.find(r => r.item === 'Severely leaning');
  const wound = remarks.find(r => r.item === 'Wound on trunk');
  const bending = remarks.find(r => r.item === 'Bending');
  assert(leaning && leaning.nos.join() === 'T12,T13', JSON.stringify(leaning));
  assert(severe && severe.nos.join() === 'T12', JSON.stringify(severe));
  assert(wound && wound.nos.join() === 'T12', JSON.stringify(wound));
  assert(bending && bending.nos.join() === 'T13', JSON.stringify(bending));
  assert(similar.length === 1 && similar[0].no === 'T12', 'only T12 has similar defects: ' + JSON.stringify(similar));
  const species = api.speciesBreakdown(trees);
  const again = api.speciesBreakdown(trees);
  deepEqual(again, species);
  assert(species.length === 1 && species[0].n === 2, JSON.stringify(species));
});

test('exact or case-only duplicates are not flagged as similar', () => {
  const rows = api.collectSimilarDefects([
    scheduleTree('T14', 'Leaning, Leaning'),
    scheduleTree('T15', 'Leaning, leaning')
  ]);
  assert(rows.length === 0, JSON.stringify(rows));
});

test('schedule and photo sources are compared separately', () => {
  const trees = [scheduleTree('T16', 'Leaning, Severely leaning')];
  const schedule = api.collectSimilarDefects(trees);
  const photos = api.collectSimilarDefectsFromPhotoAnalysis(api.analyzePhotoPaths(
    'T16_Dead branch.JPG\nT16_Minor dead branch.JPG'
  ));
  assert(schedule.length === 1 && schedule[0].wordings.includes('Leaning'), JSON.stringify(schedule));
  assert(!schedule[0].wordings.toLowerCase().includes('dead branch'), schedule[0].wordings);
  assert(photos.length === 1 && photos[0].wordings.toLowerCase().includes('dead branch'), JSON.stringify(photos));
  assert(!photos[0].wordings.toLowerCase().includes('leaning'), photos[0].wordings);
});

test('Dead branch_close up is similar to Dead branch; trailing photo index ignored', () => {
  const photos = api.collectSimilarDefectsFromPhotoAnalysis(api.analyzePhotoPaths([
    'T17_Dead branch (1).JPG',
    'T17_Dead branch_close up (2).JPG'
  ].join('\n')));
  assert(photos.length === 1 && photos[0].no === 'T17', JSON.stringify(photos));
  assert(photos[0].items.includes('Dead branch'), JSON.stringify(photos[0].items));
  assert(photos[0].items.includes('Dead branch_close up'), JSON.stringify(photos[0].items));
});

test('co-dominant branches with included bark is similar to co-dominant branches', () => {
  const rows = api.collectSimilarDefects([
    scheduleTree('T18', 'Co-dominant branches, Co-dominant branches with included bark')
  ]);
  assert(rows.length === 1, JSON.stringify(rows));
});

test('Cavity on trunk vs Cavity on roots does NOT warn', () => {
  const rows = api.collectSimilarDefects([
    scheduleTree('T19', 'Cavity on trunk, Cavity on roots')
  ]);
  assert(rows.length === 0, JSON.stringify(rows));
});

test('Removed trees are still included in the warning', () => {
  const rows = api.collectSimilarDefects([
    scheduleTree('T20', 'Removed, Leaning, Slightly leaning')
  ]);
  assert(rows.length === 1 && rows[0].no === 'T20', JSON.stringify(rows));
  assert(rows[0].items.includes('Leaning') && rows[0].items.includes('Slightly leaning'), JSON.stringify(rows[0].items));
  assert(!rows[0].items.includes('Removed'), JSON.stringify(rows[0].items));
});

test('bending + leaning + slightly leaning only flags the leaning variants', () => {
  const trees = [scheduleTree('T21', 'Bending, Leaning, Slightly leaning')];
  const similar = api.collectSimilarDefects(trees);
  const posture = api.collectPostureConflicts(trees);
  assert(similar.length === 1, JSON.stringify(similar));
  deepEqual(wordingsOf(similar[0]), ['Leaning', 'Slightly leaning']);
  assert(!(similar[0].items || []).includes('Bending'), JSON.stringify(similar[0].items));
  assert(posture.length === 1 && (posture[0].terms || []).includes('Bending'), JSON.stringify(posture));
});

test('title and html keep a warning-only note and do not suggest official names', () => {
  const rows = api.collectSimilarDefects([scheduleTree('T22', 'Leaning, Severely leaning')]);
  const title = api.similarDefectTitle(rows, []);
  assert(title === '請核對：1 棵樹的 Remarks 同時出現類似 defect', title);
  const html = api.similarDefectWarnHtml(rows, []);
  assert(html.includes('同一棵樹出現超過一項類似 defect；報告用字及統計不會自動更改。'), html);
  assert(!/最接近的表內 description/.test(html), 'suggestion column should not appear');
  const sheet = api.similarDefectSheet(rows, []);
  assert(sheet && sheet.name === '類似Defect', JSON.stringify(sheet && sheet.name));
  assert(!api.similarDefectSheet([], []), 'empty sheet should be omitted');
});

test('photo-only title uses photo rows; both sources share one tree count', () => {
  const schedule = api.collectSimilarDefects([scheduleTree('T23', 'Leaning, Slightly leaning')]);
  const photos = api.collectSimilarDefectsFromPhotoAnalysis(api.analyzePhotoPaths(
    'T23_Dead branch.JPG\nT23_Minor dead branch.JPG\nT24_Sparse foliage.JPG\nT24_Severely sparse foliage.JPG'
  ));
  const title = api.similarDefectTitle(schedule, photos);
  assert(title === '請核對：2 棵樹的 Remarks 同時出現類似 defect', title);
  const html = api.similarDefectWarnHtml(schedule, photos);
  assert(html.includes('Schedule Remarks') && html.includes('相片路徑 Remarks'), html);
});

console.log('\n' + passed + ' tests passed');
