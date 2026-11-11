import { readDeck, inferOperation, applyOperation } from './deck.js';

const workspace = document.querySelector('#workspace');
const savedKey = 'vernacular.operations.v1';
const samplePaths = {
  original: 'fixtures/launch-original.pptx',
  revision: 'fixtures/launch-revised.pptx',
  test: 'fixtures/press-preview.pptx'
};

function loadLibrary() {
  try {
    const entries = JSON.parse(localStorage.getItem(savedKey) || '[]');
    return Array.isArray(entries) ? entries : [];
  } catch {
    return [];
  }
}

const state = {
  view: 'teach',
  busy: false,
  samples: {},
  decks: {},
  operation: null,
  result: null,
  date: '2026-11-21',
  notice: null,
  library: loadLibrary()
};

const escape = value => String(value ?? '').replace(/[&<>"']/g, character => ({
  '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
}[character]));
const disabled = condition => condition ? ' disabled' : '';
const capitalize = value => value ? value[0].toUpperCase() + value.slice(1) : '';
const operationName = operation => 'Change ' + operation.field.toLowerCase();
const sampleDeck = key => state.decks[key] && state.decks[key] === state.samples[key];

function button(action, label, { primary = false, wide = false, blocked = false } = {}) {
  return '<button class="button' + (primary ? ' primary' : '') + (wide ? ' wide' : '') +
    '" type="button" data-action="' + action + '"' + disabled(state.busy || blocked) + '>' + label + '</button>';
}

function upload(key, label) {
  return '<label class="button upload">' + label + '<input type="file" aria-label="' + label + ': ' + key + '" data-upload="' + key +
    '" accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation"' +
    disabled(state.busy) + '></label>';
}

function heading(title, subtitle, action = '') {
  return '<div class="page-heading"><div><h1 tabindex="-1">' + title + '</h1><p>' + subtitle +
    '</p></div>' + action + '</div>';
}

function notice() {
  return state.notice
    ? '<p class="notice' + (state.notice.success ? ' success' : '') + '" role="' +
        (state.notice.success ? 'status' : 'alert') + '">' + escape(state.notice.text) + '</p>'
    : '';
}

function preview(deck, { sample = false, title, date, rehearsal, changed = false, result = false } = {}) {
  if (!deck) return '<div class="deck-preview empty">Choose a .pptx file</div>';
  if (sample) {
    return '<div class="deck-preview art' + (result ? ' result' : '') + '"><p class="slide-brand">FIELDNOTES AUDIO</p>' +
      '<p class="slide-title">' + escape(title) + '</p><div><p class="slide-date">Launch date: ' +
      (changed ? '<mark>' + escape(date) + '</mark>' : escape(date)) +
      '</p><p class="slide-other">Rehearsal date: ' + escape(rehearsal) + '</p></div></div>';
  }
  const lines = deck.slides ? deck.slides[0].paragraphs.map(paragraph => paragraph.text).filter(Boolean) : deck.preview;
  const slideNumber = deck.slide ?? 1;
  return '<div class="deck-preview text-preview"><h3>Slide ' + slideNumber + ' text</h3>' +
    (lines.length ? '<ul>' + lines.slice(0, 12).map(line => '<li>' + escape(line) + '</li>').join('') + '</ul>' : '<p>This slide has no text.</p>') +
    '</div>';
}

function examplePane(key, title, proposed) {
  const deck = state.decks[key];
  return '<article><div class="pane-header"><h3>' + title + '</h3>' + upload(key, deck ? 'Replace file' : 'Choose file') +
    '</div>' + preview(deck, { sample: sampleDeck(key), title: 'Launch briefing', date: proposed ? 'Nov 19' : 'Nov 12', rehearsal: 'Nov 12', changed: proposed }) +
    '<p class="file-name">' + escape(deck?.name) + '</p></article>';
}

function teachView() {
  const op = state.operation;
  const draft = op
    ? '<dl class="rule"><div><dt>Operation</dt><dd>' + escape(operationName(op)) +
      '</dd></div><div><dt>Field</dt><dd>' + escape(capitalize(op.field)) + ' · slide ' + op.example.slide +
      '</dd></div><div><dt>Example</dt><dd>' + escape(op.example.before) + ' → ' + escape(op.example.after) +
      '</dd></div><div><dt>Preserve</dt><dd>Other text, dates, and slide content</dd></div></dl>' +
      button('go-test', 'Test on another file <span aria-hidden="true">→</span>', { primary: true, wide: true })
    : '<p>Choose an original deck and a copy with one labeled date changed.</p>';
  return heading('Teach from examples', 'Show Vernacular a change you made in PowerPoint.', button('sample', 'Load sample project')) + notice() +
    '<div class="split"><div class="main-pane"><div class="comparison">' +
    examplePane('original', 'Original', false) + examplePane('revision', 'Your revision', true) +
    '</div><div class="example-links"><a href="' + samplePaths.original + '" download>Download sample original</a>' +
    '<a href="' + samplePaths.revision + '" download>Download sample revision</a></div></div>' +
    '<aside class="inspector"><h2>Proposed operation</h2>' + draft + '</aside></div>';
}

function testView() {
  const { operation: op, result, decks } = state;
  if (!op) {
    return heading('Test on a fresh file', 'Try the operation on a deck you did not use to teach it.') +
      '<div class="empty"><p>Teach an operation before running a test.</p>' + button('go-teach', 'Go to Teach', { primary: true }) + '</div>';
  }
  const test = decks.test;
  const dates = result?.preservedDates.slice(0, 6) ?? [];
  const rows = result
    ? '<table class="results"><thead><tr><th scope="col">Field</th><th scope="col">Before</th><th scope="col">After</th></tr></thead><tbody>' +
      '<tr><th scope="row">' + escape(capitalize(op.field)) + '</th><td>' + escape(result.before) +
      '</td><td class="changed">' + escape(result.after) + '</td></tr>' +
      dates.map(date => '<tr><th scope="row">' + escape(date.field) + '</th><td>' + escape(date.value) +
        '</td><td>' + escape(date.value) + '</td></tr>').join('') + '</tbody></table>' : '';
  const evidence = result
    ? '<ul class="checklist"><li><span class="check" aria-hidden="true">✓</span><span>' + escape(capitalize(op.field)) + ' updated on slide ' + result.slide +
      '</span></li><li><span class="check" aria-hidden="true">✓</span><span>' +
      (dates.length ? 'Other dates preserved' : 'Other slide content preserved') + '</span></li>' +
      '<li><span class="check" aria-hidden="true">✓</span><span>' + result.untouchedParts +
      ' other file parts unchanged</span></li></ul>' +
      button('download', 'Download editable .pptx', { wide: true }) +
      '<p class="help-text">Open the copy in PowerPoint to inspect its layout.</p>' +
      button('save', 'Save operation', { primary: true, wide: true, blocked: result.saved })
    : '<p>Choose a test deck and a new date. Vernacular will check the editable copy after making the change.</p>' +
      button('run', state.busy ? 'Running test…' : 'Run test', { primary: true, wide: true, blocked: !test || !state.date });
  return heading('Test on a fresh file', 'Try the operation on a deck you did not use to teach it.', button('sample-test', 'Use sample test deck')) +
    notice() + '<div class="split"><div class="main-pane"><div class="test-controls"><div class="action-row">' +
    upload('test', test ? 'Replace test file' : 'Choose test file') +
    '<span class="file-name">' + escape(test?.name) + '</span></div>' +
    '<label class="date-control">New date <input id="target-date" type="date" min="1900-01-01" max="9999-12-31" value="' +
    escape(state.date) + '"' + disabled(state.busy) + '></label></div>' +
    preview(result ?? test, { sample: sampleDeck('test'), title: 'Press preview', date: result?.after ?? 'Nov 14', rehearsal: 'Nov 12', changed: Boolean(result), result: true }) +
    rows + '<div class="example-links"><a href="' + samplePaths.test + '" download>Download sample test deck</a></div></div>' +
    '<aside class="inspector"><h2>Test evidence</h2>' + evidence +
    (result?.saved ? '<p class="help-text">Saved to your <button class="text-button" data-action="go-library">library</button>.</p>' : '') +
    '</aside></div>';
}

function libraryView() {
  const rows = state.library.map((entry, index) =>
    '<article class="saved-row"><div><h2>' + escape(operationName(entry.operation)) + '</h2>' +
    '<p>PowerPoint · Tested on ' + escape(entry.file) + ' · ' + escape(entry.before) + ' → ' + escape(entry.after) +
    '</p></div><div class="action-row">' + button('reuse-' + index, 'Test again') +
    button('export-' + index, 'Export for assistant', { primary: true }) + '</div></article>').join('');
  return heading('Your library', 'Reuse tested operations or export them for an assistant.') + notice() +
    (rows || '<div class="empty"><p>Saved operations will appear here after a test passes.</p>' +
      button('go-teach', 'Teach an operation', { primary: true }) + '</div>');
}

function render(focus = false) {
  document.querySelectorAll('[data-view]').forEach(tab => {
    const current = tab.dataset.view === state.view;
    tab.classList.toggle('active', current);
    tab.disabled = state.busy;
    if (current) tab.setAttribute('aria-current', 'page');
    else tab.removeAttribute('aria-current');
  });
  document.querySelector('#operation-title').textContent = state.operation ? operationName(state.operation) : 'Date operation';
  const count = document.querySelector('#library-count');
  count.textContent = state.library.length;
  count.hidden = !state.library.length;
  workspace.innerHTML = { teach: teachView, test: testView, library: libraryView }[state.view]();
  workspace.setAttribute('aria-busy', String(state.busy));
  if (focus) workspace.querySelector('h1').focus({ preventScroll: true });
}

function go(view) {
  state.notice = null;
  state.view = view;
  render(true);
}

function teach() {
  state.operation = null;
  state.result = null;
  if (state.decks.original && state.decks.revision) state.operation = inferOperation(state.decks.original, state.decks.revision);
}

async function withBusy(action) {
  if (state.busy) return;
  state.notice = null;
  state.busy = true;
  render();
  try {
    await action();
  } catch (error) {
    state.notice = { text: error.message || 'Something went wrong. Try again.' };
  } finally {
    state.busy = false;
    render();
  }
}

async function loadSamples() {
  if (!state.samples.original) {
    const decks = await Promise.all(Object.entries(samplePaths).map(async ([key, path]) => {
      const response = await fetch(path);
      if (!response.ok) throw new Error('The sample files could not be loaded. You can still choose your own files.');
      return [key, await readDeck(await response.blob(), path.split('/').pop())];
    }));
    state.samples = Object.fromEntries(decks);
  }
  state.decks = { ...state.samples };
  state.date = '2026-11-21';
  state.view = 'teach';
  teach();
}

async function chooseFile(key, file) {
  if (!/\.pptx$/i.test(file.name)) throw new Error('Choose an editable .pptx file.');
  state.decks[key] = await readDeck(file);
  state.result = null;
  if (key !== 'test') teach();
}

async function testDeck() {
  state.result = null;
  const test = state.decks.test;
  if (test === state.decks.original || test === state.decks.revision) throw new Error('Use a different deck to test the operation.');
  for (const source of [state.decks.original, state.decks.revision]) {
    if (!source || source.slides.length !== test.slides.length) continue;
    const sameSlides = source.slides.every((slide, index) => {
      const original = slide.paragraphs.map(p => p.text);
      const current = test.slides[index].paragraphs.map(p => p.text);
      return JSON.stringify(original) === JSON.stringify(current);
    });
    if (sameSlides) throw new Error('This file has the same slide text as a teaching example. Choose a fresh deck.');
  }
  state.result = await applyOperation(test, state.operation, state.date);
}

function download(bytes, name, mime) {
  const url = URL.createObjectURL(new Blob([bytes], { type: mime }));
  const link = document.createElement('a');
  link.href = url;
  link.download = name;
  link.click();
  setTimeout(() => URL.revokeObjectURL(url), 1500);
}

function save() {
  const entry = { operation: state.operation, file: state.decks.test.name, before: state.result.before, after: state.result.after };
  const library = [entry, ...state.library.filter(item =>
    item.operation.prefix !== entry.operation.prefix || item.operation.suffix !== entry.operation.suffix)];
  localStorage.setItem(savedKey, JSON.stringify(library));
  state.library = library;
  state.result.saved = true;
  state.notice = { success: true, text: 'Operation saved to your library on this browser.' };
}

async function exportOperation(entry) {
  const [script, instructions] = await Promise.all([fetch('assistant/apply.py'), fetch('assistant/SKILL.md')]);
  if (!script.ok || !instructions.ok) throw new Error('The export files could not be loaded. Try again.');
  const archive = new JSZip();
  archive.file('vernacular-operation/SKILL.md', await instructions.text());
  archive.file('vernacular-operation/apply.py', await script.text());
  archive.file('vernacular-operation/operation.json', JSON.stringify(entry.operation, null, 2) + '\n');
  download(await archive.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }), 'vernacular-operation.zip', 'application/zip');
}

document.querySelector('nav').addEventListener('click', event => {
  const tab = event.target.closest('[data-view]');
  if (tab) go(tab.dataset.view);
});
document.querySelector('#file-support').addEventListener('click', () => document.querySelector('#support-dialog').showModal());

workspace.addEventListener('change', event => {
  const input = event.target;
  if (input.matches('[data-upload]') && input.files[0]) {
    const file = input.files[0];
    void withBusy(() => chooseFile(input.dataset.upload, file));
  }
  if (input.id === 'target-date') {
    state.date = input.value;
    state.result = null;
    state.notice = null;
    render();
    workspace.querySelector('#target-date').focus();
  }
});

workspace.addEventListener('click', event => {
  const trigger = event.target.closest('[data-action]');
  if (!trigger || state.busy) return;
  const action = trigger.dataset.action;
  if (action.startsWith('go-')) return go(action.slice(3));
  if (action === 'sample') return void withBusy(loadSamples);
  if (action === 'sample-test') return void withBusy(async () => {
    if (!state.samples.test) {
      const response = await fetch(samplePaths.test);
      if (!response.ok) throw new Error('The sample test deck could not be loaded.');
      state.samples.test = await readDeck(await response.blob(), 'press-preview.pptx');
    }
    state.decks.test = state.samples.test;
    state.result = null;
  });
  if (action === 'run') return void withBusy(testDeck);
  if (action === 'download') return download(state.result.bytes, state.result.fileName, 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
  if (action === 'save') return void withBusy(save);
  if (action.startsWith('reuse-')) {
    state.operation = state.library[Number(action.slice(6))].operation;
    state.result = null;
    return go('test');
  }
  if (action.startsWith('export-')) return void withBusy(() => exportOperation(state.library[Number(action.slice(7))]));
});

void withBusy(loadSamples);
