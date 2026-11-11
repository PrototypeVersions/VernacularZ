const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const namedDate = '(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\\s+(\\d{1,2})(?:,\\s*(\\d{4}))?';
const datePattern = new RegExp('\\b(?:\\d{4}-\\d{2}-\\d{2}|' + namedDate + ')\\b', 'gi');
const encoder = new XMLSerializer();

function parseXML(xml) {
  if (/<!DOCTYPE|<!ENTITY/i.test(xml)) throw new Error('This file uses an unsupported XML declaration.');
  const document = new DOMParser().parseFromString(xml, 'application/xml');
  if (document.getElementsByTagName('parsererror').length) throw new Error('PowerPoint XML could not be read.');
  return document;
}

const elements = (root, ns, tag) => Array.from(root.getElementsByTagNameNS(ns, tag));
const normalize = text => text.trim().replace(/\s+/g, ' ').toLowerCase();
const namePart = (archive, path) => {
  const part = archive.file(path);
  if (!part) throw new Error('The PowerPoint file is missing ' + path + '.');
  return part;
};

function paragraphs(document) {
  return elements(document, A, 'p').map(paragraph => {
    const runs = elements(paragraph, A, 't');
    return { runs, text: runs.map(run => run.textContent).join('') };
  });
}

export async function readDeck(source, name = source.name) {
  if ((source.size ?? source.byteLength) > 15_000_000) throw new Error('Choose a .pptx under 15 MB.');
  let archive;
  try {
    const bytes = source instanceof Blob ? await source.arrayBuffer() : source;
    archive = await JSZip.loadAsync(bytes);
  } catch {
    throw new Error('Choose an unencrypted PowerPoint .pptx file.');
  }
  const files = Object.values(archive.files).filter(entry => !entry.dir);
  // JSZip exposes central-directory sizes on loaded ZIP entries.
  if (files.length > 3000 || files.reduce((size, entry) => size + entry._data.uncompressedSize, 0) > 70_000_000) {
    throw new Error('The uncompressed PowerPoint file is too large for this prototype.');
  }
  if (files.some(entry => /vbaProject\.bin$|^_xmlsignatures\//i.test(entry.name))) {
    throw new Error('Choose a PowerPoint file without macros or digital signatures.');
  }
  const presentation = parseXML(await namePart(archive, 'ppt/presentation.xml').async('string'));
  const relations = parseXML(await namePart(archive, 'ppt/_rels/presentation.xml.rels').async('string'));
  const links = new Map(elements(relations, REL, 'Relationship')
    .filter(relation => relation.getAttribute('Type')?.endsWith('/slide') && relation.getAttribute('TargetMode') !== 'External')
    .map(relation => [relation.getAttribute('Id'), relation.getAttribute('Target')]));
  const slideIds = elements(presentation, P, 'sldId');
  if (!slideIds.length || slideIds.length > 100) throw new Error('Choose a deck containing 1 to 100 slides.');
  const slides = await Promise.all(slideIds.map(async (id, index) => {
    const link = links.get(id.getAttributeNS(R, 'id'));
    const path = link?.startsWith('/') ? link.slice(1) : 'ppt/' + link;
    if (!/^ppt\/slides\/[^/]+\.xml$/.test(path)) throw new Error('A slide points to an unsupported file location.');
    const document = parseXML(await namePart(archive, path).async('string'));
    return { path, index, document, paragraphs: paragraphs(document) };
  }));
  return { archive, name, slides };
}

export function findDates(text) {
  return [...text.matchAll(datePattern)].map(match => ({ text: match[0], index: match.index }));
}

export function inferOperation(original, revision) {
  if (original.slides.length !== revision.slides.length) throw new Error('Use a revision with the same slides in the same order.');
  const edits = [];
  original.slides.forEach((slide, index) => {
    const revised = revision.slides[index];
    if (slide.paragraphs.length !== revised.paragraphs.length) throw new Error('Make only a date change in the revised deck and keep the same text boxes.');
    slide.paragraphs.forEach(({ text }, paragraph) => {
      const updated = revised.paragraphs[paragraph].text;
      if (updated !== text) edits.push({ slide: index + 1, before: text, after: updated });
    });
  });
  if (edits.length !== 1) throw new Error('Vernacular needs exactly one changed paragraph. This pair has ' + edits.length + '.');
  const edit = edits[0];
  const [before, after] = [edit.before, edit.after].map(text => {
    const dates = findDates(text);
    if (dates.length !== 1) throw new Error('Use one labeled date per field, such as “Launch date: Nov 12”.');
    const date = dates[0];
    return { ...date, prefix: text.slice(0, date.index), suffix: text.slice(date.index + date.text.length) };
  });
  if (before.prefix !== after.prefix || before.suffix !== after.suffix || !/\p{L}/u.test(before.prefix + before.suffix)) {
    throw new Error('Keep a stable label in the same text box as the date, such as “Launch date:”.');
  }
  const field = (before.prefix + before.suffix).trim().replace(/^[\s:·—–-]+|[\s:·—–-]+$/g, '');
  return {
    version: 1,
    operation: 'replace-powerpoint-date',
    field,
    prefix: before.prefix,
    suffix: before.suffix,
    example: { before: before.text, after: after.text, slide: edit.slide }
  };
}

function formatDate(iso, example) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(iso)) throw new Error('Choose a valid new date.');
  const [year, month, day] = iso.split('-').map(Number);
  const date = new Date(iso + 'T00:00:00Z');
  if (!year || date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) {
    throw new Error('Choose a valid new date.');
  }
  if (/^\d{4}-/.test(example)) return iso;
  const [word] = example.split(/\s/);
  const monthName = word.length > 3
    ? date.toLocaleString('en-US', { month: 'long', timeZone: 'UTC' })
    : months[month - 1];
  return monthName + ' ' + day + (/\d{4}$/.test(example) ? ', ' + year : '');
}

function findTarget(deck, operation) {
  const matches = [];
  deck.slides.forEach(slide => {
    slide.paragraphs.forEach((paragraph, index) => {
      const dates = findDates(paragraph.text);
      if (dates.length !== 1) return;
      const date = dates[0];
      const prefix = normalize(paragraph.text.slice(0, date.index));
      const suffix = normalize(paragraph.text.slice(date.index + date.text.length));
      if (prefix === normalize(operation.prefix) && suffix === normalize(operation.suffix)) {
        matches.push({ slide, paragraph, index, date });
      }
    });
  });
  if (matches.length !== 1) {
    const detail = matches.length ? 'It occurs ' + matches.length + ' times.' : 'It was not found.';
    throw new Error('The operation requires one “' + operation.field + '” field. ' + detail);
  }
  return matches[0];
}

function equalBytes(left, right) {
  return left.length === right.length && left.every((value, index) => value === right[index]);
}

export async function applyOperation(deck, operation, isoDate) {
  const { slide, paragraph, index, date } = findTarget(deck, operation);
  const replacement = formatDate(isoDate, date.text);
  if (replacement === date.text) throw new Error('Choose a new date that differs from the date in this file.');
  const edited = parseXML(encoder.serializeToString(slide.document));
  const { runs } = paragraphs(edited)[index];
  let offset = 0;
  const targetRun = runs.find(run => {
    const length = run.textContent.length;
    if (offset <= date.index && date.index + date.text.length <= offset + length) return true;
    offset += length;
    return false;
  });
  if (!targetRun) throw new Error('The date is split across formatting runs. Format the date as a single piece of text before trying again.');
  const start = date.index - offset;
  targetRun.textContent = targetRun.textContent.slice(0, start) + replacement + targetRun.textContent.slice(start + date.text.length);
  const editedXML = encoder.serializeToString(edited);
  const output = new JSZip();
  const files = Object.values(deck.archive.files).filter(entry => !entry.dir);
  for (const file of files) {
    output.file(file.name, file.name === slide.path ? editedXML : await file.async('uint8array'), { createFolders: false });
  }
  const bytes = await output.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
  const fileName = deck.name.replace(/\.pptx$/i, '') + ' — Vernacular.pptx';
  const roundTrip = await readDeck(bytes, fileName);
  const originalNames = files.map(file => file.name).sort().join('\n');
  const outputNames = Object.values(roundTrip.archive.files).filter(file => !file.dir).map(file => file.name).sort().join('\n');
  if (originalNames !== outputNames) throw new Error('The output file changed the list of package entries.');
  for (const file of files) {
    if (file.name === slide.path) continue;
    if (!equalBytes(await file.async('uint8array'), await namePart(roundTrip.archive, file.name).async('uint8array'))) {
      throw new Error('Verification found an unexpected change outside the edited slide.');
    }
  }
  const reopenedSlide = roundTrip.slides[slide.index];
  const expected = paragraph.text.slice(0, date.index) + replacement + paragraph.text.slice(date.index + date.text.length);
  if (reopenedSlide.paragraphs[index].text !== expected ||
      encoder.serializeToString(reopenedSlide.document) !== editedXML) {
    throw new Error('The result did not match the proposed change when read back.');
  }
  const preservedDates = deck.slides.flatMap(s => s.paragraphs.flatMap((p, i) => {
    if (s === slide && i === index) return [];
    return findDates(p.text).map(item => ({
      field: p.text.slice(0, item.index).trim().replace(/[:·—–-]+$/, '').trim() || 'Other date',
      value: item.text
    }));
  })).filter((item, i, items) => items.findIndex(other => other.field === item.field && other.value === item.value) === i);
  return {
    bytes,
    fileName,
    slide: slide.index + 1,
    before: date.text,
    after: replacement,
    preservedDates,
    untouchedParts: files.length - 1,
    preview: reopenedSlide.paragraphs.map(p => p.text).filter(Boolean)
  };
}
