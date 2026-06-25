#!/usr/bin/env node
/**
 * harness.js — HCD build harness (Legal System pipeline).
 *
 * Text in, one binary out. Takes the certified tokenized template
 * (hcd-template.txt) plus a resolved matter map (data.json) and emits a DRAFT
 * .docx reproducing the firm's certified Health Care Directive, using the
 * measured TR* style ramp from legal-design v1.0.
 *
 * The template and its hash are fetched from git (raw.githubusercontent.com)
 * by the pipeline — they never transit Claude's output. This harness gates on
 * the hash so a truncated, partial, or altered template cannot build.
 *
 * Subcommands:
 *   node harness.js verify [template.txt] [template.sha256]
 *     Integrity gate: recompute SHA-256 of the template and compare to the
 *     committed hash file. Exits non-zero on mismatch OR on a missing hash
 *     file (fail closed). This is the crash-vector guard.
 *   node harness.js build [template.txt] [data.json] [out.docx] [template.sha256]
 *     Verify first (hash gate), then parse, fill {{role.field}} tokens from
 *     data.json (filled red FF0000; unfilled tokens kept literal, also red),
 *     place the DRAFT marker, and write the .docx.
 */
'use strict';

const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

/** Default canonical filenames (as committed to the Agentic-harness repo). */
const DEFAULT_TEMPLATE = 'hcd-template.txt';

/**
 * Read a UTF-8 file or exit with a clear message.
 * @param {string} file
 * @returns {string}
 */
function readOrDie(file) {
  try {
    return fs.readFileSync(file, 'utf8');
  } catch (err) {
    console.error(`FATAL: cannot read ${file}: ${err.message}`);
    process.exit(2);
  }
}

/**
 * Derive the hash-file path from a template path: foo.txt -> foo.sha256.
 * @param {string} templateFile
 * @returns {string}
 */
function deriveHashPath(templateFile) {
  return templateFile.replace(/\.txt$/, '') + '.sha256';
}

/**
 * SHA-256 (hex) of a file's raw bytes.
 * @param {string} file
 * @returns {string}
 */
function sha256File(file) {
  return crypto.createHash('sha256').update(fs.readFileSync(file)).digest('hex');
}

/**
 * Integrity gate. Recompute the template's SHA-256 and compare to the
 * committed hash file. Fails closed: a missing hash file is a FAIL, never a
 * pass. The hash file may be a bare 64-char hex digest or `sha256sum` format
 * (`<hash>  <name>`); only the first whitespace-delimited token is read.
 * @param {string} file       template path
 * @param {string} [hashFile] hash path (defaults to derived <template>.sha256)
 * @returns {{ok: boolean, report: string}}
 */
function verifyTemplate(file, hashFile) {
  const text = readOrDie(file);
  const chars = text.length;
  const words = (text.match(/\S+/g) || []).length;
  const lines = text.split(/\r?\n/).length;

  const hp = hashFile || deriveHashPath(file);
  let ok = false;
  let integrity;
  if (!fs.existsSync(hp)) {
    integrity =
      `hash file: MISSING (${hp})\n` +
      `integrity: FAIL — refusing (no hash to verify against)`;
  } else {
    const expected = readOrDie(hp).trim().split(/\s+/)[0].toLowerCase();
    const actual = sha256File(file);
    ok = expected.length === 64 && expected === actual;
    integrity =
      `hash file: ${hp}\n` +
      `expected:  ${expected}\n` +
      `actual:    ${actual}\n` +
      `integrity: ${ok ? 'MATCH' : 'MISMATCH'}`;
  }

  const report =
    `template: ${file}\n` +
    `chars=${chars} words=${words} lines=${lines}\n` +
    `${integrity}\n` +
    `--- verify: ${ok ? 'PASS' : 'FAIL'} ---`;
  return { ok, report };
}

/**
 * Resolve one {{role.field}} token against the data map. Supports derived
 * *_upper fields. Returns the filled value, or the literal token if absent
 * (caller renders both red).
 * @param {object} data
 * @param {string} token  e.g. "{{client.full_name}}"
 * @returns {string}
 */
function resolveToken(data, token) {
  const inner = token.slice(2, -2).trim();
  const dot = inner.indexOf('.');
  if (dot < 0) return token;
  const role = inner.slice(0, dot);
  const field = inner.slice(dot + 1);
  const bag = data[role];
  if (field.endsWith('_upper')) {
    const base = field.slice(0, -6);
    const v = bag && bag[base];
    if (typeof v === 'string' && v.length) return v.toUpperCase();
  } else {
    const v = bag && bag[field];
    if (typeof v === 'string' && v.length) return v;
  }
  return token; // unfilled — keep literal, rendered red
}

/**
 * Split inline content into docx runs: {{tokens}} (red), ⟦BR⟧ line breaks,
 * and tabs. Plain text is black; `bold` applies to all runs.
 * @param {object} d   docx module
 * @param {object} data
 * @param {string} content
 * @param {boolean} bold
 * @returns {object[]} TextRun[]
 */
function runsFor(d, data, content, bold) {
  const parts = content.split(/(\{\{[^}]+\}\}|⟦BR⟧|\t)/);
  const runs = [];
  for (const p of parts) {
    if (p === '') continue;
    if (p === '⟦BR⟧') {
      runs.push(new d.TextRun({ text: '', break: 1 }));
    } else if (p === '\t') {
      runs.push(new d.TextRun({ children: [new d.Tab()], bold }));
    } else if (/^\{\{[^}]+\}\}$/.test(p)) {
      runs.push(new d.TextRun({ text: resolveToken(data, p), color: 'FF0000', bold }));
    } else {
      runs.push(new d.TextRun({ text: p, bold }));
    }
  }
  return runs.length ? runs : [new d.TextRun({ text: '', bold })];
}

/**
 * Parse the tokenized template into docx body children.
 * @param {object} d docx module
 * @param {object} data
 * @param {string[]} lines
 * @returns {{body: object[], footerLeft: string}}
 */
function parseBody(d, data, lines) {
  const body = [];
  let footerLeft = 'MEDICAL POWER OF ATTORNEY';
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].replace(/\s+$/, '');
    if (line.trim() === '') { body.push(new d.Paragraph({ style: 'TRBase' })); continue; }
    if (line.includes('⟦FOOTER⟧')) {
      const m = line.replace('⟦FOOTER⟧', '').replace('⟦PAGE_NUM⟧', '').trim();
      if (m) footerLeft = m;
      continue;
    }
    const tok = line.startsWith('⟦') ? line.slice(0, line.indexOf('⟧') + 1) : '';
    const content = tok ? line.slice(tok.length).replace(/^ /, '') : line;
    if (tok === '⟦TITLE⟧') {
      body.push(new d.Paragraph({ style: 'TRTitle', children: runsFor(d, data, content, false) }));
    } else if (tok === '⟦HEADING⟧') {
      body.push(new d.Paragraph({ style: 'TRBody1', indent: { firstLine: 0 }, children: runsFor(d, data, content, true) }));
    } else if (tok === '⟦PARA⟧') {
      body.push(new d.Paragraph({ style: 'TRBody1', children: runsFor(d, data, content, false) }));
    } else if (tok === '⟦CENTER⟧') {
      body.push(new d.Paragraph({ style: 'TRBody1', alignment: d.AlignmentType.CENTER, indent: { firstLine: 0 }, children: runsFor(d, data, content, false) }));
    } else if (tok === '⟦SIG_LINE⟧') {
      body.push(new d.Paragraph({ style: 'TRSigLine', children: runsFor(d, data, content, false) }));
    } else if (tok === '⟦SIG_LABEL⟧') {
      body.push(new d.Paragraph({ style: 'TRSigName', children: runsFor(d, data, content, false) }));
    } else if (tok === '⟦PAGEBREAK⟧') {
      body.push(new d.Paragraph({ children: [new d.PageBreak()] }));
    } else if (tok === '⟦ACK_ROW⟧') {
      body.push(ackRow(d, data, content));
    } else if (tok === '⟦TABLE⟧') {
      const rows = [];
      while (i + 1 < lines.length && lines[i + 1].startsWith('⟦ROW⟧')) {
        rows.push(lines[++i].slice('⟦ROW⟧'.length).replace(/^ /, ''));
      }
      body.push(altTable(d, data, rows));
    } else if (tok === '⟦ROW⟧') {
      // consumed by ⟦TABLE⟧; ignore strays
    } else {
      body.push(new d.Paragraph({ style: 'TRBody1', children: runsFor(d, data, content, false) }));
    }
  }
  return { body, footerLeft };
}

/**
 * §-rail jurat row → TRAffid paragraph with a left tab stop at 4320.
 * A bare "§" line gets one leading tab so it aligns under the rail.
 * @param {object} d @param {object} data @param {string} content
 * @returns {object} Paragraph
 */
function ackRow(d, data, content) {
  const c = content.includes('\t') ? content : '\t' + content;
  return new d.Paragraph({
    style: 'TRAffid',
    tabStops: [{ type: d.TabStopType.LEFT, position: 4320 }],
    children: runsFor(d, data, c, false),
  });
}

/** Borderless 2-column label/value cell. */
function cell(d, data, text, w) {
  const none = { style: d.BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  return new d.TableCell({
    width: { size: w, type: d.WidthType.DXA },
    borders: { top: none, bottom: none, left: none, right: none },
    children: [new d.Paragraph({ style: 'TRBase', children: runsFor(d, data, text, false) })],
  });
}

/**
 * Alternate-agent block → borderless 2-column table.
 * @param {object} d @param {object} data @param {string[]} rows  "label | value"
 * @returns {object} Table
 */
function altTable(d, data, rows) {
  const none = { style: d.BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  const trs = rows.map((r) => {
    const bar = r.indexOf('|');
    const label = bar >= 0 ? r.slice(0, bar).trim() : r.trim();
    const value = bar >= 0 ? r.slice(bar + 1).trim() : '';
    return new d.TableRow({ children: [cell(d, data, label, 1800), cell(d, data, value, 7560)] });
  });
  return new d.Table({
    width: { size: 9360, type: d.WidthType.DXA },
    borders: { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none },
    rows: trs,
  });
}

/**
 * Certified TR* paragraph style ramp (legal-design §Style ramp + Session 3).
 * @param {object} d @returns {object[]}
 */
function styleRamp(d) {
  const A = d.AlignmentType;
  return [
    { id: 'TRBase', name: 'TRBase', run: { font: 'Times New Roman', size: 24 }, paragraph: { spacing: { after: 120 } } },
    { id: 'TRBody', name: 'TRBody', basedOn: 'TRBase', paragraph: { alignment: A.JUSTIFIED, spacing: { line: 360, lineRule: d.LineRuleType.AUTO, after: 120 } } },
    { id: 'TRBody1', name: 'TRBody1', basedOn: 'TRBody', paragraph: { indent: { firstLine: 720 } } },
    { id: 'TRTitle', name: 'TRTitle', basedOn: 'TRBase', run: { bold: true }, paragraph: { alignment: A.CENTER, spacing: { after: 360 } } },
    { id: 'TRSigLine', name: 'TRSigLine', basedOn: 'TRBase', paragraph: { spacing: { before: 360 }, indent: { left: 4320 } } },
    { id: 'TRSigName', name: 'TRSigName', basedOn: 'TRBase', paragraph: { indent: { left: 4320 } } },
    { id: 'TRAffid', name: 'TRAffid', basedOn: 'TRBase', paragraph: {} },
  ];
}

/** Default footer (pages 2+): left text + right-tab page number. */
function makeFooter(d, footerLeft) {
  return new d.Footer({
    children: [new d.Paragraph({
      style: 'TRBase',
      tabStops: [{ type: d.TabStopType.RIGHT, position: 9360 }],
      children: [
        new d.TextRun({ text: footerLeft }),
        new d.TextRun({ children: [new d.Tab(), '- ', d.PageNumber.CURRENT, ' -'] }),
      ],
    })],
  });
}

/**
 * Assemble the HCD Document: DRAFT marker, body, page setup, footers.
 * @param {object} d @param {string} templateText @param {object} data
 * @returns {object} Document
 */
function buildDocx(d, templateText, data) {
  const { body, footerLeft } = parseBody(d, data, templateText.split(/\r?\n/));
  const draft = new d.Paragraph({
    style: 'TRBase', alignment: d.AlignmentType.CENTER,
    children: [new d.TextRun({ text: 'DRAFT', bold: true, color: 'FF0000' })],
  });
  return new d.Document({
    styles: { paragraphStyles: styleRamp(d) },
    sections: [{
      properties: {
        titlePage: true,
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 144 },
        },
      },
      footers: { default: makeFooter(d, footerLeft), first: new d.Footer({ children: [new d.Paragraph({})] }) },
      children: [draft, ...body],
    }],
  });
}

/**
 * build subcommand: verify-gate, then parse + fill + write the .docx.
 * @param {string} templateFile @param {string} dataFile @param {string} outFile
 * @param {string} [hashFile]
 */
async function runBuild(templateFile, dataFile, outFile, hashFile) {
  const v = verifyTemplate(templateFile, hashFile);
  if (!v.ok) {
    console.error(v.report);
    console.error('FATAL: template failed verify — refusing to build.');
    process.exit(1);
  }
  const d = require('docx');
  const templateText = readOrDie(templateFile);
  let data;
  try {
    data = JSON.parse(readOrDie(dataFile));
  } catch (err) {
    console.error(`FATAL: cannot parse ${dataFile}: ${err.message}`);
    process.exit(2);
  }
  let buf;
  try {
    buf = await d.Packer.toBuffer(buildDocx(d, templateText, data));
    fs.writeFileSync(outFile, buf);
  } catch (err) {
    console.error(`FATAL: build failed: ${err.message}`);
    process.exit(3);
  }
  console.log(`OK: wrote ${path.resolve(outFile)} (${fs.statSync(outFile).size} bytes)`);
}

/** CLI dispatch. */
function main() {
  const cmd = process.argv[2];
  const a = process.argv.slice(3);
  if (cmd === 'verify') {
    const { ok, report } = verifyTemplate(a[0] || DEFAULT_TEMPLATE, a[1]);
    console.log(report);
    process.exit(ok ? 0 : 1);
  } else if (cmd === 'build') {
    runBuild(a[0] || DEFAULT_TEMPLATE, a[1] || 'data.json', a[2] || 'HCD_DRAFT.docx', a[3]);
  } else {
    console.error('usage: node harness.js verify [template.txt] [template.sha256]');
    console.error('       node harness.js build [template.txt] [data.json] [out.docx] [template.sha256]');
    process.exit(1);
  }
}

main();
