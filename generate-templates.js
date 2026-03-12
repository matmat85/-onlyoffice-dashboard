/**
 * generate-templates.js
 *
 * Creates minimal valid blank Office Open XML template files:
 *   templates/blank.docx
 *   templates/blank.xlsx
 *   templates/blank.pptx
 *
 * Uses ONLY Node.js built-in modules — no extra dependencies required.
 * Run once:  node generate-templates.js
 */

'use strict';

const fs   = require('fs');
const path = require('path');

const TEMPLATES_DIR = path.join(__dirname, 'templates');
if (!fs.existsSync(TEMPLATES_DIR)) fs.mkdirSync(TEMPLATES_DIR, { recursive: true });

// ── Minimal CRC-32 ────────────────────────────────────────
function crc32(buf) {
  let c = ~0 >>> 0;
  for (let i = 0; i < buf.length; i++) {
    c ^= buf[i];
    for (let j = 0; j < 8; j++) c = (c >>> 1) ^ (c & 1 ? 0xEDB88320 : 0);
  }
  return (~c) >>> 0;
}

// ── Minimal ZIP writer (STORED, no compression) ───────────
function buildZip(files) {
  // files = [{ name: string, data: string|Buffer }]
  const localParts = [];
  const centralParts = [];
  let localOffset = 0;

  for (const file of files) {
    const name = Buffer.from(file.name, 'utf8');
    const data = Buffer.isBuffer(file.data) ? file.data : Buffer.from(file.data, 'utf8');
    const crc  = crc32(data);
    const size = data.length;
    const now  = dosDateTime();

    // ── Local file header ──
    const local = Buffer.alloc(30 + name.length + size);
    local.writeUInt32LE(0x04034b50, 0);   // signature
    local.writeUInt16LE(20,  4);          // version needed
    local.writeUInt16LE(0,   6);          // flags
    local.writeUInt16LE(0,   8);          // compression: STORED
    local.writeUInt16LE(now.time, 10);    // mod time
    local.writeUInt16LE(now.date, 12);    // mod date
    local.writeUInt32LE(crc,  14);        // crc32
    local.writeUInt32LE(size, 18);        // compressed size
    local.writeUInt32LE(size, 22);        // uncompressed size
    local.writeUInt16LE(name.length, 26); // filename length
    local.writeUInt16LE(0, 28);           // extra length
    name.copy(local, 30);
    data.copy(local, 30 + name.length);

    localParts.push(local);

    // ── Central directory entry ──
    const cd = Buffer.alloc(46 + name.length);
    cd.writeUInt32LE(0x02014b50, 0);       // signature
    cd.writeUInt16LE(20, 4);               // version made by
    cd.writeUInt16LE(20, 6);               // version needed
    cd.writeUInt16LE(0,  8);               // flags
    cd.writeUInt16LE(0,  10);              // compression
    cd.writeUInt16LE(now.time, 12);
    cd.writeUInt16LE(now.date, 14);
    cd.writeUInt32LE(crc,  16);
    cd.writeUInt32LE(size, 20);
    cd.writeUInt32LE(size, 24);
    cd.writeUInt16LE(name.length, 28);
    cd.writeUInt16LE(0, 30);               // extra
    cd.writeUInt16LE(0, 32);               // comment
    cd.writeUInt16LE(0, 34);               // disk start
    cd.writeUInt16LE(0, 36);               // internal attrs
    cd.writeUInt32LE(0, 38);               // external attrs
    cd.writeUInt32LE(localOffset, 42);     // local header offset
    name.copy(cd, 46);
    centralParts.push(cd);

    localOffset += local.length;
  }

  const cdBuf    = Buffer.concat(centralParts);
  const eocd     = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);   // signature
  eocd.writeUInt16LE(0, 4);
  eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(files.length, 8);
  eocd.writeUInt16LE(files.length, 10);
  eocd.writeUInt32LE(cdBuf.length, 12);
  eocd.writeUInt32LE(localOffset, 16);
  eocd.writeUInt16LE(0, 20);

  return Buffer.concat([...localParts, cdBuf, eocd]);
}

function dosDateTime() {
  const d = new Date();
  const date = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
  const time = (d.getHours() << 11) | (d.getMinutes() << 5) | Math.floor(d.getSeconds() / 2);
  return { date: date & 0xFFFF, time: time & 0xFFFF };
}

// ── XML helpers ───────────────────────────────────────────
const xml = (s) => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${s}`;
const NS_REL    = 'http://schemas.openxmlformats.org/package/2006/relationships';
const NS_CTYPE  = 'http://schemas.openxmlformats.org/package/2006/content-types';
const NS_OPC    = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

// ── DOCX ──────────────────────────────────────────────────
function buildDocx() {
  const contentTypes = xml(`<Types xmlns="${NS_CTYPE}">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);

  const rootRels = xml(`<Relationships xmlns="${NS_REL}">
<Relationship Id="rId1" Type="${NS_OPC}/officeDocument" Target="word/document.xml"/>
</Relationships>`);

  const docBody = xml(`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body><w:p/><w:sectPr/></w:body>
</w:document>`);

  const docRels = xml(`<Relationships xmlns="${NS_REL}"/>`);

  return buildZip([
    { name: '[Content_Types].xml',       data: contentTypes },
    { name: '_rels/.rels',               data: rootRels },
    { name: 'word/document.xml',         data: docBody },
    { name: 'word/_rels/document.xml.rels', data: docRels },
  ]);
}

// ── XLSX ──────────────────────────────────────────────────
function buildXlsx() {
  const NS_SS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

  const contentTypes = xml(`<Types xmlns="${NS_CTYPE}">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`);

  const rootRels = xml(`<Relationships xmlns="${NS_REL}">
<Relationship Id="rId1" Type="${NS_OPC}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`);

  const workbook = xml(`<workbook xmlns="${NS_SS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>`);

  const wbRels = xml(`<Relationships xmlns="${NS_REL}">
<Relationship Id="rId1" Type="${NS_OPC}/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`);

  const sheet = xml(`<worksheet xmlns="${NS_SS}"><sheetData/></worksheet>`);

  return buildZip([
    { name: '[Content_Types].xml',          data: contentTypes },
    { name: '_rels/.rels',                  data: rootRels },
    { name: 'xl/workbook.xml',              data: workbook },
    { name: 'xl/_rels/workbook.xml.rels',   data: wbRels },
    { name: 'xl/worksheets/sheet1.xml',     data: sheet },
  ]);
}

// ── PPTX ──────────────────────────────────────────────────
function buildPptx() {
  const NS_PML = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const NS_A   = 'http://schemas.openxmlformats.org/drawingml/2006/main';

  const contentTypes = xml(`<Types xmlns="${NS_CTYPE}">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`);

  const rootRels = xml(`<Relationships xmlns="${NS_REL}">
<Relationship Id="rId1" Type="${NS_OPC}/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`);

  const presentation = xml(`<p:presentation xmlns:p="${NS_PML}" xmlns:a="${NS_A}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<p:sldMasterIdLst/>
<p:sldSz cx="9144000" cy="6858000"/>
<p:notesSz cx="6858000" cy="9144000"/>
<p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst>
</p:presentation>`);

  const presRels = xml(`<Relationships xmlns="${NS_REL}">
<Relationship Id="rId1" Type="${NS_OPC}/slide" Target="slides/slide1.xml"/>
</Relationships>`);

  const slide = xml(`<p:sld xmlns:p="${NS_PML}" xmlns:a="${NS_A}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
</p:spTree></p:cSld>
</p:sld>`);

  const slideRels = xml(`<Relationships xmlns="${NS_REL}"/>`);

  return buildZip([
    { name: '[Content_Types].xml',                    data: contentTypes },
    { name: '_rels/.rels',                            data: rootRels },
    { name: 'ppt/presentation.xml',                   data: presentation },
    { name: 'ppt/_rels/presentation.xml.rels',        data: presRels },
    { name: 'ppt/slides/slide1.xml',                  data: slide },
    { name: 'ppt/slides/_rels/slide1.xml.rels',       data: slideRels },
  ]);
}

// ── Write ─────────────────────────────────────────────────
const targets = [
  { file: 'blank.docx',  build: buildDocx },
  { file: 'blank.xlsx',  build: buildXlsx },
  { file: 'blank.pptx',  build: buildPptx },
];

let generated = 0;
for (const { file, build } of targets) {
  const outPath = path.join(TEMPLATES_DIR, file);
  fs.writeFileSync(outPath, build());
  console.log(`✔ templates/${file}`);
  generated++;
}

console.log(`\nDone — ${generated} template(s) created in ./templates/`);
