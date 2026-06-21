// netlify/functions/generate-excel.js
//
// Generates an .xlsx workbook for the "In Production" tab with:
//   1. "All Products" sheet  — every catalogue product, grouped by category, qty=0 if not ordered
//   2. "Ordered Only" sheet  — same layout but only products that were actually ordered (qty > 0)
//   3. One sheet per client  — that client's order only, grouped by category, with a checkbox column
//
// ZERO DEPENDENCIES — no npm install needed. Uses only Node.js built-ins
// (zlib) to build the .xlsx file (a ZIP of XML parts) from scratch.
// Just drop this single file into netlify/functions/ and deploy — no
// package.json, no node_modules, works with drag-and-drop deploys.

// xlsx-writer.js
// A minimal, ZERO-DEPENDENCY .xlsx file writer.
// Uses only Node.js built-in modules (zlib for ZIP compression).
// Supports: multiple sheets, cell styles (bold, color, fill, borders, alignment),
// merged cells, column widths, frozen header rows.
//
// This avoids needing `npm install exceljs` for drag-and-drop Netlify deploys
// where no build step runs.

const zlib = require('zlib');

// ── Minimal ZIP writer (store + deflate, no external deps) ──────────
function crc32(buf) {
  let c, crc = 0xFFFFFFFF;
  const table = crc32.table || (crc32.table = (() => {
    const t = [];
    for (let n = 0; n < 256; n++) {
      c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[n] = c;
    }
    return t;
  })());
  for (let i = 0; i < buf.length; i++) {
    crc = table[(crc ^ buf[i]) & 0xFF] ^ (crc >>> 8);
  }
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

function dosDateTime(date) {
  const d = date || new Date();
  const dosTime = ((d.getHours() & 0x1F) << 11) | ((d.getMinutes() & 0x3F) << 5) | ((d.getSeconds() >> 1) & 0x1F);
  const dosDate = (((d.getFullYear() - 1980) & 0x7F) << 9) | (((d.getMonth() + 1) & 0xF) << 5) | (d.getDate() & 0x1F);
  return { dosTime, dosDate };
}

class ZipWriter {
  constructor() {
    this.entries = [];
  }
  addFile(name, content) {
    const buf = Buffer.isBuffer(content) ? content : Buffer.from(content, 'utf8');
    const compressed = zlib.deflateRawSync(buf);
    this.entries.push({ name, raw: buf, compressed, crc: crc32(buf) });
  }
  toBuffer() {
    const { dosTime, dosDate } = dosDateTime();
    const localParts = [];
    const centralParts = [];
    let offset = 0;

    for (const e of this.entries) {
      const nameBuf = Buffer.from(e.name, 'utf8');
      const useCompressed = e.compressed.length < e.raw.length;
      const method = useCompressed ? 8 : 0;
      const data = useCompressed ? e.compressed : e.raw;

      const localHeader = Buffer.alloc(30);
      localHeader.writeUInt32LE(0x04034b50, 0);
      localHeader.writeUInt16LE(20, 4); // version needed
      localHeader.writeUInt16LE(0, 6); // flags
      localHeader.writeUInt16LE(method, 8);
      localHeader.writeUInt16LE(dosTime, 10);
      localHeader.writeUInt16LE(dosDate, 12);
      localHeader.writeUInt32LE(e.crc, 14);
      localHeader.writeUInt32LE(data.length, 18);
      localHeader.writeUInt32LE(e.raw.length, 22);
      localHeader.writeUInt16LE(nameBuf.length, 26);
      localHeader.writeUInt16LE(0, 28);

      localParts.push(localHeader, nameBuf, data);

      const centralHeader = Buffer.alloc(46);
      centralHeader.writeUInt32LE(0x02014b50, 0);
      centralHeader.writeUInt16LE(20, 4); // version made by
      centralHeader.writeUInt16LE(20, 6); // version needed
      centralHeader.writeUInt16LE(0, 8); // flags
      centralHeader.writeUInt16LE(method, 10);
      centralHeader.writeUInt16LE(dosTime, 12);
      centralHeader.writeUInt16LE(dosDate, 14);
      centralHeader.writeUInt32LE(e.crc, 16);
      centralHeader.writeUInt32LE(data.length, 20);
      centralHeader.writeUInt32LE(e.raw.length, 24);
      centralHeader.writeUInt16LE(nameBuf.length, 28);
      centralHeader.writeUInt16LE(0, 30); // extra len
      centralHeader.writeUInt16LE(0, 32); // comment len
      centralHeader.writeUInt16LE(0, 34); // disk number
      centralHeader.writeUInt16LE(0, 36); // internal attrs
      centralHeader.writeUInt32LE(0, 38); // external attrs
      centralHeader.writeUInt32LE(offset, 42);

      centralParts.push(centralHeader, nameBuf);

      offset += localHeader.length + nameBuf.length + data.length;
    }

    const centralStart = offset;
    let centralSize = 0;
    for (const p of centralParts) centralSize += p.length;

    const end = Buffer.alloc(22);
    end.writeUInt32LE(0x06054b50, 0);
    end.writeUInt16LE(0, 4);
    end.writeUInt16LE(0, 6);
    end.writeUInt16LE(this.entries.length, 8);
    end.writeUInt16LE(this.entries.length, 10);
    end.writeUInt32LE(centralSize, 12);
    end.writeUInt32LE(centralStart, 16);
    end.writeUInt16LE(0, 20);

    return Buffer.concat([...localParts, ...centralParts, end]);
  }
}

// ── XML escaping ──────────────────────────────────────────────────
function esc(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function colLetter(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// ── Workbook builder ─────────────────────────────────────────────
class Workbook {
  constructor() {
    this.sheets = [];
    this.styles = []; // array of style defs, index 0 = default
    this.styleCache = new Map();
    this._addStyle({}); // index 0: default style
  }

  _addStyle(style) {
    const key = JSON.stringify(style);
    if (this.styleCache.has(key)) return this.styleCache.get(key);
    const idx = this.styles.length;
    this.styles.push(style);
    this.styleCache.set(key, idx);
    return idx;
  }

  addSheet(name) {
    const sheet = new Sheet(name, this);
    this.sheets.push(sheet);
    return sheet;
  }

  toBuffer() {
    const zip = new ZipWriter();

    zip.addFile('[Content_Types].xml', this._contentTypesXml());
    zip.addFile('_rels/.rels', this._relsXml());
    zip.addFile('docProps/core.xml', this._coreXml());
    zip.addFile('docProps/app.xml', this._appXml());
    zip.addFile('xl/workbook.xml', this._workbookXml());
    zip.addFile('xl/_rels/workbook.xml.rels', this._workbookRelsXml());
    zip.addFile('xl/styles.xml', this._stylesXml());

    this.sheets.forEach((sheet, i) => {
      zip.addFile(`xl/worksheets/sheet${i + 1}.xml`, sheet.toXml());
    });

    return zip.toBuffer();
  }

  _contentTypesXml() {
    const sheetOverrides = this.sheets
      .map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`)
      .join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
${sheetOverrides}
</Types>`;
  }

  _relsXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
  }

  _coreXml() {
    const now = new Date().toISOString();
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:creator>Matteo Orders</dc:creator>
<cp:lastModifiedBy>Matteo Orders</cp:lastModifiedBy>
<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created>
<dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified>
</cp:coreProperties>`;
  }

  _appXml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>Matteo Orders</Application>
</Properties>`;
  }

  _workbookXml() {
    const sheetsXml = this.sheets
      .map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`)
      .join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>${sheetsXml}</sheets>
</workbook>`;
  }

  _workbookRelsXml() {
    const sheetRels = this.sheets
      .map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`)
      .join('');
    const stylesRelId = this.sheets.length + 1;
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${sheetRels}
<Relationship Id="rId${stylesRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
  }

  _stylesXml() {
    // Collect unique fonts, fills, borders, number formats from styles
    const fonts = [{ name: 'Calibri', sz: 11 }]; // index 0
    const fills = [{ pattern: 'none' }, { pattern: 'gray125' }]; // 0,1 reserved
    const borders = [{}]; // index 0: none
    const fontCache = new Map([['{"name":"Calibri","sz":11}', 0]]);
    const fillCache = new Map([['{"pattern":"none"}', 0], ['{"pattern":"gray125"}', 1]]);
    const borderCache = new Map([['{}', 0]]);

    const cellXfs = this.styles.map((st) => {
      const font = { name: st.fontName || 'Arial', sz: st.fontSize || 11, b: !!st.bold, color: st.fontColor || null };
      const fontKey = JSON.stringify(font);
      let fontId = fontCache.get(fontKey);
      if (fontId === undefined) {
        fontId = fonts.length;
        fonts.push(font);
        fontCache.set(fontKey, fontId);
      }

      let fillId = 0;
      if (st.fillColor) {
        const fill = { pattern: 'solid', fgColor: st.fillColor };
        const fillKey = JSON.stringify(fill);
        fillId = fillCache.get(fillKey);
        if (fillId === undefined) {
          fillId = fills.length;
          fills.push(fill);
          fillCache.set(fillKey, fillId);
        }
      }

      let borderId = 0;
      if (st.border) {
        const border = st.border;
        const borderKey = JSON.stringify(border);
        borderId = borderCache.get(borderKey);
        if (borderId === undefined) {
          borderId = borders.length;
          borders.push(border);
          borderCache.set(borderKey, borderId);
        }
      }

      return {
        fontId, fillId, borderId,
        align: st.align || null,
      };
    });

    const fontsXml = fonts
      .map((f) => `<font><sz val="${f.sz}"/>${f.b ? '<b/>' : ''}<name val="${esc(f.name)}"/>${f.color ? `<color rgb="${f.color}"/>` : ''}</font>`)
      .join('');

    const fillsXml = fills
      .map((f) => {
        if (f.pattern === 'none') return '<fill><patternFill patternType="none"/></fill>';
        if (f.pattern === 'gray125') return '<fill><patternFill patternType="gray125"/></fill>';
        return `<fill><patternFill patternType="solid"><fgColor rgb="${f.fgColor}"/><bgColor rgb="${f.fgColor}"/></patternFill></fill>`;
      })
      .join('');

    function borderSide(side) {
      if (!side) return '';
      return `style="${side.style || 'thin'}"><color rgb="${side.color || 'FFDDDDDD'}"/>`;
    }
    const bordersXml = borders
      .map((b) => {
        const top = b.top ? `<top ${borderSide(b.top)}</top>` : '<top/>';
        const bottom = b.bottom ? `<bottom ${borderSide(b.bottom)}</bottom>` : '<bottom/>';
        const left = b.left ? `<left ${borderSide(b.left)}</left>` : '<left/>';
        const right = b.right ? `<right ${borderSide(b.right)}</right>` : '<right/>';
        return `<border>${left}${right}${top}${bottom}<diagonal/></border>`;
      })
      .join('');

    const cellXfsXml = cellXfs
      .map((xf) => {
        const alignAttr = xf.align
          ? `<alignment horizontal="${xf.align.h || 'general'}" vertical="${xf.align.v || 'bottom'}"${xf.align.wrap ? ' wrapText="1"' : ''}/>`
          : '';
        return `<xf numFmtId="0" fontId="${xf.fontId}" fillId="${xf.fillId}" borderId="${xf.borderId}" xfId="0" applyFont="1" applyFill="1" applyBorder="1"${xf.align ? ' applyAlignment="1"' : ''}>${alignAttr}</xf>`;
      })
      .join('');

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="${fonts.length}">${fontsXml}</fonts>
<fills count="${fills.length}">${fillsXml}</fills>
<borders count="${borders.length}">${bordersXml}</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="${cellXfs.length}">${cellXfsXml}</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;
  }
}

class Sheet {
  constructor(name, workbook) {
    this.name = name.substring(0, 31);
    this.workbook = workbook;
    this.cells = new Map(); // "row,col" -> {value, styleId}
    this.colWidths = {};
    this.merges = [];
    this.freezeRow = 0;
    this.maxRow = 0;
    this.maxCol = 0;
  }

  setColWidth(col, width) {
    this.colWidths[col] = width;
  }

  freezeHeaderRows(n) {
    this.freezeRow = n;
  }

  mergeCells(r1, c1, r2, c2) {
    this.merges.push([r1, c1, r2, c2]);
  }

  setCell(row, col, value, style) {
    this.maxRow = Math.max(this.maxRow, row);
    this.maxCol = Math.max(this.maxCol, col);
    const styleId = style ? this.workbook._addStyle(style) : 0;
    this.cells.set(`${row},${col}`, { value, styleId });
  }

  toXml() {
    const colsXml = Object.entries(this.colWidths)
      .map(([col, width]) => `<col min="${col}" max="${col}" width="${width}" customWidth="1"/>`)
      .join('');

    const rowsMap = new Map();
    for (const [key, cell] of this.cells.entries()) {
      const [r, c] = key.split(',').map(Number);
      if (!rowsMap.has(r)) rowsMap.set(r, []);
      rowsMap.get(r).push({ col: c, cell });
    }

    const rowNumbers = [...rowsMap.keys()].sort((a, b) => a - b);
    const rowsXml = rowNumbers
      .map((r) => {
        const cellsInRow = rowsMap.get(r).sort((a, b) => a.col - b.col);
        const cellsXml = cellsInRow
          .map(({ col, cell }) => {
            const ref = `${colLetter(col)}${r}`;
            if (cell.value === null || cell.value === undefined || cell.value === '') {
              return `<c r="${ref}" s="${cell.styleId}"/>`;
            }
            if (typeof cell.value === 'number') {
              return `<c r="${ref}" s="${cell.styleId}"><v>${cell.value}</v></c>`;
            }
            return `<c r="${ref}" s="${cell.styleId}" t="inlineStr"><is><t xml:space="preserve">${esc(cell.value)}</t></is></c>`;
          })
          .join('');
        return `<row r="${r}">${cellsXml}</row>`;
      })
      .join('');

    const mergesXml = this.merges.length
      ? `<mergeCells count="${this.merges.length}">${this.merges
          .map(([r1, c1, r2, c2]) => `<mergeCell ref="${colLetter(c1)}${r1}:${colLetter(c2)}${r2}"/>`)
          .join('')}</mergeCells>`
      : '';

    const dim = `A1:${colLetter(Math.max(this.maxCol, 1))}${Math.max(this.maxRow, 1)}`;
    const paneXml = this.freezeRow
      ? `<sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="${this.freezeRow}" topLeftCell="A${this.freezeRow + 1}" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>`
      : '<sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>';

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<dimension ref="${dim}"/>
${paneXml}
<cols>${colsXml}</cols>
<sheetData>${rowsXml}</sheetData>
${mergesXml}
</worksheet>`;
  }
}



// ── Netlify Function handler ───────────────────────────────────────
exports.handler = async (event) => {
  if (event.httpMethod !== 'POST') {
    return { statusCode: 405, body: JSON.stringify({ error: 'Method not allowed' }) };
  }

  let payload;
  try {
    payload = JSON.parse(event.body || '{}');
  } catch (e) {
    return { statusCode: 400, body: JSON.stringify({ error: 'Invalid JSON body' }) };
  }

  const { orders = [], products = [], categories = [] } = payload;

  if (!orders.length) {
    return { statusCode: 400, body: JSON.stringify({ error: 'No orders provided' }) };
  }
  if (!products.length) {
    return { statusCode: 400, body: JSON.stringify({ error: 'No products provided' }) };
  }

  const DEFAULT_CAT_ORDER = [
    { id: 'sauces', label: 'Sauces' }, { id: 'oven', label: 'Oven' },
    { id: 'desserts', label: 'Desserts' }, { id: 'small-packs', label: 'Small Packs' },
    { id: 'cheeses', label: 'Cheeses' }, { id: 'cured-meat', label: 'Cured Meat' },
    { id: 'vegg', label: 'Vegg' }, { id: 'frozen', label: 'Frozen' },
    { id: 'dry', label: 'Dry' }, { id: 'pastas', label: 'Pastas' },
    { id: 'drinks', label: 'Drinks' }, { id: 'oils', label: 'Oils' },
    { id: 'packaging', label: 'Packaging' }, { id: 'chemical', label: 'Chemical' },
    { id: 'office', label: 'Office' }, { id: 'backup', label: 'Backup' },
  ];
  const catOrder = categories.length ? categories : DEFAULT_CAT_ORDER;

  const productsByCat = {};
  for (const cat of catOrder) productsByCat[cat.id] = [];
  for (const p of products) {
    if (!productsByCat[p.cat]) productsByCat[p.cat] = [];
    productsByCat[p.cat].push(p);
  }

  const totalQtyByProduct = {};
  for (const o of orders) {
    for (const [pid, qty] of Object.entries(o.items || {})) {
      totalQtyByProduct[pid] = (totalQtyByProduct[pid] || 0) + (parseInt(qty, 10) || 0);
    }
  }

  const wb = new Workbook();

  const HEADER_FILL = 'FF1D9E75';
  const CAT_FILL = 'FFE8E8E8';
  const ZERO_COLOR = 'FFB0B0B0';
  const GREEN_COLOR = 'FF1D9E75';
  const WHITE_COLOR = 'FFFFFFFF';
  const TOTAL_BORDER = { top: { style: 'medium', color: GREEN_COLOR } };
  const ROW_BORDER = { bottom: { style: 'thin', color: 'FFDDDDDD' } };

  function buildProductQtySheet(sheet, { includeAllProducts, qtyMap, checkbox, title }) {
    const lastCol = checkbox ? 3 : 2;
    sheet.setColWidth(1, 34);
    sheet.setColWidth(2, 12);
    if (checkbox) sheet.setColWidth(3, 10);

    let r = 1;
    if (title) {
      sheet.mergeCells(r, 1, r, lastCol);
      sheet.setCell(r, 1, title, { bold: true, fontSize: 14, fontColor: GREEN_COLOR });
      r += 2;
    }

    sheet.setCell(r, 1, 'PRODUCT', { bold: true, fontColor: WHITE_COLOR, fillColor: HEADER_FILL });
    sheet.setCell(r, 2, 'QTY', { bold: true, fontColor: WHITE_COLOR, fillColor: HEADER_FILL, align: { h: 'center' } });
    if (checkbox) sheet.setCell(r, 3, 'CHECK', { bold: true, fontColor: WHITE_COLOR, fillColor: HEADER_FILL, align: { h: 'center' } });
    sheet.freezeHeaderRows(r);
    r += 1;

    let grandTotal = 0;

    for (const cat of catOrder) {
      const prods = (productsByCat[cat.id] || []).slice().sort((a, b) => a.name.localeCompare(b.name));
      const rowsForCat = includeAllProducts ? prods : prods.filter((p) => (qtyMap[p.id] || 0) > 0);
      if (!rowsForCat.length) continue;

      sheet.setCell(r, 1, cat.label.toUpperCase(), { bold: true, fillColor: CAT_FILL });
      sheet.setCell(r, 2, '', { fillColor: CAT_FILL });
      if (checkbox) sheet.setCell(r, 3, '', { fillColor: CAT_FILL });
      r += 1;

      for (const p of rowsForCat) {
        const qty = qtyMap[p.id] || 0;
        grandTotal += qty;
        const name = p.description ? `${p.name} (${p.description})` : p.name;
        const color = qty > 0 ? null : ZERO_COLOR;

        sheet.setCell(r, 1, name, { fontColor: color, border: ROW_BORDER });
        sheet.setCell(r, 2, qty, { bold: qty > 0, fontColor: color, align: { h: 'center' }, border: ROW_BORDER });
        if (checkbox) sheet.setCell(r, 3, '', { border: ROW_BORDER });
        r += 1;
      }
    }
  }

  function buildMultiClientSheet(sheet, { includeAllProducts, title }) {
    const numClients = orders.length;
    const lastCol = 2 + numClients; // PRODUCT + one col per client + TOTAL

    sheet.setColWidth(1, 34);
    for (let i = 0; i < numClients; i++) sheet.setColWidth(2 + i, 13);
    sheet.setColWidth(lastCol, 12);

    let r = 1;
    if (title) {
      sheet.mergeCells(r, 1, r, lastCol);
      sheet.setCell(r, 1, title, { bold: true, fontSize: 14, fontColor: GREEN_COLOR });
      r += 2;
    }

    sheet.setCell(r, 1, 'PRODUCT', { bold: true, fontColor: WHITE_COLOR, fillColor: HEADER_FILL });
    orders.forEach((o, i) => {
      const label = o.order_number ? `#${o.order_number} ${o.username}` : o.username;
      sheet.setCell(r, 2 + i, label, { bold: true, fontColor: WHITE_COLOR, fillColor: HEADER_FILL, align: { h: 'center', wrap: true } });
    });
    sheet.setCell(r, lastCol, 'TOTAL', { bold: true, fontColor: WHITE_COLOR, fillColor: HEADER_FILL, align: { h: 'center' } });
    sheet.freezeHeaderRows(r);
    r += 1;

    const colTotals = new Array(numClients).fill(0);
    let grandTotal = 0;

    for (const cat of catOrder) {
      const prods = (productsByCat[cat.id] || []).slice().sort((a, b) => a.name.localeCompare(b.name));
      const rowsForCat = includeAllProducts
        ? prods
        : prods.filter((p) => (totalQtyByProduct[p.id] || 0) > 0);
      if (!rowsForCat.length) continue;

      sheet.setCell(r, 1, cat.label.toUpperCase(), { bold: true, fillColor: CAT_FILL });
      for (let i = 0; i < numClients; i++) sheet.setCell(r, 2 + i, '', { fillColor: CAT_FILL });
      sheet.setCell(r, lastCol, '', { fillColor: CAT_FILL });
      r += 1;

      for (const p of rowsForCat) {
        const name = p.description ? `${p.name} (${p.description})` : p.name;
        let rowTotal = 0;

        orders.forEach((o, i) => {
          const qty = parseInt((o.items || {})[p.id], 10) || 0;
          rowTotal += qty;
          colTotals[i] += qty;
          const color = qty > 0 ? null : ZERO_COLOR;
          sheet.setCell(r, 2 + i, qty, { bold: qty > 0, fontColor: color, align: { h: 'center' }, border: ROW_BORDER });
        });

        grandTotal += rowTotal;
        const rowColor = rowTotal > 0 ? null : ZERO_COLOR;
        sheet.setCell(r, 1, name, { fontColor: rowColor, border: ROW_BORDER });
        sheet.setCell(r, lastCol, rowTotal, { bold: rowTotal > 0, fontColor: rowTotal > 0 ? GREEN_COLOR : ZERO_COLOR, align: { h: 'center' }, border: ROW_BORDER });
        r += 1;
      }
    }
  }

  // Sheet 1: All Products (one column per client + TOTAL)
  const allSheet = wb.addSheet('All Products');
  buildMultiClientSheet(allSheet, {
    includeAllProducts: true,
    title: `All Products — ${orders.length} order${orders.length !== 1 ? 's' : ''} combined`,
  });

  // Sheet 2: Ordered Only (one column per client + TOTAL)
  const orderedSheet = wb.addSheet('Ordered Only');
  buildMultiClientSheet(orderedSheet, {
    includeAllProducts: false,
    title: `Ordered Products Only — ${orders.length} order${orders.length !== 1 ? 's' : ''} combined`,
  });

  // Sheet 3+: one per client order
  const usedNames = new Set(['All Products', 'Ordered Only']);
  function safeSheetName(raw) {
    let name = raw.replace(/[\\/?*\[\]:]/g, ' ').trim().substring(0, 28) || 'Order';
    let candidate = name;
    let i = 2;
    while (usedNames.has(candidate)) {
      candidate = `${name} (${i})`.substring(0, 31);
      i += 1;
    }
    usedNames.add(candidate);
    return candidate;
  }

  for (const order of orders) {
    const qtyMap = {};
    for (const [pid, qty] of Object.entries(order.items || {})) {
      qtyMap[pid] = (qtyMap[pid] || 0) + (parseInt(qty, 10) || 0);
    }
    const label = order.order_number ? `#${order.order_number} ${order.username}` : order.username;
    const sheet = wb.addSheet(safeSheetName(label));

    const titleParts = [order.username];
    if (order.order_number) titleParts.push(`#${order.order_number}`);
    if (order.ordered_by) titleParts.push(`Placed by: ${order.ordered_by}`);

    buildProductQtySheet(sheet, {
      includeAllProducts: false,
      qtyMap,
      checkbox: true,
      title: titleParts.join('  •  '),
    });
  }

  const buffer = wb.toBuffer();

  return {
    statusCode: 200,
    headers: {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': 'attachment; filename="in-production.xlsx"',
    },
    body: buffer.toString('base64'),
    isBase64Encoded: true,
  };
};
