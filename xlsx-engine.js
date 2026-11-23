/* ==========================================================================
   AcordosHub — Pure-JS Zero-Dependency XLSX Reader & Writer Engine
   Supports reading multi-sheet .xlsx (Deflate & Stored) and exporting valid .xlsx
   ========================================================================== */

(function (global) {
  'use strict';

  /* -------------------- CRC32 Table for ZIP Writing -------------------- */
  const CRC_TABLE = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) {
      c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
    }
    CRC_TABLE[i] = c >>> 0;
  }

  function crc32(bytes) {
    let crc = 0xffffffff;
    for (let i = 0; i < bytes.length; i++) {
      crc = CRC_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
    }
    return (crc ^ 0xffffffff) >>> 0;
  }

  /* -------------------- ZIP Reader (Central Directory) -------------------- */
  async function inflateRaw(bytes) {
    if (typeof DecompressionStream !== 'undefined') {
      const ds = new DecompressionStream('deflate-raw');
      const stream = new Blob([bytes]).stream().pipeThrough(ds);
      const buf = await new Response(stream).arrayBuffer();
      return new Uint8Array(buf);
    }
    throw new Error('Seu navegador não suporta descompressão nativa de arquivos .xlsx. Atualize o navegador ou use .csv.');
  }

  async function unzipEntries(arrayBuffer) {
    const u8 = new Uint8Array(arrayBuffer);
    const dv = new DataView(arrayBuffer);

    // Locate End of Central Directory (0x06054b50)
    let eocdOffset = -1;
    const minOffset = Math.max(0, u8.length - 65557);
    for (let i = u8.length - 22; i >= minOffset; i--) {
      if (dv.getUint32(i, true) === 0x06054b50) {
        eocdOffset = i;
        break;
      }
    }
    if (eocdOffset === -1) {
      throw new Error('Arquivo .xlsx inválido ou corrompido (assinatura ZIP não encontrada).');
    }

    const totalEntries = dv.getUint16(eocdOffset + 10, true);
    const cdOffset = dv.getUint32(eocdOffset + 16, true);
    const decoder = new TextDecoder('utf-8');
    const files = {};

    let ptr = cdOffset;
    for (let i = 0; i < totalEntries; i++) {
      if (dv.getUint32(ptr, true) !== 0x02014b50) break;

      const method = dv.getUint16(ptr + 10, true);
      const compSize = dv.getUint32(ptr + 20, true);
      const nameLen = dv.getUint16(ptr + 28, true);
      const extraLen = dv.getUint16(ptr + 30, true);
      const commentLen = dv.getUint16(ptr + 32, true);
      const localHeaderOffset = dv.getUint32(ptr + 42, true);

      const name = decoder.decode(u8.subarray(ptr + 46, ptr + 46 + nameLen));
      ptr += 46 + nameLen + extraLen + commentLen;

      // Read local header to find exact data offset
      if (dv.getUint32(localHeaderOffset, true) !== 0x04034b50) continue;
      const lhNameLen = dv.getUint16(localHeaderOffset + 26, true);
      const lhExtraLen = dv.getUint16(localHeaderOffset + 28, true);
      const dataStart = localHeaderOffset + 30 + lhNameLen + lhExtraLen;
      const compData = u8.subarray(dataStart, dataStart + compSize);

      if (name.endsWith('/')) continue;

      if (method === 0) {
        files[name] = compData;
      } else if (method === 8) {
        files[name] = await inflateRaw(compData);
      }
    }

    return files;
  }

  /* -------------------- Excel Helpers -------------------- */
  function colRefToIndex(ref) {
    const letters = String(ref || '').replace(/[^A-Z]/gi, '').toUpperCase();
    let idx = 0;
    for (let i = 0; i < letters.length; i++) {
      idx = idx * 26 + (letters.charCodeAt(i) - 64);
    }
    return idx - 1;
  }

  function excelSerialToISO(serial) {
    // Excel epoch: 1899-12-30 (accounts for 1900 leap-year bug)
    const utcDays = Math.floor(serial - 25569);
    const utcMs = utcDays * 86400 * 1000;
    const d = new Date(utcMs);
    if (Number.isNaN(d.getTime())) return String(serial);
    const y = d.getUTCFullYear();
    const m = String(d.getUTCMonth() + 1).padStart(2, '0');
    const day = String(d.getUTCDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }

  function parseXml(bytes) {
    const text = new TextDecoder('utf-8').decode(bytes);
    return new DOMParser().parseFromString(text, 'application/xml');
  }

  function getDateStyleSet(files) {
    const dateStyles = new Set();
    const BUILTIN_DATE_IDS = new Set([14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47]);
    if (!files['xl/styles.xml']) return dateStyles;

    try {
      const doc = parseXml(files['xl/styles.xml']);
      const customDateIds = new Set();
      doc.querySelectorAll('numFmt').forEach(el => {
        const id = Number(el.getAttribute('numFmtId'));
        const code = String(el.getAttribute('formatCode') || '').toLowerCase();
        if (/[dmy]/.test(code) && !code.includes('r$') && !code.includes('0.00')) {
          customDateIds.add(id);
        }
      });

      const cellXfs = doc.querySelector('cellXfs');
      if (cellXfs) {
        const xfs = cellXfs.querySelectorAll('xf');
        xfs.forEach((xf, idx) => {
          const numFmtId = Number(xf.getAttribute('numFmtId') || 0);
          if (BUILTIN_DATE_IDS.has(numFmtId) || customDateIds.has(numFmtId)) {
            dateStyles.add(idx);
          }
        });
      }
    } catch (_) {}
    return dateStyles;
  }

  function getSharedStrings(files) {
    const list = [];
    if (!files['xl/sharedStrings.xml']) return list;
    const doc = parseXml(files['xl/sharedStrings.xml']);
    const sis = doc.getElementsByTagName('si');
    for (let i = 0; i < sis.length; i++) {
      const ts = sis[i].getElementsByTagName('t');
      let str = '';
      for (let j = 0; j < ts.length; j++) {
        str += ts[j].textContent || '';
      }
      list.push(str);
    }
    return list;
  }

  function getSheetEntries(files) {
    const sheets = [];
    const relMap = {};

    if (files['xl/_rels/workbook.xml.rels']) {
      const relDoc = parseXml(files['xl/_rels/workbook.xml.rels']);
      const rels = relDoc.getElementsByTagName('Relationship');
      for (let i = 0; i < rels.length; i++) {
        const id = rels[i].getAttribute('Id');
        let target = rels[i].getAttribute('Target') || '';
        target = target.replace(/^\/+/, '');
        if (!target.startsWith('xl/')) target = 'xl/' + target;
        relMap[id] = target;
      }
    }

    if (files['xl/workbook.xml']) {
      const wbDoc = parseXml(files['xl/workbook.xml']);
      const sheetNodes = wbDoc.getElementsByTagName('sheet');
      for (let i = 0; i < sheetNodes.length; i++) {
        const name = sheetNodes[i].getAttribute('name') || `Aba ${i + 1}`;
        const rId = sheetNodes[i].getAttribute('r:id') || sheetNodes[i].getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'id');
        const path = (rId && relMap[rId]) || `xl/worksheets/sheet${i + 1}.xml`;
        if (files[path]) sheets.push({ name, path });
      }
    }

    if (sheets.length === 0) {
      Object.keys(files)
        .filter(k => /^xl\/worksheets\/sheet\d+\.xml$/i.test(k))
        .sort()
        .forEach((path, idx) => sheets.push({ name: `Aba ${idx + 1}`, path }));
    }

    return sheets;
  }

  function parseWorksheet(xmlBytes, sharedStrings, dateStyles) {
    const doc = parseXml(xmlBytes);
    const rowNodes = doc.getElementsByTagName('row');
    const matrix = [];
    let maxCol = 0;

    for (let r = 0; r < rowNodes.length; r++) {
      const cells = rowNodes[r].getElementsByTagName('c');
      const rowArr = [];
      for (let c = 0; c < cells.length; c++) {
        const cell = cells[c];
        const ref = cell.getAttribute('r');
        const colIdx = ref ? colRefToIndex(ref) : c;
        if (colIdx > maxCol) maxCol = colIdx;

        const type = cell.getAttribute('t') || '';
        const styleIdx = cell.getAttribute('s') !== null ? Number(cell.getAttribute('s')) : -1;
        let value = '';

        if (type === 's') {
          const vNode = cell.getElementsByTagName('v')[0];
          const sIdx = vNode ? Number(vNode.textContent) : -1;
          value = sharedStrings[sIdx] !== undefined ? sharedStrings[sIdx] : '';
        } else if (type === 'inlineStr') {
          const ts = cell.getElementsByTagName('t');
          let txt = '';
          for (let k = 0; k < ts.length; k++) txt += ts[k].textContent || '';
          value = txt;
        } else if (type === 'b') {
          const vNode = cell.getElementsByTagName('v')[0];
          value = vNode && vNode.textContent === '1' ? 'Sim' : 'Não';
        } else if (type === 'str' || type === 'e') {
          const vNode = cell.getElementsByTagName('v')[0];
          value = vNode ? vNode.textContent : '';
        } else {
          const vNode = cell.getElementsByTagName('v')[0];
          if (vNode && vNode.textContent !== '') {
            const num = Number(vNode.textContent);
            if (!Number.isNaN(num) && dateStyles.has(styleIdx) && num > 20000 && num < 80000) {
              value = excelSerialToISO(num);
            } else {
              value = Number.isNaN(num) ? vNode.textContent : num;
            }
          }
        }
        rowArr[colIdx] = value;
      }

      if (rowArr.some(v => v !== undefined && v !== null && String(v).trim() !== '')) {
        matrix.push(rowArr);
      }
    }

    if (matrix.length === 0) return { headers: [], rows: [] };

    const rawHeaders = matrix[0];
    const headers = [];
    for (let i = 0; i <= maxCol; i++) {
      const h = rawHeaders[i] !== undefined && String(rawHeaders[i]).trim() !== ''
        ? String(rawHeaders[i]).trim()
        : `Coluna ${i + 1}`;
      headers.push(h);
    }

    // Trim trailing completely empty columns
    while (headers.length > 1) {
      const lastIdx = headers.length - 1;
      const isGeneric = headers[lastIdx] === `Coluna ${lastIdx + 1}`;
      const anyVal = matrix.slice(1).some(r => r[lastIdx] !== undefined && String(r[lastIdx]).trim() !== '');
      if (isGeneric && !anyVal) headers.pop();
      else break;
    }

    const rows = matrix.slice(1).map((r, idx) => {
      const obj = { _id: `xlsx-${idx + 2}`, _rowNumber: idx + 2 };
      headers.forEach((h, ci) => {
        obj[h] = r[ci] !== undefined && r[ci] !== null ? r[ci] : '';
      });
      return obj;
    });

    return { headers, rows };
  }

  /**
   * Reads an ArrayBuffer of a .xlsx file and returns { sheets: [{ name, headers, rows }] }
   */
  async function readXLSX(arrayBuffer) {
    const files = await unzipEntries(arrayBuffer);
    const sharedStrings = getSharedStrings(files);
    const dateStyles = getDateStyleSet(files);
    const sheetMeta = getSheetEntries(files);

    if (sheetMeta.length === 0) {
      throw new Error('Nenhuma aba válida encontrada no arquivo .xlsx.');
    }

    const sheets = sheetMeta.map(meta => {
      const parsed = parseWorksheet(files[meta.path], sharedStrings, dateStyles);
      return { name: meta.name, headers: parsed.headers, rows: parsed.rows };
    });

    return { sheets };
  }

  /* -------------------- Pure-JS XLSX Writer (STORED ZIP) -------------------- */
  function xmlEscape(val) {
    return String(val === null || val === undefined ? '' : val)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  function indexToColRef(index) {
    let s = '';
    let n = index + 1;
    while (n > 0) {
      const rem = (n - 1) % 26;
      s = String.fromCharCode(65 + rem) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  function buildStoredZip(fileEntries) {
    const encoder = new TextEncoder();
    const localParts = [];
    const centralParts = [];
    let offset = 0;

    fileEntries.forEach(({ name, content }) => {
      const nameBytes = encoder.encode(name);
      const dataBytes = typeof content === 'string' ? encoder.encode(content) : content;
      const crc = crc32(dataBytes);
      const size = dataBytes.length;

      // Local File Header (30 bytes + name)
      const lh = new ArrayBuffer(30 + nameBytes.length);
      const ldv = new DataView(lh);
      ldv.setUint32(0, 0x04034b50, true);
      ldv.setUint16(4, 20, true);
      ldv.setUint16(6, 0, true);
      ldv.setUint16(8, 0, true); // STORED
      ldv.setUint16(10, 0, true);
      ldv.setUint16(12, 0x5421, true);
      ldv.setUint32(14, crc, true);
      ldv.setUint32(18, size, true);
      ldv.setUint32(22, size, true);
      ldv.setUint16(26, nameBytes.length, true);
      ldv.setUint16(28, 0, true);
      new Uint8Array(lh, 30).set(nameBytes);

      localParts.push(new Uint8Array(lh), dataBytes);

      // Central Directory Header (46 bytes + name)
      const ch = new ArrayBuffer(46 + nameBytes.length);
      const cdv = new DataView(ch);
      cdv.setUint32(0, 0x02014b50, true);
      cdv.setUint16(4, 20, true);
      cdv.setUint16(6, 20, true);
      cdv.setUint16(8, 0, true);
      cdv.setUint16(10, 0, true);
      cdv.setUint16(12, 0, true);
      cdv.setUint16(14, 0x5421, true);
      cdv.setUint32(16, crc, true);
      cdv.setUint32(20, size, true);
      cdv.setUint32(24, size, true);
      cdv.setUint16(28, nameBytes.length, true);
      cdv.setUint16(30, 0, true);
      cdv.setUint16(32, 0, true);
      cdv.setUint16(34, 0, true);
      cdv.setUint16(36, 0, true);
      cdv.setUint32(38, 0, true);
      cdv.setUint32(42, offset, true);
      new Uint8Array(ch, 46).set(nameBytes);

      centralParts.push(new Uint8Array(ch));
      offset += 30 + nameBytes.length + size;
    });

    const cdSize = centralParts.reduce((a, b) => a + b.length, 0);
    const eocd = new ArrayBuffer(22);
    const edv = new DataView(eocd);
    edv.setUint32(0, 0x06054b50, true);
    edv.setUint16(4, 0, true);
    edv.setUint16(6, 0, true);
    edv.setUint16(8, fileEntries.length, true);
    edv.setUint16(10, fileEntries.length, true);
    edv.setUint32(12, cdSize, true);
    edv.setUint32(16, offset, true);
    edv.setUint16(20, 0, true);

    const allParts = [...localParts, ...centralParts, new Uint8Array(eocd)];
    const totalLen = allParts.reduce((a, b) => a + b.length, 0);
    const out = new Uint8Array(totalLen);
    let pos = 0;
    allParts.forEach(p => { out.set(p, pos); pos += p.length; });
    return out;
  }

  /**
   * Generates a valid .xlsx Uint8Array from headers and row objects
   */
  function writeXLSX(headers, rows, sheetName = 'Relatorio') {
    const safeSheet = String(sheetName || 'Relatorio').replace(/[\\/*?:[\]]/g, '').slice(0, 31) || 'Relatorio';

    let rowsXml = '';
    // Header row
    rowsXml += `<row r="1">` + headers.map((h, ci) =>
      `<c r="${indexToColRef(ci)}1" t="inlineStr"><is><t>${xmlEscape(h)}</t></is></c>`
    ).join('') + `</row>`;

    // Data rows
    rows.forEach((row, ri) => {
      const rNum = ri + 2;
      const cellsXml = headers.map((h, ci) => {
        const ref = `${indexToColRef(ci)}${rNum}`;
        const v = row[h];
        if (typeof v === 'number' && Number.isFinite(v)) {
          return `<c r="${ref}"><v>${v}</v></c>`;
        }
        return `<c r="${ref}" t="inlineStr"><is><t>${xmlEscape(v === undefined || v === null ? '' : v)}</t></is></c>`;
      }).join('');
      rowsXml += `<row r="${rNum}">${cellsXml}</row>`;
    });

    const sheetXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>${rowsXml}</sheetData>
</worksheet>`;

    const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="${xmlEscape(safeSheet)}" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;

    const wbRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;

    const rootRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

    const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`;

    return buildStoredZip([
      { name: '[Content_Types].xml', content: contentTypesXml },
      { name: '_rels/.rels', content: rootRelsXml },
      { name: 'xl/workbook.xml', content: workbookXml },
      { name: 'xl/_rels/workbook.xml.rels', content: wbRelsXml },
      { name: 'xl/worksheets/sheet1.xml', content: sheetXml }
    ]);
  }

  global.XLSXEngine = { readXLSX, writeXLSX };
})(window);
