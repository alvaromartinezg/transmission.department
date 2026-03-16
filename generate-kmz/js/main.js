const COLUMN_HEADERS = [
  'SITES', 'LAT', 'LONG', 'HOUSE', 'TOWER', 'IP', 'DESCRIPTION', 'FOLDER'
];

const COLUMN_KEYS = [
  'sites', 'lat', 'long', 'house', 'tower', 'ip', 'description', 'folder'
];

const INITIAL_ROWS = 200;

function createEmptyRow() {
  return ['', '', '', '', '', '', '', ''];
}

function createEmptyData(rowsCount) {
  return Array.from({ length: rowsCount }, () => createEmptyRow());
}

const container = document.getElementById('grid');

const hot = new Handsontable(container, {
  data: createEmptyData(INITIAL_ROWS),
  colHeaders: COLUMN_HEADERS,
  rowHeaders: true,
  width: '100%',
  height: 620,
  stretchH: 'all',
  manualColumnResize: true,
  manualRowResize: true,
  minSpareRows: 50,
  minRows: INITIAL_ROWS,
  contextMenu: true,
  copyPaste: true,
  licenseKey: 'non-commercial-and-evaluation',
  columns: [
    { type: 'text' },
    { type: 'numeric', numericFormat: { pattern: '0.000000' } },
    { type: 'numeric', numericFormat: { pattern: '0.000000' } },
    { type: 'text' },
    { type: 'text' },
    { type: 'text' },
    { type: 'text' },
    { type: 'text' }
  ],
  afterChange: function (changes, source) {
    if (source !== 'loadData') {
      updateStatus();
    }
  },
  afterPaste: function () {
    ensureExtraRows();
    updateStatus();
  }
});

function ensureExtraRows() {
  const usedRows = getUsedRowsCount();
  const totalRows = hot.countRows();
  const buffer = 100;

  if ((totalRows - usedRows) < buffer) {
    hot.alter('insert_row_below', totalRows - 1, buffer);
  }
}

function isRowCompletelyEmpty(row) {
  return row.every(cell => String(cell ?? '').trim() === '');
}

function getAllData() {
  return hot.getData();
}

function getUsedRows() {
  return getAllData().filter(row => !isRowCompletelyEmpty(row));
}

function getUsedRowsCount() {
  return getUsedRows().length;
}

function parseRow(row, index) {
  const obj = {};
  COLUMN_KEYS.forEach((key, i) => {
    obj[key] = String(row[i] ?? '').trim();
  });
  obj._rowNumber = index + 1;
  return obj;
}

function normalizeCoordinate(value) {
  return String(value ?? '')
    .replace(/°/g, '')
    .replace(/\s+/g, '')
    .trim();
}

function validateRow(obj) {
  const errors = [];

  if (!obj.sites) errors.push('SITES is empty');
  if (!obj.lat) errors.push('LAT is empty');
  if (!obj.long) errors.push('LONG is empty');
  if (!obj.folder) errors.push('FOLDER is empty');

  const latText = normalizeCoordinate(obj.lat);
  const lonText = normalizeCoordinate(obj.long);

  const lat = Number(latText);
  const lon = Number(lonText);

  if (obj.lat && Number.isNaN(lat)) {
    errors.push('LAT is not numeric');
  }

  if (obj.long && Number.isNaN(lon)) {
    errors.push('LONG is not numeric');
  }

  if (!Number.isNaN(lat) && obj.lat && !(lat < 0)) {
    errors.push('LAT must be negative');
  }

  if (!Number.isNaN(lon) && obj.long && !(lon < 0)) {
    errors.push('LONG must be negative');
  }

  return {
    valid: errors.length === 0,
    errors
  };
}

function getValidatedData() {
  const rows = getUsedRows();
  const parsed = rows.map((row, idx) => parseRow(row, idx));
  return parsed.map(item => ({
    ...item,
    validation: validateRow(item)
  }));
}

function updateStatus() {
  const results = getValidatedData();
  const validCount = results.filter(r => r.validation.valid).length;
  const invalidRows = results.filter(r => !r.validation.valid);

  document.getElementById('countRows').textContent = results.length;
  document.getElementById('countValid').textContent = validCount;
  document.getElementById('countInvalid').textContent = invalidRows.length;

  const errorList = document.getElementById('errorList');

  if (invalidRows.length === 0) {
    errorList.innerHTML = '<li>No errors found.</li>';
    return;
  }

  errorList.innerHTML = invalidRows.map(r => {
    return `<li><b>Row ${r._rowNumber}:</b> ${escapeHtml(r.validation.errors.join(' | '))}</li>`;
  }).join('');
}

function escapeHtml(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function groupByFolder(records) {
  const folders = {};
  for (const r of records) {
    if (!folders[r.folder]) folders[r.folder] = [];
    folders[r.folder].push(r);
  }
  return folders;
}

function generateKmlString(validRecords) {
  const folders = groupByFolder(validRecords);

  let kml = `<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2">
  <Document>
    <name>Exported</name>

    <Style id="siteBalloonStyle">
      <BalloonStyle>
        <text><![CDATA[
<div>
  <b>SITES</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$[SITES]<br>
  <b>LAT</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$[LAT]<br>
  <b>LONG</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$[LONG]<br>
  <b>HOUSE</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$[HOUSE]<br>
  <b>TOWER</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$[TOWER]<br>
  <b>IP</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;$[IP]<br>
  <b>DESCRIPTION</b>&nbsp;$[DESCRIPTION]
</div>
        ]]></text>
      </BalloonStyle>
    </Style>`;

  for (const folderName of Object.keys(folders)) {
    kml += `
    <Folder>
      <name>${escapeHtml(folderName)}</name>`;

    for (const r of folders[folderName]) {
      const latClean = normalizeCoordinate(r.lat);
      const lonClean = normalizeCoordinate(r.long);

      kml += `
      <Placemark>
        <name>${escapeHtml(r.sites)}</name>
        <Snippet maxLines="0"></Snippet>
        <styleUrl>#siteBalloonStyle</styleUrl>
        <ExtendedData>
          <Data name="SITES"><value>${escapeHtml(r.sites)}</value></Data>
          <Data name="LAT"><value>${escapeHtml(latClean)}</value></Data>
          <Data name="LONG"><value>${escapeHtml(lonClean)}</value></Data>
          <Data name="HOUSE"><value>${escapeHtml(r.house)}</value></Data>
          <Data name="TOWER"><value>${escapeHtml(r.tower)}</value></Data>
          <Data name="IP"><value>${escapeHtml(r.ip)}</value></Data>
          <Data name="DESCRIPTION"><value>${escapeHtml(r.description)}</value></Data>
        </ExtendedData>
        <Point>
          <coordinates>${lonClean},${latClean},0</coordinates>
        </Point>
      </Placemark>`;
    }

    kml += `
    </Folder>`;
  }

  kml += `
  </Document>
</kml>`;

  return kml;
}

function getExportReadyDataOrShowError() {
  const results = getValidatedData();

  if (results.length === 0) {
    alert('There is no data to export.');
    return null;
  }

  const invalidRows = results.filter(r => !r.validation.valid);

  if (invalidRows.length > 0) {
    alert('There are rows with validation errors. Please fix them before exporting.');
    updateStatus();
    return null;
  }

  return results.map(r => ({
    sites: r.sites,
    lat: r.lat,
    long: r.long,
    house: r.house,
    tower: r.tower,
    ip: r.ip,
    description: r.description,
    folder: r.folder
  }));
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function exportKML() {
  const data = getExportReadyDataOrShowError();
  if (!data) return;

  const kml = generateKmlString(data);
  const blob = new Blob([kml], {
    type: 'application/vnd.google-earth.kml+xml'
  });

  downloadBlob(blob, 'Exported.kml');
}

async function exportKMZ() {
  const data = getExportReadyDataOrShowError();
  if (!data) return;

  const kml = generateKmlString(data);
  const zip = new JSZip();
  zip.file('doc.kml', kml);

  const blob = await zip.generateAsync({
    type: 'blob',
    compression: 'DEFLATE',
    compressionOptions: { level: 9 }
  });

  downloadBlob(blob, 'Exported.kmz');
}

document.getElementById('btnClear').addEventListener('click', () => {
  const ok = confirm('Do you want to clear the entire sheet?');
  if (!ok) return;
  hot.loadData(createEmptyData(INITIAL_ROWS));
  updateStatus();
});

document.getElementById('btnExportKML').addEventListener('click', exportKML);
document.getElementById('btnExportKMZ').addEventListener('click', exportKMZ);

updateStatus();
