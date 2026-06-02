const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const { config } = require('./config');

const REQUIRED_COLUMNS = [
  '货物品名',
  '报关单号',
  '报关单号项号',
  '成交单位',
  '成交数量',
  '发票号码',
  '发票号码行号',
  '计税金额'
];

const STATUS_COLUMNS = {
  status: '处理状态',
  message: '处理说明',
  timestamp: '处理时间'
};

const COMPLETED_STATUS = '已处理';

function normalizeText(value) {
  return String(value ?? '').replace(/\s+/g, ' ').trim();
}

function normalizeCode(value) {
  return normalizeText(value).replace(/,/g, '');
}

function normalizeNumberString(value) {
  const text = normalizeText(value).replace(/,/g, '');
  if (!text) return '';
  if (!/^-?\d+(?:\.\d+)?$/.test(text)) {
    return text;
  }
  const number = Number(text);
  if (!Number.isFinite(number)) {
    return text;
  }
  return String(number);
}

function toNumber(value) {
  const normalized = normalizeNumberString(value);
  if (!normalized || !/^-?\d+(?:\.\d+)?$/.test(normalized)) {
    return NaN;
  }
  return Number(normalized);
}

function parseRemarkLineNos(values = []) {
  const numbers = [];
  for (const value of values) {
    const text = normalizeText(value);
    if (!text) continue;
    const matches = text.match(/\d+/g) || [];
    for (const match of matches) {
      const normalized = normalizeNumberString(match);
      if (normalized) {
        numbers.push(normalized);
      }
    }
  }
  return Array.from(new Set(numbers));
}

function parseInvoiceLineNos(value) {
  const text = normalizeText(value);
  if (!text) {
    return [];
  }

  const matches = text.match(/\d+/g) || [];
  return Array.from(new Set(matches
    .map(match => normalizeNumberString(match))
    .filter(Boolean)));
}

function buildColumnIndexMap(headerRow) {
  const normalizedHeaders = headerRow.map(cell => normalizeText(cell));
  const indexMap = new Map();
  normalizedHeaders.forEach((header, index) => {
    if (header && !indexMap.has(header)) {
      indexMap.set(header, index);
    }
  });

  const missing = REQUIRED_COLUMNS.filter(column => !indexMap.has(column));
  if (missing.length > 0) {
    throw new Error(`Excel 缺少必需列：${missing.join('、')}`);
  }

  const lastRequiredIndex = Math.max(...Array.from(indexMap.values()));
  const remarkIndexes = normalizedHeaders
    .map((header, index) => ({ header, index }))
    .filter(({ header, index }) => header === '备注' || /^备注\d+$/.test(header) || (index > lastRequiredIndex && !header))
    .map(item => item.index);

  return {
    normalizedHeaders,
    indexMap,
    remarkIndexes,
    statusIndexes: {
      status: indexMap.get(STATUS_COLUMNS.status) ?? -1,
      message: indexMap.get(STATUS_COLUMNS.message) ?? -1,
      timestamp: indexMap.get(STATUS_COLUMNS.timestamp) ?? -1
    }
  };
}

function listPendingExcelFiles() {
  const pendingDir = config.paths.pendingDir;
  if (!fs.existsSync(pendingDir)) {
    return [];
  }

  return fs.readdirSync(pendingDir, { withFileTypes: true })
    .filter(entry => entry.isFile())
    .filter(entry => !entry.name.startsWith('~$'))
    .map(entry => path.join(pendingDir, entry.name))
    .filter(filePath => /\.(xlsx|xls)$/i.test(filePath))
    .sort((left, right) => left.localeCompare(right, 'zh-CN'));
}

function resolvePendingExcelFile() {
  const files = listPendingExcelFiles();
  if (files.length === 0) {
    throw new Error(`待处理目录没有 Excel 文件：${config.paths.pendingDir}`);
  }
  if (files.length > 1) {
    throw new Error(`待处理目录存在多个 Excel，请只保留 1 个：${files.map(file => path.basename(file)).join('、')}`);
  }
  return files[0];
}

function getEffectiveSheetName(workbook, requestedSheetName = '') {
  const normalizedRequested = normalizeText(requestedSheetName);
  if (normalizedRequested) {
    return normalizedRequested;
  }
  return workbook.SheetNames[0] || '';
}

function valueAt(row, index) {
  return index >= 0 ? row[index] ?? '' : '';
}

function readWorkbookMatrix(filePath, requestedSheetName = '') {
  const workbook = XLSX.readFile(filePath, { cellDates: false, raw: false, sheetRows: 10000 });
  const targetSheetName = getEffectiveSheetName(workbook, requestedSheetName);
  const sheet = workbook.Sheets[targetSheetName];

  if (!sheet) {
    throw new Error(`工作表不存在：${targetSheetName}`);
  }

  // 修复：重置损坏的sheet范围
  if (sheet['!ref'] && sheet['!ref'].includes('XEI')) {
    const range = XLSX.utils.decode_range(sheet['!ref']);
    range.e.r = Math.min(range.e.r, 9999);
    sheet['!ref'] = XLSX.utils.encode_range(range);
  }

  const matrix = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    defval: '',
    blankrows: false
  });

  if (matrix.length < 2) {
    throw new Error('Excel 没有可处理的数据行');
  }

  return {
    workbook,
    targetSheetName,
    matrix
  };
}

function ensureStatusColumns(headerRow) {
  const nextHeaderRow = [...headerRow];
  const indexes = {};

  for (const [key, title] of Object.entries(STATUS_COLUMNS)) {
    let index = nextHeaderRow.findIndex(cell => normalizeText(cell) === title);
    if (index < 0) {
      index = nextHeaderRow.length;
      nextHeaderRow.push(title);
    }
    indexes[key] = index;
  }

  return {
    headerRow: nextHeaderRow,
    indexes
  };
}

function loadTaxRows(options = {}) {
  const filePath = options.filePath || resolvePendingExcelFile();
  const requestedSheetName = Object.prototype.hasOwnProperty.call(options, 'sheetName') ? options.sheetName : config.input.sheetName;
  const { targetSheetName, matrix } = readWorkbookMatrix(filePath, requestedSheetName);

  const [headerRow, ...dataRows] = matrix;
  const { indexMap, remarkIndexes, statusIndexes } = buildColumnIndexMap(headerRow);

  return dataRows
    .map((row, dataIndex) => {
      const rowNumber = dataIndex + 2;
      const goodsName = normalizeText(valueAt(row, indexMap.get('货物品名')));
      const declarationNo = normalizeCode(valueAt(row, indexMap.get('报关单号')));
      const declarationItemNo = normalizeNumberString(valueAt(row, indexMap.get('报关单号项号')));
      const dealUnit = normalizeText(valueAt(row, indexMap.get('成交单位')));
      const dealQty = normalizeNumberString(valueAt(row, indexMap.get('成交数量')));
      const invoiceNo = normalizeCode(valueAt(row, indexMap.get('发票号码')));
      const invoiceLineNo = normalizeNumberString(valueAt(row, indexMap.get('发票号码行号')));
      const parsedInvoiceLineNos = parseInvoiceLineNos(valueAt(row, indexMap.get('发票号码行号')));
      const taxAmount = normalizeNumberString(valueAt(row, indexMap.get('计税金额')));
      const remarkValues = remarkIndexes.map(index => valueAt(row, index));
      const remarkLineNos = parseRemarkLineNos(remarkValues);
      const selectedInvoiceLineNos = Array.from(new Set([...parsedInvoiceLineNos, ...remarkLineNos].filter(Boolean)));
      const processingStatus = normalizeText(valueAt(row, statusIndexes.status));
      const processingMessage = normalizeText(valueAt(row, statusIndexes.message));
      const processingTime = normalizeText(valueAt(row, statusIndexes.timestamp));

      return {
        rowNumber,
        sourceFile: filePath,
        sheetName: targetSheetName,
        goodsName,
        declarationNo,
        declarationItemNo,
        dealUnit,
        dealQty,
        invoiceNo,
        invoiceLineNo,
        selectedInvoiceLineNos,
        taxAmount,
        processingStatus,
        processingMessage,
        processingTime,
        remarkValues,
        raw: {
          goodsName: valueAt(row, indexMap.get('货物品名')),
          declarationNo: valueAt(row, indexMap.get('报关单号')),
          declarationItemNo: valueAt(row, indexMap.get('报关单号项号')),
          dealUnit: valueAt(row, indexMap.get('成交单位')),
          dealQty: valueAt(row, indexMap.get('成交数量')),
          invoiceNo: valueAt(row, indexMap.get('发票号码')),
          invoiceLineNo: valueAt(row, indexMap.get('发票号码行号')),
          taxAmount: valueAt(row, indexMap.get('计税金额')),
          processingStatus: valueAt(row, statusIndexes.status),
          processingMessage: valueAt(row, statusIndexes.message),
          processingTime: valueAt(row, statusIndexes.timestamp)
        }
      };
    })
    .filter(row => row.declarationNo && row.goodsName);
}

function updateTaxRowStatus(options = {}) {
  const filePath = options.filePath || resolvePendingExcelFile();
  const requestedSheetName = Object.prototype.hasOwnProperty.call(options, 'sheetName') ? options.sheetName : config.input.sheetName;
  const rowNumber = Number(options.rowNumber);
  if (!Number.isInteger(rowNumber) || rowNumber < 2) {
    throw new Error('更新 Excel 状态时缺少有效的行号');
  }

  const { workbook, targetSheetName, matrix } = readWorkbookMatrix(filePath, requestedSheetName);
  const { headerRow, indexes } = ensureStatusColumns(matrix[0] || []);
  const targetIndex = rowNumber - 1;
  if (targetIndex >= matrix.length) {
    throw new Error(`Excel 中不存在第 ${rowNumber} 行`);
  }

  const nextMatrix = matrix.map(row => Array.isArray(row) ? [...row] : []);
  nextMatrix[0] = headerRow;
  nextMatrix[targetIndex] = Array.isArray(nextMatrix[targetIndex]) ? [...nextMatrix[targetIndex]] : [];

  nextMatrix[targetIndex][indexes.status] = normalizeText(options.status);
  nextMatrix[targetIndex][indexes.message] = normalizeText(options.message);
  nextMatrix[targetIndex][indexes.timestamp] = normalizeText(options.timestamp || new Date().toISOString());

  workbook.Sheets[targetSheetName] = XLSX.utils.aoa_to_sheet(nextMatrix);
  XLSX.writeFile(workbook, filePath);
}

if (require.main === module) {
  const filePath = resolvePendingExcelFile();
  const rows = loadTaxRows({ filePath });
  const preview = rows.slice(0, 3).map(row => ({
    rowNumber: row.rowNumber,
    goodsName: row.goodsName,
    declarationNo: row.declarationNo,
    declarationItemNo: row.declarationItemNo,
    dealUnit: row.dealUnit,
    dealQty: row.dealQty,
    invoiceNo: row.invoiceNo,
    invoiceLineNo: row.invoiceLineNo,
    selectedInvoiceLineNos: row.selectedInvoiceLineNos,
    taxAmount: row.taxAmount,
    processingStatus: row.processingStatus,
    processingMessage: row.processingMessage,
    processingTime: row.processingTime
  }));

  console.log(JSON.stringify({
    file: path.resolve(filePath),
    sheet: rows[0]?.sheetName || config.input.sheetName || '',
    rowCount: rows.length,
    preview
  }, null, 2));
}

module.exports = {
  STATUS_COLUMNS,
  COMPLETED_STATUS,
  listPendingExcelFiles,
  resolvePendingExcelFile,
  loadTaxRows,
  updateTaxRowStatus,
  normalizeText,
  normalizeCode,
  normalizeNumberString,
  parseRemarkLineNos,
  parseInvoiceLineNos,
  toNumber
};
