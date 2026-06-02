const fs = require('fs');
const path = require('path');
const http = require('http');
const { spawn } = require('child_process');
const readline = require('readline');
const { chromium } = require('playwright');
const { config, ensureProjectDirs, validateRuntimeConfig, readJsonIfExists } = require('./config');
const { loadTaxRows, resolvePendingExcelFile, updateTaxRowStatus, COMPLETED_STATUS, normalizeText, normalizeNumberString, toNumber } = require('./tax_refund_excel');

const TABLE_ROOT_CONFIGS = [
  { rootSelector: 'table', rowSelector: 'tbody tr', cellSelector: 'th, td' },
  { rootSelector: '.ant-table', rowSelector: '.ant-table-tbody tr', cellSelector: 'th, td' },
  { rootSelector: '.el-table', rowSelector: '.el-table__body-wrapper tbody tr', cellSelector: 'th, td' },
  { rootSelector: '[role="grid"]', rowSelector: '[role="row"]', cellSelector: '[role="columnheader"], [role="gridcell"], th, td' }
];

const HEADER_ALIASES = {
  declarationNo: ['报关单号'],
  goodsName: ['商品名称', '货物品名', '品名'],
  declarationItemNo: ['项号', '报关单号项号'],
  dealUnit: ['成交单位', '发票单位', '单位'],
  dealQty: ['成交数量', '出口/进货数量', '数量'],
  invoiceNo: ['发票号码'],
  invoiceLineNo: ['发票行号', '发票号码行号'],
  taxAmount: ['计税金额']
};

function nowStamp() {
  const iso = new Date().toISOString();
  return iso.replace(/[:.]/g, '-');
}

function ensureDir(dir) {
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
}

function readState() {
  return readJsonIfExists(config.paths.stateFile, {
    runs: []
  });
}

function writeState(state) {
  fs.writeFileSync(config.paths.stateFile, JSON.stringify(state, null, 2));
}

function appendRunState(patch) {
  const state = readState();
  state.runs = Array.isArray(state.runs) ? state.runs : [];
  state.runs.push({
    createdAt: new Date().toISOString(),
    ...patch
  });
  writeState(state);
}

// ============================================================
// 致命错误判断
// ============================================================
// 判断一个错误是否为"致命错误"——即需要用户介入才能恢复的错误
// 致命错误发生时，脚本应该暂停并询问用户选择后续操作
function isCriticalError(error) {
  const message = (error?.message || error?.toString() || '').toLowerCase();

  // 会话/登录过期
  if (message.includes('登录信息已过期') ||
      message.includes('session') ||
      message.includes('登录失效')) {
    return true;
  }

  // 浏览器连接断开
  if (message.includes('target page, context or browser has been closed') ||
      message.includes('browser closed') ||
      message.includes('connection closed') ||
      message.includes('playwright connection')) {
    return true;
  }

  // 导航失败（无法到达目标页面）
  if (message.includes('navigation') ||
      message.includes('导航') ||
      message.includes('navigate')) {
    return true;
  }

  // 页面崩溃或dom不可用
  if (message.includes('dom not available') ||
      message.includes('page not found') ||
      message.includes('frame not found')) {
    return true;
  }

  // 超时过多（可能是服务器问题，但也可能需要重新登录）
  // 这类不直接判定为致命，由调用方决定

  return false;
}

// ============================================================
// 致命错误用户交互
// ============================================================
// 发生致命错误时，弹出选项让用户选择后续操作
// rowNumber：行号（数字）或阶段描述（字符串如"导航阶段"）
// 返回：'retry' | 'relogin' | 'skip' | 'quit'
async function askCriticalErrorAction(rowNumber, errorMessage, logger) {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });

  // 根据 rowNumber 类型生成不同提示
  const isRowNumber = typeof rowNumber === 'number';
  const rowHint = isRowNumber ? '第 ' + rowNumber + ' 行' : rowNumber;
  const retryHint = isRowNumber ? '重新运行脚本时，加参数 --row=' + rowNumber + ' 即可从这一行继续' : '重新运行脚本即可重新开始';

  console.log('');
  console.log('========================================');
  console.log('  发生致命错误（' + rowHint + '）');
  console.log('========================================');
  console.log('错误原因：' + errorMessage);
  console.log('');
  if (isRowNumber) {
    console.log('提示：' + retryHint);
  }
  console.log('');
  console.log('请选择后续操作：');
  console.log('  [R] 重新登录（推荐）');
  console.log('      关闭浏览器，重新扫码');
  if (isRowNumber) {
    console.log('      重新运行脚本时加参数 --row=' + rowNumber);
  }
  if (isRowNumber) {
    console.log('  [T] 重试此行');
    console.log('      不关闭浏览器，直接重试当前行');
    console.log('      适用于服务器临时抖动的情况');
  }
  console.log('  [S] 跳过此行（仅对行处理错误有效）');
  console.log('      不写 Excel 状态，继续处理下一行');
  console.log('  [Q] 退出');
  console.log('      停止运行，Excel 状态不变，可重新运行脚本继续');
  console.log('');

  return new Promise((resolve) => {
    const question = () => {
      const optionStr = isRowNumber ? 'R/T/S/Q' : 'R/S/Q';
      rl.question('请输入选择 [' + optionStr + ']：', (answer) => {
        const a = answer.trim().toUpperCase();
        if (a === 'R') {
          rl.close();
          console.log('已选择：重新登录');
          resolve('relogin');
        } else if (a === 'T' && isRowNumber) {
          rl.close();
          console.log('已选择：重试此行');
          resolve('retry');
        } else if (a === 'S') {
          rl.close();
          console.log('已选择：跳过此行');
          resolve('skip');
        } else if (a === 'Q') {
          rl.close();
          console.log('已选择：退出');
          resolve('quit');
        } else {
          console.log('输入无效，请输入 ' + optionStr + ' 中的一个');
          question();
        }
      });
    };
    question();
  });
}

function getDebugEndpointUrl(port = config.browser.remoteDebuggingPort) {
  return `http://127.0.0.1:${port}/json/version`;
}

function fetchJson(url, timeoutMs = 2000) {
  return new Promise((resolve, reject) => {
    const request = http.get(url, { timeout: timeoutMs }, (response) => {
      let body = '';
      response.setEncoding('utf8');
      response.on('data', chunk => {
        body += chunk;
      });
      response.on('end', () => {
        if (response.statusCode && response.statusCode >= 400) {
          reject(new Error(`HTTP ${response.statusCode}`));
          return;
        }
        try {
          resolve(JSON.parse(body));
        } catch (error) {
          reject(error);
        }
      });
    });

    request.on('timeout', () => {
      request.destroy(new Error('timeout'));
    });
    request.on('error', reject);
  });
}

async function getExistingDebugWsEndpoint(port = config.browser.remoteDebuggingPort) {
  try {
    const payload = await fetchJson(getDebugEndpointUrl(port), 1500);
    const wsEndpoint = normalizeText(payload?.webSocketDebuggerUrl || '');
    return wsEndpoint || '';
  } catch {
    return '';
  }
}

async function connectToExistingBrowser(logger) {
  const wsEndpoint = await getExistingDebugWsEndpoint();
  if (!wsEndpoint) {
    return null;
  }

  logger.log(`[browser] 检测到可复用的 Chrome 调试端口：${config.browser.remoteDebuggingPort}`);
  const browser = await chromium.connectOverCDP(wsEndpoint);
  const context = browser.contexts()[0] || await browser.newContext();
  const pages = context.pages();

  // 遍历所有页面，找到包含目标内容的那个
  let targetPage = pages[0];
  for (const p of pages) {
    try {
      const url = p.url() || '';
      const title = await p.title().catch(() => '');
      logger.log(`[browser] 发现页面：${title} - ${url}`);

      // 跳过登录页面
      if (url.includes('loginb') || title.includes('登录')) {
        continue;
      }

      // 使用这个页面作为目标
      targetPage = p;
      logger.log(`[browser] 选择页面：${title}`);
      break;
    } catch (e) {}
  }

  logger.log('[browser] 已重新连接到现有 Chrome 实例');
  return { browser, context, page: targetPage, mode: 'connected' };
}

async function waitForDebugWsEndpoint(timeoutMs = 15000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    const wsEndpoint = await getExistingDebugWsEndpoint();
    if (wsEndpoint) {
      return wsEndpoint;
    }
    await new Promise(resolve => setTimeout(resolve, 300));
  }
  return '';
}

async function waitForChromeLaunchOutcome(chromeProcess, timeoutMs = 15000) {
  const startedAt = Date.now();
  while (Date.now() - startedAt < timeoutMs) {
    const wsEndpoint = await getExistingDebugWsEndpoint();
    if (wsEndpoint) {
      return { wsEndpoint, exitCode: null };
    }

    const exitCode = chromeProcess.exitCode;
    if (exitCode !== null) {
      return { wsEndpoint: '', exitCode };
    }

    await new Promise(resolve => setTimeout(resolve, 300));
  }

  return { wsEndpoint: '', exitCode: chromeProcess.exitCode };
}

async function launchManagedChrome(logger) {
  const chromeArgs = [
    `--remote-debugging-port=${config.browser.remoteDebuggingPort}`,
    `--user-data-dir=${config.browser.userDataDir}`,
    ...(config.browser.args || ['--start-maximized'])
  ];

  logger.log(`[browser] 启动本地 Chrome：${config.browser.chromePath}`);
  const chromeProcess = spawn(config.browser.chromePath, chromeArgs, {
    detached: true,
    stdio: 'ignore',
    windowsHide: false
  });
  chromeProcess.unref();

  const { wsEndpoint, exitCode } = await waitForChromeLaunchOutcome(chromeProcess);
  if (!wsEndpoint) {
    if (exitCode !== null) {
      throw new Error(`Chrome 启动后立即退出（exit=${exitCode}）。这通常表示该 profile 已被现有浏览器占用，或当前实例未开启调试端口；请关闭占用 ${config.browser.userDataDir} 的 Chrome 窗口后重试`);
    }
    throw new Error(`Chrome 已启动但未暴露调试端口 ${config.browser.remoteDebuggingPort}；如果已有同 profile 浏览器在运行，请先关闭后重试`);
  }

  const browser = await chromium.connectOverCDP(wsEndpoint);
  const context = browser.contexts()[0] || await browser.newContext();
  const page = context.pages()[0] || await context.newPage();
  logger.log('[browser] 已连接到新启动的 Chrome 实例');
  return { browser, context, page, mode: 'launched' };
}

async function launchOrConnectBrowser(logger) {
  const connected = await connectToExistingBrowser(logger).catch(error => {
    logger.log(`[browser] 连接现有 Chrome 失败：${error.message}`);
    return null;
  });
  if (connected) {
    return connected;
  }

  return launchManagedChrome(logger);
}

function buildProcessedFilePath(sourcePath) {
  const parsed = path.parse(sourcePath);
  const targetDir = config.paths.processedDir;
  let targetPath = path.join(targetDir, parsed.base);
  let attempt = 1;

  while (fs.existsSync(targetPath)) {
    targetPath = path.join(targetDir, `${parsed.name}-${nowStamp()}-${attempt}${parsed.ext}`);
    attempt += 1;
  }

  return targetPath;
}

function moveProcessedExcelFile(sourcePath, logger) {
  const targetPath = buildProcessedFilePath(sourcePath);
  fs.renameSync(sourcePath, targetPath);
  logger.log(`[file] 已归档 Excel：${sourcePath} -> ${targetPath}`);
  appendRunState({
    status: 'excel_archived',
    sourceFile: sourcePath,
    archivedFile: targetPath
  });
  return targetPath;
}
function createLogger() {
  const stamp = nowStamp();
  const logFile = path.join(config.paths.logDir, `run-${stamp}.log`);
  const lines = [];

  function log(...parts) {
    const message = parts.map(part => typeof part === 'string' ? part : JSON.stringify(part)).join(' ');
    const line = `[${new Date().toISOString()}] ${message}`;
    lines.push(line);
    console.log(message);
    fs.writeFileSync(logFile, `${lines.join('\n')}\n`, 'utf8');
  }

  return { logFile, log };
}


async function waitForUi(page, timeout = 1500) {
  if (!page || page.isClosed()) {
    return;
  }

  const safeTimeout = Math.max(timeout, 0);
  await page.waitForLoadState('domcontentloaded', { timeout: Math.min(Math.max(safeTimeout, 500), 2500) }).catch(() => {});
  if (safeTimeout >= 1000) {
    await page.waitForLoadState('networkidle', { timeout: Math.min(safeTimeout, 4000) }).catch(() => {});
  }
  await page.waitForTimeout(safeTimeout).catch(() => {});
}

async function waitForSearchResultsReady(page, timeoutMs = 12000) {
  const startedAt = Date.now();

  while (Date.now() - startedAt < timeoutMs) {
    const state = await page.evaluate(() => {
      const normalize = value => String(value ?? '').replace(/\s+/g, ' ').trim();
      const isVisible = el => {
        if (!el || !el.isConnected) return false;
        const style = window.getComputedStyle(el);
        if (!style || style.display === 'none' || style.visibility === 'hidden' || Number(style.opacity) === 0) {
          return false;
        }
        const rect = el.getBoundingClientRect();
        return rect.width > 0 && rect.height > 0;
      };

      const loadingSelectors = [
        '.el-loading-mask', '.el-icon-loading', '.el-loading-spinner',
        '.layui-layer-loading', '.ant-spin-spinning',
        '[class*="loading"]', '[class*="spinner"]', '[class*="spin"]'
      ];

      const rowSelectors = [
        'tbody tr', '.ant-table-tbody tr',
        '.el-table__body-wrapper tbody tr', '[role="row"]'
      ];

      const emptyTexts = ['暂无数据', '无数据', '未查询到数据', '没有查询到数据', '无匹配数据'];

      function checkWindow(win) {
        let loadingVisible = false;
        let anyVisibleRow = false;
        let emptyVisible = false;
        try {
          loadingVisible = loadingSelectors.some(selector =>
            Array.from(win.document.querySelectorAll(selector)).some(isVisible)
          );
          anyVisibleRow = rowSelectors.some(selector =>
            Array.from(win.document.querySelectorAll(selector)).some(row => isVisible(row) && normalize(row.textContent || row.innerText || ''))
          );
          const bodyText = normalize(win.document.body?.innerText || '');
          emptyVisible = emptyTexts.some(text => bodyText.includes(text));
        } catch (e) {}
        return { loadingVisible, anyVisibleRow, emptyVisible };
      }

      let combined = checkWindow(window);
      try {
        for (let i = 0; i < window.frames.length; i++) {
          const frameState = checkWindow(window.frames[i]);
          combined.loadingVisible = combined.loadingVisible || frameState.loadingVisible;
          combined.anyVisibleRow = combined.anyVisibleRow || frameState.anyVisibleRow;
          combined.emptyVisible = combined.emptyVisible || frameState.emptyVisible;
        }
      } catch (e) {}

      return combined;
    }).catch(() => ({ loadingVisible: false, anyVisibleRow: false, emptyVisible: false }));

    if (!state.loadingVisible && (state.anyVisibleRow || state.emptyVisible)) {
      return state;
    }

    await page.waitForTimeout(600).catch(() => {});
  }

  return { loadingVisible: false, anyVisibleRow: false, emptyVisible: false, timedOut: true };
}

// 搜索结果出来后，将每页条数切换到50条，确保所有结果在一页内显示
async function setPageSizeTo50(page, logger) {
  const frames = page.frames();
  for (const frame of frames) {
    try {
      // 找到"条/页"下拉框（值为10的可选select）
      const selectLocator = frame.locator('select').filter({ has: frame.locator('option[value="10"]') });
      const count = await selectLocator.count();
      if (count > 0) {
        await selectLocator.selectOption('50');
        logger.log(`[search] 已将每页条数切换为50，等待刷新...`);
        await waitForUi(page, 1500);
        return true;
      }
    } catch (e) {}
  }
  logger.log(`[search] 未找到条/页下拉框，跳过`);
  return false;
}

async function firstVisible(candidates) {
  for (const candidate of candidates) {
    const locator = candidate.first();
    const count = await locator.count().catch(() => 0);
    if (!count) {
      continue;
    }
    const visible = await locator.isVisible().catch(() => false);
    if (visible) {
      return locator;
    }
  }
  return null;
}

function getSearchContexts(page) {
  return [page, ...page.frames()];
}

function normalizeCodeLike(value) {
  return normalizeText(value).replace(/,/g, '');
}

async function detectAnyVisibleText(page, texts, options = {}) {
  const candidates = Array.isArray(texts) ? texts : [texts];
  const exact = options.exact === true;

  for (const text of candidates) {
    for (const ctx of getSearchContexts(page)) {
      const visible = await firstVisible([
        ctx.getByText(text, { exact }),
        ctx.getByRole('button', { name: text, exact }),
        ctx.getByRole('link', { name: text, exact }),
        ctx.locator(`text=${text}`)
      ]);
      if (visible) {
        return text;
      }
    }
  }

  return '';
}

async function detectAnyVisiblePlaceholder(page, placeholders) {
  const candidates = Array.isArray(placeholders) ? placeholders : [placeholders];

  for (const placeholder of candidates) {
    for (const ctx of getSearchContexts(page)) {
      const visible = await firstVisible([
        ctx.getByPlaceholder(placeholder, { exact: false }),
        ctx.locator(`input[placeholder*="${placeholder}"], textarea[placeholder*="${placeholder}"]`)
      ]);
      if (visible) {
        return placeholder;
      }
    }
  }

  return '';
}

async function isLikelyLoginPage(page) {
  // 网站是SPA，URL不会变，只靠页面内容判断
  const loginHintText = await detectAnyVisibleText(page, [
    '电子营业执照',
    '企业业务',
    '自然人业务',
    '代理业务',
    '其他登录',
    '数字证书登录'
  ], { exact: false });
  if (loginHintText) {
    return true;
  }

  const loginPlaceholder = await detectAnyVisiblePlaceholder(page, [
    '统一社会信用代码/纳税人识别号',
    '居民身份证号码/手机号码/用户名',
    '个人用户密码'
  ]);
  return Boolean(loginPlaceholder);
}

async function hasTargetPageTitle(page) {
  return Boolean(await detectAnyVisibleText(page, ['逐项配单'], { exact: false }));
}

async function hasTargetSearchArea(page) {
  const searchField = await findDeclarationSearchField(page, { allowGenericFallback: true });
  if (!searchField) {
    return false;
  }

  const searchButton = await detectAnyVisibleText(page, [config.taxRefund.searchButtonText, '搜索', '查询'], { exact: false });
  if (!searchButton) {
    return false;
  }

  const searchFieldLabel = normalizeText(config.taxRefund.searchFieldLabel || '');
  const pageText = normalizeText(await page.locator('body').innerText().catch(() => ''));
  return Boolean(searchFieldLabel && pageText.includes(searchFieldLabel));
}


function normalizeCompareValue(field, value) {
  if (field === 'declarationNo' || field === 'invoiceNo') {
    return normalizeCodeLike(value);
  }

  if (field === 'declarationItemNo' || field === 'invoiceLineNo' || field === 'dealQty' || field === 'taxAmount') {
    return normalizeNumberString(value);
  }

  return normalizeText(value);
}

function compareFieldValue(field, left, right) {
  return normalizeCompareValue(field, left) === normalizeCompareValue(field, right);
}

function areNumbersClose(left, right, epsilon = 0.01) {
  if (!Number.isFinite(left) || !Number.isFinite(right)) {
    return false;
  }
  return Math.abs(left - right) <= epsilon;
}

function areNumberStringsEqual(left, right) {
  const leftNumber = toNumber(left);
  const rightNumber = toNumber(right);
  if (Number.isFinite(leftNumber) && Number.isFinite(rightNumber)) {
    return areNumbersClose(leftNumber, rightNumber);
  }
  return normalizeNumberString(left) === normalizeNumberString(right);
}

function findHeaderIndex(headers, aliases) {
  const normalizedHeaders = headers.map(header => normalizeText(header));
  for (const alias of aliases) {
    const index = normalizedHeaders.findIndex(header => header === normalizeText(alias));
    if (index >= 0) {
      return index;
    }
  }
  return -1;
}

function getCellValue(row, index) {
  if (!row || !Array.isArray(row.cells) || index < 0 || index >= row.cells.length) {
    return '';
  }
  return row.cells[index] ?? '';
}

function buildFieldIndexMap(headers, keys) {
  const map = {};
  for (const key of keys) {
    map[key] = findHeaderIndex(headers, HEADER_ALIASES[key] || []);
  }
  return map;
}

function listMissingFieldHeaders(fieldIndexMap) {
  return Object.entries(fieldIndexMap)
    .filter(([, index]) => index < 0)
    .map(([key]) => key);
}

async function clickTextAcrossPage(page, texts, options = {}) {
  const candidates = Array.isArray(texts) ? texts : [texts];
  const exact = options.exact !== false;

  for (const text of candidates) {
    for (const ctx of getSearchContexts(page)) {
      const locators = [
        ctx.getByRole('button', { name: text, exact }),
        ctx.getByRole('link', { name: text, exact }),
        ctx.getByText(text, { exact }),
        ctx.locator(`text=${text}`)
      ];

      const visible = await firstVisible(locators);
      if (!visible) {
        continue;
      }

      await visible.scrollIntoViewIfNeeded().catch(() => {});
      await visible.click({ timeout: 3000 }).catch(async () => {
        await visible.click({ force: true, timeout: 2000 }).catch(() => {});
      });
      return { text, url: page.url() };
    }
  }

  throw new Error(`未找到可点击文本：${candidates.join(' / ')}`);
}

async function extractTablesFromFrame(frame) {
  try {
    return await frame.evaluate((rootConfigs) => {
      const normalize = (value) => String(value ?? '').replace(/\s+/g, ' ').trim();
      const textOf = (elements, limit = 30) => Array.from(elements)
        .map(el => normalize(el.textContent || el.innerText || ''))
        .filter(Boolean)
        .slice(0, limit);
      const cellValueOf = (cell) => {
        const text = normalize(cell.textContent || cell.innerText || '');
        if (text) {
          return text;
        }
        const formValues = Array.from(cell.querySelectorAll('input, textarea, select'))
          .map(el => normalize(el.value || el.getAttribute('value') || ''))
          .filter(Boolean);
        if (formValues.length) {
          return formValues.join(' ');
        }
        return '';
      };
      const isVisible = (el) => {
        if (!el || !el.isConnected) return false;
        const style = window.getComputedStyle(el);
        if (!style || style.display === 'none' || style.visibility === 'hidden' || Number(style.opacity) === 0) {
          return false;
        }
        const rect = el.getBoundingClientRect();
        return rect.width > 0 && rect.height > 0;
      };
      const collectHeaders = (root) => {
        const selectors = [
          'thead th',
          '[role="columnheader"]',
          '.ant-table-thead th',
          '.el-table__header th'
        ];
        for (const selector of selectors) {
          const values = textOf(root.querySelectorAll(selector), 100);
          if (values.length) {
            return values;
          }
        }
        const firstRow = root.querySelector('tbody tr, .ant-table-tbody tr, .el-table__body-wrapper tbody tr, [role="row"]');
        if (!firstRow) {
          return [];
        }
        return Array.from(firstRow.querySelectorAll('th, td, [role="gridcell"], [role="columnheader"]'))
          .map(cellValueOf)
          .filter(Boolean);
      };

      const seen = new Set();
      const tables = [];

      rootConfigs.forEach((config, configIndex) => {
        const roots = Array.from(document.querySelectorAll(config.rootSelector));
        roots.forEach((root, rootIndex) => {
          if (!isVisible(root) || seen.has(root)) {
            return;
          }
          seen.add(root);

          const headers = collectHeaders(root);
          const rowElements = Array.from(root.querySelectorAll(config.rowSelector)).filter(isVisible);
          const rows = rowElements.map((rowEl, rowIndex) => {
            const cells = Array.from(rowEl.querySelectorAll(config.cellSelector))
              .map(cellValueOf);
            const actionTexts = textOf(rowEl.querySelectorAll('button, a, [role="button"], .ant-btn'), 20);
            return {
              rowIndex,
              rowText: normalize(rowEl.textContent || rowEl.innerText || ''),
              cells,
              actionTexts
            };
          }).filter(row => row.cells.some(Boolean) || row.actionTexts.length > 0);

          if (!headers.length && !rows.length) {
            return;
          }

          tables.push({
            configIndex,
            rootSelector: config.rootSelector,
            rowSelector: config.rowSelector,
            cellSelector: config.cellSelector,
            rootIndex,
            headers,
            rowCount: rows.length,
            rows: rows.slice(0, 100)
          });
        });
      });

      return tables;
    }, TABLE_ROOT_CONFIGS);
  } catch (error) {
    return [{
      error: error.message,
      headers: [],
      rowCount: 0,
      rows: []
    }];
  }
}

async function summarizeFrame(frame) {
  try {
    const tables = await extractTablesFromFrame(frame);
    return await frame.evaluate((tableSnapshot) => {
      const textOf = (elements, limit = 30) => Array.from(elements)
        .map(el => (el.textContent || el.innerText || '').replace(/\s+/g, ' ').trim())
        .filter(Boolean)
        .slice(0, limit);

      const inputs = Array.from(document.querySelectorAll('input, textarea, select')).slice(0, 80).map(el => ({
        tag: el.tagName.toLowerCase(),
        type: el.getAttribute('type') || '',
        name: el.getAttribute('name') || '',
        placeholder: el.getAttribute('placeholder') || '',
        value: (el.value || '').slice(0, 80),
        ariaLabel: el.getAttribute('aria-label') || ''
      }));

      const buttons = textOf(document.querySelectorAll('button, [role="button"], .ant-btn, a'), 80);
      const labels = textOf(document.querySelectorAll('label, th, .ant-form-item-label, .el-form-item__label'), 80);
      const headers = textOf(document.querySelectorAll('th, thead td, .ant-table-thead th, [role="columnheader"]'), 80);
      const dialogs = textOf(document.querySelectorAll('[role="dialog"], .ant-modal, .el-dialog'), 40);
      const visibleText = (document.body?.innerText || '').replace(/\s+/g, ' ').trim().slice(0, 6000);

      return {
        title: document.title || '',
        buttons,
        labels,
        headers,
        dialogs,
        inputs,
        tables: tableSnapshot,
        visibleText
      };
    }, tables);
  } catch (error) {
    return {
      error: error.message
    };
  }
}

async function collectPageSummary(page, label) {
  const frames = [];
  const allFrames = page.frames();
  for (let i = 0; i < allFrames.length; i += 1) {
    const frame = allFrames[i];
    frames.push({
      frameIndex: i,
      name: frame.name(),
      url: frame.url(),
      summary: await summarizeFrame(frame)
    });
  }

  return {
    label,
    capturedAt: new Date().toISOString(),
    url: page.url(),
    title: await page.title().catch(() => ''),
    frameCount: frames.length,
    frames
  };
}

// 调试开关：true = 关闭截图和页面快照（稳定运行时使用）
const DISABLE_SCREENSHOTS = true;

async function saveInspectionArtifacts(page, label, logger) {
  const summary = await collectPageSummary(page, label);
  if (DISABLE_SCREENSHOTS) {
    return { screenshotPath: '', summaryPath: '', summary };
  }
  const stamp = nowStamp();
  const screenshotPath = path.join(config.paths.logDir, `${label}-${stamp}.png`);
  const summaryPath = path.join(config.paths.logDir, `${label}-${stamp}.json`);
  await page.screenshot({ path: screenshotPath, fullPage: true }).catch(() => {});
  fs.writeFileSync(summaryPath, JSON.stringify(summary, null, 2), 'utf8');
  logger.log(`[inspect] 已保存页面快照：${summaryPath}`);
  logger.log(`[inspect] 已保存截图：${screenshotPath}`);
  return { screenshotPath, summaryPath, summary };
}

function stripLeadingSelectionCell(headers, cells) {
  if (!Array.isArray(cells)) {
    return [];
  }

  const firstCell = normalizeText(cells[0]);
  const looksLikeSelectionCell = firstCell === '' || firstCell === 'on';
  if (Array.isArray(headers) && cells.length === headers.length + 1 && looksLikeSelectionCell) {
    return cells.slice(1);
  }

  return [...cells];
}

function headersEchoFirstRow(currentHeaders, firstRowCells) {
  const comparableFirstRowCells = stripLeadingSelectionCell(currentHeaders, firstRowCells);
  const normalizedHeaders = (currentHeaders || []).map(header => normalizeText(header)).filter(Boolean);
  const normalizedFirstRowCells = comparableFirstRowCells.map(cell => normalizeText(cell));
  const normalizedNonEmptyFirstRowCells = normalizedFirstRowCells.filter(Boolean);

  const exactPositionMatch = normalizedHeaders.length > 0
    && normalizedHeaders.every((header, headerIndex) => header === normalizedFirstRowCells[headerIndex]);

  const compactMatch = normalizedHeaders.length > 0
    && normalizedHeaders.length === normalizedNonEmptyFirstRowCells.length
    && normalizedHeaders.every((header, headerIndex) => header === normalizedNonEmptyFirstRowCells[headerIndex]);

  return exactPositionMatch || compactMatch;
}

function hasReusableHeaderSignature(headers) {
  const normalizedHeaders = (headers || []).map(header => normalizeText(header)).filter(Boolean);
  if (normalizedHeaders.length < 4) {
    return false;
  }

  return normalizedHeaders.includes('报关单号')
    || normalizedHeaders.includes('发票号码')
    || normalizedHeaders.includes('发票行号');
}

function looksLikeDataHeaderEcho(headers) {
  const normalizedHeaders = (headers || []).map(header => normalizeText(header));
  const nonEmptyHeaders = normalizedHeaders.filter(Boolean);
  if (!nonEmptyHeaders.length) {
    return false;
  }

  if (hasReusableHeaderSignature(nonEmptyHeaders)) {
    return false;
  }

  const firstHeader = nonEmptyHeaders[0];
  return firstHeader === 'on'
    || /^\d+$/.test(firstHeader)
    || /[A-Za-z0-9]{12,}/.test(firstHeader);
}

function enrichSplitTableHeaders(tableEntries) {
  const enriched = [];

  for (let index = 0; index < tableEntries.length; index += 1) {
    const current = tableEntries[index];
    const previous = enriched[enriched.length - 1];
    const currentHeaders = current.table.headers || [];
    const currentRows = current.table.rows || [];
    const firstRowCells = currentRows[0]?.cells || [];
    const comparableFirstRowCells = stripLeadingSelectionCell(previous?.table.headers || [], firstRowCells);
    const previousHeaders = previous?.table.headers || [];

    const shouldReusePreviousHeaders = Boolean(
      previous
      && previous.frameIndex === current.frameIndex
      && previous.table.rootSelector === current.table.rootSelector
      && (previous.table.rows || []).length === 0
      && hasReusableHeaderSignature(previousHeaders)
      && currentRows.length > 0
      && (
        headersEchoFirstRow(currentHeaders, firstRowCells)
        || looksLikeDataHeaderEcho(currentHeaders)
        || comparableFirstRowCells.length === previousHeaders.length
      )
    );

    if (shouldReusePreviousHeaders) {
      enriched.push({
        ...current,
        table: {
          ...current.table,
          headers: previousHeaders
        }
      });
      continue;
    }

    enriched.push(current);
  }

  return enriched;
}
function flattenTables(summary) {
  const tables = [];
  for (const frame of summary.frames || []) {
    for (const table of frame.summary?.tables || []) {
      tables.push({
        frameIndex: frame.frameIndex,
        frameName: frame.name,
        frameUrl: frame.url,
        table
      });
    }
  }
  return enrichSplitTableHeaders(tables);
}

function describeRowMatch(match) {
  const table = match.table.table;
  return `frame=${match.table.frameIndex} root=${table.rootSelector}[${table.rootIndex}] row=${match.row.rowIndex}`;
}

function buildAlignedCellSamples(headers, row) {
  const normalizedHeaders = (headers || []).map(header => normalizeText(header));
  const cells = stripLeadingSelectionCell(headers, row?.cells || []);

  if (!normalizedHeaders.length || cells.length >= normalizedHeaders.length) {
    return [...cells];
  }

  const optionalHeaders = new Set(['规格型号', '第二单位']);
  const aligned = [];
  let cellIndex = 0;

  for (const header of normalizedHeaders) {
    const remainingHeaders = normalizedHeaders.length - aligned.length;
    const remainingCells = cells.length - cellIndex;
    const mustPadOptional = optionalHeaders.has(header) && remainingCells < remainingHeaders;

    if (mustPadOptional) {
      aligned.push('');
      continue;
    }

    aligned.push(cells[cellIndex] ?? '');
    cellIndex += 1;
  }

  return aligned;
}
function findDeclarationRowMatches(summary, expectedRow) {
  const matches = [];
  for (const tableEntry of flattenTables(summary)) {
    const { headers, rows } = tableEntry.table;
    const fieldIndexMap = buildFieldIndexMap(headers, ['declarationNo', 'goodsName', 'declarationItemNo', 'dealUnit', 'dealQty']);
    if (listMissingFieldHeaders(fieldIndexMap).length > 0) {
      continue;
    }

    for (const row of rows || []) {
      const alignedCells = buildAlignedCellSamples(headers, row);
      const values = {
        declarationNo: alignedCells[fieldIndexMap.declarationNo],
        goodsName: alignedCells[fieldIndexMap.goodsName],
        declarationItemNo: alignedCells[fieldIndexMap.declarationItemNo],
        dealUnit: alignedCells[fieldIndexMap.dealUnit],
        dealQty: alignedCells[fieldIndexMap.dealQty]
      };
      const isMatch = compareFieldValue('declarationNo', values.declarationNo, expectedRow.declarationNo)
        && compareFieldValue('goodsName', values.goodsName, expectedRow.goodsName)
        && compareFieldValue('declarationItemNo', values.declarationItemNo, expectedRow.declarationItemNo)
        && compareFieldValue('dealUnit', values.dealUnit, expectedRow.dealUnit)
        && compareFieldValue('dealQty', values.dealQty, expectedRow.dealQty);

      if (isMatch) {
        matches.push({
          table: tableEntry,
          row,
          fieldIndexMap,
          values
        });
      }
    }
  }
  return matches;
}

function normalizeSelectionTableHeaders(headers) {
  const normalizedHeaders = [...(headers || [])];
  if (!normalizedHeaders.length) {
    return normalizedHeaders;
  }

  const firstHeader = normalizeText(normalizedHeaders[0]);
  if (firstHeader === 'on') {
    normalizedHeaders[0] = '';
  }

  return normalizedHeaders;
}

function isLikelyInvoiceSelectionTable(headers) {
  const normalizedHeaders = normalizeSelectionTableHeaders(headers);
  return findHeaderIndex(normalizedHeaders, HEADER_ALIASES.invoiceNo) >= 0
    && findHeaderIndex(normalizedHeaders, HEADER_ALIASES.invoiceLineNo) >= 0
    && findHeaderIndex(normalizedHeaders, HEADER_ALIASES.goodsName) >= 0
    && findHeaderIndex(normalizedHeaders, ['开票日期']) >= 0
    && findHeaderIndex(normalizedHeaders, ['供货方税号']) >= 0
    && findHeaderIndex(normalizedHeaders, ['单位']) >= 0
    && findHeaderIndex(normalizedHeaders, ['数量']) >= 0
    && findHeaderIndex(normalizedHeaders, HEADER_ALIASES.taxAmount) >= 0
    && findHeaderIndex(normalizedHeaders, ['操作']) < 0;
}

function findInvoiceSelectionCandidates(summary) {
  return flattenTables(summary).map(tableEntry => {
    const headers = normalizeSelectionTableHeaders(tableEntry.table.headers || []);
    return {
      ...tableEntry,
      table: {
        ...tableEntry.table,
        headers
      }
    };
  }).filter(tableEntry => {
    const headers = tableEntry.table.headers || [];
    return findHeaderIndex(headers, HEADER_ALIASES.invoiceNo) >= 0
      && findHeaderIndex(headers, HEADER_ALIASES.invoiceLineNo) >= 0
      && findHeaderIndex(headers, HEADER_ALIASES.goodsName) >= 0
      && findHeaderIndex(headers, ['开票日期']) >= 0
      && findHeaderIndex(headers, ['供货方税号']) >= 0
      && findHeaderIndex(headers, ['单位']) >= 0
      && findHeaderIndex(headers, ['数量']) >= 0
      && findHeaderIndex(headers, HEADER_ALIASES.taxAmount) >= 0;
  });
}

function getInvoiceSelectionIndexMap(headers) {
  return {
    invoiceNo: findHeaderIndex(headers, HEADER_ALIASES.invoiceNo),
    invoiceLineNo: findHeaderIndex(headers, HEADER_ALIASES.invoiceLineNo),
    goodsName: findHeaderIndex(headers, HEADER_ALIASES.goodsName),
    unit: findHeaderIndex(headers, ['单位']),
    qty: findHeaderIndex(headers, ['数量']),
    taxAmount: findHeaderIndex(headers, HEADER_ALIASES.taxAmount),
    issueDate: findHeaderIndex(headers, ['开票日期']),
    vendorTaxNo: findHeaderIndex(headers, ['供货方税号'])
  };
}

function hasRequiredInvoiceSelectionIndexes(indexMap) {
  return Object.values(indexMap).every(index => index >= 0);
}

function getAlignedCellValue(row, headers, index) {
  const alignedCells = buildAlignedCellSamples(headers, row);
  if (index < 0 || index >= alignedCells.length) {
    return '';
  }
  return alignedCells[index] ?? '';
}

function findInvoiceSelectionMatches(summary, expectedRow) {
  const expectedLineNos = Array.from(new Set((expectedRow.selectedInvoiceLineNos || []).map(value => normalizeNumberString(value)).filter(Boolean)));
  const matches = [];

  for (const tableEntry of findInvoiceSelectionCandidates(summary)) {
    const { headers, rows } = tableEntry.table;
    const indexMap = getInvoiceSelectionIndexMap(headers);
    if (!hasRequiredInvoiceSelectionIndexes(indexMap)) {
      continue;
    }

    for (const row of rows || []) {
      const sameInvoice = compareFieldValue('invoiceNo', getAlignedCellValue(row, headers, indexMap.invoiceNo), expectedRow.invoiceNo);
      const currentLineNo = normalizeNumberString(getAlignedCellValue(row, headers, indexMap.invoiceLineNo));
      if (sameInvoice && expectedLineNos.includes(currentLineNo)) {
        matches.push({
          table: tableEntry,
          row,
          currentLineNo,
          indexMap
        });
      }
    }
  }

  return matches;
}

function findInvoiceRowMatches(summary, expectedRow) {
  return findInvoiceSelectionMatches(summary, expectedRow);
}

function findSelectedDetailCandidates(summary) {
  return flattenTables(summary).filter(tableEntry => {
    const headers = tableEntry.table.headers || [];
    return findHeaderIndex(headers, HEADER_ALIASES.invoiceNo) >= 0
      && findHeaderIndex(headers, HEADER_ALIASES.invoiceLineNo) >= 0
      && findHeaderIndex(headers, HEADER_ALIASES.goodsName) >= 0
      && findHeaderIndex(headers, HEADER_ALIASES.dealQty) >= 0
      && findHeaderIndex(headers, HEADER_ALIASES.taxAmount) >= 0
      && findHeaderIndex(headers, HEADER_ALIASES.dealUnit) >= 0;
  });
}

function validateSelectedInvoiceRows(detailTableEntry, expectedRow) {
  const headers = detailTableEntry.table.headers || [];
  const fieldIndexMap = buildFieldIndexMap(headers, ['invoiceNo', 'invoiceLineNo', 'goodsName', 'dealQty', 'dealUnit', 'taxAmount']);
  const missing = listMissingFieldHeaders(fieldIndexMap);
  if (missing.length > 0) {
    throw new Error(`保存前核对缺少关键列：${missing.join('、')}`);
  }

  const expectedLineSet = new Set((expectedRow.selectedInvoiceLineNos || []).map(value => normalizeNumberString(value)).filter(Boolean));
  const matchedRows = (detailTableEntry.table.rows || []).filter(row => {
    const invoiceNo = getAlignedCellValue(row, headers, fieldIndexMap.invoiceNo);
    const invoiceLineNo = getAlignedCellValue(row, headers, fieldIndexMap.invoiceLineNo);
    return compareFieldValue('invoiceNo', invoiceNo, expectedRow.invoiceNo)
      && expectedLineSet.has(normalizeNumberString(invoiceLineNo));
  });

  if (!matchedRows.length) {
    throw new Error('保存前核对未找到已选发票明细行');
  }

  const issues = [];
  const selectedLineSet = new Set();
  let totalQty = 0;
  let totalTax = 0;

  for (const row of matchedRows) {
    const goodsName = getAlignedCellValue(row, headers, fieldIndexMap.goodsName);
    const unit = getAlignedCellValue(row, headers, fieldIndexMap.dealUnit);
    const lineNo = getAlignedCellValue(row, headers, fieldIndexMap.invoiceLineNo);
    const qty = getAlignedCellValue(row, headers, fieldIndexMap.dealQty);
    const taxAmount = getAlignedCellValue(row, headers, fieldIndexMap.taxAmount);

    selectedLineSet.add(normalizeNumberString(lineNo));

    if (!compareFieldValue('goodsName', goodsName, expectedRow.goodsName)) {
      issues.push(`商品名称不一致：网页=${goodsName} Excel=${expectedRow.goodsName}`);
    }
    if (!compareFieldValue('dealUnit', unit, expectedRow.dealUnit)) {
      issues.push(`单位不一致：网页=${unit} Excel=${expectedRow.dealUnit}`);
    }

    totalQty += toNumber(qty) || 0;
    totalTax += toNumber(taxAmount) || 0;
  }

  const selectedLineNos = Array.from(selectedLineSet).filter(Boolean).sort();
  const expectedLineNos = Array.from(expectedLineSet).filter(Boolean).sort();
  if (selectedLineNos.join('|') !== expectedLineNos.join('|')) {
    issues.push(`发票行号不一致：网页=${selectedLineNos.join(',')} Excel=${expectedLineNos.join(',')}`);
  }

  const totalQtyText = String(totalQty);
  const totalTaxText = String(totalTax);
  const qtyMatches = areNumberStringsEqual(totalQtyText, expectedRow.dealQty);
  const taxMatches = areNumberStringsEqual(totalTaxText, expectedRow.taxAmount);

  const nonAdjustableIssues = issues.filter(issue => !issue.startsWith('数量') && !issue.startsWith('计税金额'));

  return {
    matchedRows,
    fieldIndexMap,
    headers,
    issues,
    nonAdjustableIssues,
    qtyMatches,
    taxMatches,
    totals: {
      qty: totalQty,
      tax: totalTax
    },
    expected: {
      qty: toNumber(expectedRow.dealQty),
      tax: toNumber(expectedRow.taxAmount)
    },
    canAutoAdjust: nonAdjustableIssues.length === 0 && (!qtyMatches || !taxMatches)
  };
}

function getFrameByIndex(page, frameIndex) {
  return page.frames()[frameIndex] || page.mainFrame();
}

function getTableLocator(page, tableEntry) {
  const frame = getFrameByIndex(page, tableEntry.frameIndex);
  return frame.locator(tableEntry.table.rootSelector).nth(tableEntry.table.rootIndex);
}

function getRowLocator(page, match) {
  const tableLocator = getTableLocator(page, match.table);
  return tableLocator.locator(match.table.table.rowSelector).nth(match.row.rowIndex);
}

function getNearbyTableIndexes(rootIndex, maxIndex, radius = 4) {
  const indexes = [rootIndex];
  for (let step = 1; step <= radius; step += 1) {
    indexes.push(rootIndex - step);
    indexes.push(rootIndex + step);
  }
  return indexes.filter(index => index >= 0 && index < maxIndex);
}

async function findVisibleCheckboxInRow(rowLocator) {
  return firstVisible([
    rowLocator.locator('input[type="checkbox"]'),
    rowLocator.locator('.ant-checkbox-input'),
    rowLocator.locator('[role="checkbox"]'),
    rowLocator.locator('.layui-form-checkbox'),
    rowLocator.locator('.laytable-cell-checkbox'),
    rowLocator.locator('td:first-child label, td:first-child .layui-unselect')
  ]);
}

async function findInvoiceCheckboxTarget(page, match) {
  const primaryRowLocator = getRowLocator(page, match);
  const primaryCheckbox = await findVisibleCheckboxInRow(primaryRowLocator);
  if (primaryCheckbox) {
    return {
      rowLocator: primaryRowLocator,
      checkbox: primaryCheckbox,
      tableRootIndex: match.table.table.rootIndex
    };
  }

  const frame = getFrameByIndex(page, match.table.frameIndex);
  const tableLocator = frame.locator(match.table.table.rootSelector);
  const tableCount = await tableLocator.count().catch(() => 0);
  const nearbyIndexes = getNearbyTableIndexes(match.table.table.rootIndex, tableCount);

  for (const tableRootIndex of nearbyIndexes) {
    if (tableRootIndex === match.table.table.rootIndex) {
      continue;
    }

    const rowLocator = tableLocator.nth(tableRootIndex).locator(match.table.table.rowSelector).nth(match.row.rowIndex);
    const checkbox = await findVisibleCheckboxInRow(rowLocator);
    if (checkbox) {
      return {
        rowLocator,
        checkbox,
        tableRootIndex
      };
    }
  }

  return {
    rowLocator: primaryRowLocator,
    checkbox: null,
    tableRootIndex: match.table.table.rootIndex
  };
}

async function clickRowAction(page, match, actionTexts, logger) {
  const rowLocator = getRowLocator(page, match);
  await rowLocator.scrollIntoViewIfNeeded().catch(() => {});

  const candidates = [];
  for (const text of actionTexts) {
    candidates.push(rowLocator.getByRole('button', { name: text, exact: false }));
    candidates.push(rowLocator.getByRole('link', { name: text, exact: false }));
    candidates.push(rowLocator.getByText(text, { exact: false }));
    candidates.push(rowLocator.locator(`text=${text}`));
  }

  const visible = await firstVisible(candidates);
  if (!visible) {
    throw new Error(`匹配行已找到，但未定位到行内动作按钮：${actionTexts.join(' / ')}`);
  }

  await visible.click({ timeout: 3000 }).catch(async () => {
    await visible.click({ force: true, timeout: 2000 }).catch(() => {});
  });
  logger.log(`[row ${match.row.rowIndex}] 已点击行内动作：${actionTexts.join(' / ')}`);
}

async function findDeclarationSearchField(page, options = {}) {
  const allowGenericFallback = options.allowGenericFallback !== false;
  const labels = [config.taxRefund.searchFieldLabel, '报关单号'];
  for (const ctx of getSearchContexts(page)) {
    for (const label of labels) {
      const locator = await firstVisible([
        ctx.getByLabel(label, { exact: false }),
        ctx.getByPlaceholder(label, { exact: false }),
        ctx.getByRole('textbox', { name: label, exact: false }),
        ctx.locator(`input[placeholder*="${label}"], input[name*="${label}"], input[aria-label*="${label}"]`).first()
      ]);
      if (locator) {
        return locator;
      }
    }

    if (!allowGenericFallback) {
      continue;
    }

    const textInputs = ctx.locator('input[type="text"], input:not([type]), .ant-input');
    const visible = await firstVisible([textInputs]);
    if (visible) {
      return visible;
    }
  }
  return null;
}

async function findSearchField(page) {
  return findDeclarationSearchField(page, { allowGenericFallback: true });
}


async function searchDeclarationRow(page, row, logger) {
  let foundCorrectResults = false;
  for (let attempt = 0; attempt < 5; attempt++) {
    // Re-obtain search field each attempt to avoid stale locator
    const searchField = await findSearchField(page);
    if (!searchField) {
      if (attempt === 4) {
        throw new Error('未找到"报关单号"搜索输入框');
      }
      logger.log(`[row ${row.rowNumber}] 重试 ${attempt + 1}：未找到搜索框，等待后重试`);
      await waitForUi(page, 2000);
      continue;
    }

    await searchField.scrollIntoViewIfNeeded().catch(() => {});
    await searchField.click({ timeout: 2000 }).catch(() => {});
    await searchField.fill('');
    await searchField.fill(row.declarationNo);
    logger.log(`[row ${row.rowNumber}] 已输入报关单号：${row.declarationNo}${attempt > 0 ? `（重试 ${attempt + 1}）` : ''}`);

    logger.log(`[row ${row.rowNumber}] 点击搜索按钮...`);
    await clickTextAcrossPage(page, [config.taxRefund.searchButtonText, '搜索', '查询'], { exact: false }).catch(async () => {
      await searchField.press('Enter').catch(() => {});
    });
    logger.log(`[row ${row.rowNumber}] 搜索按钮已点击，等待结果...`);

    await waitForUi(page, 2500);
    logger.log(`[row ${row.rowNumber}] 等待搜索结果加载...`);
    const searchState = await waitForSearchResultsReady(page, 15000);
    logger.log(`[row ${row.rowNumber}] 搜索结果状态：loading=${searchState.loadingVisible ? 'Y' : 'N'} rows=${searchState.anyVisibleRow ? 'Y' : 'N'} empty=${searchState.emptyVisible ? 'Y' : 'N'}${searchState.timedOut ? ' timeout=Y' : ''}`);

    // Read text from ALL frames using Playwright frame API
    logger.log(`[row ${row.rowNumber}] 读取页面文本...`);
    const allFrameTexts = [];
    const frames = page.frames();
    logger.log(`[row ${row.rowNumber}] 共 ${frames.length} 个frames`);
    for (const frame of frames) {
      try {
        // Use evaluate instead of locator API for reliable frame text access
        const frameText = await frame.evaluate(() => document.body ? document.body.innerText : '').catch(() => '');
        if (frameText) allFrameTexts.push(frameText);
      } catch (e) {
        // Frame not accessible, skip
      }
    }
    const allFrameText = allFrameTexts.join(' | ');
    logger.log(`[row ${row.rowNumber}] 页面文本读取完成`);

    const declNoInLocator = await page.locator(`text=${row.declarationNo}`).first().isVisible({ timeout: 3000 }).catch(() => false);
    const targetInText = allFrameText.includes(row.declarationNo);
    const targetInLocator = declNoInLocator;

    if (targetInText || targetInLocator) {
      foundCorrectResults = true;
      logger.log(`[row ${row.rowNumber}] 搜索验证通过（text=${targetInText} locator=${targetInLocator}${attempt > 0 ? `（第${attempt + 1}次重试成功）` : ''}）`);
      break;
    }

    // Log what numbers ARE visible for debugging
    const visibleNumbers = (allFrameText.match(/\d{15,}/g) || []);
    const uniqueNumbers = [...new Set(visibleNumbers)];
    logger.log(`[row ${row.rowNumber}] 重试 ${attempt + 1}：未找到 ${row.declarationNo}，页面上可见号码=${uniqueNumbers.join(',') || '无'}`);
  }

  if (!foundCorrectResults) {
    throw new Error(`搜索报关单号 ${row.declarationNo} 后结果未包含该报关单号（服务器响应不稳定）`);
  }
}

function getInvoiceSearchKeyword(row) {
  const invoiceNo = normalizeCodeLike(row?.invoiceNo || '');
  if (!invoiceNo) {
    return '';
  }

  if (invoiceNo.length >= 8) {
    return invoiceNo.slice(-8);
  }

  if (invoiceNo.length >= 4) {
    return invoiceNo.slice(-4);
  }

  return invoiceNo;
}

async function findInvoiceDialogSearchField(page) {
  const labels = ['发票号码8位或后4位', '发票号码'];

  for (const ctx of getSearchContexts(page)) {
    for (const label of labels) {
      const locator = await firstVisible([
        ctx.getByPlaceholder(label, { exact: false }),
        ctx.getByLabel(label, { exact: false }),
        ctx.getByRole('textbox', { name: label, exact: false }),
        ctx.locator(`input[placeholder*="${label}"], input[name*="fphm"], input[aria-label*="${label}"]`).first()
      ]);
      if (locator) {
        return locator;
      }
    }
  }

  return null;
}

async function clickInvoiceDialogSearchTrigger(searchField) {
  return searchField.evaluate(input => {
    const normalize = value => String(value ?? '').replace(/\s+/g, ' ').trim();
    const isVisible = el => Boolean(el && (el.offsetWidth || el.offsetHeight || el.getClientRects().length));
    const textOf = el => normalize(el.innerText || el.textContent || el.getAttribute('aria-label') || el.getAttribute('title') || '');
    const classNameOf = el => normalize(el.className || '');
    const isSearchLike = el => {
      const text = textOf(el);
      const className = classNameOf(el);
      return text.includes('查询')
        || text.includes('搜索')
        || /search|layui-icon-search|icon-search|sousuo/i.test(className);
    };

    let current = input;
    for (let depth = 0; current && depth < 6; depth += 1, current = current.parentElement) {
      const candidates = Array.from(current.querySelectorAll('button, a, [role="button"], .layui-btn, .el-button, .ant-btn, .layui-icon, .anticon, .el-icon-search, [class*="search"]'));
      const target = candidates.find(el => el !== input && isVisible(el) && isSearchLike(el));
      if (target) {
        target.click();
        return textOf(target) || classNameOf(target) || 'search-trigger';
      }
    }

    return '';
  }).catch(() => '');
}

async function searchInvoiceRowsInDialog(page, row, logger) {
  const searchField = await findInvoiceDialogSearchField(page);
  if (!searchField) {
    throw new Error('未找到发票弹窗中的“发票号码8位或后4位”输入框');
  }

  const keyword = getInvoiceSearchKeyword(row);
  if (!keyword) {
    throw new Error('当前 Excel 行缺少可用于检索的发票号码');
  }

  await searchField.scrollIntoViewIfNeeded().catch(() => {});
  await searchField.click({ timeout: 2000 }).catch(() => {});
  await searchField.fill('');
  await searchField.fill(keyword);

  const clickedLabel = await clickInvoiceDialogSearchTrigger(searchField);
  if (!clickedLabel) {
    await searchField.press('Enter').catch(() => {});
  }

  logger.log(`[row ${row.rowNumber}] 发票弹窗已按发票号检索：原始=${row.invoiceNo} 检索值=${keyword}${clickedLabel ? ` 按钮=${clickedLabel}` : ' 触发=Enter'}`);
  await waitForUi(page, 2500);
}

async function ensureManualPrepPage(connection, logger) {
  if (connection.mode === 'connected') {
    const currentUrl = connection.page.url() || 'about:blank';
    logger.log(`[open] 复用已打开的 Chrome；当前页：${currentUrl}`);

    // 不关闭任何页面！用户要求不关浏览器
    // 但需要找到有左侧菜单的页面（gt4-area 或 loginb）作为导航起始页
    logger.log(`[open] 查找包含左侧菜单的页面...`);
    let targetPage = connection.page;
    const allPages = connection.context.pages();
    for (const p of allPages) {
      if (p.isClosed()) continue;
      const url = p.url() || '';
      // 跳过明显的登录扫码页
      if (url.includes('tpass') || url.includes('/#/login')) {
        continue;
      }
      // 检查页面frames是否有左侧菜单内容
      const frames = p.frames();
      let hasLeftMenu = false;
      for (const f of frames) {
        try {
          const text = await f.evaluate(() => document.body ? document.body.innerText : '').catch(() => '');
          if (text.includes('特色办税') && text.includes('地方特色')) {
            hasLeftMenu = true;
            break;
          }
        } catch (e) {}
      }
      if (hasLeftMenu) {
        targetPage = p;
        logger.log(`[open] 找到有左侧菜单的页面：${url}`);
        break;
      }
    }

    // 检查是否已登录（遍历所有页面）
    // 已登录的标志：左侧菜单（特色办税）或逐项配单页面（报关单号）
    let isLoggedIn = false;
    for (const p of connection.context.pages()) {
      if (p.isClosed()) continue;
      for (const f of p.frames()) {
        try {
          const text = await f.evaluate(() => document.body ? document.body.innerText : '').catch(() => '');
          // 左侧菜单页面
          if ((text.includes('特色办税') || text.includes('地方特色'))) {
            isLoggedIn = true;
            targetPage = p;
            logger.log(`[open] 检测到已登录（左侧菜单）：${p.url()}`);
            await closeOtherPages(targetPage, logger);
            break;
          }
          // 逐项配单页面（有搜索框）
          if (text.includes('逐项配单') && text.includes('报关单号')) {
            isLoggedIn = true;
            targetPage = p;
            logger.log(`[open] 检测到已登录（逐项配单页）：${p.url()}`);
            await closeOtherPages(targetPage, logger);
            break;
          }
        } catch (e) {}
      }
      if (isLoggedIn) break;
    }

    if (!isLoggedIn) {
      logger.log(`[open] 检测到未登录，等待扫码登录...`);
      // 等待登录完成（遍历所有页面）
      for (let i = 0; i < 60; i++) {
        await new Promise(r => setTimeout(r, 1000));
        for (const p of connection.context.pages()) {
          if (p.isClosed()) continue;
          for (const f of p.frames()) {
            try {
              const text = await f.evaluate(() => document.body ? document.body.innerText : '').catch(() => '');
              if ((text.includes('特色办税') || text.includes('地方特色'))) {
                isLoggedIn = true;
                targetPage = p;
                logger.log(`[open] 登录完成（左侧菜单）`);
                await closeOtherPages(targetPage, logger);
                break;
              }
              if (text.includes('逐项配单') && text.includes('报关单号')) {
                isLoggedIn = true;
                targetPage = p;
                logger.log(`[open] 登录完成（逐项配单页）`);
                await closeOtherPages(targetPage, logger);
                break;
              }
            } catch (e) {}
          }
          if (isLoggedIn) break;
        }
        if (isLoggedIn) break;
      }
    }

    connection.page = targetPage;
    return targetPage;
  }

  await connection.page.goto(config.urls.home, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await waitForUi(connection.page, 1500);
  logger.log(`[open] 已打开首页：${connection.page.url()}`);
  await saveInspectionArtifacts(connection.page, 'home-before-manual', logger);
  return connection.page;
}

// 导航到逐项配单页面的完整流程
async function navigateToTargetPage(page, logger) {
  logger.log(`[nav] 开始导航到逐项配单页面`);

  // 第一阶段：在当前页面点击"地方特色"，然后在gt4-area页面点击 特色办税 → 单一窗口
  page = await navPhase1_Click地方特色ThenNavigate(page, logger);

  // 第二阶段：在 cktsfw/main 退税桌面点击 在线申报
  page = await navPhase2_退税桌面_在线申报(page, logger);

  // 第三阶段：在出口退税在线申报页面点击 免退税申报 → 逐项配单
  page = await navPhase3_在线申报(page, logger);

  // 第四阶段：等待逐项配单表单加载完成
  await waitFor逐项配单Form(page, logger);

  return page;
}

// 第一阶段：点击"地方特色"打开gt4-area，然后在gt4-area点击 特色办税 → 单一窗口
async function navPhase1_Click地方特色ThenNavigate(page, logger) {
  logger.log(`[nav-1] 查找地方特色菜单位置...`);

  // 找到包含地方特色的frame
  let navFrame = null;
  for (const f of page.frames()) {
    try {
      const text = await f.locator('body').innerText({ timeout: 2000 }).catch(() => '');
      if (text.includes('地方特色') && text.includes('上海图吉')) {
        navFrame = f;
        break;
      }
    } catch (e) {}
  }

  if (!navFrame) {
    logger.log(`[nav-1] 未找到包含地方特色的frame，等待...`);
    for (let i = 0; i < 30; i++) {
      await page.waitForTimeout(1000);
      for (const f of page.frames()) {
        try {
          const text = await f.locator('body').innerText({ timeout: 2000 }).catch(() => '');
          if (text.includes('地方特色') && text.includes('上海图吉')) {
            navFrame = f;
            break;
          }
        } catch (e) {}
      }
      if (navFrame) break;
    }
  }

  if (!navFrame) {
    logger.log(`[nav-1] 无法找到导航frame`);
    return page;
  }

  // 1. 点击"地方特色" (div.menuItem)
  logger.log(`[nav-1] 点击 地方特色`);
  try {
    await navFrame.locator('div.menuItem', { hasText: '地方特色' }).click({ timeout: 5000 });
  } catch (e) {
    logger.log(`[nav-1] 点击地方特色失败: ${e.message}`);
  }

  // 2. 等待gt4-area新页面出现
  logger.log(`[nav-1] 等待gt4-area页面出现...`);
  let gt4Page = null;
  for (let i = 0; i < 30; i++) {
    await page.waitForTimeout(1000);
    const allPages = page.context().pages();
    gt4Page = allPages.find(p => p.url().includes('gt4-area'));
    if (gt4Page) break;
  }

  if (!gt4Page) {
    logger.log(`[nav-1] 未找到gt4-area页面`);
    return page;
  }
  logger.log(`[nav-1] 找到gt4-area页面`);

  // 在gt4-area页面操作
  const frames = gt4Page.frames();
  let gt4Frame = null;
  for (const f of frames) {
    try {
      const text = await f.locator('body').innerText({ timeout: 2000 }).catch(() => '');
      if (text.includes('特色办税') && text.includes('地方特色')) {
        gt4Frame = f;
        break;
      }
    } catch (e) {}
  }

  if (!gt4Frame) {
    logger.log(`[nav-1] 未找到gt4-area的菜单frame`);
    return gt4Page;
  }

  // 等待特色办税出现
  logger.log(`[nav-1] 等待特色办税菜单...`);
  for (let i = 0; i < 20; i++) {
    try {
      const text = await gt4Frame.locator('body').innerText({ timeout: 1000 }).catch(() => '');
      if (text.includes('特色办税')) {
        break;
      }
    } catch (e) {}
    await gt4Page.waitForTimeout(500);
  }

  // 3. 点击"特色办税" (span.leftSub)
  logger.log(`[nav-1] 点击 特色办税`);
  try {
    await gt4Frame.locator('span.leftSub', { hasText: '特色办税' }).click({ timeout: 5000 });
  } catch (e) {
    logger.log(`[nav-1] 点击特色办税失败`);
  }
  await gt4Page.waitForTimeout(2000);

  // 4. 等待单一窗口出现
  logger.log(`[nav-1] 等待单一窗口菜单元件...`);
  for (let i = 0; i < 20; i++) {
    try {
      const text = await gt4Frame.locator('body').innerText({ timeout: 1000 }).catch(() => '');
      if (text.includes('单一窗口出口退（免）税办理')) {
        break;
      }
    } catch (e) {}
    await gt4Page.waitForTimeout(500);
  }

  // 5. 点击"单一窗口出口退（免）税办理" (div.mxth)
  logger.log(`[nav-1] 点击 单一窗口出口退（免）税办理`);
  try {
    await gt4Frame.locator('div.mxth', { hasText: '单一窗口出口退（免）税办理' }).click({ timeout: 5000 });
  } catch (e) {
    logger.log(`[nav-1] 点击单一窗口失败`);
  }
  await gt4Page.waitForTimeout(3000);

  // 6. 等待cktsfw页面出现
  logger.log(`[nav-1] 等待退税桌面(cktsfw)页面出现...`);
  let cktsfwPage = null;
  for (let i = 0; i < 30; i++) {
    await page.waitForTimeout(1000);
    const allPages = page.context().pages();
    cktsfwPage = allPages.find(p => p.url().includes('cktsfw'));
    if (cktsfwPage) break;
  }

  if (cktsfwPage) {
    logger.log(`[nav-1] 找到退税桌面页面`);
    // 关闭其他页面，只留cktsfw
    await closeOtherPages(cktsfwPage, logger);
    return cktsfwPage;
  }

  logger.log(`[nav-1] 未找到退税桌面页面`);
  return gt4Page;
}

// 第二阶段：在退税桌面点击 在线申报
async function navPhase2_退税桌面_在线申报(page, logger) {
  logger.log(`[nav-2] 在退税桌面点击 在线申报...`);

  // 等待退税桌面页面加载
  let mainFrame = null;
  for (let i = 0; i < 30; i++) {
    const frames = page.frames();
    for (const f of frames) {
      try {
        const text = await f.locator('body').innerText({ timeout: 2000 }).catch(() => '');
        if (text.includes('退税桌面') && text.includes('常用快捷')) {
          mainFrame = f;
          break;
        }
      } catch (e) {}
    }
    if (mainFrame) break;
    await page.waitForTimeout(1000);
  }

  if (!mainFrame) {
    logger.log(`[nav-2] 未找到退税桌面主菜单frame`);
    return page;
  }

  // 等待在线申报出现
  logger.log(`[nav-2] 等待在线申报选项...`);
  for (let i = 0; i < 20; i++) {
    try {
      const text = await mainFrame.locator('body').innerText({ timeout: 1000 }).catch(() => '');
      if (text.includes('在线申报')) {
        break;
      }
    } catch (e) {}
    await page.waitForTimeout(500);
  }

  // 点击"在线申报" (在常用快捷区域)
  logger.log(`[nav-2] 点击 在线申报`);
  try {
    await mainFrame.locator('text=在线申报').click({ timeout: 5000 });
  } catch (e) {
    logger.log(`[nav-2] 点击在线申报失败`);
  }
  await page.waitForTimeout(3000);

  // 等待出口退税在线申报页面出现
  logger.log(`[nav-2] 等待出口退税在线申报页面...`);
  let jssoPage = null;
  for (let i = 0; i < 30; i++) {
    await page.waitForTimeout(1000);
    const allPages = page.context().pages();
    jssoPage = allPages.find(p => p.url().includes('jsso'));
    if (jssoPage) break;
  }

  if (jssoPage) {
    logger.log(`[nav-2] 找到出口退税在线申报页面`);
    await closeOtherPages(jssoPage, logger);
    return jssoPage;
  }

  logger.log(`[nav-2] 未找到出口退税在线申报页面`);
  return page;
}

async function navPhase3_在线申报(page, logger) {
  logger.log(`[nav-3] 等待出口退税在线申报菜单加载...`);

  // 等待包含菜单的frame加载
  let menuFrame = null;
  for (let attempt = 0; attempt < 10; attempt++) {
    const frames = page.frames();
    for (const f of frames) {
      try {
        const text = await f.locator('body').innerText({ timeout: 2000 }).catch(() => '');
        if (text.includes('免退税申报') && text.includes('退税申报') && text.includes('出口退税在线申报')) {
          menuFrame = f;
          logger.log(`[nav-3] 找到出口退税在线申报菜单 (frame: ${f.name()})`);
          break;
        }
      } catch (e) {}
    }
    if (menuFrame) break;
    await page.waitForTimeout(1000);
  }

  if (!menuFrame) {
    logger.log(`[nav-3] 未找到出口退税在线申报菜单，持续等待...`);
    let lastLogTime = Date.now();
    while (!menuFrame) {
      await page.waitForTimeout(1000);
      const frames = page.frames();
      for (const f of frames) {
        try {
          const text = await f.locator('body').innerText({ timeout: 500 }).catch(() => '');
          if (text.includes('免退税申报') && text.includes('退税申报')) {
            menuFrame = f;
            logger.log(`[nav-3] 找到出口退税在线申报菜单`);
            break;
          }
        } catch (e) {}
      }
      if (!menuFrame && Date.now() - lastLogTime > 15000) {
        logger.log(`[nav-3] 等待菜单出现...（当前frames数: ${frames.length}）`);
        lastLogTime = Date.now();
      }
    }
  }

  // 等待免退税申报可见
  logger.log(`[nav-3] 等待免退税申报菜单可见...`);
  await menuFrame.locator('text=免退税申报').waitFor({ state: 'visible', timeout: 10000 }).catch(() => {});

  // 6. 点击 免退税申报 (左侧菜单)
  logger.log(`[nav-3] 点击 免退税申报`);
  try {
    await menuFrame.locator('text=免退税申报').click({ timeout: 5000 });
    logger.log(`[nav-3] 点击成功`);
  } catch (e) {
    logger.log(`[nav-3] locator点击免退税申报失败，尝试JS点击: ${e.message}`);
    try {
      await menuFrame.evaluate(() => {
        const els = Array.from(document.querySelectorAll('span, div, a'));
        const el = els.find(el => el.textContent.trim() === '免退税申报');
        if (el) el.click();
      });
    } catch (e2) {
      logger.log(`[nav-3] JS点击也失败: ${e2.message}`);
    }
  }
  await page.waitForTimeout(2000);

  // 等待逐项配单出现（展开后的子菜单）
  logger.log(`[nav-3] 等待逐项配单出现...`);
  let 逐项配单找到 = false;
  for (let i = 0; i < 20; i++) {
    try {
      const text = await menuFrame.locator('body').innerText({ timeout: 1000 }).catch(() => '');
      if (text.includes('逐项配单')) {
        逐项配单找到 = true;
        logger.log(`[nav-3] 逐项配单已出现`);
        break;
      }
    } catch (e) {}
    await page.waitForTimeout(500);
  }

  if (!逐项配单找到) {
    logger.log(`[nav-3] 等待逐项配单超时，尝试直接点击`);
  }

  // 7. 点击 逐项配单 (展开后的子菜单)
  logger.log(`[nav-3] 点击 逐项配单`);
  try {
    await menuFrame.locator('text=逐项配单').click({ timeout: 5000 });
    logger.log(`[nav-3] 点击成功`);
  } catch (e) {
    logger.log(`[nav-3] locator点击逐项配单失败，尝试JS点击: ${e.message}`);
    try {
      await menuFrame.evaluate(() => {
        const els = Array.from(document.querySelectorAll('span, div, a'));
        const el = els.find(el => el.textContent.trim() === '逐项配单');
        if (el) el.click();
      });
    } catch (e2) {
      logger.log(`[nav-3] JS点击也失败: ${e2.message}`);
    }
  }

  // 点击后等待一段时间让表单加载
  logger.log(`[nav-3] 等待表单加载...`);
  await page.waitForTimeout(5000);

  // 检查是否打开了新页面
  const newPage = await switchToTargetPage(page, logger);
  if (newPage) {
    logger.log(`[nav-3] 检测到新页面，切換过去`);
    return newPage;
  }

  return page;
}
async function waitFor逐项配单Form(page, logger) {
  logger.log(`[nav-4] 等待逐项配单表单加载...`);

  for (let retry = 0; retry < 10; retry++) {
    await page.waitForTimeout(2000);

    // 遍历所有页面和所有frames，找目标表单
    const allPages = page.context().pages();
    for (const p of allPages) {
      if (p.isClosed()) continue;
      const frames = p.frames();
      for (const f of frames) {
        try {
          const text = await f.evaluate(() => document.body ? document.body.innerText : '').catch(() => '');
          if (text.includes('报关单号') && (text.includes('搜索') || text.includes('查询'))) {
            logger.log(`[nav-4] 逐项配单表单已加载`);
            return p; // 返回所在页面
          }
        } catch (e) {}
      }
    }

    // 也检查嵌套iframe
    for (const p of allPages) {
      if (p.isClosed()) continue;
      for (const f of p.frames()) {
        try {
          const iframes = await f.locator('iframe').all();
          for (const ifr of iframes) {
            const childFrame = await f.frame(ifr);
            if (childFrame) {
              const text = await childFrame.evaluate(() => document.body ? document.body.innerText : '').catch(() => '');
              if (text.includes('报关单号')) {
                logger.log(`[nav-4] 逐项配单表单已加载 (在iframe中)`);
                return;
              }
            }
          }
        } catch (e) {}
      }
    }

    logger.log(`[nav-4] 等待表单加载... (${retry + 1}/10)`);
  }

  logger.log(`[nav-4] 表单加载超时，将继续执行`);
}

// 切换到目标页面（不关闭任何页面，只返回目标页面引用）
// 用户要求不关闭浏览器，此函数只查找并返回目标页面引用，不做任何关闭操作
async function switchToTargetPage(page, logger) {
  try {
    const context = page.context();
    const pages = await context.pages();
    logger.log(`[nav] 共有 ${pages.length} 个页面，开始排查目标页面...`);

    // 遍历所有页面，找包含"逐项配单"+"报关单号"搜索表单的页面
    let targetPage = null;
    for (const p of pages) {
      const url = p.url() || '';
      const title = await p.title().catch(() => '');
      logger.log(`[nav] 检查页面：${url.substring(0, 70)}`);

      // 跳过登录页
      if (url.includes('loginb') || title.includes('登录')) {
        logger.log(`[nav]   -> 跳过（登录页）`);
        continue;
      }

      // 检查所有frames，要求"逐项配单"+"报关单号"同时存在（搜索表单）
      let foundFrame = null;
      for (const f of p.frames()) {
        try {
          const text = await f.evaluate(() => document.body ? document.body.innerText : '').catch(() => '');
          if (text.includes('逐项配单') && text.includes('报关单号')) {
            logger.log(`[nav]   -> frame "${f.name()}" 含"逐项配单"+"报关单号"，选中`);
            foundFrame = f;
            break;
          }
        } catch (e) {}
      }
      if (foundFrame) {
        targetPage = p;
        break;
      }
    }

    // 如果没找到含搜索表单的页面，记录警告
    if (!targetPage) {
      logger.log(`[nav] 警告：未找到含"逐项配单"+"报关单号"搜索表单的页面`);
      for (const p of pages) {
        if (p === page) continue;
        const url = p.url() || '';
        if (!url.includes('loginb')) {
          logger.log(`[nav] 备用页面：${p.url().substring(0, 60)}`);
          targetPage = p;
          break;
        }
      }
    } else {
      logger.log(`[nav] 目标页面确定：${targetPage.url().substring(0, 70)}`);
    }

    return targetPage;
  } catch (e) {
    logger.log(`[nav] 切换页面失败：${e.message}`);
  }
  return null;
}

// 关闭所有其他标签页，只保留指定页面
// 注意：只能关闭标签页，不能关闭浏览器窗口本身
async function closeOtherPages(page, logger) {
  const context = page.context();
  const allPages = context.pages();
  const toClose = allPages.filter(p => !p.isClosed() && p !== page);
  if (toClose.length === 0) {
    logger.log(`[nav] 只有 1 个标签页，无需关闭`);
    return;
  }
  logger.log(`[nav] 关闭 ${toClose.length} 个多余标签页...`);
  for (const p of toClose) {
    try {
      const url = p.url().substring(0, 60);
      await p.close();
      logger.log(`[nav] 已关闭：${url}`);
    } catch (e) {
      logger.log(`[nav] 关闭页面失败：${e.message}`);
    }
  }
}


async function detectTargetPageReady(page) {
  if (await isLikelyLoginPage(page)) {
    return false;
  }

  if (await hasTargetPageTitle(page)) {
    return true;
  }

  return hasTargetSearchArea(page);
}

async function findReadyTargetPage(context) {
  for (const page of context.pages()) {
    if (page.isClosed()) {
      continue;
    }
    if (await detectTargetPageReady(page)) {
      return page;
    }

    // 检查页面中的所有frames
    const frames = page.frames();
    for (const frame of frames) {
      try {
        if (await hasTargetSearchArea(frame)) {
          logger.log(`[detect] 在 frame "${frame.name()}" 中找到目标页面`);
          return page; // 返回主页面，不是frame
        }
      } catch (e) {}
    }
  }
  return null;
}

async function waitForTargetPageReady(context, logger, timeoutMs = 900000) {
  const startedAt = Date.now();
  let lastNoticeAt = 0;
  logger.log('[manual] 请你手动登录并进入“逐项配单”页面，脚本检测到目标页后会自动继续');

  while (Date.now() - startedAt < timeoutMs) {
    const targetPage = await findReadyTargetPage(context);
    if (targetPage) {
      logger.log(`[manual] 已检测到目标页面：${targetPage.url()}`);
      // 关闭其他多余的页面
      await closeOtherPages(targetPage, logger);
      return targetPage;
    }

    if (Date.now() - lastNoticeAt > 30000) {
      const currentPages = context.pages().filter(page => !page.isClosed()).map(page => page.url()).filter(Boolean);
      logger.log(`[manual] 等待手动进入“逐项配单”页面。当前页面：${currentPages.join(' | ') || 'about:blank'}`);
      lastNoticeAt = Date.now();
    }

    for (const page of context.pages()) {
      if (!page.isClosed()) {
        await waitForUi(page, 800).catch(() => {});
      }
    }
  }

  throw new Error('长时间未检测到“逐项配单”页面，请确认你已手动登录并打开目标页');
}

function buildRunContext({ sourceFile, rows, totalRowsInSheet = rows.length }) {
  return {
    sourceFile,
    sheetName: rows[0]?.sheetName || config.input.sheetName || '',
    totalRows: rows.length,
    totalRowsInSheet
  };
}

function shouldSkipRow(row) {
  const status = normalizeText(row?.processingStatus);
  return status === COMPLETED_STATUS || status === '已跳过';
}

function buildRowStatusReason(row, error) {
  const message = normalizeText(error?.message || error || '');
  if (!message) {
    return '未知错误';
  }

  if (message.includes('未找到完全匹配的报关单行')) {
    return `${message}（可能已手工处理、网页状态已变化，或当前结果与 Excel 不一致）`;
  }

  if (message.includes('未在发票选择区域找到匹配发票')) {
    return `${message}（当前网页发票行与 Excel 备注/行号不一致）`;
  }

  return message;
}

function writeRowStatus(row, status, message, logger) {
  updateTaxRowStatus({
    filePath: row.sourceFile,
    sheetName: row.sheetName,
    rowNumber: row.rowNumber,
    status,
    message,
    timestamp: new Date().toISOString()
  });
  logger.log(`[row ${row.rowNumber}] 已写入状态：${status}${message ? ` - ${message}` : ''}`);
}

async function inspectTargetPage(page, rows, logger, runContext) {
  logger.log(`[inspect] 当前 Excel 行数：${rows.length}`);
  if (rows[0]) {
    logger.log(`[inspect] 首行样例：报关单号=${rows[0].declarationNo}，品名=${rows[0].goodsName}，发票号=${rows[0].invoiceNo}`);
  }

  await saveInspectionArtifacts(page, 'target-page', logger);
  appendRunState({
    mode: 'inspect',
    status: 'ready_for_dom_mapping',
    currentUrl: page.url(),
    excelFile: runContext.sourceFile,
    sheetName: runContext.sheetName,
    totalRows: runContext.totalRows
  });
}

async function selectInvoiceRows(page, row, logger) {
  await clickTextAcrossPage(page, [config.taxRefund.selectInvoiceButtonText, '选择发票信息'], { exact: false });
  await waitForUi(page, 1200);
  await searchInvoiceRowsInDialog(page, row, logger);

  const artifacts = await saveInspectionArtifacts(page, `row-${row.rowNumber}-invoice-dialog`, logger);
  const matches = findInvoiceRowMatches(artifacts.summary, row);
  const expectedCount = Array.from(new Set((row.selectedInvoiceLineNos || []).map(value => normalizeNumberString(value)).filter(Boolean))).length;

  if (!matches.length) {
    const candidates = findInvoiceSelectionCandidates(artifacts.summary)
      .map(tableEntry => {
        const headers = tableEntry.table.headers || [];
        const indexMap = getInvoiceSelectionIndexMap(headers);
        return (tableEntry.table.rows || []).map(candidateRow => ({
          where: `frame=${tableEntry.frameIndex} root=${tableEntry.table.rootSelector}[${tableEntry.table.rootIndex}] row=${candidateRow.rowIndex}`,
          invoiceNo: getAlignedCellValue(candidateRow, headers, indexMap.invoiceNo),
          invoiceLineNo: getAlignedCellValue(candidateRow, headers, indexMap.invoiceLineNo),
          goodsName: getAlignedCellValue(candidateRow, headers, indexMap.goodsName),
          unit: getAlignedCellValue(candidateRow, headers, indexMap.unit),
          qty: getAlignedCellValue(candidateRow, headers, indexMap.qty),
          taxAmount: getAlignedCellValue(candidateRow, headers, indexMap.taxAmount)
        }));
      })
      .flat()
      .slice(0, 10);

    appendRunState({
      mode: 'run',
      status: 'invoice_row_not_found',
      rowNumber: row.rowNumber,
      declarationNo: row.declarationNo,
      invoiceNo: row.invoiceNo,
      candidates
    });

    throw new Error(`未在发票选择区域找到匹配发票：发票号=${row.invoiceNo} 行号=${(row.selectedInvoiceLineNos || []).join(',')}`);
  }

  const selectedLineNos = new Set();
  for (const match of matches) {
    const target = await findInvoiceCheckboxTarget(page, match);
    const rowLocator = target.rowLocator;
    await rowLocator.scrollIntoViewIfNeeded().catch(() => {});

    const checkbox = target.checkbox;

    if (!checkbox) {
      throw new Error(`已找到发票行，但未找到可勾选控件：${describeRowMatch(match)} 附近表 rootIndex=${target.tableRootIndex}`);
    }

    const checkedBefore = await checkbox.evaluate(el => {
      if ('checked' in el) {
        return Boolean(el.checked);
      }
      return el.getAttribute('aria-checked') === 'true' || el.classList?.contains('layui-form-checked');
    }).catch(() => false);
    if (!checkedBefore) {
      await checkbox.click({ force: true }).catch(async () => {
        await rowLocator.locator('td, th').first().click({ force: true }).catch(async () => {
          await rowLocator.click({ force: true }).catch(() => {});
        });
      });
      await waitForUi(page, 300);
    }
    selectedLineNos.add(match.currentLineNo);
    logger.log(`[row ${row.rowNumber}] 已勾选发票行号：${match.currentLineNo}`);
  }

  if (expectedCount && selectedLineNos.size !== expectedCount) {
    throw new Error(`发票行号勾选数量不完整：期望 ${expectedCount}，实际 ${selectedLineNos.size}`);
  }

  const confirmed = await clickConfirmInvoiceSelection(page);
  if (!confirmed) {
    await clickTextAcrossPage(page, [config.taxRefund.confirmSelectButtonText, '选择', '确定'], { exact: false });
  }
  await waitForUi(page, 1200);
  await saveInspectionArtifacts(page, `row-${row.rowNumber}-invoice-selected`, logger);
}

function resolveDomCellIndex(row, headers, alignedIndex) {
  const cells = row?.cells || [];
  const normalizedHeaders = (headers || []).map(header => normalizeText(header)).filter(Boolean);
  const hasLeadingSelectionCell = cells.length === normalizedHeaders.length + 1 && normalizeText(cells[0]) === '';
  return hasLeadingSelectionCell ? alignedIndex + 1 : alignedIndex;
}

async function clickConfirmInvoiceSelection(page) {
  for (const ctx of getSearchContexts(page)) {
    const candidates = [
      ctx.getByRole('button', { name: /^选择(?:\(\d+\))?$/ }),
      ctx.getByRole('link', { name: /^选择(?:\(\d+\))?$/ }),
      ctx.getByText(/^选择(?:\(\d+\))?$/),
      ctx.locator('button, a, .layui-btn, .el-button').filter({ hasText: /^选择(?:\(\d+\))?$/ })
    ];

    const visible = await firstVisible(candidates);
    if (!visible) {
      continue;
    }

    await visible.scrollIntoViewIfNeeded().catch(() => {});
    await visible.click({ timeout: 3000 }).catch(async () => {
      await visible.click({ force: true, timeout: 2000 }).catch(() => {});
    });
    return true;
  }

  return false;
}

async function setEditableCellValue(page, match, headerIndex, nextValue, headers = []) {
  const rowLocator = getRowLocator(page, match);
  const domCellIndex = resolveDomCellIndex(match.row, headers, headerIndex);
  const cellLocator = rowLocator.locator(match.table.table.cellSelector).nth(domCellIndex);
  await cellLocator.scrollIntoViewIfNeeded().catch(() => {});

  let input = cellLocator.locator('input, textarea').first();
  let hasInput = await input.count().catch(() => 0);

  if (!hasInput) {
    await cellLocator.click({ force: true }).catch(() => {});
    await waitForUi(page, 500);
    input = cellLocator.locator('input, textarea').first();
    hasInput = await input.count().catch(() => 0);
  }

  if (!hasInput) {
    throw new Error('自动修正时未定位到可编辑输入框');
  }

  await input.fill(String(nextValue));
  await input.blur().catch(() => {});
}

async function autoAdjustInvoiceRows(page, row, detailTableEntry, validation, logger) {
  if (!validation.canAutoAdjust) {
    throw new Error('当前差异不属于允许自动修正范围');
  }

  const qtyIndex = validation.fieldIndexMap.dealQty;
  const taxIndex = validation.fieldIndexMap.taxAmount;
  if (qtyIndex < 0 || taxIndex < 0) {
    throw new Error('自动修正缺少数量或计税金额列');
  }

  const rowsWithNumbers = validation.matchedRows.map(currentRow => ({
    row: currentRow,
    qtyText: getAlignedCellValue(currentRow, validation.headers, qtyIndex),
    taxText: getAlignedCellValue(currentRow, validation.headers, taxIndex),
    qty: toNumber(getAlignedCellValue(currentRow, validation.headers, qtyIndex)),
    tax: toNumber(getAlignedCellValue(currentRow, validation.headers, taxIndex))
  })).filter(item => Number.isFinite(item.qty) && Number.isFinite(item.tax));

  if (!rowsWithNumbers.length) {
    throw new Error(`自动修正未找到可计算的发票明细行：${JSON.stringify(validation.matchedRows.map(currentRow => ({
      cells: currentRow.cells,
      qtyText: getAlignedCellValue(currentRow, validation.headers, qtyIndex),
      taxText: getAlignedCellValue(currentRow, validation.headers, taxIndex)
    })) )}`);
  }

  const maxQtyRow = rowsWithNumbers.reduce((best, current) => current.qty > best.qty ? current : best, rowsWithNumbers[0]);
  const deltaQty = validation.expected.qty - validation.totals.qty;
  const deltaTax = validation.expected.tax - validation.totals.tax;
  const nextQty = maxQtyRow.qty + deltaQty;
  const nextTax = maxQtyRow.tax + deltaTax;
  const normalizedNextQty = areNumbersClose(nextQty, Math.round(nextQty)) ? Math.round(nextQty) : Number(nextQty.toFixed(5));
  const normalizedNextTax = Number(nextTax.toFixed(2));

  if (normalizedNextQty < 0 || normalizedNextTax < 0) {
    throw new Error(`自动修正后出现负数：数量=${normalizedNextQty} 计税金额=${normalizedNextTax}`);
  }

  const rowMatch = {
    table: detailTableEntry,
    row: maxQtyRow.row
  };

  await setEditableCellValue(page, rowMatch, qtyIndex, normalizedNextQty, validation.headers);
  await setEditableCellValue(page, rowMatch, taxIndex, normalizedNextTax, validation.headers);
  await waitForUi(page, 500);

  logger.log(`[row ${row.rowNumber}] 已自动修正最大数量行：数量 ${maxQtyRow.qty} -> ${normalizedNextQty}，计税金额 ${maxQtyRow.tax} -> ${normalizedNextTax}`);
  await saveInspectionArtifacts(page, `row-${row.rowNumber}-after-auto-adjust`, logger);
}

async function verifyBeforeSave(page, row, logger) {
  const artifacts = await saveInspectionArtifacts(page, `row-${row.rowNumber}-before-save`, logger);
  const candidates = findSelectedDetailCandidates(artifacts.summary);
  if (!candidates.length) {
    throw new Error('保存前未找到可核对的发票明细表格，请根据 logs 中的 JSON 补充真实 DOM 映射');
  }

  let lastError = null;
  for (const candidate of candidates) {
    try {
      const validation = validateSelectedInvoiceRows(candidate, row);
      if (!validation.qtyMatches || !validation.taxMatches) {
        if (!validation.canAutoAdjust) {
          throw new Error(`保存前核对失败：${validation.issues.join('；') || '数量/计税金额不一致'}`);
        }

        logger.log(`[row ${row.rowNumber}] 检测到数量/计税金额差异，尝试自动修正`);
        await autoAdjustInvoiceRows(page, row, candidate, validation, logger);

        const refreshedArtifacts = await saveInspectionArtifacts(page, `row-${row.rowNumber}-after-adjust-verify`, logger);
        const refreshedCandidates = findSelectedDetailCandidates(refreshedArtifacts.summary);
        const refreshedCandidate = refreshedCandidates.find(item => item.frameIndex === candidate.frameIndex && item.table.rootSelector === candidate.table.rootSelector && item.table.rootIndex === candidate.table.rootIndex) || refreshedCandidates[0];
        const refreshedValidation = validateSelectedInvoiceRows(refreshedCandidate, row);
        if (!refreshedValidation.qtyMatches || !refreshedValidation.taxMatches || refreshedValidation.nonAdjustableIssues.length > 0) {
          throw new Error(`自动修正后仍未通过核对：${refreshedValidation.issues.join('；') || '数量/计税金额仍不一致'}`);
        }
        return;
      }

      if (validation.nonAdjustableIssues.length > 0) {
        throw new Error(`保存前核对失败：${validation.nonAdjustableIssues.join('；')}`);
      }

      logger.log(`[row ${row.rowNumber}] 保存前核对通过：数量=${validation.totals.qty}，计税金额=${validation.totals.tax}`);
      return;
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError || new Error('保存前核对失败');
}

async function saveCurrentForm(page, row, logger) {
  await clickTextAcrossPage(page, [config.taxRefund.saveButtonText, '保存'], { exact: false });
  await waitForUi(page, 2500);
  await saveInspectionArtifacts(page, `row-${row.rowNumber}-after-save`, logger);
  logger.log(`[row ${row.rowNumber}] 已点击保存`);
}

async function forceCloseDialogs(page, row, logger) {
  // 强制关闭弹窗：尝试多种方法
  const frames = page.frames();

  // 方法1：用JS直接关闭弹窗（最有效）
  try {
    for (const frame of frames) {
      // 尝试用JS点击关闭按钮
      const closeResult = await frame.evaluate(() => {
        // 尝试多种关闭方式
        const selectors = [
          'button.ant-modal-close', '.ant-modal-close', '[aria-label="关闭"]',
          '.el-dialog__close', '.ui-dialog-close', 'button[class*="close"]',
          '.ant-btn-close', 'button.close', '[class*="modal-close"]'
        ];
        for (const sel of selectors) {
          const btn = document.querySelector(sel);
          if (btn) { btn.click(); return 'clicked:' + sel; }
        }
        // 尝试按ESC
        document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
        return 'esc-dispatched';
      }).catch(() => 'js-failed');
      if (closeResult !== 'js-failed') {
        logger.log(`[row ${row.rowNumber}] JS关闭弹窗: ${closeResult}`);
      }
    }
  } catch (e) {
    logger.log(`[row ${row.rowNumber}] JS关闭失败: ${e.message}`);
  }

  // 方法2：按Escape多次（最安全）
  for (let i = 0; i < 8; i++) {
    await page.keyboard.press('Escape');
    await waitForUi(page, 200);
  }

  // 方法3：点击各frame中的关闭按钮
  for (const frame of frames) {
    try {
      // 尝试多种关闭文本
      for (const closeText of ['关闭', '×', 'x', 'X', '取消']) {
        const closeBtn = frame.locator('text="' + closeText + '"').first();
        if (await closeBtn.isVisible({ timeout: 500 }).catch(() => false)) {
          await closeBtn.click({ timeout: 1000, force: true }).catch(() => {});
          logger.log(`[row ${row.rowNumber}] 在 frame "${frame.name()}" 点击了 "${closeText}"`);
          await waitForUi(page, 500);
        }
      }
    } catch (e) {}
  }

  // 方法4：点击遮罩层关闭（弹窗外的深色背景区域）
  try {
    // 尝试点击 ant-design 的 mask/backdrop
    const mask = page.locator('.ant-modal-mask, .ant-backdrop, .el-dialog__mask, .v-modal, [class*="mask"], [class*="backdrop"]').first();
    if (await mask.isVisible({ timeout: 500 }).catch(() => false)) {
      await mask.click({ timeout: 1000, force: true }).catch(() => {});
      logger.log(`[row ${row.rowNumber}] 点击了遮罩层`);
      await waitForUi(page, 500);
    }
  } catch (e) {}

  // 方法5：点击弹窗背景区域（弹窗外部）
  try {
    const modalWrapper = page.locator('.ant-modal-wrap, .el-dialog__wrapper, .v-modal__wrapper, [class*="modal-wrap"]').first();
    if (await modalWrapper.isVisible({ timeout: 500 }).catch(() => false)) {
      // 点击弹窗外部区域
      await modalWrapper.click({ position: { x: 10, y: 10 }, timeout: 1000, force: true }).catch(() => {});
      logger.log(`[row ${row.rowNumber}] 点击了弹窗外部区域`);
      await waitForUi(page, 500);
    }
  } catch (e) {}

  // 方法6：点击每个frame的右上角关闭按钮（×）
  for (const frame of frames) {
    try {
      // 尝试点击frame右上角的×按钮
      const closeX = frame.locator('[aria-label="关闭"], .ant-modal-close, .el-dialog__close, .ui-dialog-close, button[class*="close"], [class*="close"]').first();
      if (await closeX.isVisible({ timeout: 500 }).catch(() => false)) {
        await closeX.click({ timeout: 1000, force: true }).catch(() => {});
        logger.log(`[row ${row.rowNumber}] 点击了frame "${frame.name()}" 的×关闭按钮`);
        await waitForUi(page, 500);
      }
    } catch (e) {}
  }

  await waitForUi(page, 800);
}

async function ensureBackToList(page, row, logger) {
  // Always try to close dialogs first by clicking 关闭 in ALL frames
  for (let attempt = 0; attempt < 3; attempt++) {
    // 先尝试强制关闭
    await forceCloseDialogs(page, row, logger);

    // Try clicking 关闭 in each frame
    const frames = page.frames();
    for (const frame of frames) {
      try {
        const closeBtn = frame.locator('text=关闭').first();
        if (await closeBtn.isVisible({ timeout: 1000 }).catch(() => false)) {
          await closeBtn.click({ timeout: 2000 }).catch(() => {});
          logger.log(`[row ${row.rowNumber}] 在 frame "${frame.name()}" 点击了关闭`);
        }
      } catch (e) {
        // Frame not accessible
      }
    }

    // Try Escape key globally
    await page.keyboard.press('Escape').catch(() => {});
    await waitForUi(page, 1500);

    // Now check if we can find the search field AND no modal is visible
    const searchField = await findSearchField(page);
    const hasModal = await checkVisibleModal(page);

    if (searchField && !hasModal) {
      if (attempt === 0) {
        logger.log(`[row ${row.rowNumber}] 已在列表页`);
      } else {
        logger.log(`[row ${row.rowNumber}] 关闭弹窗后回到列表页`);
      }
      return;
    }

    // Not on list page - try 返回 button
    logger.log(`[row ${row.rowNumber}] 未找到搜索框或弹窗未关闭，点击返回（第 ${attempt + 1} 次）`);
    await clickTextAcrossPage(page, [config.taxRefund.backButtonText, '返回'], { exact: false }).catch(() => {});
    await waitForUi(page, 1500);

    const searchFieldAfterBack = await findSearchField(page);
    const hasModalAfterBack = await checkVisibleModal(page);
    if (searchFieldAfterBack && !hasModalAfterBack) {
      logger.log(`[row ${row.rowNumber}] 点击返回后回到列表页`);
      return;
    }
  }

  // 弹窗无法关闭，尝试刷新iframe内容
  logger.log(`[row ${row.rowNumber}] 弹窗无法关闭，尝试重新加载iframe内容...`);

  // 点击"免退税申报"再点"逐项配单"来重新加载iframe
  await clickTextAcrossPage(page, ['免退税申报'], { exact: false }).catch(() => {});
  await waitForUi(page, 2000);
  await clickTextAcrossPage(page, ['逐项配单'], { exact: false }).catch(() => {});
  await waitForUi(page, 3000);

  // 检查是否恢复正常
  const searchFieldAfterReload = await findSearchField(page);
  const hasModalAfterReload = await checkVisibleModal(page);
  if (searchFieldAfterReload && !hasModalAfterReload) {
    logger.log(`[row ${row.rowNumber}] iframe重新加载成功，回到列表页`);
    return;
  }

  // 还不行，再试一次
  await clickTextAcrossPage(page, ['免退税申报'], { exact: false }).catch(() => {});
  await waitForUi(page, 2000);
  await clickTextAcrossPage(page, ['逐项配单'], { exact: false }).catch(() => {});
  await waitForUi(page, 3000);

  const searchFieldAfterRetry = await findSearchField(page);
  const hasModalAfterRetry = await checkVisibleModal(page);
  if (searchFieldAfterRetry && !hasModalAfterRetry) {
    logger.log(`[row ${row.rowNumber}] iframe重新加载成功（第2次），回到列表页`);
    return;
  }

  // 还是不行，报告错误让用户处理
  logger.log(`[row ${row.rowNumber}] 无法自动恢复，请在浏览器中手动操作后重试`);
  await saveInspectionArtifacts(page, `row-${row.rowNumber}-ensure-back-failed`, logger);
  throw new Error('页面无法自动恢复，请手动操作后重试');
}

async function checkVisibleModal(page) {
  try {
    // 检查主页面是否有可见的弹窗
    const modalSelectors = [
      '.ant-modal', '.el-dialog', '.v-modal', '.modal',
      '[role="dialog"]', '[aria-modal="true"]'
    ];
    for (const sel of modalSelectors) {
      const modal = page.locator(sel).first();
      if (await modal.isVisible({ timeout: 500 }).catch(() => false)) {
        return true;
      }
    }

    // 检查各iframe里是否有弹窗
    const frames = page.frames();
    for (const frame of frames) {
      try {
        for (const sel of modalSelectors) {
          const modal = frame.locator(sel).first();
          if (await modal.isVisible({ timeout: 300 }).catch(() => false)) {
            return true;
          }
        }
        // 检查iframe里的关闭按钮是否可见
        const closeBtn = frame.locator('text=关闭').first();
        if (await closeBtn.isVisible({ timeout: 300 }).catch(() => false)) {
          return true;
        }
      } catch (e) {}
    }

    return false;
  } catch (e) {
    return false;
  }
}

async function checkSessionExpired(page) {
  // 检测"用户信息失效"等session过期提示
  const sessionExpiredTexts = [
    '用户信息失效',
    '请重新登录',
    '登录已过期',
    'session过期',
    '登录超时'
  ];

  try {
    for (const ctx of getSearchContexts(page)) {
      for (const text of sessionExpiredTexts) {
        const found = await ctx.getByText(text, { exact: false }).isVisible({ timeout: 500 }).catch(() => false);
        if (found) {
          return true;
        }
      }
    }
  } catch (e) {}
  return false;
}

// 页面识别：检测当前在哪个页面
async function identifyCurrentPage(page, logger) {
  const result = {
    pageName: 'unknown',
    menuPath: [],
    hasSearchArea: false,
    isLoginPage: false
  };

  try {
    const url = page.url() || '';
    const title = await page.title().catch(() => '');

    // 检测是否登录页面
    if (url.includes('loginb') || title.includes('登录') || url.includes('#/login')) {
      result.isLoginPage = true;
      result.pageName = '登录页';
      logger.log(`[page-identify] 当前页面：登录页`);
      return result;
    }

    // 获取所有frames的文本内容来识别页面
    const frames = page.frames();
    const allTexts = [];
    for (const frame of frames) {
      try {
        const frameText = await frame.locator('body').innerText({ timeout: 2000 }).catch(() => '');
        if (frameText) allTexts.push({ frame: frame.name(), text: frameText });
      } catch (e) {}
    }

    // 检测逐项配单页面（有搜索框+报关单号）
    for (const { frame, text } of allTexts) {
      if (text.includes('报关单号') && text.includes('搜索')) {
        result.pageName = '逐项配单';
        result.hasSearchArea = true;
        result.menuPath = ['地方特色', '特色办税', '单一窗口出口退（免）税办理', '在线申报', '免退税申报', '逐项配单'];
        logger.log(`[page-identify] 当前页面：逐项配单（frame: ${frame}）`);
        return result;
      }
    }

    // 检测当前在哪个菜单（从叶子节点往回检测）
    const menuItems = [
      { name: '免退税申报', keywords: ['免退税申报'] },
      { name: '在线申报', keywords: ['在线申报'] },
      { name: '单一窗口出口退（免）税办理', keywords: ['单一窗口出口退（免）税办理', '单一窗口'] },
      { name: '特色办税', keywords: ['特色办税'] },
      { name: '地方特色', keywords: ['地方特色'] }
    ];

    for (const { name, keywords } of menuItems) {
      for (const { frame, text } of allTexts) {
        for (const keyword of keywords) {
          if (text.includes(keyword)) {
            result.pageName = name;
            const idx = menuItems.findIndex(m => m.name === name);
            result.menuPath = menuItems.slice(idx).map(m => m.name);
            logger.log(`[page-identify] 当前页面：${name}（frame: ${frame}）`);
            return result;
          }
        }
      }
    }

    logger.log(`[page-identify] 当前页面：未知 | URL: ${url}`);
    return result;
  } catch (e) {
    logger.log(`[page-identify] 页面识别失败：${e.message}`);
    return result;
  }
}

async function runRowFlowSkeleton(page, row, logger) {
  logger.log(`========== [row ${row.rowNumber}] 开始处理 ==========`);
  logger.log(`报关单号=${row.declarationNo} 项号=${row.declarationItemNo} 发票号=${row.invoiceNo}`);

  appendRunState({
    mode: 'run',
    status: 'row_started',
    rowNumber: row.rowNumber,
    declarationNo: row.declarationNo,
    invoiceNo: row.invoiceNo
  });

  // STEP 1: 确保在列表页
  logger.log(`[STEP 1/5] 确保在搜索列表页`);
  await ensureBackToList(page, row, logger);

  // STEP 2: 搜索报关单
  logger.log(`[STEP 2/5] 搜索报关单号：${row.declarationNo}`);
  await searchDeclarationRow(page, row, logger);

  // STEP 3: 查找并点击配单按钮
  const searchArtifacts = await saveInspectionArtifacts(page, `row-${row.rowNumber}-after-search`, logger);
  const matches = findDeclarationRowMatches(searchArtifacts.summary, row);
  if (!matches.length) {
    const candidates = flattenTables(searchArtifacts.summary)
      .map(tableEntry => {
        const { headers, rows } = tableEntry.table;
        const fieldIndexMap = buildFieldIndexMap(headers, ['declarationNo', 'goodsName', 'declarationItemNo', 'dealUnit', 'dealQty']);
        if (listMissingFieldHeaders(fieldIndexMap).length > 0) {
          return [];
        }
        return (rows || []).map(candidateRow => {
          const alignedCells = buildAlignedCellSamples(headers, candidateRow);
          return {
            where: `frame=${tableEntry.frameIndex} root=${tableEntry.table.rootSelector}[${tableEntry.table.rootIndex}] row=${candidateRow.rowIndex}`,
            declarationNo: alignedCells[fieldIndexMap.declarationNo] ?? '',
            goodsName: alignedCells[fieldIndexMap.goodsName] ?? '',
            declarationItemNo: alignedCells[fieldIndexMap.declarationItemNo] ?? '',
            dealUnit: alignedCells[fieldIndexMap.dealUnit] ?? '',
            dealQty: alignedCells[fieldIndexMap.dealQty] ?? ''
          };
        });
      })
      .flat()
      .slice(0, 10);

    appendRunState({
      mode: 'run',
      status: 'declaration_row_not_found',
      rowNumber: row.rowNumber,
      declarationNo: row.declarationNo,
      invoiceNo: row.invoiceNo,
      candidates
    });

    logger.log(`[STEP 3/5] 报关单行查找结果：未找到匹配行`);
    throw new Error(`网页中未找到完全匹配的报关单行：报关单号=${row.declarationNo}，品名=${row.goodsName}`);
  }
  if (matches.length > 1) {
    logger.log(`[STEP 3/5] 报关单行查找结果：找到 ${matches.length} 条匹配`);
    throw new Error(`网页中找到 ${matches.length} 条完全匹配的报关单行，无法唯一确定目标：${matches.map(describeRowMatch).join('；')}`);
  }
  logger.log(`[STEP 3/5] 报关单行查找结果：找到 1 条匹配`);
  logger.log(`  匹配值=${JSON.stringify(matches[0].values)}`);

  logger.log(`[STEP 3/5] 点击"配单"按钮`);
  await clickRowAction(page, matches[0], [config.taxRefund.rowActionText, '配单'], logger);
  await waitForUi(page, 2500);
  await saveInspectionArtifacts(page, `row-${row.rowNumber}-after-action`, logger);

  // STEP 4: 选择发票并保存
  logger.log(`[STEP 4/5] 选择发票并核对`);

  // 检查session是否过期
  const expired = await checkSessionExpired(page);
  if (expired) {
    throw new Error('用户登录信息已过期，请关闭浏览器重新登录后再运行脚本');
  }

  await selectInvoiceRows(page, row, logger);
  await verifyBeforeSave(page, row, logger);
  await saveCurrentForm(page, row, logger);
  logger.log(`[STEP 5/5] 保存完成，准备返回列表页`);
  await ensureBackToList(page, row, logger);

  appendRunState({
    mode: 'run',
    status: 'row_completed',
    rowNumber: row.rowNumber,
    declarationNo: row.declarationNo,
    invoiceNo: row.invoiceNo
  });
  logger.log(`========== [row ${row.rowNumber}] 处理完成 ==========`);
}

async function cleanup(connection, logger) {
  if (!connection || !connection.browser) {
    return;
  }

  try {
    if (typeof connection.browser.disconnect === 'function') {
      await connection.browser.disconnect();
      logger?.log('[browser] 已断开 CDP 连接，浏览器窗口保持打开');
    } else {
      logger?.log('[browser] 不执行关闭，浏览器窗口保持打开');
    }
  } catch (error) {
    logger?.log(`[browser] 断开连接时已忽略关闭异常：${error.message}`);
  }
}

function parseRequestedRowNumber(argv = process.argv.slice(2)) {
  const rowArg = argv.find(arg => arg.startsWith('--row='));
  if (!rowArg) {
    return null;
  }

  const value = Number(rowArg.slice('--row='.length));
  if (!Number.isInteger(value) || value < 2) {
    throw new Error('参数 --row 必须是 Excel 中大于等于 2 的行号');
  }

  return value;
}

function filterRowsForRun(rows, requestedRowNumber) {
  if (requestedRowNumber) {
    const matchedRow = rows.find(row => row.rowNumber === requestedRowNumber);
    if (!matchedRow) {
      throw new Error(`Excel 中未找到第 ${requestedRowNumber} 行可处理数据`);
    }
    return [matchedRow];
  }

  return rows.filter(row => !shouldSkipRow(row));
}

function finalizeExcelIfFullyProcessed(sourceFile, runContext, logger) {
  const refreshedRows = loadTaxRows({ filePath: sourceFile });
  const unresolvedRows = refreshedRows.filter(row => !shouldSkipRow(row));
  if (unresolvedRows.length > 0) {
    logger.log(`[file] 仍有 ${unresolvedRows.length} 行未完成，Excel 保留在待处理目录`);
    return null;
  }

  const archivedFile = moveProcessedExcelFile(runContext.sourceFile, logger);
  appendRunState({
    mode: 'run',
    status: 'completed',
    excelFile: runContext.sourceFile,
    archivedFile,
    sheetName: runContext.sheetName,
    totalRows: runContext.totalRows
  });
  logger.log(`[file] 当前批次已全部完成，Excel 已归档到已处理目录：${archivedFile}`);
  return archivedFile;
}

async function main() {
  ensureProjectDirs();
  const errors = validateRuntimeConfig();
  if (errors.length > 0) {
    throw new Error(errors.join('；'));
  }

  const inspectOnly = process.argv.includes('--inspect') || !process.argv.includes('--run');
  const requestedRowNumber = parseRequestedRowNumber(process.argv.slice(2));
  const sourceFile = resolvePendingExcelFile();
  const allRows = loadTaxRows({ filePath: sourceFile });
  const rows = filterRowsForRun(allRows, requestedRowNumber);
  const runContext = buildRunContext({ sourceFile, rows, totalRowsInSheet: allRows.length });
  const logger = createLogger();
  logger.log(`运行模式：${inspectOnly ? 'inspect' : 'run'}`);
  logger.log(`Excel：${runContext.sourceFile}`);
  logger.log(`Sheet：${runContext.sheetName || '首个工作表'}`);
  logger.log(`本次待处理 ${runContext.totalRows} 行；工作表总行数 ${runContext.totalRowsInSheet}`);
  if (requestedRowNumber) {
    logger.log(`仅处理 Excel 第 ${requestedRowNumber} 行`);
  }
  if (!rows.length) {
    logger.log('没有需要处理的 Excel 行：已处理行会自动跳过');
    return;
  }

  const connection = await launchOrConnectBrowser(logger);
  const { context, mode } = connection;
  logger.log(`[browser] 浏览器控制模式：${mode}`);
  logger.log(`[browser] 调试端口：${config.browser.remoteDebuggingPort}`);

  try {
  let targetPage = null;
  if (mode === 'connected') {
    const pages = context.pages();
    logger.log(`[browser] 当前共有 ${pages.length} 个标签页`);
    // 找到第一个非空白页作为目标
    for (const p of pages) {
      if (p.isClosed()) continue;
      const url = p.url() || '';
      if (url === 'about:blank') continue;
      targetPage = p;
      logger.log(`[browser] 选定目标页面：${url.substring(0, 70)}`);
      break;
    }
    if (!targetPage) {
      targetPage = pages[0];
    }
  }

  if (!targetPage) {
    throw new Error('未找到浏览器页面，请确保已打开 Chrome 并登录到逐项配单页面');
  }

  // 验证目标页面是否就绪（逐项配单页面）
  const isReady = await detectTargetPageReady(targetPage);
  if (!isReady) {
    throw new Error('当前页面不是逐项配单页面，请先手动登录并导航到"逐项配单"页面后再运行脚本');
  }
  logger.log(`[browser] 目标页面已就绪：${targetPage.url().substring(0, 70)}`);
  await setPageSizeTo50(targetPage, logger);
  await saveInspectionArtifacts(targetPage, 'target-page-ready', logger);

    if (inspectOnly) {
      await inspectTargetPage(targetPage, rows, logger, runContext);
      logger.log('探查模式完成。请把 logs 目录中的页面结构文件作为后续联调依据。');
      logger.log('浏览器窗口会保持打开，方便你继续查看页面。');
      return;
    }

    const summary = {
      success: 0,
      failed: 0,
      skippedCompleted: allRows.length - rows.length,
      rowResults: []
    };

    // 逐行处理，每行独立捕获错误
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      try {
        await runRowFlowSkeleton(targetPage, row, logger);
        writeRowStatus(row, COMPLETED_STATUS, '自动处理完成', logger);
        summary.success += 1;
        summary.rowResults.push({ rowNumber: row.rowNumber, status: COMPLETED_STATUS, message: '自动处理完成' });
        logger.log(`========== [row ${row.rowNumber}] 处理成功 ==========`);
      } catch (error) {
        const reason = buildRowStatusReason(row, error);
        logger.log(`[row ${row.rowNumber}] 处理失败：${reason}`);

        // 判断是否为致命错误（需要用户介入）
        if (isCriticalError(error)) {
          logger.log(`[row ${row.rowNumber}] 发生致命错误：${reason}`);
          const action = await askCriticalErrorAction(row.rowNumber, reason, logger);

          if (action === 'relogin') {
            // 重新登录：标记当前行为"待核对"后退出，让用户重新启动
            writeRowStatus(row, '待核对', reason + '（致命错误，已退出，--row=' + row.rowNumber + ' 重试）', logger);
            logger.log('请关闭浏览器后，重新运行脚本，并加参数 --row=' + row.rowNumber + ' 从这一行继续');
            try {
              if (connection.browser && typeof connection.browser.disconnect === 'function') {
                connection.browser.disconnect();
              }
            } catch (e) {}
            process.exit(0);
          } else if (action === 'retry') {
            // 重试此行：不标记失败，不计入成功，重新处理同一行
            logger.log(`[row ${row.rowNumber}] 重新处理此行...`);
            i--; // 循环结束后 i++，这里减一使得下一次循环还是同一行
          } else if (action === 'skip') {
            // 跳过此行，不写 Excel 状态，只记日志，继续下一行
            logger.log(`[row ${row.rowNumber}] 跳过此行，继续处理下一行（日志可查）`);
          } else {
            // quit - 退出
            logger.log('用户选择退出，Excel 状态不变，可重新运行脚本继续');
            process.exit(0);
          }
        } else {
          // 服务器抖动等普通错误：不写 Excel 状态，只记日志，继续下一行
          // 用户重新运行时会自动重试该行
          logger.log(`[row ${row.rowNumber}] 服务器错误：${reason}，继续处理下一行（重新运行脚本可重试）`);
        }
      }
    }

    appendRunState({
      mode: 'run',
      status: 'batch_completed',
      excelFile: runContext.sourceFile,
      sheetName: runContext.sheetName,
      totalRows: runContext.totalRows,
      totalRowsInSheet: runContext.totalRowsInSheet,
      summary
    });

    logger.log(`[summary] 成功 ${summary.success} 行，失败 ${summary.failed} 行，已跳过 ${summary.skippedCompleted} 行`);

    if (requestedRowNumber) {
      const archivedFile = finalizeExcelIfFullyProcessed(sourceFile, runContext, logger);
      if (!archivedFile) {
        logger.log(`已完成单行重试，未归档 Excel：第 ${requestedRowNumber} 行`);
      }
      return;
    }

    finalizeExcelIfFullyProcessed(sourceFile, runContext, logger);
  } catch (error) {
    appendRunState({
      mode: inspectOnly ? 'inspect' : 'run',
      status: 'failed',
      error: error.message,
      excelFile: runContext.sourceFile,
      sheetName: runContext.sheetName,
      totalRows: runContext.totalRows
    });
    logger.log(`[error] ${error.message}`);

    // 致命错误询问用户后续操作
    if (isCriticalError(error)) {
      const action = await askCriticalErrorAction('导航/初始化', error.message, logger);
      if (action === 'relogin') {
        logger.log('请关闭浏览器后，重新运行脚本。');
        try {
          if (connection.browser && typeof connection.browser.disconnect === 'function') {
            connection.browser.disconnect();
          }
        } catch (e) {}
        process.exit(0);
      }
      // S 或 Q 的情况：fallthrough 到下面的浏览器保持打开
    }

    logger.log('浏览器会保持打开，方便人工检查。');
    throw error;
  } finally {
    await cleanup(connection, logger);
  }
}

if (require.main === module) {
  main().catch(error => {
    console.error(`启动失败：${error.message}`);
    process.exitCode = 1;
  });
}

module.exports = {
  main,
  waitForUi,
  clickTextAcrossPage,
  collectPageSummary,
  saveInspectionArtifacts,
  findDeclarationRowMatches,
  findInvoiceRowMatches,
  validateSelectedInvoiceRows
};


