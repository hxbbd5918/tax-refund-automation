const path = require('path');
const fs = require('fs');

// 打包后通过环境变量指定根目录
const ROOT_DIR = process.env.TAX_ROOT_DIR || __dirname;
const LOCAL_CONFIG_FILE = path.join(ROOT_DIR, 'local.config.json');
const LOG_DIR = path.join(ROOT_DIR, 'logs');
const STATE_DIR = path.join(ROOT_DIR, 'state');
const STATE_FILE = path.join(STATE_DIR, 'tax_refund_state.json');
const PROFILE_DIR = path.join(ROOT_DIR, '.chrome-profile-tax');
const PENDING_DIR = path.join(ROOT_DIR, '待处理');
const PROCESSED_DIR = path.join(ROOT_DIR, '已处理');

function readJsonIfExists(filePath, fallback = {}) {
  if (!fs.existsSync(filePath)) return fallback;
  try {
    return JSON.parse(fs.readFileSync(filePath, 'utf8'));
  } catch {
    return fallback;
  }
}

function deepMerge(base, override) {
  if (!override || typeof override !== 'object') return base;
  const output = { ...base };
  for (const [key, value] of Object.entries(override)) {
    if (value && typeof value === 'object' && !Array.isArray(value) && output[key] && typeof output[key] === 'object' && !Array.isArray(output[key])) {
      output[key] = deepMerge(output[key], value);
    } else {
      output[key] = value;
    }
  }
  return output;
}

function firstExistingPath(paths) {
  for (const candidate of paths) {
    if (candidate && fs.existsSync(candidate)) return candidate;
  }
  return '';
}

function ensureDir(dir) {
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
}

const localConfig = readJsonIfExists(LOCAL_CONFIG_FILE, {});

const defaultConfig = {
  urls: {
    home: 'https://etax.shanghai.chinatax.gov.cn:8443/#/'
  },
  browser: {
    chromePath: process.env.TAX_CHROME_PATH || firstExistingPath([
      'C:/Program Files/Google/Chrome/Application/chrome.exe',
      'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe',
      path.join(process.env.LOCALAPPDATA || '', 'Google/Chrome/Application/chrome.exe')
    ]),
    userDataDir: process.env.TAX_USER_DATA_DIR || PROFILE_DIR,
    remoteDebuggingPort: Number(process.env.TAX_REMOTE_DEBUGGING_PORT || 9222),
    headless: false,
    args: [
      '--start-maximized'
    ]
  },
  input: {
    sheetName: process.env.TAX_EXCEL_SHEET || ''
  },
  paths: {
    rootDir: ROOT_DIR,
    logDir: LOG_DIR,
    stateDir: STATE_DIR,
    stateFile: STATE_FILE,
    pendingDir: PENDING_DIR,
    processedDir: PROCESSED_DIR,
    localConfigFile: LOCAL_CONFIG_FILE
  },
  taxRefund: {
    loginButtonText: '登录',
    loginSuccessTexts: ['退出登录', '安全退出', '退出', '用户中心'],
    menuPath: ['地方特色', '特色办税', '单一窗口出口退（免）税办理', '在线申报', '免退税申报', '逐项配单'],
    searchFieldLabel: '报关单号',
    searchButtonText: '搜索',
    rowActionText: '配单',
    selectInvoiceButtonText: '选择发票信息',
    confirmSelectButtonText: '选择',
    saveButtonText: '保存',
    backButtonText: '返回'
  }
};

const config = deepMerge(defaultConfig, localConfig);

function ensureProjectDirs() {
  ensureDir(config.paths.logDir);
  ensureDir(config.paths.stateDir);
  ensureDir(config.paths.pendingDir);
  ensureDir(config.paths.processedDir);
  ensureDir(config.browser.userDataDir);
}

function validateRuntimeConfig() {
  const errors = [];
  if (!config.browser.chromePath) {
    errors.push('未找到 Chrome 安装路径');
  }
  return errors;
}

module.exports = {
  config,
  ensureProjectDirs,
  validateRuntimeConfig,
  readJsonIfExists,
  deepMerge
};
