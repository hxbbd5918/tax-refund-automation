# 出口退税申报自动化工具

基于 Playwright 驱动 Chrome,从 Excel 读取报关单数据,自动在上海电子税务局
"逐项配单"页面完成 **搜索 → 配单 → 选发票 → 核对 → 保存** 全流程。

> 设计原则:稳定优先,异常即停,保留浏览器现场供人工接管。

---

## 功能特性

- **登录态复用**:首次扫码登录后,Chrome 用户目录持久化,后续无需重复登录
- **断点续跑**:运行状态写入 `state/tax_refund_state.json`,中断后可继续
- **三种启动模式**(`启动.bat`):
  - `[C]` 继续处理 —— 自动跳过已完成行
  - `[N]` 重新开始 —— 清空状态从头跑
  - `[R]` 从指定行 —— 手动指定从第 N 行开始
- **严格匹配**:报关单号、商品名称、项号、成交单位、成交数量全一致才点"配单"
- **自动修正**:仅当出口数量/计税金额不一致时,自动调整最大行使合计匹配;其他字段差异直接停止
- **致命错误识别**:登录失效/浏览器关闭/导航失败 → 立即暂停并等待人工处理

---

## 技术栈

- **Node.js** + CommonJS
- **Playwright** — 驱动本地 Chrome(CDP 连接,端口 9222)
- **xlsx** — 读取 Excel
- **Electron**(可选)— `npm start` 启动 UI 入口
- 本地 JSON 状态文件 + 日志文件

---

## 环境要求

- Node.js ≥ 18
- Google Chrome(默认路径自动探测,也可通过 `TAX_CHROME_PATH` 环境变量指定)
- Windows(`启动.bat` 仅 Windows;脚本本身跨平台)

---

## 安装

```bash
git clone https://github.com/hxbbd5918/tax-refund-automation.git
cd tax-refund-automation
npm install
```

可选:复制示例配置覆盖默认行为

```bash
cp local.config.json.example local.config.json
```

`local.config.json` 字段会与 `config.js` 内默认值做深合并,常见用法:

```json
{
  "input": { "sheetName": "Sheet1" },
  "browser": { "chromePath": "D:/Chrome/chrome.exe" }
}
```

---

## 使用

### 1. 准备 Excel

把 **唯一一个** `.xlsx` 文件放入项目下的 `待处理/` 目录(运行时自动创建)。

Excel 表头须包含:

| 列名 | 说明 |
| --- | --- |
| 报关单号 | 必填 |
| 报关单号项号(或"项号") | 必填 |
| 货物品名(或"商品名称"/"品名") | 必填 |
| 成交单位(或"发票单位"/"单位") | 必填 |
| 成交数量(或"出口/进货数量"/"数量") | 必填 |
| 发票号码 | 必填 |
| 发票号码行号(或"发票行号") | 必填 |
| 计税金额 | 必填 |
| 备注 | 可选,数字会被解析为额外发票行号 |

### 2. 启动

**双击** `启动.bat`,选择运行模式即可。

或在命令行手动启动:

```bash
# 继续处理(默认从未完成行接力)
node tax_refund_matcher.js --run

# 从指定行开始
node tax_refund_matcher.js --run --row=5

# 仅探查页面结构,不写入
node tax_refund_matcher.js --inspect
```

### 3. 首次扫码

脚本启动后会自动打开税务局首页,**人工完成扫码登录**,然后回车继续。
登录态保存在 `.chrome-profile-tax/` 目录,之后启动会自动登录。

---

## 项目结构

```text
tax-refund-automation/
├─ config.js                     # 配置加载(默认值 + local.config.json 深合并)
├─ tax_refund_matcher.js         # 主流程(打开浏览器 / 配单 / 选发票 / 保存)
├─ tax_refund_excel.js           # Excel 读写、状态校验、备注解析
├─ gen_bat.js                    # 重新生成启动.bat(GBK 编码)
├─ 启动.bat                       # Windows 一键启动菜单(C/N/R)
├─ local.config.json.example     # 配置模板
├─ 待处理/                        # 放待处理的 .xlsx(运行时自动创建)
├─ 已处理/                        # 处理完成的 .xlsx 自动迁入
├─ logs/                          # 运行日志
├─ state/tax_refund_state.json   # 运行状态(断点续跑依据)
└─ .chrome-profile-tax/          # Chrome 用户目录(登录态)
```

---

## 环境变量

| 变量 | 说明 |
| --- | --- |
| `TAX_ROOT_DIR` | 覆盖根目录(打包后场景使用) |
| `TAX_CHROME_PATH` | 自定义 Chrome 路径 |
| `TAX_USER_DATA_DIR` | 自定义 Chrome 用户目录 |
| `TAX_REMOTE_DEBUGGING_PORT` | CDP 端口(默认 9222) |
| `TAX_EXCEL_SHEET` | 指定 Excel 工作表名 |

---

## 异常处理

异常发生时脚本会:

1. **停止执行**,不再处理下一行
2. **保留浏览器窗口**,方便人工核对/接管
3. 在 `state/tax_refund_state.json` 写入失败原因和行号
4. 在 `logs/` 留下完整日志和截图

修复后,用 `--row=N` 从问题行继续即可。

---

## 注意事项

- 本工具仅自动化重复操作,**不替代人工审核**,建议每批先小样本验证
- `.chrome-profile-tax/`、`待处理/`、`已处理/`、`state/`、`logs/` 都被 `.gitignore` 排除,**不会被提交**
- Excel 业务数据敏感,请勿提交到公共仓库

---

## 许可

私有项目,仅作内部使用。
