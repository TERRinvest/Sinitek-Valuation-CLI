# Sinitek CLI Bridge

这个项目把 Excel 里的"携宁云估值"插件能力包装成命令行入口，避免在 Excel GUI 里手工点按钮。主入口是 `sinitek.ps1`，直接在 Windows PowerShell 中运行。

注意：这是 Windows-only 工具。Excel COM 和携宁插件必须运行在 Windows 侧。

## 快速开始

### 1. 放行 PowerShell 执行策略（一次性）

Windows 默认禁止运行 `.ps1` 脚本。在 PowerShell 中执行：

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

### 2. 配置账号密码

推荐在当前 PowerShell 会话中设置环境变量：

```powershell
$env:SINITEK_USERNAME = "your.name@domainname.com"
$env:SINITEK_PASSWORD = "your-password"
```

如果希望之后新开的终端也能自动读取，可以写入当前 Windows 用户环境变量：

```powershell
[Environment]::SetEnvironmentVariable("SINITEK_USERNAME", "your.name@domainname.com", "User")
[Environment]::SetEnvironmentVariable("SINITEK_PASSWORD", "your-password", "User")
```

### 3. 运行

一键提取股票历史数据并另存为新的估值模型：

```powershell
.\sinitek.ps1 -Action produce -Stock 600519
```

上面的命令会打开默认模型模板，按配置中的账号和参数从携宁云提取 `600519` 的历史数据，更新模型、生成 output sheet，并自动另存为新的模型文件。`-Stock` 可替换为目标股票代码，命令成功时会回显 `Artifact=<xlsx路径>`，便于找到生成的文件。

生成 output sheet 时，CLI 会直接从用户名环境变量的邮箱域名提取输出表后缀。例如 `your.name@domainname.com` 会使用 `@domainname.com`，不需要在 YAML 中单独配置。

敲入命令到生成文件大约需要50s。

### 免执行策略快捷方式

如果不想执行 `Set-ExecutionPolicy`，也可以通过 `sinitek.cmd` 调用（内部带 `-ExecutionPolicy Bypass`）：

```cmd
sinitek.cmd -Action produce -Stock 600519
```

或在 PowerShell 命令中显式指定：

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\sinitek.ps1 -Action produce -Stock 600519
```

> **注意**：不建议通过 `cmd /c` 调用 `sinitek.cmd` 传递含引号的参数（如 `-Workbook`），cmd.exe 的嵌套引号解析有固有限制。直接在 PowerShell 中调用 `sinitek.ps1` 可避免此问题。

## 运行环境

- Windows + Microsoft Excel，支持 COM 自动化。
- 已安装携宁云估值 Excel 插件，默认路径为 `C:\Sinitek\SinitekExcelAddin`。
- 插件目录中需要存在 `SinitekExcel.dll` 和 `Newtonsoft.Json.dll`。
- 当前目录中需要有估值模型模板，例如 `Sinitek_Model_Ashare_V12.xlsx`。
- 推荐在 Windows PowerShell 5.1 或 PowerShell 7 中运行。PowerShell 7 会自动转发到 Windows PowerShell 5.1。

## 依赖项

| 类别 | 依赖 | 说明 |
| --- | --- | --- |
| 操作系统 | Windows 桌面环境 | 必需。需要能启动本机 Excel COM，不支持原生 Linux/headless 环境。 |
| Office | Microsoft Excel | 必需。CLI 会通过 COM 自动化打开和保存 xlsx。 |
| PowerShell | Windows PowerShell 5.1 | 必需。`sinitek.ps1` 最终在 `powershell.exe` 下运行。 |
| PowerShell | PowerShell 7 / `pwsh` | 可选。入口会自动转发到 Windows PowerShell 5.1。 |
| 携宁插件 | `C:\Sinitek\SinitekExcelAddin\SinitekExcel.dll` | 必需。核心插件 DLL。 |
| 携宁插件 | `C:\Sinitek\SinitekExcelAddin\Newtonsoft.Json.dll` | 必需。随携宁插件安装的 JSON 依赖。 |
| Office Interop | `office.dll` | 必需。脚本会从常见 Office/GAC 路径自动查找。 |
| Office Interop | `Microsoft.Office.Interop.Excel.dll` | 必需。脚本会从常见 Office 安装路径自动查找。 |
| COM Add-in | `Extensibility.dll` | 必需。用于调用插件的 `OnConnection` 初始化流程。 |
| 模型文件 | 携宁云估值 xlsx 模板 | 必需。需要包含 `ModelVersion`、`ModelType` 等自定义文档属性。 |
| 账号 | 携宁云账号和密码 | 必需。建议通过 `SINITEK_USERNAME`、`SINITEK_PASSWORD` 环境变量提供。 |
| 网络 | `https://cloudmodel.sinitek.com` | 必需。股票搜索、权限校验、数据更新会访问携宁云接口。 |

不需要安装 Node.js、Python、Visual Studio 或额外包管理器。C# bridge 由 PowerShell 的 `Add-Type` 在运行时编译。

脚本当前会自动查找这些 Office/COM 依赖路径：

```text
C:\Program Files\Microsoft Office\root\vfs\ProgramFilesX86\Microsoft Office\Office16\DCF\office.dll
C:\Program Files\Microsoft Office\root\Office16\ADDINS\PowerPivot Excel Add-in\OFFICE.dll
C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\OFFICE.DLL

C:\Program Files\Microsoft Office\root\vfs\ProgramFilesX86\Microsoft Office\Office16\DCF\Microsoft.Office.Interop.Excel.dll
C:\Program Files\Microsoft Office\root\Office16\ADDINS\PowerPivot Excel Add-in\Microsoft.Office.Interop.Excel.dll
C:\Program Files\Microsoft Office\root\Office16\ADDINS\Microsoft Power Query for Excel Integrated\bin\Microsoft.Office.Interop.Excel.dll

C:\Windows\assembly\GAC\Extensibility\7.0.3300.0__b03f5f7f11d50a3a\extensibility.dll
```

如果这些路径都不存在，CLI 会报 `Required file not found`，需要先修复 Office/插件安装环境，或在 `sinitek.ps1` 中补充本机实际路径。

## 文件说明

| 文件 | 作用 |
| --- | --- |
| `sinitek.ps1` | 主入口，读取 YAML、加载插件 DLL、调用 C# bridge。 |
| `SinitekCliBridge.cs` | C# 桥接层，负责 Excel COM、插件调用、股票搜索和文件保存。 |
| `sinitek.cmd` | 免执行策略快捷方式，内部调用 `sinitek.ps1` 并带 `-ExecutionPolicy Bypass`。 |
| `sinitek.yaml` | 本机默认配置。 |
| `sinitek.yaml.example` | 配置样例，可复制后修改。 |
| `Sinitek_Model_Ashare_V12.xlsx` | 当前估值模型模板。 |

## 初始化配置

复制配置样例：

```powershell
Copy-Item .\sinitek.yaml.example .\sinitek.yaml
```

配置账号密码。YAML 里不直接保存密码，只保存环境变量名：

```yaml
auth:
  username_env: SINITEK_USERNAME
  password_env: SINITEK_PASSWORD
```

在当前 PowerShell 会话中设置：

```powershell
$env:SINITEK_USERNAME = "your.name@domainname.com"
$env:SINITEK_PASSWORD = "your-password"
```

也可以写入当前 Windows 用户环境变量，之后新开的终端会自动读取：

```powershell
[Environment]::SetEnvironmentVariable("SINITEK_USERNAME", "your.name@domainname.com", "User")
[Environment]::SetEnvironmentVariable("SINITEK_PASSWORD", "your-password", "User")
```

如果不想设置环境变量，也可以临时传参：

```powershell
.\sinitek.ps1 -Action login -Username "your.name@domainname.com" -Password "your-password"
```

注意：CLI 每次执行都是新进程，`login` action 主要用于校验账号密码，不等于后续命令已经登录。实际执行需要每次能从环境变量或命令参数拿到账号密码。

## YAML 字段

```yaml
workbook: ./Sinitek_Model_Ashare_V12.xlsx
output_dir: ./output

auth:
  username_env: SINITEK_USERNAME
  password_env: SINITEK_PASSWORD

defaults:
  history_year: 4
  forecast_year: 3
  currency_unit: "0.000001"
  update_directory: true
  update_src_data: true
  migrate: false
  add_output: false
  segment_dimension: "product"
  peer_stock: ""
  prediction_scope: "all"
  prediction_rows: ""
  prediction_indicators: ""
  prediction_method: ""
  prediction_settings: ""
  timeout_seconds: 300
```

| 字段 | 含义 |
| --- | --- |
| `workbook` | 默认打开的估值模型 xlsx。 |
| `output_dir` | 变更类操作没有指定 `-OutWorkbook` 时，自动保存到这个目录。 |
| `auth.username_env` | 用户名所在环境变量名；生成 output sheet 时会从该用户名的邮箱域名自动提取输出表后缀。 |
| `auth.password_env` | 密码所在环境变量名。 |
| `defaults` | 命令行未显式传参时使用的本机默认配置。旧版 `fallback` 仍兼容，但新配置推荐使用 `defaults`。 |
| `history_year` | 更新模型时的历史年数。CLI 会同步写入"目录"页 `D3`。作为默认值，仅在命令行未传 `-HistoryYear` 时生效。 |
| `forecast_year` | 更新模型时的预测年数。CLI 会同步写入"目录"页 `D4` 并由模型公式更新 `D7`。作为默认值，仅在命令行未传 `-ForecastYear` 时生效。 |
| `currency_unit` | 货币单位倍率。CLI 会同步写入"目录"页的单位单元格，使报表显示和公式缩放生效。当前示例默认 `0.000001`，即百万元。 |
| `update_directory` | `update` 是否更新目录。 |
| `update_src_data` | `update` 是否更新财务源数据、附注、经营数据、可比分析等。 |
| `migrate` | 是否执行历史数据迁移。 |
| `add_output` | 更新完成后是否顺带生成 output sheet。 |
| `segment_dimension` | 业务分部口径，决定"公司经营数据"sheet 按哪种维度拆分业务板块。可选值：`"industry"`、`"product"`（默认）、`"region"`。详见下方小节。 |
| `peer_stock` | 可比公司 GSDM 列表，逗号分隔；留空时复刻插件逻辑自动选择。 |
| `prediction_scope` | `predict` 的批量范围，支持 `all`、`sales`、`capital`。 |
| `prediction_rows` | `predict` 的批量行号筛选，逗号/分号分隔。 |
| `prediction_indicators` | `predict` 的批量财务指标名筛选，逗号/分号分隔。 |
| `prediction_method` | `predict` 的批量预测方法，例如 `latest1`、`avg3`、`weighted3`。 |
| `prediction_settings` | `predict` 的精确设置，格式为 `指标名=方法`、`scope:行号=方法` 或 `控件名=方法`。 |
| `timeout_seconds` | CLI 单次执行的总超时秒数，默认 300；设为 `0` 可关闭总超时监督。 |

YAML 解析器只支持当前这种简单结构：顶层字段和一层嵌套字段，缩进使用两个空格。

### 业务分部口径（segment_dimension）

`segment_dimension` 决定"公司经营数据"sheet 按哪种维度拆分业务板块。

| 可选值 | 分部维度 | 说明 |
| --- | --- | --- |
| `"industry"` | 按行业 | 分部数据写入年度列（Q/S 等奇数列），如"消费级"、"企业级"。适用于公司按行业分类披露经营数据的场景。 |
| `"product"` | 按产品 | 分部数据仅写入半年度列（L 列），如"芯片"、"算法"。适用于公司按产品线分类披露的场景。**当前默认值。** |
| `"region"` | 按地区 | 分部数据按地区维度拆分，如"华东"、"华南"、"海外"。适用于以地域为主要经营维度的公司。 |

CLI 内部会将英文值映射为插件所需的 `CompanyManagementType` 数字（industry=1, product=2, region=3）和 `CompanyManagementName` 中文字符串（按行业/按产品/按地区）。

命令行覆盖：

```powershell
.\sinitek.ps1 -Action produce -Stock 688343 -SegmentDimension industry
```

也可以在 YAML 中配置默认值：

```yaml
defaults:
  segment_dimension: "industry"
```

**选择建议**：与目标公司年报中"报告分部"披露口径保持一致。如果年报按行业分部披露（如云天励飞），使用 `industry`；如果按产品分部披露，使用 `product`；如果按地区分部披露，使用 `region`。

### 货币单位

`currency_unit` 支持以下值。CLI 会按插件原"货币单位"下拉框的逻辑，同步写入"目录"页 `D2`、`D5`、`D6`，因此后续报表中引用 `目录!D5` / `目录!D6` 的显示单位和缩放公式会一起变化。

| `currency_unit` | 货币单位 | 股数单位 |
| --- | --- | --- |
| `1` | 元 | 股 |
| `0.001` | 千元 | 千股 |
| `0.0001` | 万元 | 万股 |
| `0.000001` | 百万元 | 百万股 |
| `0.00000001` | 亿元 | 亿股 |

也可以通过命令行临时覆盖：

```powershell
.\sinitek.ps1 -Action produce -Stock 600519 -CurrencyUnit 1
```

## 命令

所有命令都通过 `-Action` 选择动作。当前公开 action 如下：

| Action | 作用 | 修改 workbook | 常用参数 |
| --- | --- | --- | --- |
| `inspect` | 检查模型、插件和登录状态 | 否 | `-Workbook` |
| `login` | 校验携宁云账号密码 | 否 | `-Username`、`-Password` |
| `output` | 生成 output sheet 并保存 | 是 | `-Stock`、`-CurrencyUnit`、`-OutWorkbook` |
| `update` | 更新模型历史数据和相关表 | 是 | `-Stock`、`-HistoryYear`、`-ForecastYear`、`-PeerStock` |
| `produce` | 一键更新并生成 output sheet | 是 | `-Stock`、`-CurrencyUnit`、`-TimeoutSeconds` |
| `predict` | 执行"预测数据设置" | 是 | `-PredictionMethod`、`-PredictionIndicators`、`-PredictionSettings` |

### inspect

检查当前 workbook、插件版本、模型版本、登录状态和关键自定义属性：

```powershell
.\sinitek.ps1 -Action inspect
```

### login

校验账号密码是否可用：

```powershell
.\sinitek.ps1 -Action login
```

也可以临时传入账号密码：

```powershell
.\sinitek.ps1 -Action login -Username "your.name@domainname.com" -Password "your-password"
```

`login` 只用于校验；CLI 每次执行都是新进程，后续 `update`、`output`、`produce` 仍需要能从环境变量或命令参数读取账号密码。

### output

导出 output sheet 并保存为副本：

```powershell
.\sinitek.ps1 -Action output -Stock 600519 -OutWorkbook .\output\maotai-output.xlsx
```

`output` 会在导出前先按 `currency_unit` / `-CurrencyUnit` 同步报表单位。需要临时切换单位时直接传参：

```powershell
.\sinitek.ps1 -Action output -Stock 600519 -CurrencyUnit 0.0001
```

### update

更新模型历史数据、附注、经营数据、可比分析等，并保存副本：

```powershell
.\sinitek.ps1 -Action update -Stock 600519 -OutWorkbook .\output\maotai-updated.xlsx
```

需要临时调整历史年数、预测年数或货币单位时，可以直接覆盖：

```powershell
.\sinitek.ps1 -Action update -Stock 600519 -HistoryYear 5 -ForecastYear 4 -CurrencyUnit 0.0001
```

可比公司默认按原插件逻辑处理：如果当前 workbook 已经是同一主公司，继续复用模型现有 `PeerStock`；切换主公司时，用主公司的 `Gsdm` 请求携宁云 `/api/company/analysis/gsdms`，把返回的推荐可比公司 `gsdm` 写入 `PeerStock`。需要手工指定时传逗号分隔的 GSDM：

```powershell
.\sinitek.ps1 -Action update -Stock 600519 -PeerStock "000858.SZ,000568.SZ,600809.SH"
```

日常推荐在 `sinitek.yaml` 中配置 `output_dir`，然后让 CLI 自动生成输出文件名，避免覆盖历史结果：

```powershell
.\sinitek.ps1 -Action update -Stock 600519
```

### produce

一键输出最终产物：登录校验、更新数据、生成 output sheet、保存副本。

```powershell
.\sinitek.ps1 -Action produce -Stock 600519
```

`produce` 会强制包含 output sheet，并会按 `currency_unit` / `-CurrencyUnit` 同步报表单位；输出路径遵循 `-OutWorkbook`、`output_dir` 或 `-Save` 的保存规则。命令成功时会回显 `Artifact=<xlsx路径>`，便于后续脚本直接读取产物位置。

CLI 默认有 300 秒总超时，覆盖 PowerShell、Excel COM 和插件调用链。需要放宽时可以传 `-TimeoutSeconds`：

```powershell
.\sinitek.ps1 -Action produce -Stock 600519 -TimeoutSeconds 600
```

传 `-TimeoutSeconds 0` 可关闭总超时监督。

### predict

`predict` 对应插件 `btnSet` / "预测数据设置"功能。CLI 会打开模型，写入预测设置自定义属性，并调用插件预测设置表单的公式写入逻辑，不弹 GUI。

按范围批量设置：

```powershell
.\sinitek.ps1 -Action predict -PredictionScope sales -PredictionMethod avg3 -OutWorkbook .\output\maotai-predict.xlsx
```

按财务指标名批量设置：

```powershell
.\sinitek.ps1 -Action predict -PredictionIndicators "研发费用率,所得税税率,应收账款周转天数" -PredictionMethod weighted3 -OutWorkbook .\output\maotai-predict.xlsx
```

精确到单个指标、每个指标单独指定方法：

```powershell
.\sinitek.ps1 -Action predict -PredictionSettings "研发费用率=avg3,所得税税率=weighted2,应收账款周转天数=zero" -OutWorkbook .\output\maotai-predict.xlsx
```

底层控件名仍可直接调用：

```powershell
.\sinitek.ps1 -Action predict -PredictionSettings "ASalescmb_31=avg3,ACapitalcmb_15=weighted2"
```

`-PredictionMethod` 支持 `latest1`、`avg2`、`avg3`、`weighted2`、`weighted3`、`custom`、`zero`，也支持插件下拉框索引 `0`-`6`。`-PredictionScope` 支持 `all`、`sales`、`capital`；当前 A 股模板的预测设置表单实际暴露 `sales` 和 `capital` 两组指标。`-PredictionRows` 和 `-PredictionIndicators` 都为空时，表示所选范围内全部指标。

| 参数 | 用法 |
| --- | --- |
| `-PredictionScope` | 批量范围：`all`、`sales`、`capital`。 |
| `-PredictionRows` | 按 Excel 行号筛选，逗号/分号分隔，例如 `22,31,48`。 |
| `-PredictionIndicators` | 按财务指标名筛选，逗号/分号分隔，例如 `研发费用率,所得税税率`。 |
| `-PredictionSettings` | 精确设置，格式为 `指标名=方法`、`scope:行号=方法` 或 `控件名=方法`。 |

当前 A 股模板映射如下，行号来自插件预测设置表单实际绑定的 Excel 行：

| Scope | Sheet | 行号 | 控件名 | 推荐指标名 | 可用别名示例 |
| --- | --- | ---: | --- | --- | --- |
| `sales` | 销售预测 | 22 | `ASalescmb_22` | 税金及附加/营业收入 | 税金及附加率 |
| `sales` | 销售预测 | 25 | `ASalescmb_25` | 销售费用/营业收入 | 销售费用率 |
| `sales` | 销售预测 | 28 | `ASalescmb_28` | 管理费用/营业收入 | 管理费用率 |
| `sales` | 销售预测 | 31 | `ASalescmb_31` | 研发费用/营业收入 | 研发费用率 |
| `sales` | 销售预测 | 34 | `ASalescmb_34` | 其他收益/营业收入 | 其他收益率 |
| `sales` | 销售预测 | 37 | `ASalescmb_37` | 投资收益/营业收入 | 投资收益率 |
| `sales` | 销售预测 | 40 | `ASalescmb_40` | 资产处置收益/营业收入 | 资产处置收益率 |
| `sales` | 销售预测 | 42 | `ASalescmb_42` | 营业外收入 | 营业外收入 |
| `sales` | 销售预测 | 43 | `ASalescmb_43` | 营业外支出 | 营业外支出 |
| `sales` | 销售预测 | 48 | `ASalescmb_48` | 所得税税率 | 所得税率 |
| `sales` | 销售预测 | 51 | `ASalescmb_51` | 少数股东损益/净利润 | 少数股东损益率 |
| `capital` | 资产预测 | 11 | `ACapitalcmb_11` | 应收票据周转天数 | 应收票据/营业收入 |
| `capital` | 资产预测 | 15 | `ACapitalcmb_15` | 应收账款周转天数 | 应收账款/营业收入 |
| `capital` | 资产预测 | 19 | `ACapitalcmb_19` | 应收款项融资/营业收入 | 应收款项融资占收入比 |
| `capital` | 资产预测 | 22 | `ACapitalcmb_22` | 预付账款/营业成本 | 预付账款占成本比 |
| `capital` | 资产预测 | 25 | `ACapitalcmb_25` | 其他应收款周转天数 | 其他应收款/营业收入 |
| `capital` | 资产预测 | 29 | `ACapitalcmb_29` | 存货周转天数 | 存货/营业成本 |
| `capital` | 资产预测 | 33 | `ACapitalcmb_33` | 合同资产/营业收入 | 合同资产占收入比 |
| `capital` | 资产预测 | 39 | `ACapitalcmb_39` | 应付票据周转天数 | 应付票据/营业成本 |
| `capital` | 资产预测 | 43 | `ACapitalcmb_43` | 应付账款周转天数 | 应付账款/营业成本 |
| `capital` | 资产预测 | 47 | `ACapitalcmb_47` | 预收账款/营业收入 | 预收账款占收入比 |
| `capital` | 资产预测 | 50 | `ACapitalcmb_50` | 合同负债/营业收入 | 合同负债占收入比 |
| `capital` | 资产预测 | 53 | `ACapitalcmb_53` | 应付职工薪酬/营业成本 | 应付职工薪酬占成本比 |
| `capital` | 资产预测 | 56 | `ACapitalcmb_56` | 应交税费/营业收入 | 应交税费占收入比 |
| `capital` | 资产预测 | 59 | `ACapitalcmb_59` | 其他应付款/营业成本 | 其他应付款占成本比 |

命令执行后回显 `Applied=` 是底层控件名，`AppliedIndicators=` 是财务指标名，二者可用于复核设置是否命中预期行。

## 参数覆盖规则

- 命令行参数优先级高于 `sinitek.yaml`。
- 没有显式传 `-Config` 时，默认读取脚本目录下的 `sinitek.yaml`。
- 显式传相对路径时，相对当前工作目录解析。
- YAML 中的 `workbook` 和 `output_dir` 相对 YAML 文件所在目录解析。
- YAML 中的 `defaults` 是命令行参数的默认值；如果同时存在 `defaults` 和旧版 `fallback`，优先使用 `defaults`。
- `-Stock` 是股票代码入口，目前不在 YAML 中定义。
- `-Gsdm` 可手动传公司代码；不传时会通过股票搜索结果解析，必要时用 `SecurityCode` 兜底。
- `-PeerStock` 可手动覆盖可比公司，传入逗号分隔的 GSDM；不传时使用自动选择逻辑。
- `-TimeoutSeconds` 默认 300 秒，可用 YAML 的 `defaults.timeout_seconds` 或命令行覆盖；命令行显式传参优先。

## 变更类 action 的保存规则

以下 action 会修改 workbook 状态：

- `output`
- `update`
- `produce`
- `predict`

这些 action 必须满足以下条件之一：

- 传 `-OutWorkbook <path>` 保存到指定文件。
- 配置或传入 `-OutputDir <dir>` 自动生成输出文件，命名格式为 `<原始Workbook文件名>-<Action>-<Stock>-<yyyyMMdd-HHmmss>.xlsx`，其中 `Action` 统一首字母大写，`Stock` 为纯数字股票代码；`predict` 未提供股票代码时会省略 `Stock` 段。
- 传 `-Save` 写回原文件。

如果都没有提供，CLI 会拒绝执行，避免无意修改原模板。

## 故障排查

没有账号密码：

```text
ERROR: Stock search requires login token. Set SINITEK_USERNAME/SINITEK_PASSWORD, or pass -Username and -Password.
```

处理方式：设置 `SINITEK_USERNAME` 和 `SINITEK_PASSWORD`，或在命令中传 `-Username/-Password`。

云端接口超时：

```text
ERROR: Stock search timed out after 15 seconds: https://cloudmodel.sinitek.com/api/stock
```

处理方式：检查网络、代理、VPN 或携宁云服务状态。CLI 已设置 15 秒超时，避免无限挂起。

总执行超时：

```text
ERROR: Action 'produce' timed out after 300 seconds.
```

处理方式：先检查网络、Excel 是否弹出对话框、插件是否卡在更新过程。CLI 超时后会终止本次子 PowerShell，并按本次启动的 Excel PID 尝试清理 Excel 进程，退出码为 `124`。确实需要更久时，传 `-TimeoutSeconds 600` 或在 `sinitek.yaml` 的 `defaults.timeout_seconds` 中调整。

提示缺少参数：

```text
ERROR: Stock search response did not contain a stock list. Server message: 缺少参数 ...
```

通常表示 token、模型版本或模型类型没有传到云端。确认账号密码可用，并且 `workbook` 指向有效的携宁云估值模板。

检查是否残留 Excel 进程：

```powershell
Get-Process EXCEL -ErrorAction SilentlyContinue
```

正常情况下 CLI 会关闭自己启动的 Excel 进程。

## 实现说明

CLI 会在当前进程内绕开插件对 `Microsoft.Office.Core.DocumentProperties` 的直接强转，改用 late-binding 读写 workbook 自定义属性。这是为了解决隐藏 Excel COM 自动化场景下的 `E_NOINTERFACE` 问题；不会修改 `C:\Sinitek\SinitekExcelAddin` 中的插件 DLL。

货币单位同步逻辑复刻了插件中"货币单位"下拉框的行为：根据 `currency_unit` 更新"目录"页 `D2`、`D5`、`D6`，并同时写入 `CurrencyUnit` 自定义属性和插件运行时参数。
历史年数和预测年数也会同步到"目录"页 `D3`、`D4`，避免只写自定义属性导致报表仍沿用模板内置年数。