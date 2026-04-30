# Sinitek CLI Bridge

这个项目把 Excel 里的“携宁云估值”插件能力包装成命令行入口，避免在 Excel GUI 里手工点按钮。主实现统一走 `sinitek.ps1`，外层保留 `sinitek.cmd` 和 `sinitek.sh` 适配不同 shell。

注意：这是 Windows-only 工具。Excel COM 和携宁插件必须运行在 Windows 侧；Git Bash/WSL 只能作为命令入口，不能作为原生 Linux 运行环境。

## 快速开始

首次运行前需要先提供携宁云账号和密码。推荐在当前 PowerShell 会话中设置环境变量：

```powershell
$env:SINITEK_USERNAME = "your.name@domainname.com"
$env:SINITEK_PASSWORD = "your-password"
```

如果希望之后新开的终端也能自动读取，可以写入当前 Windows 用户环境变量：

```powershell
[Environment]::SetEnvironmentVariable("SINITEK_USERNAME", "your.name@domainname.com", "User")
[Environment]::SetEnvironmentVariable("SINITEK_PASSWORD", "your-password", "User")
```

一键提取股票历史数据并另存为新的估值模型：

```powershell
.\sinitek.cmd -Action produce -Stock 600519
```

上面的命令会打开默认模型模板，按配置中的账号和参数从携宁云提取 `600519` 的历史数据，更新模型、生成 output sheet，并自动另存为新的模型文件。`-Stock` 可替换为目标股票代码，命令成功时会回显 `Artifact=<xlsx路径>`，便于找到生成的文件。

生成 output sheet 时，CLI 会直接从用户名环境变量的邮箱域名提取输出表后缀。例如 `your.name@domainname.com` 会使用 `@domainname.com`，不需要在 YAML 中单独配置。

敲入命令到生成文件大约需要50s。

## 运行环境

- Windows + Microsoft Excel，支持 COM 自动化。
- 已安装携宁云估值 Excel 插件，默认路径为 `C:\Sinitek\SinitekExcelAddin`。
- 插件目录中需要存在 `SinitekExcel.dll` 和 `Newtonsoft.Json.dll`。
- 当前目录中需要有估值模型模板，例如 `Sinitek_Model_Ashare_V12.xlsx`。
- 推荐在 cmd、Windows PowerShell 5.1 或 PowerShell 7 中运行。PowerShell 7 会自动转发到 Windows PowerShell 5.1。
- 不支持在原生 Linux 中运行 Excel COM。`sinitek.sh` 只是从 bash 调用 Windows 的 `powershell.exe`。

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
| Shell | `cmd.exe` | 可选。使用 `sinitek.cmd` 时需要。 |
| Shell | Git Bash / WSL bash | 可选。使用 `sinitek.sh` 时需要，但仍调用 Windows `powershell.exe`。 |

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
| `sinitek.cmd` | Windows 推荐入口，cmd、PowerShell、pwsh 都可以调用。 |
| `sinitek.sh` | Git Bash/WSL 风格入口，内部仍调用 Windows 的 `powershell.exe`。 |
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
.\sinitek.cmd -Action login -Username "your.name@domainname.com" -Password "your-password"
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
  company_management_type: "2"
  peer_stock: ""
  timeout_seconds: 300
```

| 字段 | 含义 |
| --- | --- |
| `workbook` | 默认打开的估值模型 xlsx。 |
| `output_dir` | 变更类操作没有指定 `-OutWorkbook` 时，自动保存到这个目录。 |
| `auth.username_env` | 用户名所在环境变量名；生成 output sheet 时会从该用户名的邮箱域名自动提取输出表后缀。 |
| `auth.password_env` | 密码所在环境变量名。 |
| `defaults` | 命令行未显式传参时使用的本机默认配置。旧版 `fallback` 仍兼容，但新配置推荐使用 `defaults`。 |
| `history_year` | 更新模型时的历史年数。CLI 会同步写入“目录”页 `D3`。作为默认值，仅在命令行未传 `-HistoryYear` 时生效。 |
| `forecast_year` | 更新模型时的预测年数。CLI 会同步写入“目录”页 `D4` 并由模型公式更新 `D7`。作为默认值，仅在命令行未传 `-ForecastYear` 时生效。 |
| `currency_unit` | 货币单位倍率。CLI 会同步写入“目录”页的单位单元格，使报表显示和公式缩放生效。当前示例默认 `0.000001`，即百万元。 |
| `update_directory` | `update` 是否更新目录。 |
| `update_src_data` | `update` 是否更新财务源数据、附注、经营数据、可比分析等。 |
| `migrate` | 是否执行历史数据迁移。 |
| `add_output` | 更新完成后是否顺带生成 output sheet。 |
| `company_management_type` | 公司经营数据口径，当前默认 `"2"`。 |
| `peer_stock` | 可比公司 GSDM 列表，逗号分隔；留空时复刻插件逻辑自动选择。 |
| `timeout_seconds` | CLI 单次执行的总超时秒数，默认 300；设为 `0` 可关闭总超时监督。 |

YAML 解析器只支持当前这种简单结构：顶层字段和一层嵌套字段，缩进使用两个空格。

### 货币单位

`currency_unit` 支持以下值。CLI 会按插件原“货币单位”下拉框的逻辑，同步写入“目录”页 `D2`、`D5`、`D6`，因此后续报表中引用 `目录!D5` / `目录!D6` 的显示单位和缩放公式会一起变化。

| `currency_unit` | 货币单位 | 股数单位 |
| --- | --- | --- |
| `1` | 元 | 股 |
| `0.001` | 千元 | 千股 |
| `0.0001` | 万元 | 万股 |
| `0.000001` | 百万元 | 百万股 |
| `0.00000001` | 亿元 | 亿股 |

也可以通过命令行临时覆盖：

```powershell
.\sinitek.cmd -Action produce -Stock 600519 -CurrencyUnit 1
```

## 常用命令

检查当前模型和插件状态：

```powershell
.\sinitek.cmd -Action inspect
```

校验账号密码：

```powershell
.\sinitek.cmd -Action login
```

导出 output sheet：

```powershell
.\sinitek.cmd -Action output -Stock 600519 -OutWorkbook .\output\maotai-output.xlsx
```

`output` 会在导出前先按 `currency_unit` / `-CurrencyUnit` 同步报表单位。

直接更新模型并保存副本：

```powershell
.\sinitek.cmd -Action update -Stock 600519 -OutWorkbook .\output\maotai-updated.xlsx
```

需要临时切换报表单位时，可以直接传 `-CurrencyUnit`：

```powershell
.\sinitek.cmd -Action update -Stock 600519 -CurrencyUnit 0.0001
```

可比公司默认按原插件逻辑处理：如果当前 workbook 已经是同一主公司，继续复用模型现有 `PeerStock`；切换主公司时，用主公司的 `Gsdm` 请求携宁云 `/api/company/analysis/gsdms`，把返回的推荐可比公司 `gsdm` 写入 `PeerStock`。需要手工指定时传逗号分隔的 GSDM：

```powershell
.\sinitek.cmd -Action update -Stock 600519 -PeerStock "000858.SZ,000568.SZ,600809.SH"
```

一键输出最终产物（登录、更新数据、生成 output sheet、保存副本）：

```powershell
.\sinitek.cmd -Action produce -Stock 600519
```

`produce` 会强制包含 output sheet，并会按 `currency_unit` / `-CurrencyUnit` 同步报表单位；输出路径仍遵循 `-OutWorkbook`、`output_dir` 或 `-Save` 的保存规则。
命令成功时会回显 `Artifact=<xlsx路径>`，便于后续脚本直接读取产物位置。

CLI 默认有 300 秒总超时，覆盖整个 PowerShell、Excel COM 和插件调用链。需要放宽时可以传 `-TimeoutSeconds`：

```powershell
.\sinitek.cmd -Action produce -Stock 600519 -TimeoutSeconds 600
```

传 `-TimeoutSeconds 0` 可关闭总超时监督。

日常推荐在 `sinitek.yaml` 中配置 `output_dir`，然后让 CLI 自动生成输出文件名，避免覆盖历史结果：

```powershell
.\sinitek.cmd -Action update -Stock 600519
```

只有需要固定文件名时，再显式传 `-OutWorkbook`。

强制保存回原模板：

```powershell
.\sinitek.cmd -Action update -Stock 600519 -Save
```

谨慎使用 `-Save`，它会写回原 workbook。

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

这些 action 必须满足以下条件之一：

- 传 `-OutWorkbook <path>` 保存到指定文件。
- 配置或传入 `-OutputDir <dir>` 自动生成输出文件，命名格式为 `<原始Workbook文件名>-<Action>-<Stock>-<yyyyMMdd-HHmmss>.xlsx`，其中 `Action` 统一首字母大写，`Stock` 为纯数字股票代码。
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

## Bash/WSL 入口

`sinitek.sh` 只是 bash 包装器，最终仍调用 Windows 的 `powershell.exe`：

```bash
./sinitek.sh -Action inspect
```

调用链是：

```text
bash -> powershell.exe -> Excel COM -> 携宁云估值插件
```

因此它不代表原生 WSL/Linux 支持。即使从 WSL 发起命令，也必须满足：

- Windows 侧能调用 `powershell.exe`。
- Windows 侧已安装 Microsoft Excel。
- Windows 侧已安装携宁云估值插件。
- 工作簿路径能被 Windows PowerShell 正确访问。

如果主要在 Windows PowerShell 中使用，推荐直接用：

```powershell
.\sinitek.cmd -Action update -Stock 600519
```

## 实现说明

CLI 会在当前进程内绕开插件对 `Microsoft.Office.Core.DocumentProperties` 的直接强转，改用 late-binding 读写 workbook 自定义属性。这是为了解决隐藏 Excel COM 自动化场景下的 `E_NOINTERFACE` 问题；不会修改 `C:\Sinitek\SinitekExcelAddin` 中的插件 DLL。

货币单位同步逻辑复刻了插件中“货币单位”下拉框的行为：根据 `currency_unit` 更新“目录”页 `D2`、`D5`、`D6`，并同时写入 `CurrencyUnit` 自定义属性和插件运行时参数。
历史年数和预测年数也会同步到“目录”页 `D3`、`D4`，避免只写自定义属性导致报表仍沿用模板内置年数。
