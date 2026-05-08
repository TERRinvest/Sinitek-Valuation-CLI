param(
    [ValidateSet('inspect', 'login', 'output', 'update', 'produce', 'predict')]
    [string]$Action = 'inspect',

    [string]$Config = '.\sinitek.yaml',
    [string]$Workbook = '.\Sinitek_Model_Ashare_V12.xlsx',
    [string]$OutWorkbook = '',
    [string]$OutputDir = '',
    [switch]$Save,
    [switch]$Visible,

    [string]$Username = $env:SINITEK_USERNAME,
    [string]$Password = $env:SINITEK_PASSWORD,

    [string]$Stock = '',
    [string]$Gsdm = '',
    [string]$StockName = '',
    [int]$HistoryYear = 5,
    [int]$ForecastYear = 5,
    [string]$CurrencyUnit = '0.000001',
    [string]$CompanyManagementType = '2',
    [string]$CompanyManagementName = '',
    [string]$PeerStock = '',
    [string]$PredictionScope = 'all',
    [string]$PredictionRows = '',
    [string]$PredictionIndicators = '',
    [string]$PredictionMethod = '',
    [string]$PredictionSettings = '',
    [bool]$UpdateDirectory = $true,
    [bool]$UpdateSrcData = $true,
    [bool]$Migrate = $false,
    [bool]$AddOutput = $false,
    [ValidateRange(0, 86400)]
    [int]$TimeoutSeconds = 300,

    [switch]$NoTimeoutSupervisor,
    [string]$ExcelPidFile = ''
)

$ErrorActionPreference = 'Stop'
$Utf8NoBom = New-Object System.Text.UTF8Encoding $false
[Console]::InputEncoding = $Utf8NoBom
[Console]::OutputEncoding = $Utf8NoBom
$OutputEncoding = $Utf8NoBom

if ($PSVersionTable.PSEdition -eq 'Core') {
    $ForwardArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
    foreach ($Entry in $PSBoundParameters.GetEnumerator()) {
        $Name = '-' + $Entry.Key
        $Value = $Entry.Value
        if ($Value -is [System.Management.Automation.SwitchParameter]) {
            if ($Value.IsPresent) {
                $ForwardArgs += $Name
            }
        }
        elseif ($Value -is [bool]) {
            $ForwardArgs += ($Name + ':' + ($(if ($Value) { '$true' } else { '$false' })))
        }
        else {
            $ForwardArgs += $Name
            $ForwardArgs += [string]$Value
        }
    }
    & powershell.exe @ForwardArgs
    exit $LASTEXITCODE
}

$ExplicitParams = @{}
$InitialBoundParameters = @{}
foreach ($Key in $PSBoundParameters.Keys) {
    $ExplicitParams[$Key] = $true
    $InitialBoundParameters[$Key] = $PSBoundParameters[$Key]
}

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$BridgePath = Join-Path $Root 'SinitekCliBridge.cs'
$AddinDir = 'C:\Sinitek\SinitekExcelAddin'
$SinitekDll = Join-Path $AddinDir 'SinitekExcel.dll'
$NewtonsoftDll = Join-Path $AddinDir 'Newtonsoft.Json.dll'

function First-ExistingPath {
    param([string[]]$Paths)
    foreach ($Path in $Paths) {
        if (Test-Path -LiteralPath $Path) {
            return $Path
        }
    }
    throw "None of these paths exist: $($Paths -join ', ')"
}

function Test-ExplicitParam {
    param([string]$Name)
    return $ExplicitParams.ContainsKey($Name)
}

function Resolve-OptionalPath {
    param(
        [string]$Path,
        [string]$BasePath
    )
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return ''
    }
    if ([IO.Path]::IsPathRooted($Path)) {
        return [IO.Path]::GetFullPath($Path)
    }
    return [IO.Path]::GetFullPath((Join-Path $BasePath $Path))
}

function Format-ActionForFileName {
    param([string]$ActionName)
    if ([string]::IsNullOrWhiteSpace($ActionName)) {
        return 'Action'
    }
    $LowerAction = $ActionName.Trim().ToLowerInvariant()
    return [Globalization.CultureInfo]::InvariantCulture.TextInfo.ToTitleCase($LowerAction)
}

function Normalize-StockCodeForFileName {
    param([string]$StockCode)
    if ([string]::IsNullOrWhiteSpace($StockCode)) {
        return ''
    }
    $Match = [regex]::Match($StockCode.Trim(), '\d+')
    if ($Match.Success) {
        return $Match.Value
    }
    return ''
}

function Get-XlsxCustomProperty {
    param(
        [string]$Path,
        [string]$Name
    )
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        return ''
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    $Zip = $null
    $Reader = $null
    try {
        $Zip = [IO.Compression.ZipFile]::OpenRead($Path)
        $Entry = $Zip.GetEntry('docProps/custom.xml')
        if (-not $Entry) {
            return ''
        }

        $Reader = [IO.StreamReader]::new($Entry.Open(), [Text.Encoding]::UTF8)
        [xml]$Xml = $Reader.ReadToEnd()
        $NamespaceManager = New-Object Xml.XmlNamespaceManager($Xml.NameTable)
        $NamespaceManager.AddNamespace('cp', 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties')
        $Node = $Xml.SelectSingleNode("//cp:property[@name='$Name']", $NamespaceManager)
        if ($Node -and $Node.FirstChild) {
            return [string]$Node.FirstChild.InnerText
        }
    }
    catch {
        return ''
    }
    finally {
        if ($Reader) {
            $Reader.Dispose()
        }
        if ($Zip) {
            $Zip.Dispose()
        }
    }
    return ''
}

function Get-StockCodeForFileName {
    param(
        [string]$StockCode,
        [string]$WorkbookPath
    )

    $NormalizedStock = Normalize-StockCodeForFileName $StockCode
    if (-not [string]::IsNullOrWhiteSpace($NormalizedStock)) {
        return $NormalizedStock
    }

    foreach ($PropertyName in @('StkCode', 'index_stkcode', 'base_stkcode', 'GSCode')) {
        $PropertyValue = Get-XlsxCustomProperty -Path $WorkbookPath -Name $PropertyName
        $NormalizedStock = Normalize-StockCodeForFileName $PropertyValue
        if (-not [string]::IsNullOrWhiteSpace($NormalizedStock)) {
            return $NormalizedStock
        }
    }

    throw "Cannot determine stock code for output filename. Pass -Stock or -OutWorkbook."
}

function ConvertFrom-SimpleYamlValue {
    param([string]$Value)
    $Text = $Value.Trim()
    if ($Text.Length -ge 2) {
        if (($Text.StartsWith('"') -and $Text.EndsWith('"')) -or ($Text.StartsWith("'") -and $Text.EndsWith("'"))) {
            return $Text.Substring(1, $Text.Length - 2)
        }
    }
    if ($Text -match '(?i)^(true|false)$') {
        return [bool]::Parse($Text)
    }
    if ($Text -match '^-?\d+$') {
        return [int]$Text
    }
    return $Text
}

function Read-SimpleYaml {
    param([string]$Path)

    $ConfigData = @{}
    $Section = $null
    foreach ($RawLine in Get-Content -LiteralPath $Path -Encoding UTF8) {
        $Line = $RawLine -replace "`t", '    '
        if ($Line.Trim().Length -eq 0 -or $Line.TrimStart().StartsWith('#')) {
            continue
        }

        $Line = $Line -replace '\s+#.*$', ''
        if ($Line -notmatch '^(\s*)([A-Za-z0-9_-]+)\s*:\s*(.*)$') {
            throw "Unsupported YAML line in ${Path}: $RawLine"
        }

        $Indent = $Matches[1].Length
        $Key = $Matches[2]
        $Value = $Matches[3]
        if ($Indent -eq 0) {
            if ([string]::IsNullOrWhiteSpace($Value)) {
                $ConfigData[$Key] = @{}
                $Section = $Key
            }
            else {
                $ConfigData[$Key] = ConvertFrom-SimpleYamlValue $Value
                $Section = $null
            }
            continue
        }

        if ($Indent -ne 2 -or [string]::IsNullOrWhiteSpace($Section)) {
            throw "Only one nested YAML level with two-space indentation is supported: $RawLine"
        }
        $ConfigData[$Section][$Key] = ConvertFrom-SimpleYamlValue $Value
    }
    return $ConfigData
}

function Set-FromConfig {
    param(
        [hashtable]$ConfigData,
        [string]$ConfigKey,
        [string]$VariableName,
        [string]$ExplicitName = $VariableName,
        [switch]$PathValue,
        [string]$PathBase = ''
    )
    if ((Test-ExplicitParam $ExplicitName) -or -not $ConfigData.ContainsKey($ConfigKey)) {
        return
    }

    $Value = $ConfigData[$ConfigKey]
    if ($PathValue) {
        $Value = Resolve-OptionalPath -Path ([string]$Value) -BasePath $PathBase
    }
    Set-Variable -Name $VariableName -Value $Value -Scope 1
}

function Set-FromDefaultConfig {
    param(
        [hashtable]$ConfigData,
        [string]$ConfigKey,
        [string]$VariableName,
        [string]$ExplicitName = $VariableName
    )
    if (Test-ExplicitParam $ExplicitName) {
        return
    }

    foreach ($Section in @('defaults', 'fallback')) {
        if ($ConfigData.ContainsKey($Section) -and $ConfigData[$Section].ContainsKey($ConfigKey)) {
            Set-Variable -Name $VariableName -Value $ConfigData[$Section][$ConfigKey] -Scope 1
            return
        }
    }
}

function Get-CliExceptionMessage {
    param([Exception]$Exception)

    while ($Exception.InnerException -and (
        $Exception -is [System.Management.Automation.MethodInvocationException] -or
        $Exception -is [System.Reflection.TargetInvocationException])) {
        $Exception = $Exception.InnerException
    }

    return $Exception.Message
}

function ConvertTo-WindowsCommandLineArgument {
    param([string]$Argument)

    if ($null -eq $Argument) {
        return '""'
    }
    if ($Argument.Length -gt 0 -and $Argument -notmatch '[\s"]') {
        return $Argument
    }

    $Result = '"'
    $Backslashes = 0
    foreach ($Character in $Argument.ToCharArray()) {
        if ($Character -eq '\') {
            $Backslashes++
            continue
        }

        if ($Character -eq '"') {
            if ($Backslashes -gt 0) {
                $Result += ('\' * ($Backslashes * 2))
                $Backslashes = 0
            }
            $Result += '\"'
            continue
        }

        if ($Backslashes -gt 0) {
            $Result += ('\' * $Backslashes)
            $Backslashes = 0
        }
        $Result += $Character
    }

    if ($Backslashes -gt 0) {
        $Result += ('\' * ($Backslashes * 2))
    }
    $Result += '"'
    return $Result
}

function Join-WindowsCommandLineArguments {
    param([string[]]$Arguments)

    return (($Arguments | ForEach-Object { ConvertTo-WindowsCommandLineArgument $_ }) -join ' ')
}

function Write-RedirectedFile {
    param(
        [string]$Path,
        [switch]$ErrorStream
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        return
    }

    try {
        $Text = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)
    }
    catch {
        return
    }
    if ($Text.Length -eq 0) {
        return
    }

    if ($ErrorStream) {
        [Console]::Error.Write($Text)
    }
    else {
        [Console]::Out.Write($Text)
    }
}

function Read-ExcelPidFile {
    param([string]$Path)

    $ProcessIds = @()
    if (-not (Test-Path -LiteralPath $Path)) {
        return $ProcessIds
    }

    foreach ($Line in Get-Content -LiteralPath $Path -ErrorAction SilentlyContinue) {
        $PidText = ([string]$Line).Trim()
        if ([string]::IsNullOrWhiteSpace($PidText)) {
            continue
        }
        $ParsedPid = 0
        if ([int]::TryParse($PidText, [ref]$ParsedPid) -and $ParsedPid -gt 0) {
            $ProcessIds += $ParsedPid
        }
    }

    return $ProcessIds | Select-Object -Unique
}

function Stop-ProcessTree {
    param([int]$ProcessId)

    try {
        $Children = Get-CimInstance Win32_Process -Filter "ParentProcessId = $ProcessId" -ErrorAction SilentlyContinue
        foreach ($Child in $Children) {
            Stop-ProcessTree -ProcessId ([int]$Child.ProcessId)
        }
    }
    catch {
    }

    try {
        Stop-Process -Id $ProcessId -Force -ErrorAction SilentlyContinue
    }
    catch {
    }
}

function Remove-TimeoutTempDirectory {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
        return
    }

    $ResolvedPath = [IO.Path]::GetFullPath($Path)
    $ResolvedTemp = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
    $Leaf = Split-Path -Leaf $ResolvedPath
    if ($ResolvedPath.StartsWith($ResolvedTemp, [StringComparison]::OrdinalIgnoreCase) -and $Leaf.StartsWith('sinitek-cli-', [StringComparison]::OrdinalIgnoreCase)) {
        Remove-Item -LiteralPath $ResolvedPath -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Invoke-WithTimeoutSupervisor {
    param(
        [int]$TimeoutSeconds,
        [string]$ActionName
    )

    $TempRoot = Join-Path ([IO.Path]::GetTempPath()) ('sinitek-cli-' + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $TempRoot -Force | Out-Null

    $StdoutPath = Join-Path $TempRoot 'stdout.txt'
    $StderrPath = Join-Path $TempRoot 'stderr.txt'
    $ChildExcelPidFile = Join-Path $TempRoot 'excel.pid'

    $PowerShellExe = Join-Path $PSHOME 'powershell.exe'
    $ChildArgs = @(
        '-NoProfile',
        '-ExecutionPolicy',
        'Bypass',
        '-File',
        $PSCommandPath,
        '-NoTimeoutSupervisor',
        '-ExcelPidFile',
        $ChildExcelPidFile
    )

    foreach ($Entry in $InitialBoundParameters.GetEnumerator()) {
        if ($Entry.Key -in @('NoTimeoutSupervisor', 'ExcelPidFile')) {
            continue
        }

        $Name = '-' + $Entry.Key
        $Value = $Entry.Value
        if ($Value -is [System.Management.Automation.SwitchParameter]) {
            if ($Value.IsPresent) {
                $ChildArgs += $Name
            }
        }
        elseif ($Value -is [bool]) {
            $BoolValue = if ($Value) { '$true' } else { '$false' }
            $ChildArgs += ($Name + ':' + $BoolValue)
        }
        else {
            $ChildArgs += $Name
            $ChildArgs += [string]$Value
        }
    }

    $ArgumentString = Join-WindowsCommandLineArguments -Arguments $ChildArgs
    $Process = $null
    try {
        $Process = Start-Process `
            -FilePath $PowerShellExe `
            -ArgumentList $ArgumentString `
            -WorkingDirectory (Get-Location).Path `
            -RedirectStandardOutput $StdoutPath `
            -RedirectStandardError $StderrPath `
            -WindowStyle Hidden `
            -PassThru

        if ($Process.WaitForExit($TimeoutSeconds * 1000)) {
            $Process.WaitForExit()
            $ExitCode = $Process.ExitCode
            $StderrText = ''
            try {
                if (Test-Path -LiteralPath $StderrPath) {
                    $StderrText = [IO.File]::ReadAllText($StderrPath, [Text.Encoding]::UTF8)
                }
            }
            catch {
                $StderrText = ''
            }
            Write-RedirectedFile -Path $StdoutPath
            Write-RedirectedFile -Path $StderrPath -ErrorStream
            if (($null -eq $ExitCode -or $ExitCode -eq 0) -and $StderrText -match '(?m)^ERROR:') {
                $ExitCode = 1
            }
            if ($null -eq $ExitCode) {
                $ExitCode = 1
            }
            exit $ExitCode
        }

        $ExcelProcessIds = Read-ExcelPidFile -Path $ChildExcelPidFile
        Stop-ProcessTree -ProcessId $Process.Id
        try {
            $Process.WaitForExit(5000) | Out-Null
        }
        catch {
        }
        foreach ($ExcelProcessId in $ExcelProcessIds) {
            try {
                Stop-Process -Id $ExcelProcessId -Force -ErrorAction SilentlyContinue
            }
            catch {
            }
        }

        Start-Sleep -Milliseconds 200
        Write-RedirectedFile -Path $StdoutPath
        Write-RedirectedFile -Path $StderrPath -ErrorStream
        [Console]::Error.WriteLine("ERROR: Action '$ActionName' timed out after $TimeoutSeconds seconds.")
        exit 124
    }
    finally {
        if ($Process -and -not $Process.HasExited) {
            Stop-ProcessTree -ProcessId $Process.Id
        }
        Remove-TimeoutTempDirectory -Path $TempRoot
    }
}

$ConfigBase = $Root
if (Test-ExplicitParam 'Config') {
    $ConfigPath = Resolve-OptionalPath -Path $Config -BasePath (Get-Location).Path
}
else {
    $ConfigPath = Resolve-OptionalPath -Path $Config -BasePath $Root
}

if (-not [string]::IsNullOrWhiteSpace($ConfigPath) -and (Test-Path -LiteralPath $ConfigPath)) {
    $ConfigBase = Split-Path -Parent $ConfigPath
    $ConfigData = Read-SimpleYaml -Path $ConfigPath

    Set-FromConfig -ConfigData $ConfigData -ConfigKey 'workbook' -VariableName 'Workbook' -PathValue -PathBase $ConfigBase
    Set-FromConfig -ConfigData $ConfigData -ConfigKey 'output_dir' -VariableName 'OutputDir' -PathValue -PathBase $ConfigBase

    Set-FromDefaultConfig -ConfigData $ConfigData -ConfigKey 'history_year' -VariableName 'HistoryYear'
    Set-FromDefaultConfig -ConfigData $ConfigData -ConfigKey 'forecast_year' -VariableName 'ForecastYear'
    Set-FromDefaultConfig -ConfigData $ConfigData -ConfigKey 'currency_unit' -VariableName 'CurrencyUnit'
    Set-FromDefaultConfig -ConfigData $ConfigData -ConfigKey 'update_directory' -VariableName 'UpdateDirectory'
    Set-FromDefaultConfig -ConfigData $ConfigData -ConfigKey 'update_src_data' -VariableName 'UpdateSrcData'
    Set-FromDefaultConfig -ConfigData $ConfigData -ConfigKey 'migrate' -VariableName 'Migrate'
    Set-FromDefaultConfig -ConfigData $ConfigData -ConfigKey 'add_output' -VariableName 'AddOutput'
    Set-FromDefaultConfig -ConfigData $ConfigData -ConfigKey 'company_management_type' -VariableName 'CompanyManagementType'
    Set-FromDefaultConfig -ConfigData $ConfigData -ConfigKey 'company_management_name' -VariableName 'CompanyManagementName'
    Set-FromDefaultConfig -ConfigData $ConfigData -ConfigKey 'peer_stock' -VariableName 'PeerStock'
    Set-FromDefaultConfig -ConfigData $ConfigData -ConfigKey 'prediction_scope' -VariableName 'PredictionScope'
    Set-FromDefaultConfig -ConfigData $ConfigData -ConfigKey 'prediction_rows' -VariableName 'PredictionRows'
    Set-FromDefaultConfig -ConfigData $ConfigData -ConfigKey 'prediction_indicators' -VariableName 'PredictionIndicators'
    Set-FromDefaultConfig -ConfigData $ConfigData -ConfigKey 'prediction_method' -VariableName 'PredictionMethod'
    Set-FromDefaultConfig -ConfigData $ConfigData -ConfigKey 'prediction_settings' -VariableName 'PredictionSettings'
    Set-FromDefaultConfig -ConfigData $ConfigData -ConfigKey 'timeout_seconds' -VariableName 'TimeoutSeconds'

    if (-not (Test-ExplicitParam 'Username') -and $ConfigData.ContainsKey('auth') -and $ConfigData['auth'].ContainsKey('username_env')) {
        $ResolvedUsername = [Environment]::GetEnvironmentVariable([string]$ConfigData['auth']['username_env'])
        if (-not [string]::IsNullOrEmpty($ResolvedUsername)) {
            $Username = $ResolvedUsername
        }
    }
    if (-not (Test-ExplicitParam 'Password') -and $ConfigData.ContainsKey('auth') -and $ConfigData['auth'].ContainsKey('password_env')) {
        $ResolvedPassword = [Environment]::GetEnvironmentVariable([string]$ConfigData['auth']['password_env'])
        if (-not [string]::IsNullOrEmpty($ResolvedPassword)) {
            $Password = $ResolvedPassword
        }
    }
}

try {
    $TimeoutSeconds = [int]$TimeoutSeconds
}
catch {
    throw "TimeoutSeconds must be an integer between 0 and 86400."
}
if ($TimeoutSeconds -lt 0 -or $TimeoutSeconds -gt 86400) {
    throw "TimeoutSeconds must be between 0 and 86400."
}

if (-not $NoTimeoutSupervisor.IsPresent -and $TimeoutSeconds -gt 0) {
    Invoke-WithTimeoutSupervisor -TimeoutSeconds $TimeoutSeconds -ActionName $Action
}

if (-not [string]::IsNullOrWhiteSpace($ExcelPidFile)) {
    $env:SINITEK_EXCEL_PID_FILE = $ExcelPidFile
}
else {
    [Environment]::SetEnvironmentVariable('SINITEK_EXCEL_PID_FILE', $null, 'Process')
}

if (Test-ExplicitParam 'Workbook') {
    $Workbook = Resolve-OptionalPath -Path $Workbook -BasePath (Get-Location).Path
}
else {
    $Workbook = Resolve-OptionalPath -Path $Workbook -BasePath $ConfigBase
}
if (Test-ExplicitParam 'OutWorkbook') {
    $OutWorkbook = Resolve-OptionalPath -Path $OutWorkbook -BasePath (Get-Location).Path
}

$OfficeDll = First-ExistingPath @(
    'C:\Program Files\Microsoft Office\root\vfs\ProgramFilesX86\Microsoft Office\Office16\DCF\office.dll',
    'C:\Program Files\Microsoft Office\root\Office16\ADDINS\PowerPivot Excel Add-in\OFFICE.dll',
    'C:\Windows\assembly\GAC_MSIL\office\15.0.0.0__71e9bce111e9429c\OFFICE.DLL'
)

$ExcelInteropDll = First-ExistingPath @(
    'C:\Program Files\Microsoft Office\root\vfs\ProgramFilesX86\Microsoft Office\Office16\DCF\Microsoft.Office.Interop.Excel.dll',
    'C:\Program Files\Microsoft Office\root\Office16\ADDINS\PowerPivot Excel Add-in\Microsoft.Office.Interop.Excel.dll',
    'C:\Program Files\Microsoft Office\root\Office16\ADDINS\Microsoft Power Query for Excel Integrated\bin\Microsoft.Office.Interop.Excel.dll'
)

$ExtensibilityDll = First-ExistingPath @(
    'C:\Windows\assembly\GAC\Extensibility\7.0.3300.0__b03f5f7f11d50a3a\extensibility.dll'
)

foreach ($RequiredPath in @($BridgePath, $SinitekDll, $NewtonsoftDll, $OfficeDll, $ExcelInteropDll, $ExtensibilityDll)) {
    if (-not (Test-Path -LiteralPath $RequiredPath)) {
        throw "Required file not found: $RequiredPath"
    }
}

[Reflection.Assembly]::LoadFrom($OfficeDll) | Out-Null
[Reflection.Assembly]::LoadFrom($ExcelInteropDll) | Out-Null
[Reflection.Assembly]::LoadFrom($NewtonsoftDll) | Out-Null
[Reflection.Assembly]::LoadFrom($SinitekDll) | Out-Null
[Reflection.Assembly]::LoadFrom($ExtensibilityDll) | Out-Null

$TrustedPlatformAssemblies = @()
$TrustedPlatformAssemblyString = [AppContext]::GetData('TRUSTED_PLATFORM_ASSEMBLIES')
if ($TrustedPlatformAssemblyString) {
    $TrustedPlatformAssemblies = $TrustedPlatformAssemblyString -split [IO.Path]::PathSeparator
}

$References = @(
    $TrustedPlatformAssemblies
    'System'
    'System.Core'
    'System.Xml'
    'System.Xml.Linq'
    'System.Drawing'
    'System.Windows.Forms'
    $OfficeDll
    $ExcelInteropDll
    $ExtensibilityDll
    $NewtonsoftDll
    $SinitekDll
) | Where-Object { $_ } | Select-Object -Unique

Add-Type -Path $BridgePath -ReferencedAssemblies $References

$WorkbookPath = if ($Workbook) { (Resolve-Path -LiteralPath $Workbook).Path } else { '' }

$MutatingActions = @('output', 'update', 'produce', 'predict')
if ($MutatingActions -contains $Action) {
    if (-not $Save.IsPresent -and [string]::IsNullOrWhiteSpace($OutWorkbook) -and -not [string]::IsNullOrWhiteSpace($OutputDir)) {
        $OutputBasePath = if (Test-ExplicitParam 'OutputDir') { (Get-Location).Path } else { $ConfigBase }
        $OutputBase = Resolve-OptionalPath -Path $OutputDir -BasePath $OutputBasePath
        $WorkbookName = if ($WorkbookPath) { [IO.Path]::GetFileNameWithoutExtension($WorkbookPath) } else { 'workbook' }
        $ActionName = Format-ActionForFileName $Action
        $StockSegment = ''
        try {
            $StockSegment = Get-StockCodeForFileName -StockCode $Stock -WorkbookPath $WorkbookPath
        }
        catch {
            if ($Action -ne 'predict') {
                throw
            }
        }
        $StockPart = if ([string]::IsNullOrWhiteSpace($StockSegment)) { '' } else { "-$StockSegment" }
        $Timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $OutWorkbook = Join-Path $OutputBase "$WorkbookName-$ActionName$StockPart-$Timestamp.xlsx"
    }
    if (-not $Save.IsPresent -and [string]::IsNullOrWhiteSpace($OutWorkbook)) {
        throw "Mutating action '$Action' requires -OutWorkbook, -OutputDir, or -Save."
    }
}

try {
    switch ($Action) {
        'inspect' {
            [SinitekCliBridge]::Inspect($WorkbookPath, $Visible.IsPresent)
        }
        'login' {
            [SinitekCliBridge]::Login($Username, $Password)
        }
        'output' {
            [SinitekCliBridge]::OutputDirect(
                $WorkbookPath,
                $OutWorkbook,
                $Save.IsPresent,
                $Visible.IsPresent,
                $Stock,
                $CurrencyUnit,
                $Username,
                $Password
            )
        }
        'predict' {
            $Result = [SinitekCliBridge]::PredictionSettingsDirect(
                $WorkbookPath,
                $OutWorkbook,
                $Save.IsPresent,
                $Visible.IsPresent,
                $PredictionScope,
                $PredictionRows,
                $PredictionIndicators,
                $PredictionMethod,
                $PredictionSettings
            )
            $Artifact = if (-not [string]::IsNullOrWhiteSpace($OutWorkbook)) {
                $OutWorkbook
            }
            elseif ($Save.IsPresent) {
                $WorkbookPath
            }
            else {
                ''
            }
            if (-not [string]::IsNullOrWhiteSpace($Artifact)) {
                $Result + [Environment]::NewLine + "Artifact=" + $Artifact
            }
            else {
                $Result
            }
        }
        'update' {
            [SinitekCliBridge]::UpdateDirect(
                $WorkbookPath,
                $OutWorkbook,
                $Save.IsPresent,
                $Visible.IsPresent,
                $Stock,
                $Gsdm,
                $StockName,
                $HistoryYear,
                $ForecastYear,
                $CurrencyUnit,
                $CompanyManagementType,
                $CompanyManagementName,
                $PeerStock,
                $UpdateDirectory,
                $UpdateSrcData,
                $Migrate,
                $AddOutput,
                $Username,
                $Password
            )
        }
        'produce' {
            $Result = [SinitekCliBridge]::UpdateDirect(
                $WorkbookPath,
                $OutWorkbook,
                $Save.IsPresent,
                $Visible.IsPresent,
                $Stock,
                $Gsdm,
                $StockName,
                $HistoryYear,
                $ForecastYear,
                $CurrencyUnit,
                $CompanyManagementType,
                $CompanyManagementName,
                $PeerStock,
                $UpdateDirectory,
                $UpdateSrcData,
                $Migrate,
                $true,
                $Username,
                $Password
            )
            $Artifact = if (-not [string]::IsNullOrWhiteSpace($OutWorkbook)) {
                $OutWorkbook
            }
            elseif ($Save.IsPresent) {
                $WorkbookPath
            }
            else {
                ''
            }
            if (-not [string]::IsNullOrWhiteSpace($Artifact)) {
                $Result + [Environment]::NewLine + "Artifact=" + $Artifact
            }
            else {
                $Result
            }
        }
    }
}
catch {
    [Console]::Error.WriteLine("ERROR: " + (Get-CliExceptionMessage $_.Exception))
    exit 1
}
