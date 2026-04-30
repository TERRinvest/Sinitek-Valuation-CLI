param(
    [ValidateSet('inspect', 'login', 'stock-search', 'button', 'handler', 'output', 'update-direct', 'produce')]
    [string]$Action = 'inspect',

    [string]$Config = '.\sinitek.yaml',
    [string]$Workbook = '.\Sinitek_Model_Ashare_V12.xlsx',
    [string]$OutWorkbook = '',
    [string]$OutputDir = '',
    [switch]$Save,
    [switch]$Visible,

    [string]$ButtonId = '',
    [string]$HandlerType = '',

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
    [bool]$UpdateDirectory = $true,
    [bool]$UpdateSrcData = $true,
    [bool]$Migrate = $false,
    [bool]$AddOutput = $false,
    [int]$Count = 10
)

$ErrorActionPreference = 'Stop'

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
foreach ($Key in $PSBoundParameters.Keys) {
    $ExplicitParams[$Key] = $true
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

function Set-FromNestedConfig {
    param(
        [hashtable]$ConfigData,
        [string]$Section,
        [string]$ConfigKey,
        [string]$VariableName,
        [string]$ExplicitName = $VariableName
    )
    if (Test-ExplicitParam $ExplicitName) {
        return
    }
    if (-not $ConfigData.ContainsKey($Section) -or -not $ConfigData[$Section].ContainsKey($ConfigKey)) {
        return
    }
    Set-Variable -Name $VariableName -Value $ConfigData[$Section][$ConfigKey] -Scope 1
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
    Set-FromDefaultConfig -ConfigData $ConfigData -ConfigKey 'count' -VariableName 'Count'

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

$MutatingActions = @('button', 'handler', 'output', 'update-direct', 'produce')
if ($MutatingActions -contains $Action) {
    if (-not $Save.IsPresent -and [string]::IsNullOrWhiteSpace($OutWorkbook) -and -not [string]::IsNullOrWhiteSpace($OutputDir)) {
        $OutputBasePath = if (Test-ExplicitParam 'OutputDir') { (Get-Location).Path } else { $ConfigBase }
        $OutputBase = Resolve-OptionalPath -Path $OutputDir -BasePath $OutputBasePath
        $WorkbookName = if ($WorkbookPath) { [IO.Path]::GetFileNameWithoutExtension($WorkbookPath) } else { 'workbook' }
        $Timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $OutWorkbook = Join-Path $OutputBase "$WorkbookName-$Action-$Timestamp.xlsx"
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
        'stock-search' {
            [SinitekCliBridge]::StockSearch($WorkbookPath, $Stock, $Count, $Username, $Password)
        }
        'button' {
            [SinitekCliBridge]::InvokeButton(
                $WorkbookPath,
                $OutWorkbook,
                $Save.IsPresent,
                $Visible.IsPresent,
                $ButtonId,
                $Username,
                $Password
            )
        }
        'handler' {
            [SinitekCliBridge]::InvokeHandler(
                $WorkbookPath,
                $OutWorkbook,
                $Save.IsPresent,
                $Visible.IsPresent,
                $HandlerType,
                $Username,
                $Password
            )
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
        'update-direct' {
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
