param(
    [Parameter(Mandatory)][string]$Directory,
    [string]$Version = '1.0.11.0',
    [ValidateSet('oneview','github')][string]$UpdateChannel = 'oneview',
    [string]$Wix,
    [string]$Csc = 'F:\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\Roslyn\csc.exe'
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
$Directory = (Resolve-Path $Directory).Path
if (!$Wix) { $Wix = Join-Path $repo 'work\tools\wix\wix.exe' }
$work = Join-Path $repo ('work\msi-check-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory $work | Out-Null
$installer = New-Object -ComObject WindowsInstaller.Installer
function Read-Rows($database, [string]$table, [string[]]$columns) {
    $quoted = $columns | ForEach-Object { ([char]96) + $_ + ([char]96) }
    $view = $database.OpenView('SELECT ' + ($quoted -join ',') + ' FROM ' + ([char]96) + $table + ([char]96))
    [void]$view.Execute()
    try {
        while ($record = $view.Fetch()) {
            $row = @{}
            for ($i = 0; $i -lt $columns.Count; $i++) { $row[$columns[$i]] = $record.StringData($i + 1) }
            [pscustomobject]$row
        }
    } finally { [void]$view.Close() }
}
function Require([bool]$condition, [string]$message) { if (!$condition) { throw $message } }
$codes = @()
foreach ($arch in @('x86','x64')) {
    $msi = Join-Path $Directory "ExcelNavigator-$Version-$arch.msi"
    $db = $installer.OpenDatabase($msi,0)
    $properties = @{}
    Read-Rows $db 'Property' @('Property','Value') | ForEach-Object { $properties[$_.Property] = $_.Value }
    Require ($properties.ProductVersion -eq ($Version -replace '\.0$','')) 'MSI version mismatch'
    Require ($properties.UpgradeCode -eq '{A14219F5-2541-441C-BAEC-606015D32B55}') 'Upgrade family changed'
    Require ($properties.ALLUSERS -eq '1') 'Must install per-machine'
    Require ($properties.REBOOT -eq 'ReallySuppress' -and $properties.MSIRESTARTMANAGERCONTROL -eq 'Disable') 'Must not restart Excel or Windows'
    $codes += $properties.ProductCode
    $template = $db.SummaryInformation(0).Property(7)
    Require ($template -eq $(if ($arch -eq 'x86') {'Intel;2052'} else {'x64;2052'})) 'MSI architecture mismatch'
    $dirs = @(Read-Rows $db 'Directory' @('Directory','Directory_Parent','DefaultDir'))
    $expected = if ($arch -eq 'x64') { 'ProgramFiles64Folder' } else { 'ProgramFilesFolder' }
    Require (($dirs | Where-Object Directory -eq INSTALLFOLDER).Directory_Parent -eq $expected) 'Not a trusted Program Files installation'
    $regs = @(Read-Rows $db 'Registry' @('Root','Key','Name','Value','Component_'))
    $registrationCount = if ($arch -eq 'x64') { 2 } else { 1 }
    Require ($regs.Count -eq 4 * $registrationCount -and @($regs | Where-Object { $_.Root -ne '2' -or $_.Key -ne 'Software\Microsoft\Office\Excel\Addins\ExcelNavigatorPane' }).Count -eq 0) 'Unexpected registry writes'
    foreach ($group in ($regs | Group-Object Component_)) {
        Require ($group.Count -eq 4) 'Incomplete host registration'
        Require (($group.Group | Where-Object Name -eq Manifest).Value -eq 'file:///[INSTALLFOLDER]ExcelNavigatorPane.vsto|vstolocal') 'Manifest must load locally'
        Require (($group.Group | Where-Object Name -eq LoadBehavior).Value -eq '#3') 'Wrong load behavior'
    }
    $component = Read-Rows $db 'Component' @('Component','Attributes') | Where-Object Component -eq AddinRegistration
    Require ((([int]$component.Attributes -band 256) -ne 0) -eq ($arch -eq 'x64')) 'Wrong registry view'
    if ($arch -eq 'x64') {
        $compat = @(Read-Rows $db 'Component' @('Component','Attributes') | Where-Object Component -eq AddinRegistration32)
        Require ($compat.Count -eq 1 -and ([int]$compat[0].Attributes -band 256) -eq 0) 'Missing 32-bit registration beside 64-bit registration'
        $features = @(Read-Rows $db 'FeatureComponents' @('Feature_','Component_'))
        Require (@($features | Where-Object { $_.Feature_ -eq 'Main' -and $_.Component_ -eq 'AddinRegistration32' }).Count -eq 1) '32-bit registration must participate in install and uninstall'
    }
    $sequence = @{}
    Read-Rows $db 'InstallExecuteSequence' @('Action','Sequence') | ForEach-Object { $sequence[$_.Action] = [int]$_.Sequence }
    Require ($sequence.RemoveExistingProducts -gt $sequence.InstallInitialize -and $sequence.RemoveExistingProducts -lt $sequence.InstallFiles) 'Upgrade must remove old version within rollback transaction'
    $upgrade = @(Read-Rows $db 'Upgrade' @('UpgradeCode','VersionMin','VersionMax','Attributes','ActionProperty'))
    Require (@($upgrade | Where-Object { $_.ActionProperty -eq 'WIX_DOWNGRADE_DETECTED' -and $_.VersionMin -eq $properties.ProductVersion }).Count -eq 1) 'Missing downgrade detection'
    $conditions = @(Read-Rows $db 'LaunchCondition' @('Condition','Description'))
    if ($arch -eq 'x86') { Require (@($conditions | Where-Object Condition -eq 'Installed OR NOT VersionNT64').Count -eq 1) '32-bit MSI must not omit the 64-bit host on 64-bit Windows' }
    Require (@($conditions | Where-Object Condition -eq 'NOT WIX_DOWNGRADE_DETECTED').Count -eq 1) 'Missing downgrade rejection'
    $session = $installer.OpenPackage($msi,1)
    # AppSearch only reads registry; never run InstallInitialize or INSTALL in this check.
    Require ($session.DoAction('AppSearch') -eq 1) 'Dependency search failed'
    $runtime = $session.Property('VSTORUNTIME')
    Require (@($conditions | Where-Object Condition -match 'VSTORUNTIME').Count -eq 1) 'Missing VSTO dependency gate'
    foreach ($value in @('', '#461814','#528039','#528040','#533320')) {
        $session.Property('NET48') = $value
        Require (($session.EvaluateCondition('NET48 >= "#528040"') -eq 1) -eq ($value -in @('#528040','#533320'))) 'NET48 boundary check failed'
    }
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($session)
    $out = Join-Path $work $arch
    & $Wix msi decompile $msi -x $out -o (Join-Path $out 'Product.wxs')
    if ($LASTEXITCODE -ne 0) { throw 'MSI extraction failed' }
    $files = @(Read-Rows $db 'File' @('File','FileName'))
    foreach ($file in $files) {
        $name = ($file.FileName -split '\|')[-1]
        Copy-Item (Join-Path $out ('File\' + $file.File)) (Join-Path $out $name)
        if ($arch -eq 'x64') { Require ((Get-FileHash (Join-Path $out $name)).Hash -eq (Get-FileHash (Join-Path $work "x86\$name")).Hash) "x86/x64 payload mismatch: $name" }
    }
    foreach ($name in @('ExcelNavigatorPane.vsto','ExcelNavigatorPane.dll.manifest')) {
        [xml]$manifest = Get-Content (Join-Path $out $name) -Raw
        Require ($null -ne $manifest.SelectSingleNode('//*[local-name()="Signature"]')) 'Unsigned VSTO manifest'
        Require ($manifest.SelectSingleNode('/*/*[local-name()="assemblyIdentity"]').version -eq $Version) 'VSTO version mismatch'
        Require ($null -eq $manifest.SelectSingleNode('//*[local-name()="deploymentProvider"]')) 'Unexpected ClickOnce online source'
        foreach ($dependency in $manifest.SelectNodes('//*[local-name()="dependentAssembly"][@codebase]')) {
            $local = Join-Path $out $dependency.codebase
            $digest = $dependency.SelectSingleNode('.//*[local-name()="DigestValue"]')
            if ($digest) {
                Require (Test-Path $local) "Missing signed dependency: $local"
                $sha = [Security.Cryptography.SHA256]::Create()
                try { $actual = [Convert]::ToBase64String($sha.ComputeHash([IO.File]::ReadAllBytes($local))) } finally { $sha.Dispose() }
                Require ($digest.InnerText -eq $actual) 'Signed dependency hash mismatch'
            }
        }
    }
    [xml]$application = Get-Content (Join-Path $out 'ExcelNavigatorPane.dll.manifest') -Raw
    Require ((Get-Item (Join-Path $out 'ExcelNavigatorPane.dll')).VersionInfo.FileVersion -eq $Version) 'DLL file version must advance with MSI upgrades'
    Require ($application.SelectSingleNode('//*[local-name()="update"]').enabled -eq 'false') 'ClickOnce updates must be disabled'
    Write-Output "PASS: $arch MSI tables, local payload hashes, dependency gate (installed VSTO=$runtime), upgrade transaction"
}
Require ($codes[0] -ne $codes[1]) 'Architecture packages must have different ProductCodes'
$check = Join-Path $work 'UpdateCheck.exe'
& $Csc /nologo /target:exe /main:UpdateCheck /r:System.Net.Http.dll /r:System.Xml.Linq.dll /r:System.Windows.Forms.dll /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll "/out:$check" "/resource:$repo\Properties\UpdateSettings.xml,ExcelNavigatorPane.UpdateSettings.xml" (Join-Path $repo 'UpdateChecker.cs') (Join-Path $repo 'installer\SetupLauncher.cs') (Join-Path $PSScriptRoot 'UpdateCheck.cs')
if ($LASTEXITCODE -ne 0) { throw 'Update check compilation failed' }
& $check $Directory (Join-Path $work 'x64\ExcelNavigatorPane.dll') $UpdateChannel
if ($LASTEXITCODE -ne 0) { throw 'Update checks failed' }
$hostCheck = Join-Path $work 'WpsCompatibilityCheck.exe'
& $Csc /nologo /target:exe /r:System.Windows.Forms.dll /r:System.Drawing.dll "/out:$hostCheck" (Join-Path $PSScriptRoot 'WpsCompatibilityCheck.cs')
if ($LASTEXITCODE -ne 0) { throw 'Host compatibility check compilation failed' }
& $hostCheck (Join-Path $work 'x64\ExcelNavigatorPane.dll')
if ($LASTEXITCODE -ne 0) { throw 'Host compatibility check failed' }
Write-Output "PASS: offline packaging checks. Real install/load/upgrade/uninstall: NOT_RUN. Evidence: $work"
$migrationCheck = Join-Path $work 'RegistrationMigrationCheck.exe'
& $Csc /nologo /target:exe /main:RegistrationMigrationCheck /r:System.Windows.Forms.dll /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll "/out:$migrationCheck" (Join-Path $repo 'installer\SetupLauncher.cs') (Join-Path $PSScriptRoot 'RegistrationMigrationCheck.cs')
if ($LASTEXITCODE -ne 0) { throw 'Registration migration check compilation failed' }
& $migrationCheck
if ($LASTEXITCODE -ne 0) { throw 'Registration migration check failed' }
