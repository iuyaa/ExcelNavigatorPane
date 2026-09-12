param(
    [string]$Version = '1.0.14.0',
    [ValidateSet('oneview','github')][string]$UpdateChannel = 'oneview',
    [string]$ChannelsFile = (Join-Path $PSScriptRoot 'UpdateChannels.json'),
    [string]$UpdateBaseUrl,
    [string]$OutputDirectory,
    [string]$CertificateThumbprint,
    [string]$MSBuild = 'F:\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe',
    [string]$Wix,
    [string]$BootstrapperPath = 'C:\Program Files (x86)\Microsoft SDKs\ClickOnce Bootstrapper'
)
$ErrorActionPreference = 'Stop'
$UpdateChannel = $UpdateChannel.ToLowerInvariant()
# MSI compares three version fields; reserve the fourth field as zero.
if ($Version -notmatch '^\d+\.\d+\.\d+\.0$') { throw 'MSI versions must have the form major.minor.build.0.' }
$parsed = [version]$Version
if ($parsed.ToString() -ne $Version) { throw 'Version fields must not contain leading zeros.' }
if ($parsed.Major -gt 255 -or $parsed.Minor -gt 255 -or $parsed.Build -gt 65535 -or $parsed -lt [version]'1.0.1.0') { throw 'Version exceeds MSI limits or predates MSI migration.' }
$repo = Split-Path $PSScriptRoot -Parent
if (!$Wix) { $Wix = Join-Path $repo 'work\tools\wix\wix.exe' }
if (!(Test-Path $Wix)) { throw 'Install build tool: dotnet tool install wix --version 4.0.6 --tool-path work/tools/wix' }
$channels = Get-Content -LiteralPath $ChannelsFile -Raw | ConvertFrom-Json
$manifestUrl = $channels.$UpdateChannel.manifestUrl
if ($UpdateBaseUrl) {
    if ($UpdateChannel -ne 'oneview') { throw 'UpdateBaseUrl only applies to the oneview channel.' }
    $manifestUrl = $UpdateBaseUrl.TrimEnd('/') + '/latest.xml'
}
$uri = $null
if (![uri]::TryCreate($manifestUrl, [UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -ne 'https' -or
    $uri.UserInfo -or $uri.Query -or $uri.Fragment -or $manifestUrl -match '[;\r\n]') { throw 'A fixed HTTPS update address is required.' }
if ($UpdateChannel -eq 'github') {
    if ($uri.Host -ne 'github.com' -or $uri.AbsolutePath -notmatch '^/[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+/releases\.atom$') { throw 'A GitHub releases feed address is required.' }
} elseif (!$uri.AbsolutePath.EndsWith('/latest.xml')) { throw 'A latest.xml address is required.' }
$certificate = if ($CertificateThumbprint) { Get-Item "Cert:\CurrentUser\My\$CertificateThumbprint" } else {
    $candidates = @(Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert | Where-Object {
        $_.Subject -eq 'CN=Excel Navigator Internal Test' -and $_.HasPrivateKey -and $_.NotAfter -gt (Get-Date)
    })
    if ($candidates.Count -ne 1) { throw 'Specify the existing release CertificateThumbprint; update signing keys must not change implicitly.' }
    $candidates[0]
}
if (!$certificate -or !$certificate.HasPrivateKey -or $certificate.NotAfter -le (Get-Date)) { throw 'Existing release signing certificate required; do not rotate the update key implicitly.' }
$rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($certificate)
$stage = Join-Path $repo ('work\installer\msi-' + [guid]::NewGuid().ToString('N'))
$app = Join-Path $stage 'app'; $payload = Join-Path $stage 'payload'
$dist = if ($OutputDirectory) { [IO.Path]::GetFullPath($OutputDirectory) } else { Join-Path $repo "dist\releases\$Version\$UpdateChannel" }
New-Item -ItemType Directory -Path $app,$payload,$dist -Force | Out-Null
$settingsFile = Join-Path $stage 'UpdateSettings.xml'
$settings = [xml]::new(); $node = $settings.CreateElement('updates')
$node.SetAttribute('version', $Version); $node.SetAttribute('channel', $UpdateChannel); $node.SetAttribute('manifestUrl', $uri.AbsoluteUri)
$key = $settings.CreateElement('publicKey'); $key.InnerText = $rsa.ToXmlString($false)
[void]$node.AppendChild($key); [void]$settings.AppendChild($node); $settings.Save($settingsFile)
function Read-AddinRegistration {
    $snapshot = foreach ($view in @([Microsoft.Win32.RegistryView]::Registry32, [Microsoft.Win32.RegistryView]::Registry64)) {
        $root = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $view)
        $registration = $root.OpenSubKey('Software\Microsoft\Office\Excel\Addins\ExcelNavigatorPane')
        try {
            $values = [ordered]@{}
            if ($registration) { foreach ($name in ($registration.GetValueNames() | Sort-Object)) { $values[$name] = @($registration.GetValueKind($name).ToString(), $registration.GetValue($name)) } }
            [ordered]@{ View = $view.ToString(); Exists = $null -ne $registration; Values = $values }
        } finally { if ($registration) { $registration.Dispose() }; $root.Dispose() }
    }
    ConvertTo-Json -InputObject $snapshot -Depth 5 -Compress
}
$registrationBefore = Read-AddinRegistration
# Remove only the build's staging-path inclusion entry; never register/unregister the active add-in.
& $MSBuild (Join-Path $repo 'ExcelNavigatorPane.csproj') '/t:Rebuild;RemoveOfficeAddInSecurity' /p:Configuration=Release /p:Platform=AnyCPU /p:BuildInstaller=true "/p:OutputPath=$app\" "/p:ApplicationVersion=$Version" "/p:UpdateSettingsFile=$settingsFile" "/p:ManifestCertificateThumbprint=$($certificate.Thumbprint)" /p:ManifestKeyFile= /p:UpdateEnabled=false /p:BootstrapperEnabled=false /nologo /verbosity:minimal
if ($LASTEXITCODE -ne 0) { throw 'VSTO local payload build failed.' }
if ((Read-AddinRegistration) -cne $registrationBefore) { throw 'VSTO build changed the existing add-in registration; packaging stopped.' }
if ((Get-Item (Join-Path $app 'ExcelNavigatorPane.dll')).VersionInfo.FileVersion -ne $Version) {
    throw 'AssemblyFileVersion in Properties/AssemblyInfo.cs must match the MSI release version so Windows Installer replaces the DLL.'
}
$appFiles = @(Get-ChildItem $app -File | Where-Object { $_.Name -match '\.(dll|manifest|vsto|config)$' })
if ($appFiles.Count -lt 4) { throw 'Incomplete VSTO payload.' }
$components = @(); $references = @(); $index = 0
foreach ($file in $appFiles) {
    $id = 'Payload' + $index++; $escaped = [Security.SecurityElement]::Escape($file.FullName)
    $components += "<Component Id=`"$id`" Guid=`"*`"><File Id=`"File$id`" Source=`"$escaped`" KeyPath=`"yes`" /></Component>"
    $references += "<ComponentRef Id=`"$id`" />"
}
$include = Join-Path $stage 'Payload.wxi'
"<Include xmlns=`"http://wixtoolset.org/schemas/v4/wxs`">$($components -join '')</Include>" | Set-Content $include -Encoding utf8
$groups = Join-Path $stage 'PayloadGroup.wxs'
"<Wix xmlns=`"http://wixtoolset.org/schemas/v4/wxs`"><Fragment><ComponentGroup Id=`"PayloadComponents`">$($references -join '')</ComponentGroup></Fragment></Wix>" | Set-Content $groups -Encoding utf8
foreach ($arch in @('x86','x64')) {
    $folder = Join-Path $payload $arch; New-Item -ItemType Directory $folder -Force | Out-Null
    $msiName = "ExcelNavigator-$Version-$arch.msi"; $msi = Join-Path $folder $msiName
    $programFiles = if ($arch -eq 'x64') { 'ProgramFiles64Folder' } else { 'ProgramFilesFolder' }
    & $Wix build (Join-Path $PSScriptRoot 'Product.wxs') $groups -arch $arch -d "MsiVersion=$($parsed.Major).$($parsed.Minor).$($parsed.Build)" -d "ProgramFiles=$programFiles" -d "PayloadInclude=$include" -o $msi
    if ($LASTEXITCODE -ne 0) { throw "MSI build failed: $arch" }
    & $MSBuild (Join-Path $PSScriptRoot 'Bootstrapper.proj') /t:Build "/p:BootstrapOutput=$folder" "/p:BootstrapperPath=$BootstrapperPath" /nologo /verbosity:minimal
    if ($LASTEXITCODE -ne 0 -or !(Test-Path (Join-Path $folder 'setup.exe'))) { throw 'Microsoft prerequisite bootstrapper failed.' }
    Copy-Item $msi (Join-Path $dist $msiName)
}
Add-Type -AssemblyName System.IO.Compression.FileSystem
$payloadZip = Join-Path $stage 'payload.zip'
# Exclude WiX debugging metadata rather than distributing build paths.
foreach ($arch in @('x86','x64')) {
    $debug = Join-Path $payload "$arch\ExcelNavigator-$Version-$arch.wixpdb"
    if (Test-Path -LiteralPath $debug) { Remove-Item -LiteralPath $debug }
}
[IO.Compression.ZipFile]::CreateFromDirectory($payload, $payloadZip)
$versionSource = Join-Path $stage 'InstallerVersion.cs'
"[assembly: System.Reflection.AssemblyVersion(`"$Version`")]`n[assembly: System.Reflection.AssemblyFileVersion(`"$Version`")]" | Set-Content $versionSource -Encoding utf8
$exe = Join-Path $dist "ExcelNavigator-Setup-$Version.exe"
& "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe" /nologo /target:winexe /platform:anycpu /optimize+ /r:System.Windows.Forms.dll /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll "/resource:$payloadZip,payload.zip" "/out:$exe" (Join-Path $PSScriptRoot 'SetupLauncher.cs') $versionSource
if ($LASTEXITCODE -ne 0) { throw 'EXE launcher build failed.' }
$signature = Set-AuthenticodeSignature -LiteralPath $exe -Certificate $certificate -HashAlgorithm SHA256
if ($signature.SignerCertificate.Thumbprint -ne $certificate.Thumbprint) { throw 'EXE signature missing.' }
$hash = (Get-FileHash $exe -Algorithm SHA256).Hash.ToLowerInvariant()
"$hash  $([IO.Path]::GetFileName($exe))" | Set-Content ($exe + '.sha256') -Encoding ascii
$fileKey = "releases/$Version/$([IO.Path]::GetFileName($exe))"
$message = "ExcelNavigatorPane`n$Version`n$fileKey`n$hash"
$release = [xml]::new(); $entry = $release.CreateElement('release')
$entry.SetAttribute('product','ExcelNavigatorPane'); $entry.SetAttribute('version',$Version)
$entry.SetAttribute('file',$fileKey); $entry.SetAttribute('sha256',$hash)
$signed = $rsa.SignData([Text.Encoding]::UTF8.GetBytes($message), [Security.Cryptography.HashAlgorithmName]::SHA256, [Security.Cryptography.RSASignaturePadding]::Pkcs1)
$entry.SetAttribute('signature',[Convert]::ToBase64String($signed)); [void]$release.AppendChild($entry)
$release.Save((Join-Path $dist 'latest.xml'))
$rsa.ToXmlString($false) | Set-Content (Join-Path $dist 'update-public-key.xml') -Encoding utf8
(Get-Content (Join-Path $PSScriptRoot '安装说明.txt') -Raw).Replace('{VERSION}',$Version).Replace('{CHANNEL}',$UpdateChannel) | Set-Content (Join-Path $dist '安装说明.txt') -Encoding utf8
$check = Join-Path $stage 'extraction-check'
$process = Start-Process $exe -ArgumentList @('--extract-only', ('"' + $check + '"')) -WindowStyle Hidden -PassThru -Wait
if ($process.ExitCode -ne 0) { throw 'EXE extraction check failed.' }
foreach ($file in Get-ChildItem $payload -Recurse -File) {
    $relative = $file.FullName.Substring($payload.Length + 1)
    if ((Get-FileHash $file.FullName).Hash -ne (Get-FileHash (Join-Path $check $relative)).Hash) { throw "EXE payload mismatch: $relative" }
}
Write-Output "PASS: x86/x64 MSI, signed update metadata, EXE extraction hashes. Output: $dist"
Write-Output "Payload: $app. Certificate trust stores unchanged. Actual install/upgrade: NOT_RUN."
