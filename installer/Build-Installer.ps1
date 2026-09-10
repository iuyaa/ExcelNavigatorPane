param(
    [string]$Version = '1.0.0.2',
    [string]$UpdateBaseUrl,
    [string]$OutputDirectory,
    [string]$CertificateThumbprint,
    [string]$MSBuild = 'F:\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe'
)
$ErrorActionPreference = 'Stop'
if ($Version -notmatch '^\d+\.\d+\.\d+\.\d+$') { throw 'Version must contain four numeric components.' }
$parsedVersion = [version]$Version
if ($parsedVersion -lt [version]'1.0.0.0') { throw 'Version must be at least 1.0.0.0.' }
$manifestUrl = ''
if ($UpdateBaseUrl) {
    $updateUri = $null
    if (![uri]::TryCreate($UpdateBaseUrl, [UriKind]::Absolute, [ref]$updateUri) -or
        $updateUri.Scheme -ne 'https' -or $updateUri.UserInfo -or $updateUri.Query -or $updateUri.Fragment -or
        $UpdateBaseUrl -match '[;\r\n]') { throw 'UpdateBaseUrl must be a fixed HTTPS directory URL without credentials, query or fragment.' }
    $UpdateBaseUrl = $updateUri.AbsoluteUri.TrimEnd('/') + '/'
    $manifestUrl = $UpdateBaseUrl + 'ExcelNavigatorPane.vsto'
}
$repo = Split-Path $PSScriptRoot -Parent
if ($CertificateThumbprint) {
    $certificate = Get-Item "Cert:\CurrentUser\My\$CertificateThumbprint"
} else {
    $certificate = Get-ChildItem Cert:\CurrentUser\My -CodeSigningCert | Where-Object {
        $_.Subject -eq 'CN=Excel Navigator Internal Test' -and $_.HasPrivateKey -and $_.NotAfter -gt (Get-Date)
    } | Sort-Object NotAfter -Descending | Select-Object -First 1
    if (!$certificate) {
        # Creates a signing identity only; never adds it to Root or TrustedPublisher.
        $certificate = New-SelfSignedCertificate -Type CodeSigningCert -Subject 'CN=Excel Navigator Internal Test' -CertStoreLocation Cert:\CurrentUser\My -NotAfter (Get-Date).AddYears(1)
    }
}
$thumbprint = $certificate.Thumbprint
if (!$certificate.HasPrivateKey -or $certificate.NotAfter -le (Get-Date)) { throw 'A valid signing certificate with a private key is required.' }
$stage = Join-Path $repo ('work\installer\build-' + [guid]::NewGuid().ToString('N'))
$publish = Join-Path $stage 'publish'
$dist = if ($OutputDirectory) { [IO.Path]::GetFullPath($OutputDirectory) } else { Join-Path $repo 'dist' }
New-Item -ItemType Directory -Path $publish, $dist -Force | Out-Null
$settingsFile = Join-Path $stage 'UpdateSettings.xml'
$settingsXml = New-Object System.Xml.XmlDocument
$settingsNode = $settingsXml.CreateElement('updates')
$settingsNode.SetAttribute('version', $Version)
$settingsNode.SetAttribute('manifestUrl', $manifestUrl)
[void]$settingsXml.AppendChild($settingsNode)
$settingsXml.Save($settingsFile)
$updateEnabled = if ($manifestUrl) { 'true' } else { 'false' }
& $MSBuild (Join-Path $repo 'ExcelNavigatorPane.csproj') /t:Rebuild,Publish /p:Configuration=Release /p:Platform=AnyCPU "/p:PublishDir=$publish\" "/p:PublishUrl=$publish\" "/p:ApplicationVersion=$Version" "/p:UpdateSettingsFile=$settingsFile" "/p:ManifestCertificateThumbprint=$thumbprint" /p:ManifestKeyFile= "/p:UpdateEnabled=$updateEnabled" /p:UpdateInterval=0 /p:UpdateIntervalUnits=days "/p:InstallUrl=$UpdateBaseUrl" /p:BootstrapperEnabled=true "/p:IsWebBootstrapper=$updateEnabled" /p:Install=true /nologo /verbosity:minimal
if ($LASTEXITCODE -ne 0) { throw 'ClickOnce publish failed.' }

Add-Type -AssemblyName System.IO.Compression.FileSystem
$payload = Join-Path $stage 'payload.zip'
[System.IO.Compression.ZipFile]::CreateFromDirectory($publish, $payload)
Copy-Item $payload (Join-Path $dist "ExcelNavigator-Publish-$Version.zip")
$exe = Join-Path $dist "ExcelNavigator-Setup-$Version.exe"
$framework = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319'
$versionSource = Join-Path $stage 'InstallerVersion.cs'
"[assembly: System.Reflection.AssemblyVersion(`"$Version`")]`n[assembly: System.Reflection.AssemblyFileVersion(`"$Version`")]" | Set-Content $versionSource -Encoding utf8
& (Join-Path $framework 'csc.exe') /nologo /target:winexe /platform:anycpu /optimize+ /r:System.Windows.Forms.dll /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll "/resource:$payload,payload.zip" "/out:$exe" (Join-Path $PSScriptRoot 'SetupLauncher.cs') $versionSource
if ($LASTEXITCODE -ne 0) { throw 'Single-file launcher compilation failed.' }
$signature = Set-AuthenticodeSignature -LiteralPath $exe -Certificate $certificate -HashAlgorithm SHA256
if (!$signature.SignerCertificate -or $signature.SignerCertificate.Thumbprint -ne $thumbprint) { throw 'Installer signing failed.' }

# Execute only the extraction check, never install or alter Excel registration on the build machine.
$check = Join-Path $stage 'extraction-check'
$process = Start-Process -FilePath $exe -ArgumentList @('--extract-only', ('"' + $check + '"')) -WindowStyle Hidden -PassThru -Wait
if ($process.ExitCode -ne 0) { throw 'Installer extraction check failed.' }
foreach ($file in Get-ChildItem $publish -Recurse -File) {
    $relative = $file.FullName.Substring($publish.Length + 1)
    if ((Get-FileHash $file.FullName).Hash -ne (Get-FileHash (Join-Path $check $relative)).Hash) { throw "Payload mismatch: $relative" }
}
$instructions = (Get-Content (Join-Path $PSScriptRoot '安装说明.txt') -Raw).Replace('{VERSION}', $Version)
$updateNote = if ($manifestUrl) { "已配置启动时更新：$manifestUrl" } else { '本包尚未配置更新地址，自动更新未启用。' }
$instructions.Replace('{UPDATE_STATUS}', $updateNote) | Set-Content (Join-Path $dist '安装说明.txt') -Encoding utf8
$hash = Get-FileHash $exe -Algorithm SHA256
"$($hash.Hash)  $([IO.Path]::GetFileName($exe))" | Set-Content ($exe + '.sha256') -Encoding ascii
Write-Output "PASS: publish, signed EXE, extraction and all payload hashes. Installer: $exe"
Write-Output "Certificate trust status on this machine: $($signature.Status). Clean-machine installation: NOT_RUN."
