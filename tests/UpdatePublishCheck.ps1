param([Parameter(Mandatory)][string]$Archive, [Parameter(Mandatory)][string]$Version, [string]$UpdateBaseUrl)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [IO.Compression.ZipFile]::OpenRead((Resolve-Path $Archive))
try {
    function Read-Entry([string]$Name) {
        $entry = $zip.Entries | Where-Object { $_.FullName.Replace('\', '/') -eq $Name } | Select-Object -First 1
        if (!$entry) { throw "Missing entry: $Name" }
        $reader = [IO.StreamReader]::new($entry.Open())
        try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
    }
    [xml]$deployment = Read-Entry 'ExcelNavigatorPane.vsto'
    $identity = $deployment.SelectSingleNode('/*/*[local-name()="assemblyIdentity"]')
    if ($identity.version -ne $Version) { throw 'Incorrect published version' }
    $dependent = $deployment.SelectSingleNode('//*[local-name()="dependentAssembly"]')
    $appPath = $dependent.codebase.Replace('\', '/')
    if (!$appPath.Contains($Version.Replace('.', '_'))) { throw 'Version directory mismatch' }
    [xml]$application = Read-Entry $appPath
    $update = $application.SelectSingleNode('//*[local-name()="update"]')
    $enabled = if ($UpdateBaseUrl) { 'true' } else { 'false' }
    if ($update.enabled -ne $enabled -or $update.HasChildNodes) { throw 'Wrong startup update policy' }
    foreach ($doc in @($deployment, $application)) {
        if (!$doc.SelectSingleNode('//*[local-name()="Signature"]')) { throw 'Unsigned manifest' }
    }
    $setup = $zip.Entries | Where-Object Name -eq 'setup.exe'
    $memory = [IO.MemoryStream]::new()
    $input = $setup.Open()
    try { $input.CopyTo($memory); $text = [Text.Encoding]::Unicode.GetString($memory.ToArray()) }
    finally { $input.Dispose(); $memory.Dispose() }
    if ($UpdateBaseUrl -and !$text.Contains($UpdateBaseUrl.TrimEnd('/') + '/')) { throw 'Bootstrapper does not install from the online update source' }
    Write-Output "PASS: signed manifests, version $Version, startup updates=$enabled, bootstrapper source and versioned publish directory"
} finally { $zip.Dispose() }
