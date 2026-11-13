param([Parameter(Mandatory)][string]$Fixtures, [string]$Wix)
$ErrorActionPreference='Stop'
$repo=Split-Path $PSScriptRoot -Parent
if(!$Wix){$Wix=Join-Path $repo 'work\tools\wix\wix.exe'}
$Fixtures=(Resolve-Path -LiteralPath $Fixtures).Path
$id=[guid]::NewGuid().ToString('N')
$evidence=Join-Path $repo ('work\versioned-msi-check-'+$id)
$target=Join-Path $env:LOCALAPPDATA ('ExcelNavigatorPane\Tests\VersionedMsi\'+$id)
$testKey='Software\ExcelNavigatorPane\Tests\VersionedMsi\'+$id
$family=[guid]::NewGuid().ToString().ToUpperInvariant()
$legacyFamily=[guid]::NewGuid().ToString().ToUpperInvariant()
$products=@{};$packages=@{};$locked=$null
$installer=New-Object -ComObject WindowsInstaller.Installer
New-Item -ItemType Directory -Path $evidence -Force | Out-Null
function Require($condition,[string]$reason){if(!$condition){throw $reason}}
function Msi([string]$operation,[string]$version,[string]$extra=''){
    $subject=if($operation -eq 'install'){'/i "'+$packages[$version]+'"'}else{'/x '+$products[$version]}
    $log=Join-Path $evidence ($operation+'-'+$version+'-'+[guid]::NewGuid().ToString('N')+'.log')
    $process=Start-Process -FilePath (Join-Path $env:SystemRoot 'System32\msiexec.exe') -ArgumentList ($subject+' /qn /norestart REBOOT=ReallySuppress MSIRESTARTMANAGERCONTROL=Disable /l*v "'+$log+'" '+$extra) -WindowStyle Hidden -PassThru
    Require ($process.WaitForExit(60000)) 'Isolated MSI timeout; inspect before retry'
    return $process.ExitCode
}
try{
    foreach($version in @('legacy','1.0.26.0','1.0.27.0','1.0.28.0','1.0.29.0')){
        $code='{'+[guid]::NewGuid().ToString().ToUpperInvariant()+'}';$products[$version]=$code
        $sourceVersion=if($version -eq 'legacy'){'1.0.26.0'}elseif($version -eq '1.0.29.0'){'1.0.28.0'}else{$version}
        $source=Join-Path $Fixtures ('Versions\'+$sourceVersion)
        Require (Test-Path -LiteralPath (Join-Path $source 'ExcelNavigatorPane.dll')) 'Signed fixture missing'
        $packageSource=Join-Path $evidence ($version+'.wxs')
        if($version -eq 'legacy'){
            $payload=[Security.SecurityElement]::Escape((Join-Path $source 'ExcelNavigatorPane.dll'))
            $location=[Security.SecurityElement]::Escape($target)
            $componentGuid=[guid]::NewGuid().ToString().ToUpperInvariant()
            $xml=@"
<Wix xmlns="http://wixtoolset.org/schemas/v4/wxs"><Package Name="ExcelNavigator isolated legacy $id" Manufacturer="Test" Version="1.0.26" ProductCode="$code" UpgradeCode="$legacyFamily" Scope="perUser"><MediaTemplate EmbedCab="yes"/><Property Id="ARPSYSTEMCOMPONENT" Value="1"/><Property Id="INSTALLFOLDER" Value="$location"/><StandardDirectory Id="LocalAppDataFolder"><Directory Id="INSTALLFOLDER"><Component Id="Legacy" Guid="$componentGuid"><File Source="$payload"/><RegistryValue Root="HKCU" Key="$testKey\legacy" Name="Installed" Value="1" Type="integer" KeyPath="yes"/></Component></Directory></StandardDirectory><Feature Id="Main"><ComponentRef Id="Legacy"/></Feature></Package></Wix>
"@
            $xml | Set-Content $packageSource
        }else{
            # Clone production sequencing/upgrade rules, substituting only identity, per-user scope and test directory.
            [xml]$xml=Get-Content (Join-Path $repo 'installer\Product.wxs') -Raw
            $package=$xml.Wix.Package;$package.SetAttribute('Name',"ExcelNavigator isolated versioned $id $version")
            $package.SetAttribute('Scope','perUser');$package.SetAttribute('ProductCode',$code);$package.SetAttribute('UpgradeCode',$family)
            foreach($upgrade in $package.Upgrade){$upgrade.Id=if($upgrade.Id -eq 'A14219F5-2541-441C-BAEC-606015D32B55'){$legacyFamily}else{$family}}
            $property=$xml.CreateElement('Property',$xml.DocumentElement.NamespaceURI);$property.SetAttribute('Id','INSTALLFOLDER');$property.SetAttribute('Value',$target);[void]$package.AppendChild($property)
            if($version -eq '1.0.29.0'){
                $action=$xml.CreateElement('CustomAction',$xml.DocumentElement.NamespaceURI)
                foreach($entry in @{Id='TestFailure';Directory='SystemFolder';ExeCommand='"[SystemFolder]cmd.exe" /d /c exit 1';Execute='deferred';Return='check';Impersonate='yes'}.GetEnumerator()){$action.SetAttribute($entry.Key,$entry.Value)}
                [void]$package.AppendChild($action)
                $sequence=$package.InstallExecuteSequence;$step=$xml.CreateElement('Custom',$xml.DocumentElement.NamespaceURI)
                $step.SetAttribute('Action','TestFailure');$step.SetAttribute('After','InstallFiles');$step.SetAttribute('Condition','NOT Installed');[void]$sequence.AppendChild($step)
            }
            $xml.Save($packageSource)
            $components=@();$refs=@();$n=0
            foreach($file in Get-ChildItem -LiteralPath $source -File){
                $component='Payload'+($n++);$escaped=[Security.SecurityElement]::Escape($file.FullName)
                $componentGuid=[guid]::NewGuid().ToString().ToUpperInvariant()
                $components+="<Component Id=`"$component`" Guid=`"$componentGuid`"><File Source=`"$escaped`"/><RegistryValue Root=`"HKCU`" Key=`"$testKey\$version\$component`" Name=`"Installed`" Value=`"1`" Type=`"integer`" KeyPath=`"yes`"/></Component>"
                $refs+="<ComponentRef Id=`"$component`"/>"
            }
            $include=Join-Path $evidence ($version+'.wxi');"<Include xmlns=`"http://wixtoolset.org/schemas/v4/wxs`">$($components -join '')</Include>" | Set-Content $include
            $group=Join-Path $evidence ($version+'-group.wxs');"<Wix xmlns=`"http://wixtoolset.org/schemas/v4/wxs`"><Fragment><ComponentGroup Id=`"PayloadComponents`">$($refs -join '')</ComponentGroup></Fragment></Wix>" | Set-Content $group
        }
        $msi=Join-Path $evidence ($version+'.msi');$packages[$version]=$msi
        if($version -eq 'legacy'){& $Wix build $packageSource -arch x64 -o $msi *> (Join-Path $evidence ($version+'-build.log'))}
        else{& $Wix build $packageSource $group -arch x64 -d "MsiVersion=$($version -replace '\.0$','')" -d "ReleaseVersion=$version" -d 'ProgramFiles=LocalAppDataFolder' -d "PayloadInclude=$include" -o $msi *> (Join-Path $evidence ($version+'-build.log'))}
        if($LASTEXITCODE -ne 0){Get-Content (Join-Path $evidence ($version+'-build.log')) -Tail 12;throw 'Isolated package build failed'}
    }
    $products | ConvertTo-Json | Set-Content (Join-Path $evidence 'products.json')
    Require ((Msi install legacy) -eq 0) 'Legacy install failed'
    Require ((Msi install '1.0.27.0') -eq 1603) 'Bare payload MSI was not rejected'
    Require ($installer.ProductState($products.legacy) -eq 5) 'Rejected direct MSI disturbed legacy'
    Require ((Msi install '1.0.27.0' 'EXCELNAV_SETUP=1') -eq 0) 'Migration install failed'
    Require ($installer.ProductState($products.legacy) -eq -1) 'Legacy was not removed by production upgrade rule'
    $old=Join-Path $target 'Versions\1.0.27.0\ExcelNavigatorPane.dll'
    $before=(Get-FileHash -LiteralPath $old).Hash
    $locked=[IO.File]::Open($old,'Open','Read','Read')
    Require ((Msi install '1.0.28.0' 'EXCELNAV_SETUP=1') -eq 0) 'Live versioned update failed'
    Require ($installer.ProductState($products.'1.0.27.0') -eq 5 -and (Get-FileHash -LiteralPath $old).Hash -eq $before) 'New MSI removed or changed held old payload'
    $repairFile=Join-Path $target 'Versions\1.0.28.0\ExcelNavigatorPane.dll'
    $repairHash=(Get-FileHash -LiteralPath $repairFile).Hash
    Remove-Item -LiteralPath $repairFile
    Require ((Msi install '1.0.28.0' 'EXCELNAV_SETUP=1 REINSTALL=ALL REINSTALLMODE=vomus') -eq 0) 'Same-version repair failed'
    Require ((Get-FileHash -LiteralPath $repairFile).Hash -eq $repairHash) 'Repair did not restore the missing payload'
    Require ((Msi install '1.0.26.0' 'EXCELNAV_SETUP=1') -eq 1603) 'Downgrade was accepted'
    Require ((Msi install '1.0.29.0' 'EXCELNAV_SETUP=1') -eq 1603) 'Injected install failure was accepted'
    Require ($installer.ProductState($products.'1.0.29.0') -eq -1 -and !(Test-Path (Join-Path $target 'Versions\1.0.29.0\ExcelNavigatorPane.dll'))) 'Failed MSI did not roll back'
    'PASS: production MSI rules: direct-payload gate, legacy migration, locked-old-version coexistence, downgrade rejection and failed-install rollback'
}finally{
    if($locked){$locked.Dispose()}
    foreach($version in @($products.Keys)){
        if($installer.ProductState($products[$version]) -ne -1){Require ((Msi uninstall $version 'EXCELNAV_SETUP=1') -eq 0) 'Test uninstall failed'}
        Require ($installer.ProductState($products[$version]) -eq -1) 'Test MSI product left behind'
    }
    $resolved=[IO.Path]::GetFullPath($target)
    $expected=[IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA 'ExcelNavigatorPane\Tests\VersionedMsi'))+[IO.Path]::DirectorySeparatorChar
    Require ($resolved.StartsWith($expected,[StringComparison]::OrdinalIgnoreCase)) 'Unsafe test cleanup path'
    if(Test-Path -LiteralPath $resolved){Remove-Item -LiteralPath $resolved -Recurse}
    if(Test-Path -LiteralPath ('HKCU:\'+$testKey)){Remove-Item -LiteralPath ('HKCU:\'+$testKey) -Recurse}
    'Evidence: '+$evidence
}
