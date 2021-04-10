#Requires -Module Carbon
function Start-WinGetSandbox {
  # The reason why this function so good is because I didn't write it, Felipe Santos (@felipecrs) did.
  # I added a couple helper parameters for the rest of this module, which is why this is here. 
  # Parse arguments

  Param(
    [Parameter(Position = 0, HelpMessage = "The Manifest to install in the Sandbox.")]
    [String] $Manifest,
    [Parameter(Position = 1, HelpMessage = "The script to run in the Sandbox.")]
    [ScriptBlock] $Script,
    [Parameter(HelpMessage = "The folder to map in the Sandbox.")]
    [String] $MapFolder = $pwd,
    [Parameter(HelpMessage = "Automatically run manifest and send output to file.")]
    [Switch] $auto
  )

  $ErrorActionPreference = "Stop"

  $mapFolder = (Resolve-Path -Path $MapFolder).Path

  if (-Not (Test-Path -Path $mapFolder -PathType Container)) {
    Write-Error -Category InvalidArgument -Message 'The provided MapFolder is not a folder.'
  }

  # Validate manifest file

  if (-Not [String]::IsNullOrWhiteSpace($Manifest)) {
    Write-Host '--> Validating Manifest'

    if (-Not (Test-Path -Path $Manifest)) {
      throw 'The Manifest file does not exist.'
    }

    $out = winget.exe validate $Manifest
    if (-Not $?) {
      throw 'Manifest validation failed.'
    }

    Write-Host
  }

  # Check if Windows Sandbox is enabled

  if (-Not (Get-Command 'WindowsSandbox' -ErrorAction SilentlyContinue)) {
    Write-Error -Category NotInstalled -Message @'
Windows Sandbox does not seem to be available. Check the following URL for prerequisites and further details:
https://docs.microsoft.com/en-us/windows/security/threat-protection/windows-sandbox/windows-sandbox-overview
You can run the following command in an elevated PowerShell for enabling Windows Sandbox:
$ Enable-WindowsOptionalFeature -Online -FeatureName 'Containers-DisposableClientVM'
'@
  }

  # Close Windows Sandbox

  $sandbox = Get-Process 'WindowsSandboxClient' -ErrorAction SilentlyContinue
  if ($sandbox) {
    Write-Host '--> Closing Windows Sandbox'

    $sandbox | Stop-Process
    Start-Sleep -Seconds 5

    Write-Host
  }
  Remove-Variable sandbox

  # Set dependencies

  $desktopAppInstaller = @{
    fileName = 'Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.appxbundle'
    url      = 'https://github.com/microsoft/winget-cli/releases/download/v-0.2.10771-preview/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.appxbundle'
    hash     = '11ecd121b5a19e07a545e84bc4dc182bd64a6233c9de137e10e3016d1527fc1e'
  }

  $vcLibs = @{
    fileName = 'Microsoft.VCLibs.140.00_14.0.27810.0_x64__8wekyb3d8bbwe.Appx'
    url      = 'https://raw.githubusercontent.com/felipecassiors/winget-pkgs/da8548d90369eb8f69a4738dc1474caaffb58e12/Tools/SandboxTest_Temp/Microsoft.VCLibs.140.00_14.0.27810.0_x64__8wekyb3d8bbwe.Appx'
    hash     = 'fe660c46a3ff8462d9574902e735687e92eeb835f75ec462a41ef76b54ef13ed'
  }

  $vcLibsUwp = @{
    fileName = 'Microsoft.VCLibs.x64.14.00.Desktop.appx'
    url      = 'https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx'
    hash     = '6602159c341bafea747d0edf15669ac72df8817299fbfaa90469909e06794256'
  }
  $settingsFile = @{
    fileName = 'settings.json'
    url      = 'https://gist.github.com/jedieaston/28db9c14a50f18bc9731a14b2b1fd265/raw/a2b117acae3ecdf0fd25f71bda7b3fc0af9921be/settings.json'
    hash     = '30DEFCF69EDAA7724FDBDDEBCA0CAD4BC027DDE0C2D349B7414972571EBAB94E'
  }

  $dependencies = @($desktopAppInstaller, $vcLibs, $vcLibsUwp, $settingsFile)

  # Initialize Temp Folder

  $tempFolderName = 'SandboxTest'
  $tempFolder = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath $tempFolderName

  New-Item $tempFolder -ItemType Directory -ErrorAction SilentlyContinue | Out-Null

  Get-ChildItem $tempFolder -Recurse -Exclude $dependencies.fileName | Remove-Item -Force -Recurse

  if (-Not [String]::IsNullOrWhiteSpace($Manifest)) {
    Copy-Item -Path $Manifest -Destination $tempFolder -Recurse
  }

  # Download dependencies

  Write-Host '--> Downloading dependencies'

  $desktopInSandbox = 'C:\Users\WDAGUtilityAccount\Desktop'

  $WebClient = New-Object System.Net.WebClient
  foreach ($dependency in $dependencies) {
    $dependency.file = Join-Path -Path $tempFolder -ChildPath $dependency.fileName
    $dependency.pathInSandbox = Join-Path -Path $desktopInSandbox -ChildPath (Join-Path -Path $tempFolderName -ChildPath $dependency.fileName)

    # Only download if the file does not exist, or its hash does not match.
    if (-Not ((Test-Path -Path $dependency.file -PathType Leaf) -And $dependency.hash -eq $(get-filehash $dependency.file).Hash)) {
      Write-Host @"
    - Downloading:
      $($dependency.url)
"@

      try {
        $WebClient.DownloadFile($dependency.url, $dependency.file)
      }
      catch {
        throw "Error downloading $($dependency.url) ."
      }
      if (-not ($dependency.hash -eq $(get-filehash $dependency.file).Hash)) {
        throw 'Hashes do not match, try gain.'
      }
    }
  }

  Write-Host

  # Create Bootstrap script

  # See: https://stackoverflow.com/a/14382047/12156188
  $bootstrapPs1Content = @'
function Update-Environment {
  $locations = 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment',
               'HKCU:\Environment'
  $locations | ForEach-Object {
    $k = Get-Item $_
    $k.GetValueNames() | ForEach-Object {
      $name  = $_
      $value = $k.GetValue($_)
      if ($userLocation -and $name -ieq 'PATH') {
        $Env:Path += ";$value"
      } else {
        Set-Item -Path Env:$name -Value $value
      }
    }
    $userLocation = $true
  }
}
'@

  $bootstrapPs1Content += @"
Write-Host @'
--> Installing WinGet
'@
Add-AppxPackage -Path '$($desktopAppInstaller.pathInSandbox)' -DependencyPath '$($vcLibsUwp.pathInSandbox)'
Copy-Item '$($settingsFile.pathInSandbox)' 'C:\Users\WDAGUtilityAccount\AppData\Local\Packages\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe\LocalState\settings.json'
Write-Host @'
Tip: you can type 'Update-Environment' to update your environment variables, such as after installing a new software.
'@
"@

  if (-Not [String]::IsNullOrWhiteSpace($Manifest)) {
    $manifestFileName = Split-Path $Manifest -Leaf
    if (-Not (Test-Path -Path $Manifest -PathType leaf)) {
        # It's a multi file manifest!!!  
        $manifestFileName += "\"
    }
    $manifestPathInSandbox = Join-Path -Path $desktopInSandbox -ChildPath (Join-Path -Path $tempFolderName -ChildPath ($manifestFileName))
    Write-Host $manifestPathInSandbox
    if ($auto) {
      $bootstrapPs1Content += @"
Write-Host @'
--> Installing the Manifest $manifestFileName
'@
Write-Host '$manifestPathInSandbox';
winget install -m '$manifestPathInSandbox' | Out-File .\tmp.log
Write-Host @'
--> Refreshing environment variables
'@
Update-Environment;
Write-Host @'
--> Getting list of installed applications...
'@
winget list | Add-Content .\tmp.log; 
New-Item .\done;
"@
    }
    else {
      $bootstrapPs1Content += @"
Write-Host @'
--> Installing the Manifest $manifestFileName
'@
Write-Host '$manifestPathInSandbox';
winget install -m '$manifestPathInSandbox'
Write-Host @'
--> Refreshing environment variables
'@
Update-Environment;
Write-Host @'
--> Getting list of installed applications...
'@
winget list;
"@
    }
  }

  if (-Not [String]::IsNullOrWhiteSpace($Script)) {
    $bootstrapPs1Content += @"
Write-Host @'
--> Running the following script:
{
$Script
}
'@
$Script
"@
  }

  $bootstrapPs1Content += @"
Write-Host
"@

  $bootstrapPs1FileName = 'Bootstrap.ps1'
  $bootstrapPs1Content | Out-File (Join-Path -Path $tempFolder -ChildPath $bootstrapPs1FileName)

  # Create Wsb file

  $bootstrapPs1InSandbox = Join-Path -Path $desktopInSandbox -ChildPath (Join-Path -Path $tempFolderName -ChildPath $bootstrapPs1FileName)
  $mapFolderInSandbox = Join-Path -Path $desktopInSandbox -ChildPath (Split-Path -Path $mapFolder -Leaf)

  $sandboxTestWsbContent = @"
<Configuration>
  <MappedFolders>
    <MappedFolder>
      <HostFolder>$tempFolder</HostFolder>
      <ReadOnly>true</ReadOnly>
    </MappedFolder>
    <MappedFolder>
      <HostFolder>$mapFolder</HostFolder>
    </MappedFolder>
  </MappedFolders>
  <LogonCommand>
  <Command>PowerShell Start-Process PowerShell -WindowStyle Maximized -WorkingDirectory '$mapFolderInSandbox' -ArgumentList '-ExecutionPolicy Bypass -NoExit -NoLogo -File $bootstrapPs1InSandbox'</Command>
  </LogonCommand>
</Configuration>
"@

  $sandboxTestWsbFileName = 'SandboxTest.wsb'
  $sandboxTestWsbFile = Join-Path -Path $tempFolder -ChildPath $sandboxTestWsbFileName
  $sandboxTestWsbContent | Out-File $sandboxTestWsbFile

  Write-Host @"
--> Starting Windows Sandbox, and:
    - Mounting the following directories:
      - $tempFolder as read-only
      - $mapFolder as read-and-write
    - Installing WinGet
"@

  if (-Not [String]::IsNullOrWhiteSpace($Manifest)) {
    Write-Host @"
    - Installing the Manifest $manifestFileName
    - Refreshing environment variables
    - Getting a list of installed applications and their product codes
"@
  }

  if (-Not [String]::IsNullOrWhiteSpace($Script)) {
    Write-Host @"
    - Running the following script:
{
$Script
}
"@
  }

  Write-Host

  WindowsSandbox $SandboxTestWsbFile  

}

function Get-WinGetManifestType {
  # Helper function. Given a folder, we see if it contains a multi-file or singleton manifest.
  param (
      [Parameter(Position = 0, Mandatory = $true, HelpMessage="Manifest to check on.")]
      [string]$manifestFolder
  )
  $ErrorActionPreference = "Stop"
  if(Test-Path -Path $manifestFolder -PathType Leaf) {
      throw "This isn't a folder, this is a file!"
  }
  
  foreach($i in (Get-ChildItem -Path $manifestFolder)) {
      $theSplitName = $i.Name.Split(".")
      if ($theSplitName.length -ge 3) {
         $manifest = Get-Content ($manifestFolder + "\" + $i) | ConvertFrom-Yaml -Ordered
         if (($manifest.ManifestType.ToLower() -eq "version") -or $manifest.ManifestType.ToLower() -eq "singleton") {
          break
         }
      }
  }
  if ($manifest.ManifestType.ToLower() -eq "version") {
      return "multifile"
  }
  elseif ($manifest.ManifestType.ToLower() -eq "singleton") {
      return "singleton"
  }
  else {
      throw "Unknown manifest type: " + $manifest.ManifestType
  }
}

function Get-URLFileHash {
    param (
        [Parameter(mandatory=$true)] 
        [string]$url
        )
    $ProgressPreference = 'SilentlyContinue' 
    $ErrorActionPreference = 'Stop'
    Invoke-WebRequest -UseBasicParsing -OutFile $env:TEMP\installer $url
    $hash = Get-FileHash $env:TEMP\installer
    Remove-Item $env:TEMP\installer
    Write-Host "The hash for the file at " $url "is "$hash.hash
    Set-Clipboard $hash.hash
    Write-Host "It has been written to the clipboard."
}

function Get-GitHubReleases {
    # Requires the GitHub CLI!
  Param(
   [Parameter(mandatory=$true, Position = 0, HelpMessage = "The repository to get the releases for, in the form user/repo.")]
   [String] $repo
  )
  $env:GH_REPO=$repo; gh release list -L 5; Remove-Item env:\GH_REPO
}

function Test-WinGetManifest {
  Param(
 [Parameter(mandatory=$true, Position = 0, HelpMessage = "The Manifest to test.")]
 [String] $manifest,
 [Parameter(HelpMessage="Keep the log file no matter what.")]
 [Switch]$keepLog,
 [Parameter(HelpMessage="Don't stop the function after 30 minutes.")]
 [Switch]$noStop
) 
  $ErrorActionPreference = 'Stop'
  Remove-Item ".\tmp.log" -ErrorAction "SilentlyContinue"
  Remove-Item ".\done" -ErrorAction "SilentlyContinue"
  $howManySeconds = 0
  Start-WinGetSandbox $manifest -auto
  Write-Host "Waiting for installation of" $manifest "to complete."
  while ((Test-Path -PathType Leaf ".\done") -ne $true) {
      # Write-Host "Waiting for file..."
      Start-Sleep -s 1
      if ($noStop -ne $true) {
        $howManySeconds += 1
        if ($howManySeconds -ge 1800) {
           Write-Host "Script timed out after 30 minutes. The sandbox will continue, but I'll stop looking for the log file."
           Write-Host "Next time, try the -noStop parameter or fix your manifest :D"
           return $false
        }
    }
  }
  Remove-Item ".\done" -ErrorAction "SilentlyContinue"

  $str = Select-String -Path ".\tmp.log" -Pattern "Successfully installed"
  if (-Not [string]::IsNullOrEmpty($str)) {
      $sandbox = Get-Process 'WindowsSandboxClient' -ErrorAction SilentlyContinue
      if ($sandbox) {
          Write-Host '--> Closing Windows Sandbox'

          $sandbox | Stop-Process
          Start-Sleep -Seconds 3

          Write-Host
          Remove-Variable sandbox
      }
      if ($keepLog -ne $true) {
          Remove-Item ".\tmp.log"
      }
      return $true
  }
  else {
      $err = (Select-String -Path ".\tmp.log" -Pattern "Installer failed" | Select-Object Line).Line
      Write-Host "Uh-oh! $err. The sandbox will stay open so you can investigate."
      # Get-Content ".\tmp.log"
      return $false
  }
}
function Update-WinGetManifest {
    # Now with support for 1.0 manifests! I hope.
    param (
        [Parameter(Mandatory=$true, Position=0, HelpMessage="Path to manifest to be updated.")]
        [string] $oldManifestFolder,
        [Parameter(Mandatory=$true, Position=1, HelpMessage="The new version number.")]
        [string] $newVersion,
        [Parameter(HelpMessage="If this manifest only supports a single architecture, the link to the new installer.")]
        [string] $newURL,
        [Parameter(HelpMessage="If this manifest supports multiple architectures, a hashtable of all of the URLs.")]
        [hashtable] $urlMap = @{},
        [Parameter(HelpMessage="Get the Product Code for each installer.")]
        [switch] $productCode,
        [Parameter(HelpMessage="Run the new manifest in a Windows Sandbox after it is created.")]
        [switch] $test,
        [Parameter(HelpMessage="Run the new manifest in a Windows Sandbox and shut it down when installation completes successfully.")]
        [switch] $silentTest,
        [Parameter(HelpMessage="Attempt to auto replace the version number in the Installer URLs.")]
        [switch] $autoReplaceURL,
        [Parameter(HelpMessage="Run New-WinGetCommit if this function completes successfully.")]
        [switch] $commit
    )
    # Just checking to ensure Carbon is available.
    Import-Module 'Carbon'
    $ProgressPreference = "SilentlyContinue"
    $ErrorActionPreference = "Stop"
    # Get the manifest's content.
    $type = Get-WinGetManifestType $oldManifestFolder
    if(($urlMap.Count -gt 1) -and ($type -eq "singleton")) {
      # If we were passed new architectures and were using a singleton manifest, we need a multifile one now.
      $converted = $true
      Write-Host "You're adding multiple architectures to a singleton manifest!" -Foreground Yellow
      Write-Host "Creating a multifile version of $oldManifestFolder..." -ForegroundColor Yellow
      $oldManifestFolder = Convert-WinGetSingletonToMultiFile $oldManifestFolder
      $type = "multifile"
      Write-Host "Conversion successful. Let's proceed." -ForegroundColor Yellow
    }
    else {
      $converted = $false
    }
    $newManifest = @{}
    foreach ($i in (Get-ChildItem -Path $oldManifestFolder)) {
        $content = (Get-Content ($oldManifestFolder + "\" + $i) | ConvertFrom-Yaml -Ordered)
        $newManifest.add($content.ManifestType, $content)
    }
    # Now for the updating!
    # Get the installers array and the PackageIdentifier
    if ($type -eq "multifile") {
        $installers = $newManifest.Installer.Installers
        $packageIdentifier = $newManifest.Installer.PackageIdentifier
        $oldVersion = $newManifest.Version.PackageVersion
    }
    else {
        $installers = $newManifest.singleton.Installers
        $packageIdentifier = $newManifest.singleton.PackageIdentifier
        $oldVersion = $newManifest.singleton.PackageVersion
    }
    # Make sure all architectures are lowercase in new manifest
    foreach($i in $installers) {
        $i.Architecture = $i.Architecture.ToLower()
    }
    foreach($i in $urlMap.Keys) {
        # Add non existing architectures.
        $inManifest = $false
        foreach($j in $installers) {
            if ($i.ToLower() -eq $j.Architecture) {
                $inManifest = $true
                break
            }
        }
        if(-Not $inManifest) {
            # Copy everything but the arch from the existing architecture.
            # Found at: https://www.powershellgallery.com/packages/PSTK/1.2.1/Content/Public%5CCopy-OrderedHashtable.ps1
            $temp = New-Object -TypeName System.Collections.Specialized.OrderedDictionary
            $MemoryStream     = New-Object -TypeName System.IO.MemoryStream
            $BinaryFormatter  = New-Object -TypeName System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
            $BinaryFormatter.Serialize($MemoryStream, $installers[0])
            $MemoryStream.Position = 0
            $temp = $BinaryFormatter.Deserialize($MemoryStream)
            $MemoryStream.Close()
            $temp.Architecture = $i.ToLower()
            $installers.add($temp)
        }
    }
    if (($installers.Length -eq 1) -And (-Not [String]::IsNullOrEmpty($newURL))) {
        $urlMap = @{$installers[0].Architecture = $newURL}
    }
    elseif ($autoReplaceURL) {
        $urlMap = @{}
        foreach ($i in $installers) {
            $urlMap[$i.Architecture] = $i.InstallerUrl -Replace $oldVersion, $newVersion
            Write-Host "Auto-replace for arch" $i.Architecture "resulted in URL "$urlMap[$i.Architecture] -ForegroundColor Yellow
        }
    }
    if ($urlMap.Count -ne $installers.Length) {
        foreach($i in $installers) {
            if (-Not ($urlMap.ContainsKey($i.Architecture))) {
                # The user didn't specify this URL.
                Write-Host "What is the Installer URL for architecture "$i.Architecture"?"
                $urlMap[$i.Architecture] = Read-Host $i.InstallerUrl
            }
        }
        # If the user provided new architectures, we don't need to worry about it.
    }
    
    # Let's download the installers and get the needed info.
    foreach($i in $installers) {
        $url = $urlMap[$i.Architecture]
        $i.InstallerUrl = $url
        Write-Host 'Downloading installer for architecture'$i.Architecture'...' -ForegroundColor Yellow
        Invoke-WebRequest -UseBasicParsing -OutFile $env:TEMP\installer $url
        $i.InstallerSha256 = (Get-FileHash $env:TEMP\installer).Hash
        # Get Product Code if necessary
        if($i.Contains("ProductCode") -Or ($productCode)) {
            $i.ProductCode = '{' + (((Get-MSI $env:TEMP\installer).ProductCode).ToString()).ToUpper() + '}'
        }
    }
    # put the new installers array in the right place.
    if ($type -eq "multifile") {
        $newManifest.Installer.Installers = $installers
    }
    else {
        $newManifest.singleton.Installers = $installers
    }
    Write-Host "New installer downloads complete!" -ForegroundColor Green
    # Begin writing the files. This'll be ugly.
    # First, set the version numbers in every place they exist in a multifile manifest.
    if ($type -eq "multifile") {
        foreach($i in $newManifest.Keys) {
            $newManifest[$i]["PackageVersion"] = $newVersion
        }
    }
    else {
        $newManifest.singleton.PackageVersion = $newVersion
    }
    # Now let's get these back to YAML.
    $path = ".\" + $newVersion + "\"
    New-Item -Type Directory $path -Force 

    foreach($i in $newManifest.Keys) {
        if (($newManifest[$i].ManifestType.ToLower() -eq "singleton") -or ($newManifest[$i].ManifestType.ToLower() -eq "version")) {
            $fileName = $path + $packageIdentifier + ".yaml"
        }
        elseif($newManifest[$i].ManifestType.ToLower() -eq "installer") {
            $fileName = $path + $packageIdentifier + ".installer.yaml"
        }
        elseif(($newManifest[$i].ManifestType.ToLower() -eq "locale") -Or ($newManifest[$i].ManifestType.ToLower() -eq "defaultlocale")) {
            $fileName = $path + $packageIdentifier + ".locale." + $newManifest[$i].PackageLocale + ".yaml"
        }
        else {
            throw $newManifest[$i].ManifestType.ToLower() + " is an unknown type."
        }
        # Add schema info for IntelliSense
        $fileContent = '# yaml-language-server: $schema=https://aka.ms/winget-manifest.' + $newManifest[$i].ManifestType.ToLower() + '.1.0.0.schema.json' + "`r`n"
        # And the manifest...
        $fileContent += ($newManifest[$i] | ConvertTo-Yaml).replace("'", '"')
        [System.Environment]::CurrentDirectory = (Get-Location).Path
        [System.IO.File]::WriteAllLines($fileName, $fileContent)
        # $fileContent | Out-File -Encoding "utf8" -FilePath $fileName
        Write-Host $fileName "written." -ForegroundColor Green
    }
    # Get rid of the temporary manifest we created if we needed a conversion.
    if($converted) {
      Remove-Item -Recurse $oldManifestFolder 
    }
    # Validate this thing.
    winget validate $path | Out-Null
    if($LASTEXITCODE -ne 0) {
        throw "Manifest validation failed."
    }
    if ($test) {
        Start-WinGetSandbox $path
        return $true
    }
    elseif ($silentTest) {
      $testSuccess = Test-WinGetManifest $path
      if ($testSuccess) {
        Write-Host "Manifest successfully installed in Windows Sandbox!"
        if ($commit) {
          New-WinGetCommit $path
        }
        return $true
      }
      else {
        Write-Host "Manifest install failed/timed out. For more info check .\tmp.log."
        return $false
      }
    }
    else {
      return $true
    }

}
function Convert-WinGetSingletonToMultiFile {
    # Given a folder that contains a WinGet singleton manifest, this function creates a 
    # bare minimum multifile manifest in a different folder.
    param (
        [Parameter(Position=0, Mandatory=$true, HelpMessage="The manifest you wish to convert.")]
        [string]$oldManifestFolder,
        [Parameter(HelpMessage="Replace the old manifest with this new one.")]
        [switch]$overwrite
    )
    $ErrorActionPreference = "Stop"
    $versionSchema = Invoke-WebRequest "https://aka.ms/winget-manifest.version.1.0.0.schema.json" | ConvertFrom-Json
    $localeSchema = Invoke-WebRequest "https://aka.ms/winget-manifest.defaultlocale.1.0.0.schema.json" | ConvertFrom-Json
    $installerSchema = Invoke-WebRequest "https://aka.ms/winget-manifest.installer.1.0.0.schema.json" | ConvertFrom-Json
    $requiredKeys = $versionSchema.required + $localeSchema.required + $installerSchema.required
    $type = Get-WinGetManifestType $oldManifestFolder
    if ($type -ne "singleton") {
        throw "This folder does not contain a singleton manifest."
    }
    $currentManifest = (Get-ChildItem -Path $oldManifestFolder)[0] | Get-Content | ConvertFrom-Yaml -Ordered
    $newManifest = @{}
    # Add required metadata.
    $currentManifest["PackageVersion"] = [string]$currentManifest["PackageVersion"]
    $newManifest["version"] = [Ordered]@{
        PackageIdentifier = $currentManifest["PackageIdentifier"];
        PackageVersion = $currentManifest["PackageVersion"];
        DefaultLocale = $currentManifest["PackageLocale"];
        ManifestType = "version";
        ManifestVersion = "1.0.0";
    }
    $newManifest["defaultLocale"] = [Ordered]@{
        PackageIdentifier = $currentManifest["PackageIdentifier"];
        PackageVersion = $currentManifest["PackageVersion"];
        PackageLocale = $currentManifest["PackageLocale"];
        Publisher = $currentManifest["Publisher"];
        PackageName = $currentManifest["PackageName"];
        License = $currentManifest["License"];
        ShortDescription = $currentManifest["ShortDescription"];
        ManifestType = "defaultLocale";
        ManifestVersion = "1.0.0";
    }
    $newManifest["installer"] = [Ordered]@{
        PackageIdentifier = $currentManifest["PackageIdentifier"];
        PackageVersion = $currentManifest["PackageVersion"];
        Installers = $currentManifest["Installers"];
        ManifestType = "installer";
        ManifestVersion = "1.0.0";
    }
    # Now, put any keys we missed in the right spots.
    $extraKeys = [Ordered]@{}
    $extraKeys["version"] = [Ordered]@{}
    $extraKeys["defaultLocale"] = [Ordered]@{}
    $extraKeys["installer"] = [Ordered]@{}
    foreach($i in $currentManifest.Keys) {
        if (-Not ($requiredKeys -contains $i) ) {
            if ($i -in $versionSchema.properties.PSObject.Properties.Name) {
                $extraKeys["version"].add($i, $currentManifest[$i])
            }
            elseif($i -in $localeSchema.properties.PSObject.Properties.Name) {
                $extraKeys["defaultLocale"].add($i, $currentManifest[$i])
            }
            elseif($i -in $installerSchema.properties.PSObject.Properties.Name) {
                $extraKeys["installer"].add($i, $currentManifest[$i])
            }
        }
    }
    # Now, add the extra keys back to the new Manifest.
    # Since you can't add multiple keys at a certain place in a ordered dict in PowerShell, I had to do this.
    $count = 3
    foreach($i in $extraKeys["version"].Keys) {
        $newManifest["version"].Insert($count, $i, $extraKeys["version"][$i])
        $count++;
    }
    $count = 3
    foreach($i in $extraKeys["defaultLocale"].Keys) {
        $newManifest["defaultLocale"].Insert($count, $i, $extraKeys["defaultLocale"][$i])
        $count++;
    }
    $count = 2
    foreach($i in $extraKeys["installer"].Keys) {
        $newManifest["installer"].Insert($count, $i, $extraKeys["installer"][$i])
        $count++;
    }
    # Now we can write the files.
    $path = ".\" + $currentManifest["PackageVersion"] + "-multiFile\"
    New-Item -Type Directory $path -Force | Out-Null

    # Copied (with modifications) from Update-WinGetManifest. I should break this out into a function...
    foreach($i in $newManifest.Keys) {
        if ($newManifest[$i].ManifestType.ToLower() -eq "version") {
            $fileName = $path + $currentManifest["PackageIdentifier"] + ".yaml"
        }
        elseif($newManifest[$i].ManifestType.ToLower() -eq "installer") {
            $fileName = $path + $currentManifest["PackageIdentifier"] + ".installer.yaml"
        }
        elseif($newManifest[$i].ManifestType.ToLower() -eq "defaultlocale") {
            $fileName = $path + $currentManifest["PackageIdentifier"] + ".locale." + $newManifest[$i].PackageLocale + ".yaml"
        }
        else {
            throw $newManifest[$i].ManifestType.ToLower() + " is an unknown type."
        }
        # Add schema info for IntelliSense
        $fileContent = '# yaml-language-server: $schema=https://aka.ms/winget-manifest.' + $newManifest[$i].ManifestType.ToLower() + '.1.0.0.schema.json' + "`r`n"
        # And the manifest...
        $fileContent += ($newManifest[$i] | ConvertTo-Yaml).replace("'", '"')
        $fileContent | Out-File -Encoding "utf8" -FilePath $fileName
        Write-Host $fileName "written." -ForegroundColor Green
    }
    winget validate $path | Out-Null
    if($LASTEXITCODE -ne 0) {
        throw "Manifest validation failed. Check the source manifest to ensure it's good and try again."
    }
    if($overwrite) {
      Remove-Item -Recurse $oldManifestFolder
      Move-Item $path $oldManifestFolder
      $path = $oldManifestFolder
    }
    Write-Host "All done with the conversion. The new manifest can be found in "$path"." -ForegroundColor Green
    Write-Host "Please check it before committing." -ForegroundColor Green
    return $path
}

function New-WinGetCommit {
  # Autocreates a new commit. Don't use this unless you're really lazy.
  # Make sure that you have an upstream branch set too (to microsoft/winget-pkgs), or creating a new branch may fail.
  param(
    [Parameter(mandatory=$true, Position=0, HelpMessage="The manifest to commit.")]
    [string] $manifest,
    [Parameter(HelpMessage="Use the currently checked out branch, instead of making a new one.")]
    [switch] $currentBranch,
    [Parameter(HelpMessage="The commit message, if you don't want the automatically generated one.")]
    [string] $customMessage
  )
  $ErrorActionPreference = "Stop"
  if(Test-Path -Path $manifest -PathType Leaf) {
      throw "This isn't a folder, this is a file!"
  }
  # $content = Get-Content $manifest | ConvertFrom-Yaml
  foreach($i in (Get-ChildItem -Path $manifest)) {
      $theSplitName = $i.Name.Split(".")
      if ($theSplitName.length -ge 3) {
         $content = Get-Content ($manifest + "\" + $i) | ConvertFrom-Yaml -Ordered
         if ($content.ManifestType.ToLower() -eq "version") {
            $content = Get-Content ($manifest + "\" + $content.PackageIdentifier + ".locale." + $content.DefaultLocale + ".yaml") | ConvertFrom-Yaml -Ordered
            break
         }
         elseif($content.ManifestType.ToLower() -eq "singleton") {
           break
         }
         
      }
  }
  if([string]::IsNullOrEmpty($customMessage) -or [string]::IsNullOrWhiteSpace($customMessage)) {
   $commitMessage = "Added " + $content.PackageName + " version " + $content.PackageVersion + "."
  }
  else {
    $commitMessage = $customMessage
  }
  if (-Not $currentBranch) {
    $branchName = $content.PackageIdentifier + "-" + $content.PackageVersion
    git fetch --all
    git checkout -b "$branchName" upstream/master
    if($LASTEXITCODE -ne 0) {
      # The branch already exists.
      git checkout "$branchName"
    }
  }
  git add "$manifest"
  git commit -m $commitMessage
}
function Get-WinGetApplicationCurrentVersion {
  # Uses the winget cli to get the current version of an app in the repo. 
  # Useful for seeing if something needs an update.
  param (
     [Parameter(Position=0, Mandatory=$true, HelpMessage="The manifest ID to check.")]
     [string]$id 
  )

  winget source update | Out-Null
  $littleManifest = (winget show $id)
  if ($LASTEXITCODE -ne 0) {
    throw "Couldn't find manifest " + $id
  }
  try { $version = ($littleManifest | Select-Object -Skip 2 | Out-String | ConvertFrom-Yaml).version }
  catch { $version = ($littleManifest | Select-Object -Skip 2 | Out-String | ConvertFrom-Yaml).version }
  return $version
}