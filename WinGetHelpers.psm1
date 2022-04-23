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
    [Switch] $auto,
    [Parameter(HelpMessage = "Check to make sure these values match the ARP table values")]
    [Hashtable] $metadata = $null
  )

  if ($null -ne $metadata) {
    # Ugly hack to unravel the metadata. Fix this easton.
    $displayVersion = $metadata.PackageVersion
    $displayName = $metadata.PackageName
    $publisher = $metadata.Publisher
  }

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

  $apiJson = (Invoke-WebRequest 'https://api.github.com/repos/microsoft/winget-cli/releases' -UseBasicParsing | ConvertFrom-Json)[0]


  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
  $WebClient = New-Object System.Net.WebClient

  function Get-LatestUrl {
    (($apiJson).assets | Where-Object { $_.name -match '^Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle$' }).browser_download_url
  }

  function Get-LatestHash {
    $shaUrl = (($apiJson).assets | Where-Object { $_.name -match '^Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.txt$' }).browser_download_url
    $shaFile = Join-Path -Path $env:TEMP -ChildPath 'Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.txt'
    $WebClient.DownloadFile($shaUrl, $shaFile)
    Get-Content $shaFile
}

  # Hide the progress bar of Invoke-WebRequest
  # $oldProgressPreference = $ProgressPreference
  $ProgressPreference = 'SilentlyContinue'

  $desktopAppInstaller = @{
    fileName = 'Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle'
    url      = $(Get-LatestUrl)
    hash     = $(Get-LatestHash)
  }
  $uiLibsUwp = @{
    fileName = 'Microsoft.UI.Xaml.2.7.zip'
    url = 'https://www.nuget.org/api/v2/package/Microsoft.UI.Xaml/2.7.0'
    hash = "422FD24B231E87A842C4DAEABC6A335112E0D35B86FAC91F5CE7CF327E36A591"
  }

  $vcLibsUwp = @{
    fileName = 'Microsoft.VCLibs.x64.14.00.Desktop.appx'
    url      = 'https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx'
    hash     = 'A39CEC0E70BE9E3E48801B871C034872F1D7E5E8EEBE986198C019CF2C271040'
  }
  $settingsFile = @{
    fileName = 'settings.json'
    url      = 'https://gist.githubusercontent.com/jedieaston/28db9c14a50f18bc9731a14b2b1fd265/raw/c0370686c60b6cca8566a93112f7689a55d34d67/settings.json'
    hash     = 'BBDC9B3CA350576FD292AB7D7DF224B8D2B8E1DF387F4C776B57A9D1D0D1BED5'
  }

  $dependencies = @($desktopAppInstaller, $vcLibsUwp, $settingsFile, $uiLibsUwp)

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
        throw 'Hashes do not match, try gain. (Expected ' + $dependency.hash + ', got ' + $(get-filehash $dependency.file).Hash
      }
    }
  }

  Write-Host

  # Create Bootstrap script

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
Expand-Archive '$($uiLibsUwp.pathInSandbox)' ~\Desktop\xaml\
Add-AppxPackage ~\Desktop\xaml\tools\AppX\x64\Release\Microsoft.UI.Xaml.2.7.appx
Add-AppxPackage -Path '$($desktopAppInstaller.pathInSandbox)' -DependencyPath '$($vcLibsUwp.pathInSandbox)'
Copy-Item '$($settingsFile.pathInSandbox)' 'C:\Users\WDAGUtilityAccount\AppData\Local\Packages\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe\LocalState\settings.json'
Write-Host @'
Tip: you can type 'Update-Environment' to update your environment variables, such as after installing a new software.
Write-Host @'
--> Changing winget settings for testing (disabling msstore and enabling local manifest installation).
'@
winget settings --Enable LocalManifestFiles
winget source remove msstore

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
  if ($null -ne $metadata ) {
   $bootstrapPs1Content += @"
  Write-Host --> Checking the ARP table. 
    if (Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate  | Where-Object DisplayVersion -eq '$displayVersion' | Where-Object DisplayName -eq '$displayName' | Where-Object Publisher -eq '$publisher') {
       if(Test-Path .\tmp.log) {
         'ARP check went great!' | Add-Content .\tmp.log ;
         Write-Host HEY!
       }
       else {
         Write-Host 'ARP check went great!' -ForegroundColor Green;
       }
      }
    elseif (Get-ItemProperty HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate  | Where-Object DisplayVersion -eq '$displayVersion' | Where-Object DisplayName -eq '$displayName' | Where-Object Publisher -eq '$publisher') {
       if(Test-Path .\tmp.log) {
         'ARP check went great!' | Add-Content .\tmp.log ;
         Write-Host HEY!
       }
       else {
         Write-Host 'ARP check went great!' -ForegroundColor Green;
       }
    }
    elseif (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate  | Where-Object DisplayVersion -eq '$displayVersion' | Where-Object DisplayName -eq '$displayName' | Where-Object Publisher -eq '$publisher') {
       if(Test-Path .\tmp.log) {
         'ARP check went great!' | Add-Content .\tmp.log ;
         Write-Host HEY!
       }
       else {
         Write-Host 'ARP check went great!' -ForegroundColor Green;
       }
    }
    else {
        if(Test-Path .\tmp.log) {
         'ARP mismatch detected.' | Add-Content .\tmp.log ;
       }
       else {
         Write-Host 'ARP mismatch detected.' -ForegroundColor Red;
       }
  }
  ;
"@
  }
  if ($auto) {
    $bootstrapPs1Content += @"
      New-Item .\done;
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
  <vGPU>Enable</vGPU>
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
    [Parameter(Position = 0, Mandatory = $true, HelpMessage = "Manifest to check on.")]
    [string]$manifestFolder
  )
  $ErrorActionPreference = "Stop"
  if (Test-Path -Path $manifestFolder -PathType Leaf) {
    throw "This isn't a folder, this is a file!"
  }

  foreach ($i in (Get-ChildItem -Path $manifestFolder -File)) {
    $manifest = Get-Content ($manifestFolder + "\" + $i) | ConvertFrom-Yaml -Ordered
    if (($manifest.ManifestType.ToLower() -eq "version") -or $manifest.ManifestType.ToLower() -eq "singleton") {
      break
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
  <#
      .SYNOPSIS
          Given a URL, this function returns the SHA256 hash of the file located at that address.
      .DESCRIPTION
          Given a URL, this function returns the SHA256 hash of the file located at that address.
          Optionally, with the -Clipboard parameter, the hash will be written to the clipboard so that it can be pasted
          into another application.
      .INPUTS
          Nothing can be piped (yet).
      .OUTPUTS
          A SHA256 hash of the file at the URL.
      .EXAMPLE
          Get-URLFileHash https://dl.google.com/tag/s/dl/chrome/install/googlechromestandaloneenterprise64.msi

          Returns the SHA256 hash of the googlechromesandaloneenterprise64.msi.
      #>
  param (
    # The URL to get the hash for.
    [Parameter(mandatory = $true)]
    [string]$url,
    # Write the SHA256 hash to the clipboard.
    [Parameter(HelpMessage = "Write to clipboard?")]
    [switch]$clipboard
  )
  $ProgressPreference = 'SilentlyContinue'
  $ErrorActionPreference = 'Stop'
  Invoke-WebRequest -UseBasicParsing -OutFile $env:TEMP\installer $url
  $hash = Get-FileHash $env:TEMP\installer
  Remove-Item $env:TEMP\installer
  if ($clipboard) {
    Set-Clipboard $hash.hash
    Write-Host "The hash has been written to the clipboard."
  }
  return $hash.hash
}

function Get-GitHubReleases {
  # Requires the GitHub CLI!
  Param(
    [Parameter(mandatory = $true, Position = 0, HelpMessage = "The repository to get the releases for, in the form user/repo.")]
    [String] $repo
  )
  $env:GH_REPO = $repo; gh release list -L 5; Remove-Item env:\GH_REPO
}

function Test-WinGetManifest {
  Param(
    [Parameter(mandatory = $true, Position = 0, HelpMessage = "The Manifest to test.")]
    [String] $manifest,
    [Parameter(HelpMessage = "Keep the log file no matter what.")]
    [Switch]$keepLog,
    [Parameter(HelpMessage = "Don't stop the function after 30 minutes.")]
    [Switch]$noStop,
    [Parameter(HelpMessage = "Check these values against the ARP table.")]
    [hashtable] $metadata = $null
  )
  $ErrorActionPreference = 'Stop'
  Remove-Item ".\tmp.log" -ErrorAction "SilentlyContinue"
  Remove-Item ".\done" -ErrorAction "SilentlyContinue"
  $howManySeconds = 0
  Start-WinGetSandbox $manifest -auto -metadata $metadata
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
    if ($null -ne $metadata) {
      $str = Select-String -Path ".\tmp.log" -Pattern "ARP mismatch detected."
      if (-Not [string]::IsNullOrEmpty($str)) {
        Write-Host "Uh-oh! You wanted me to look for ARP errors, and I found one." -ForegroundColor Red
        Write-Host "The sandbox will stay open so you can investigate."
        return $false
      }
    }
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
    [Parameter(Mandatory = $true, Position = 0, HelpMessage = "Path to manifest to be updated.")]
    [string] $oldManifestFolder,
    [Parameter(Mandatory = $true, Position = 1, HelpMessage = "The new version number.")]
    [string] $newVersion,
    [Parameter(HelpMessage = "If this manifest only supports a single architecture, the link to the new installer.")]
    [string] $newURL,
    [Parameter(HelpMessage = "If this manifest supports multiple architectures, a hashtable of all of the URLs.")]
    [hashtable] $urlMap = @{},
    [Parameter(HelpMessage = "Get the Product Code for each installer.")]
    [switch] $productCode,
    [Parameter(HelpMessage = "Run the new manifest in a Windows Sandbox after it is created.")]
    [switch] $test,
    [Parameter(HelpMessage = "Run the new manifest in a Windows Sandbox and shut it down when installation completes successfully.")]
    [switch] $silentTest,
    [Parameter(HelpMessage = "Attempt to auto replace the version number in the Installer URLs.")]
    [switch] $autoReplaceURL,
    [Parameter(HelpMessage = "Run New-WinGetCommit if this function completes successfully.")]
    [switch] $commit,
    [Parameter(HelpMessage = "Check the applicable metadata values against the ARP table.")]
    [switch] $metadataCheck,
    [Parameter(HelpMessage = "The release date for this manifest (default: today")]
    [string] $releaseDate = (Get-Date -f "yyyy-MM-dd"),
    [Parameter(HelpMessage = "Remove the old manifest and replace it with this one.")]
    [switch] $update
  )
  # Just checking to ensure Carbon is available.
  Import-Module 'Carbon'
  $ProgressPreference = "SilentlyContinue"
  $ErrorActionPreference = "Stop"
  # Get the manifest's content.
  $type = Get-WinGetManifestType $oldManifestFolder
  if (($urlMap.Count -gt 1) -and ($type -eq "singleton")) {
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
  $arpValues = $null
  foreach ($i in (Get-ChildItem -Path $oldManifestFolder)) {
    $content = (Get-Content ($oldManifestFolder + "\" + $i) -Encoding UTF8 | ConvertFrom-Yaml -Ordered)
    if ($content.ManifestType -eq 'locale') {
      $newManifest.add(($content.ManifestType + "-" + $content.PackageLocale), $content)
    }
    else {
      $newManifest.add($content.ManifestType, $content)
    }
  }
  # Now for the updating!
  # Get the installers array and the PackageIdentifier
  if ($type -eq "multifile") {
    $installers = $newManifest.Installer.Installers
    $packageIdentifier = $newManifest.Installer.PackageIdentifier
    $oldVersion = $newManifest.Version.PackageVersion
    if ($checkArp) {
      $arpValues = @{}
      $arpValues.PackageVersion = $newVersion
      $arpValues.PackageName = $newManifest.defaultLocale.PackageName
      $arpValues.Publisher = $newManifest.defaultLocale.Publisher
    }
  }
  else {
    $installers = $newManifest.singleton.Installers
    $packageIdentifier = $newManifest.singleton.PackageIdentifier
    $oldVersion = $newManifest.singleton.PackageVersion
    if ($checkArp) {
      $arpValues = @{}
      $arpValues.PackageVersion = $newVersion
      $arpValues.PackageName = $newManifest.singleton.PackageName
      $arpValues.Publisher = $newManifest.singleton.Publisher
    }
  }
  # Make sure all architectures are lowercase in new manifest
  foreach ($i in $installers) {
    $i.Architecture = $i.Architecture.ToLower()
  }
  # Make sure ProductCode is denormalized.
  if ($type -eq "multifile") {
    if ($newManifest.Installer.Contains("ProductCode")) {
      foreach ($i in $installers) {
        $i.ProductCode = $newManifest.Installer.ProductCode
      }
      $newManifest.Installer.Remove("ProductCode")
    }
  }
  if ($type -eq "singleton") {
    if ($newManifest.singleton.Contains("ProductCode")) {
      foreach ($i in $installers) {
        $i.ProductCode = $newManifest.singleton.ProductCode
      }
      $newManifest.singleton.Remove("ProductCode")
    }
  }
  foreach ($i in $urlMap.Keys) {
    # Add non existing architectures.
    $inManifest = $false
    foreach ($j in $installers) {
      if ($i.ToLower() -eq $j.Architecture) {
        $inManifest = $true
        break
      }
    }
    if (-Not $inManifest) {
      # Copy everything but the arch from the existing architecture.
      # Found at: https://www.powershellgallery.com/packages/PSTK/1.2.1/Content/Public%5CCopy-OrderedHashtable.ps1
      $temp = New-Object -TypeName System.Collections.Specialized.OrderedDictionary
      $MemoryStream = New-Object -TypeName System.IO.MemoryStream
      $BinaryFormatter = New-Object -TypeName System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
      $BinaryFormatter.Serialize($MemoryStream, $installers[0])
      $MemoryStream.Position = 0
      $temp = $BinaryFormatter.Deserialize($MemoryStream)
      $MemoryStream.Close()
      $temp.Architecture = $i.ToLower()
      $installers.add($temp)
    }
  }
  if (($installers.Length -eq 1) -And (-Not [String]::IsNullOrEmpty($newURL))) {
    $urlMap = @{$installers[0].Architecture = $newURL }
  }
  # elseif ($autoReplaceURL) {
  #     $urlMap = @{}
  #     foreach ($i in $installers) {
  #         $urlMap[$i.Architecture] = $i.InstallerUrl -Replace $oldVersion, $newVersion
  #         Write-Host "Auto-replace for arch" $i.Architecture "resulted in URL "$urlMap[$i.Architecture] -ForegroundColor Yellow
  #     }
  # }
  if ($urlMap.Count -ne $installers.Length -And (-Not $autoReplaceURL)) {
    foreach ($i in $installers) {
      if (-Not ($urlMap.ContainsKey($i.Architecture))) {
        # The user didn't specify this URL.
        Write-Host "What is the Installer URL for architecture "$i.Architecture"?"
        $urlMap[$i.Architecture] = Read-Host $i.InstallerUrl
      }
    }
    # If the user provided new architectures, we don't need to worry about it.
  }

  # Let's download the installers and get the needed info.
  foreach ($i in $installers) {
    if ($autoReplaceURL) {
      $url = $i.InstallerUrl -Replace $oldVersion, $newVersion
      Write-Host "Auto-replace for arch" $i.Architecture "resulted in URL " $url -ForegroundColor Yellow
    }
    else {
      $url = $urlMap[$i.Architecture]
    }
    $i.InstallerUrl = $url
    Write-Host 'Downloading installer for architecture'$i.Architecture'...' -ForegroundColor Yellow
    Invoke-WebRequest -UseBasicParsing -UserAgent "winget/1.0" -OutFile $env:TEMP\installer $url
    $i.InstallerSha256 = (Get-FileHash $env:TEMP\installer).Hash
    # if ($type -eq "multifile") {
    #   $isMsi = ((($newManifest.Installer.InstallerType) -And ($newManifest.Installer.InstallerType -eq "msi")) -Or ($i.InstallerType.ToLower() -eq "msi") -Or ($i.InstallerType.ToLower -eq "burn"))
    # }
    # else {
    #   $isMsi = ((($newManifest.singleton.InstallerType) -And ($newManifest.singleton.InstallerType.ToLower() -eq "msi") -Or ($i.InstallerType.ToLower() -eq "msi") -Or ($i.InstallerType.ToLower -eq "burn")))
    # }
    # Get Product Code if necessary
    if (($i.Contains("ProductCode") -Or ($productCode))) {
      try {
        $i.ProductCode = '{' + (((Get-CMsi $env:TEMP\installer).ProductCode).ToString()).ToUpper() + '}'
      }
      catch {
        Write-Host -ForegroundColor Red "The file doesn't look to be an MSI but has a ProductCode. Please manually verify the ProductCode to make sure it is correct."
        $i.ProductCode = $i.ProductCode -replace $oldVersion, $newVersion
      }
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
    foreach ($i in $newManifest.Keys) {
      $newManifest[$i]["PackageVersion"] = $newVersion
    }
  }
  else {
    $newManifest.singleton.PackageVersion = $newVersion
  }

  # Check if ReleaseNotes/ReleaseDate/ReleaseNotesURL is set.
  if ($type -eq "multifile")
  {
    if ($newManifest.defaultLocale.Contains("ReleaseNotes") -or $newManifest.defaultLocale.Contains("ReleaseNotesUrl") -or $newManifest.installer.Contains("ReleaseDate"))
    {
      Write-Host -ForegroundColor Yellow "Warning! Something about the ReleaseNotes was set in the previous manifest. Please manually correct it."
    }
    if ($newManifest.installer.Contains("ReleaseDate"))
    {
      $newManifest.installer["ReleaseDate"] = (Get-Date ([datetime]::Parse($releaseDate)) -f "yyyy-MM-dd")
    }
  }
  else
  {
    if ($newManifest.singleton.Contains("ReleaseNotes") -or $newManifest.singleton.Contains("ReleaseNotesUrl") -or $newManifest.singleton.Contains("ReleaseDate"))
    {
      Write-Host -ForegroundColor Yellow "Warning! Something about the ReleaseNotes was set in the previous manifest. Please manually correct it."
    }
    if ($newManifest.singleton.Contains("ReleaseDate"))
    {
      $newManifest.singleton["ReleaseDate"] = (Get-Date ([datetime]::Parse($releaseDate)) -f "yyyy-MM-dd")
    }
  }

  # Now let's get these back to YAML.
  $path = ".\" + $newVersion + "\"
  New-Item -Type Directory $path -Force | Out-Null

  foreach ($i in $newManifest.Keys) {
    if (($newManifest[$i].ManifestType.ToLower() -eq "singleton") -or ($newManifest[$i].ManifestType.ToLower() -eq "version")) {
      $fileName = $path + $packageIdentifier + ".yaml"
    }
    elseif ($newManifest[$i].ManifestType.ToLower() -eq "installer") {
      $fileName = $path + $packageIdentifier + ".installer.yaml"
    }
    elseif (($newManifest[$i].ManifestType.ToLower() -eq "locale") -Or ($newManifest[$i].ManifestType.ToLower() -eq "defaultlocale")) {
      $fileName = $path + $packageIdentifier + ".locale." + $newManifest[$i].PackageLocale + ".yaml"
    }
    else {
      throw $newManifest[$i].ManifestType.ToLower() + " is an unknown type."
    }
    # Add schema info for IntelliSense
    $fileContent = '# yaml-language-server: $schema=https://aka.ms/winget-manifest.' + $newManifest[$i].ManifestType.ToLower() + '.' + $newManifest[$i].ManifestVersion.ToLower() + '.schema.json' + "`r`n"
    # And the manifest...
    if (($newManifest[$i].ManifestType.ToLower() -ne "locale") -And ($newManifest[$i].ManifestType.ToLower() -ne "defaultlocale")) {
      $fileContent += ($newManifest[$i] | ConvertTo-Yaml).replace("'", '"')
    }
    else {
      $fileContent += $newManifest[$i] | ConvertTo-Yaml
    }
    [System.Environment]::CurrentDirectory = (Get-Location).Path
    [System.IO.File]::WriteAllLines($fileName, $fileContent)
    # $fileContent | Out-File -Encoding "utf8" -FilePath $fileName
    Write-Host $fileName "written." -ForegroundColor Green
    if (Get-Content $fileName | Select-String "Â") {
      Write-Host ("UTF-8 Corruption detected in {0}" -f $fileName) -ForegroundColor Red
      Write-Host ("Please resolve this manually before I continue.")
      pause
    }
  }
  # Get rid of the temporary manifest we created if we needed a conversion.
  if ($converted) {
    Remove-Item -Recurse $oldManifestFolder
  }
  # Validate this thing.
  winget validate $path | Out-Null
  if ($LASTEXITCODE -ne 0) {
    throw "Manifest validation failed."
  }
  if ($test) {
    if ($metadataCheck) {
      Start-WinGetSandbox $path -metadata $arpValues
    }
    else {
      Start-WinGetSandbox $path
    }
    return $true
  }
  elseif ($silentTest) {
    if ($metadataCheck) {
      $testSuccess = Test-WinGetManifest $path -metadata $arpValues
    }
    else {
      $testSuccess = Test-WinGetManifest $path
    }
    if ($testSuccess) {
      Write-Host "Manifest successfully installed in Windows Sandbox!"
      if ($commit) {
        if ($update)
        {
          New-WinGetCommit $path -update $oldManifestFolder
        }
        else {
          New-WinGetCommit $path
        }
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
    [Parameter(Position = 0, Mandatory = $true, HelpMessage = "The manifest you wish to convert.")]
    [string]$oldManifestFolder,
    [Parameter(HelpMessage = "Replace the old manifest with this new one.")]
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
    PackageVersion    = $currentManifest["PackageVersion"];
    DefaultLocale     = $currentManifest["PackageLocale"];
    ManifestType      = "version";
    ManifestVersion   = "1.0.0";
  }
  $newManifest["defaultLocale"] = [Ordered]@{
    PackageIdentifier = $currentManifest["PackageIdentifier"];
    PackageVersion    = $currentManifest["PackageVersion"];
    PackageLocale     = $currentManifest["PackageLocale"];
    Publisher         = $currentManifest["Publisher"];
    PackageName       = $currentManifest["PackageName"];
    License           = $currentManifest["License"];
    ShortDescription  = $currentManifest["ShortDescription"];
    ManifestType      = "defaultLocale";
    ManifestVersion   = "1.0.0";
  }
  $newManifest["installer"] = [Ordered]@{
    PackageIdentifier = $currentManifest["PackageIdentifier"];
    PackageVersion    = $currentManifest["PackageVersion"];
    Installers        = $currentManifest["Installers"];
    ManifestType      = "installer";
    ManifestVersion   = "1.0.0";
  }
  # Now, put any keys we missed in the right spots.
  $extraKeys = [Ordered]@{}
  $extraKeys["version"] = [Ordered]@{}
  $extraKeys["defaultLocale"] = [Ordered]@{}
  $extraKeys["installer"] = [Ordered]@{}
  foreach ($i in $currentManifest.Keys) {
    if (-Not ($requiredKeys -contains $i) ) {
      if ($i -in $versionSchema.properties.PSObject.Properties.Name) {
        $extraKeys["version"].add($i, $currentManifest[$i])
      }
      elseif ($i -in $localeSchema.properties.PSObject.Properties.Name) {
        $extraKeys["defaultLocale"].add($i, $currentManifest[$i])
      }
      elseif ($i -in $installerSchema.properties.PSObject.Properties.Name) {
        $extraKeys["installer"].add($i, $currentManifest[$i])
      }
    }
  }
  # Now, add the extra keys back to the new Manifest.
  # Since you can't add multiple keys at a certain place in a ordered dict in PowerShell, I had to do this.
  $count = 3
  foreach ($i in $extraKeys["version"].Keys) {
    $newManifest["version"].Insert($count, $i, $extraKeys["version"][$i])
    $count++;
  }
  $count = 3
  foreach ($i in $extraKeys["defaultLocale"].Keys) {
    $newManifest["defaultLocale"].Insert($count, $i, $extraKeys["defaultLocale"][$i])
    $count++;
  }
  $count = 2
  foreach ($i in $extraKeys["installer"].Keys) {
    $newManifest["installer"].Insert($count, $i, $extraKeys["installer"][$i])
    $count++;
  }
  # Now we can write the files.
  $path = ".\" + $currentManifest["PackageVersion"] + "-multiFile\"
  New-Item -Type Directory $path -Force | Out-Null

  # Copied (with modifications) from Update-WinGetManifest. I should break this out into a function...
  foreach ($i in $newManifest.Keys) {
    if ($newManifest[$i].ManifestType.ToLower() -eq "version") {
      $fileName = $path + $currentManifest["PackageIdentifier"] + ".yaml"
    }
    elseif ($newManifest[$i].ManifestType.ToLower() -eq "installer") {
      $fileName = $path + $currentManifest["PackageIdentifier"] + ".installer.yaml"
    }
    elseif ($newManifest[$i].ManifestType.ToLower() -eq "defaultlocale") {
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
  if ($LASTEXITCODE -ne 0) {
    throw "Manifest validation failed. Check the source manifest to ensure it's good and try again."
  }
  if ($overwrite) {
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
    [Parameter(mandatory = $true, Position = 0, HelpMessage = "The manifest to commit.")]
    [string] $manifest,
    [Parameter(HelpMessage = "Use the currently checked out branch, instead of making a new one.")]
    [switch] $currentBranch,
    [Parameter(HelpMessage = "The commit message, if you don't want the automatically generated one.")]
    [string] $customMessage,
    [Parameter(HelpMessage = "Update this manifest with the new one (delete the old one).")]
    [string] $update,
    [Parameter(HelpMessage = "Remove this manifest.")]
    [switch] $remove
  )
  $ErrorActionPreference = "Stop"
  if (Test-Path -Path $manifest -PathType Leaf) {
    throw "This isn't a folder, this is a file!"
  }
  # $content = Get-Content $manifest | ConvertFrom-Yaml
  foreach ($i in (Get-ChildItem -Path $manifest)) {
    $theSplitName = $i.Name.Split(".")
    if ($theSplitName.length -ge 3) {
      $content = Get-Content ($manifest + "\" + $i) | ConvertFrom-Yaml -Ordered
      if ($content.ManifestType.ToLower() -eq "version") {
        $content = Get-Content ($manifest + "\" + $content.PackageIdentifier + ".locale." + $content.DefaultLocale + ".yaml") | ConvertFrom-Yaml -Ordered
        break
      }
      elseif ($content.ManifestType.ToLower() -eq "singleton") {
        break
      }

    }
  }
  if ([string]::IsNullOrEmpty($customMessage) -or [string]::IsNullOrWhiteSpace($customMessage)) {
    if($remove)
    {
      $commitMessage = "Removed non-working manifest for " + $content.PackageName + " version " + $content.PackageVersion + "."
    }
    elseif ([string]::IsNullOrEmpty($update)) {
      $commitMessage = "Added " + $content.PackageName + " version " + $content.PackageVersion + "."
    }
    
    else {
      $commitMessage = "Updated " + $content.PackageName + " to version " + $content.PackageVersion + "."
    }
  }
  else {
    $commitMessage = $customMessage
  }
  if (-Not $currentBranch) {
    $branchName = $content.PackageIdentifier + "-" + $content.PackageVersion
    if ($remove)
    {
      $branchName += "-remove"
    }
    git fetch --all | Out-Null
    git checkout -b "$branchName" upstream/master | Out-Null
    if ($LASTEXITCODE -ne 0) {
      # The branch already exists.
      git checkout "$branchName"
    }
  }
  if ($remove) {
    git rm -r "$manifest" | Out-Null
    if ($LASTEXITCODE -ne 0)
    {
      throw "Delete failed."
    }
  }
  else {
    git add "$manifest" | Out-Null
  }
  if (-Not [string]::IsNullOrEmpty($update))
  {
    git rm -r $update | Out-Null
    if ($LASTEXITCODE -ne 0)
    {
      throw "Delete failed."
    }
  }
  git commit -m $commitMessage | Out-Null
  if ($LASTEXITCODE -ne 0) {
    throw "Commit failed."
  }
}
function Get-WinGetApplicationCurrentVersion {
  # Uses the winget cli to get the current version of an app in the repo.
  # Useful for seeing if something needs an update.
  param (
    [Parameter(Position = 0, Mandatory = $true, HelpMessage = "The manifest ID to check.")]
    [string]$id
  )

  winget source update | Out-Null
  $littleManifest = (winget show $id --source winget)
  if ($LASTEXITCODE -ne 0) {
    throw "Couldn't find manifest " + $id
  }
  try { $version = ($littleManifest | Select-Object -Skip 1 | Out-String | ConvertFrom-Yaml).version }
  catch { 
    $version = ($littleManifest | Select-Object -Skip 2 | Out-String | ConvertFrom-Yaml).version 
  }
  return $version
}

function Get-WinGetManifestArpMetadata {
  param (
    [Parameter(Mandatory = $true, Position = 0, HelpMessage = "Path to manifest to be updated.")]
    [string] $manifestFolder
  )
  $manifest = @{}
  foreach ($i in (Get-ChildItem -Path $manifestFolder)) {
    $content = (Get-Content ($manifestFolder + "\" + $i) | ConvertFrom-Yaml -Ordered)
    $manifest.add($content.ManifestType, $content)
  }
  $type = Get-WinGetManifestType $manifestFolder
  if ($type -eq "multifile") {
    $arpValues = @{}
    $arpValues.PackageVersion = $manifest.defaultLocale.PackageVersion
    $arpValues.PackageName = $manifest.defaultLocale.PackageName
    $arpValues.Publisher = $manifest.defaultLocale.Publisher
  }
  else {
    $arpValues = @{}
    $arpValues.PackageVersion = $manifest.singleTon.PackageVersion
    $arpValues.PackageName = $manifest.singleton.PackageName
    $arpValues.Publisher = $manifest.singleton.Publisher
  }
  return $arpValues


}
