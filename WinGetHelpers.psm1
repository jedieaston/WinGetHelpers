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

    if (-Not (Test-Path -Path $Manifest -PathType Leaf)) {
      throw 'The Manifest file does not exist.'
    }

    winget.exe validate $Manifest
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
    url      = 'https://github.com/microsoft/winget-cli/releases/download/v-0.2.10191-preview/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.appxbundle'
    hash     = '2B521E128D7FB368A685432EFE6864473857183C9A886E5725EA32B6C84AF8E1'
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

  Get-ChildItem $tempFolder -Recurse -Exclude $dependencies.fileName | Remove-Item -Force

  if (-Not [String]::IsNullOrWhiteSpace($Manifest)) {
    Copy-Item -Path $Manifest -Destination $tempFolder
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
    $manifestPathInSandbox = Join-Path -Path $desktopInSandbox -ChildPath (Join-Path -Path $tempFolderName -ChildPath $manifestFileName)
    if ($auto) {
      $bootstrapPs1Content += @"
Write-Host @'
--> Installing the Manifest $manifestFileName
'@
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
function Get-WinGetApplication {
    param (
        [Parameter(Mandatory = $true)]
        [string]$packageName
    )
    $repositoryUrlRoot = "https://raw.githubusercontent.com/Microsoft/winget-pkgs/master/manifests/"
    $ErrorActionPreference = 'Stop'
    # Sorry for the mess.
    # Get the manifest ID and version(Publisher.Name)
    winget source update
    $littleManifest = (winget show $packageName)
    if ($LASTEXITCODE -ne 0) {throw "Couldn't find package $package."}
    $package = ($littleManifest | Select-Object -Skip 1 | Out-String).Split('[')[1].Split(']')[0]
    try { $version = ($littleManifest | Select-Object -Skip 2 | Out-String | ConvertFrom-Yaml).version }
    catch { $version = ($littleManifest | Select-Object -Skip 2 | Out-String | ConvertFrom-Yaml).version }
    # Now we can get the full manifest.
    $publisher,$appName = $package.Split('.')
    $manifestFilePath = $repositoryUrlRoot + ($publisher)+ "/" +($appName) + "/" + $version + ".yaml"
    Write-Host $manifestFilePath
    $manifest = (Invoke-WebRequest $manifestFilePath).Content | Out-String | ConvertFrom-Yaml
    $manifest.appName = $appName
    # Getting around odd manifests. This will probably break when multiple installer types are allowed.
    if ($null -ne $manifest.Installers.InstallerType) 
    {
      $manifest.InstallerType = $manifest.Installers.InstallerType
    }
    return $manifest
}
function Get-URLFileHash {
    param (
        [Parameter(mandatory=$true)] 
        [string]$url
        )
    $ProgressPreference = 'SilentlyContinue' 
    $ErrorActionPreference = 'Stop'
    Invoke-WebRequest -UseBasicParsing -OutFile $env:TEMP\fart $url
    $hash = Get-FileHash $env:TEMP\fart
    Remove-Item $env:TEMP\fart
    Write-Host "The hash for the file at " $url "is "$hash.hash
    Set-Clipboard $hash.hash
    Write-Host "It has been written to the clipboard."
}

function Get-WinGetManifestInstallerHash {
  Param(
   [Parameter(mandatory=$true, Position = 0, HelpMessage = "The Manifest to get hash for.")]
   [String] $manifest
  )
  $ErrorActionPreference = 'Stop'
  $url = (Get-Content $manifest | ConvertFrom-Yaml).Installers.URL
  $ProgressPreference = 'SilentlyContinue' 
  Invoke-WebRequest -UseBasicParsing -OutFile $env:TEMP\fart $url
  $hash = Get-FileHash $env:TEMP\fart
  Remove-Item $env:TEMP\fart
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
function Assert-WinGetManifestStatus {
  param (
    [Parameter(Mandatory=$true)]
    [string]$id
  )
  $ErrorActionPreference = 'Stop'
  $manifest = Get-WingetApplication $id
  if ($null -eq $manifest)
  {
    Write-Host "That package couldn't be found in the community repo."
  }
  try {
    $ProgressPreference = 'SilentlyContinue' 
    Invoke-WebRequest -UseBasicParsing -OutFile $env:TEMP\fart $manifest.Installers.Url
    $hash = (Get-FileHash $env:TEMP\fart).Hash.ToLower()
    if ($hash -ne (($manifest.Installers.Sha256).ToLower())) {
      Write-Host "$id hash does not match installer hash."
      Write-Host "hash is: " $hash.ToUpper()
    }
    else {
      Write-Host "$id hash matches installer hash!"
    }
    Remove-Item $env:TEMP\fart
  }
  catch {
    Write-Host "unable to verify hash for $id ."
  }
}
function Get-WinGetManifestProductCode {
  # Carbon must be installed for this to work!
  # Vars
  Param(
   [Parameter(mandatory=$true, Position = 0, HelpMessage = "The Manifest to get product code for.")]
   [String] $manifest
  )
  $ErrorActionPreference = 'Stop'
  Import-Module 'Carbon'
  $url = (Get-Content $manifest | ConvertFrom-Yaml).Installers.URL
  $ProgressPreference = 'SilentlyContinue' 
  Invoke-WebRequest -UseBasicParsing -OutFile $env:TEMP\fart $url
  $out = ((Get-MSI $env:TEMP\fart).ProductCode).ToString()
  write-host $out
  Remove-Item $env:TEMP\fart
  $outPretty = "{" + $out.ToUpper() + "}"
  Write-Host "The product code for " $manifest " is " $outPretty "."
  Write-Host "It's in your clipboard."
  Set-Clipboard $outPretty
}
function Test-WinGetManifest {
  Param(
 [Parameter(mandatory=$true, Position = 0, HelpMessage = "The Manifest to test.")]
 [String] $manifest,
 [Parameter(HelpMessage="Keep the log file no matter what.")]
 [Switch]$keepLog,
 [Parameter(HelpMessage="Don't stop the function after 10 minutes.")]
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
  if ($null -ne $str) {
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
      Get-Content ".\tmp.log"
      return $false
  }
}
function Update-WinGetManifest {
    param (
        [Parameter(mandatory=$true, Position=0, HelpMessage="The manifest to update.")]
        [string] $manifest,
        [Parameter(mandatory=$true, Position=1, HelpMessage="The new version number.")]
        [string] $newVersion,
        [Parameter(mandatory=$false, Position=2, HelpMessage="The URL for the new installer.")]
        [string] $newURL,
        [switch] $productCode,
        [switch] $test,
        [switch] $silentTest,
        [switch] $autoReplaceURL,
        [switch] $overwrite
    )
    Import-Module 'Carbon'
    $ProgressPreference = "SilentlyContinue"
    $ErrorActionPreference = 'Stop'
    $content = Get-Content $manifest | ConvertFrom-Yaml -Ordered
    $oldVersion = $content.Version
    $content.Version = $newVersion
    if ($autoReplaceURL -and ($oldVersion -ne "latest")) {
        $content.Installers[0].Url = $content.Installers[0].Url -replace  $oldVersion, $newVersion
        
        Write-Host "Auto replaced URL resulted in: " $content.Installers[0].Url
    }
    elseif ($newURL.length -ne 0) {
        $content.Installers[0].Url = $newURL
    }
    else {
        Write-Host "What is the Installer URL for the new version?"
        $content.Installers[0].Url = Read-Host $content.Installers[0].Url
    }
    # Get hash.
    Write-Host "Downloading installer, please stand by..."
    $ProgressPreference = 'SilentlyContinue' 
    Invoke-WebRequest -UseBasicParsing -OutFile $env:TEMP\fart $content.Installers[0].Url
    $content.Installers[0].Sha256 = (Get-FileHash $env:TEMP\fart).Hash
    # Get Product Code if necessary.
    if ($productCode) {
        $content.Installers[0].ProductCode = '{' + (((Get-MSI $env:TEMP\fart).ProductCode).ToString()).ToUpper() + '}'
    }
    $content | ConvertTo-Yaml | Write-Host
    $fileName =  (".\" + $content.Version + ".yaml")
    if ($overwrite) {
      # Delete the old manifest, we're overwriting!
      Remove-Item $manifest
    }
    ($content | ConvertTo-Yaml).replace("'", '"') | Out-File -FilePath $fileName
    Write-Host $fileName " written."
    winget validate $fileName
    if ($test) {
       Start-WinGetSandbox $fileName
       return $true
    }
    elseif ($silentTest) {
      $testSuccess = Test-WinGetManifest $fileName
      if ($testSuccess) {
        Write-Host "Manifest successfully installed in Windows Sandbox!"
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
function New-WinGetCommit {
  # Don't use this unless you're really lazy.
  param(
    [Parameter(mandatory=$true, Position=0, HelpMessage="The manifest to commit.")]
    [string] $manifest
  )
  $ErrorActionPreference = "Stop"
  $content = Get-Content $manifest | ConvertFrom-Yaml
  $commitMessage = "Added " + $content.name + " version " + $content.Version + "."
  git add "$manifest"
  git commit -m $commitMessage
}
