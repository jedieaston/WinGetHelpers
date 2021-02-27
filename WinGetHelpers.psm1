#Requires -Module Carbon
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

function Update-WinGetManifest {
    param (
        [Parameter(mandatory=$true, Position=0, HelpMessage="The manifest to update.")]
        [string] $manifest,
        [Parameter(mandatory=$true, Position=1, HelpMessage="The new version number.")]
        [string] $newVersion,
        [Parameter(mandatory=$false, Position=2, HelpMessage="The URL for the new installer.")]
        [string] $newURL,
        [switch] $productCode,
        [switch] $runSandbox,
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
    if ($runSandbox) {
        try {
          Start-WinGetSandbox (".\" + $content.Version + ".yaml")
        }
        catch {
          Write-Host "For -runSandbox to work, you need the Start-WinGetSandbox cmdlet. Check microsoft/winget-pkgs#827 for more information."
        }
    }
}