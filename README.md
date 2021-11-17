# WinGetHelpers!

__Now updated for 1.0 manifests__

I hastily made this PowerShell module, that contains a collection of helper functions to make creation/submission of manifests easy for [microsoft/winget-pkgs](https://github.com/microsoft/winget-pkgs). Maybe it will be useful for other people too.

I'm not good at PowerShell, which is why it looks like it's a bit of a mess. But it works on my machine (Windows 10 with WinGet installed) and I'm happy to share it with the community.

## Dependencies

This PowerShell module requires the following dependencies:

| Dependency | Why it's needed |
| --- | --- |
| Carbon | For getting the Product Code from MSI files |
| powershell-yaml | To parse the YAML files |

You can install both of these at once with the following command:
```
Install-Module powershell-yaml; Install-Module -Name Carbon
```

## Cmdlets/Functions in this module

This module contains the following cmdlets/functions:

| Cmdlet | Description |
| --- | --- |
| Get-URLFileHash | Gets the SHA256 hash of a file from a URL, optionally with a `-Clipboard` parameter to copy the hash to the clipboard |
| Get-GitHubReleases | Gets the latest 5 releases from a GitHub repository |
| Test-WinGetManifest | Tests the generated manifest in Sandbox to ensure that it installs correctly and doesn't have any ARP errors |
| Update-WinGetManifest | Generate a manifest for an updated package and optionally test it in Windows Sandbox |
| Convert-WinGetSingletonToMultiFile | Converts a singleton manifest to a multi-file manifest |
| New-WinGetCommit | Commit the modified manifest to the repository to submit a Pull Request |
| Get-WinGetApplicationCurrentVersion | Uses WinGet CLI to get the current version of an application |
| Get-WinGetManifestArpMetadata | Gets the ARP metadata from a manifest |
