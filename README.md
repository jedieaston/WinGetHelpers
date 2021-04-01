# WinGetHelpers!

* Now updated for 1.0 manifests! *


These are some cmdlets I hastely made to help support microsoft/winget-pkgs. Maybe they'll help you too. 

I'm not very good at PowerShell, which is why it looks so bad, but all of these work on my machine(tm) so as long as you are on Windows 10 and have winget installed,
they should all work for you too (except Get-GitHubReleases, which requires the GitHub CLI. I'll fix that eventually.) I'll add comments when I get time.

# Requirements

You'll need two dependencies: [Carbon](http://get-carbon.org/) (for getting the product code from MSI files) and [powershell-yaml](https://github.com/cloudbase/powershell-yaml) (to read and write YAML files). 

So before you download this module and put it in your modules folder, make sure to do ` Install-Module powershell-yaml; Install-Module -Name 'Carbon' `

