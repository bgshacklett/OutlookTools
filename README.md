# OutlookTools

This is a fork of the OutlookTools repository by Aman Dhally <amandhally@gmail.com>. The original repo exists at <https://github.com/AmanDhally/OutlookTools>.

## Installation

With PowerShell:
```
$uri = 'https://raw.githubusercontent.com/bgshacklett/OutlookTools/master'
$files = @('OutlookTools.psd1','OutlookTools.psm1')
$docs = [Environment]::GetFolderPath('MyDocuments')
$PSmodules = Join-Path -Path "$docs" -ChildPath 'WindowsPowerShell\Modules'


if (!(Test-Path -Path "$PSmodules\OutlookTools")) {
    New-Item -Path "$PSmodules" -Name 'OutlookTools' -ItemType Directory 
}

foreach ($file in $files) {
    $source = "$uri/$file"
    $dest = "$PSmodules\OutlookTools\$file"

    (New-Object System.Net.WebClient).DownloadFile($source, $dest)
}

Remove-Variable -Name uri,files,docs,PSmodules,source,dest
```

## Usage

```
Import-Module -Name OutlookTools
```
