DSCR_Font
====

PowerShell DSC Resource to add / remove Font.

## Install
You can install Resource through [PowerShell Gallery](https://www.powershellgallery.com/packages/DSCR_Font/).
```Powershell
Install-Module -Name DSCR_Font
```

## Resources
* **cFont**
PowerShell DSC Resource to add / remove Font.

## Properties
### cFont
+ [string] **Ensure** (Write):
    + Specify installation state of the font.
    + The default value is Present. { Present | Absent }

+ [string] **FontName** (Key):
    + The name of the font.
    + You should specify the exact font name. To check the font name, install the font manually and refer to the `HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts` registry key.

+ [string] **FontFile** (Require):
    + The path of the font file.
    + You can specify remote file path. (UNC/http/https/ftp)
    + Make sure the extension of the file path is { .ttf | .ttc | .otf | .fon}

+ [PSCredential] **Credential** (Write):
    + The credential for access to remote source if needed.

## Examples
+ **Example 1**: Install the "Noto Serif" font
```Powershell
Configuration Example1
{
    Import-DscResource -ModuleName DSCR_Font
    cFont Add_NotoSerif
    {
        Ensure   = 'Present'
        FontName = 'Noto Serif (TrueType)'
        FontFile = 'C:\NotoSerif-Regular.ttf'
    }
}
```

+ **Example 2**: Remove the "Noto Sans" font
```Powershell
Configuration Example2
{
    Import-DscResource -ModuleName DSCR_Font
    cFont Remove_NotoSans
    {
        Ensure   = 'Absent'
        FontName = 'Noto Sans Regular (TrueType)'
        FontFile = 'NotoSans-Regular.ttf'
    }
}
```

## ChangeLog
### v0.7.0
+ Fix compatibility with Windows 10 1809 or later.
