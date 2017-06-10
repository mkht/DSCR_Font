$CSharpCode = @'
using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;

namespace FontResource
{
    public class AddRemoveFonts
    {
        private static IntPtr HWND_BROADCAST = new IntPtr(0xffff);

        [DllImport("gdi32.dll")]
        static extern int AddFontResource(string lpFilename);

        [DllImport("gdi32.dll")]
        static extern int RemoveFontResource(string lpFileName);

        [DllImport("user32.dll",CharSet=CharSet.Auto)]
        private static extern int SendMessage(IntPtr hWnd, WM wMsg, IntPtr wParam, IntPtr lParam);

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool PostMessage(IntPtr hWnd, WM Msg, IntPtr wParam, IntPtr lParam);

        public static bool PostFontChangedMessage() {
            bool posted = PostMessage(HWND_BROADCAST, WM.FONTCHANGE, IntPtr.Zero, IntPtr.Zero);
            return posted;
        }

        public static int AddFont(string fontFilePath) {
            FileInfo fontFile = new FileInfo(fontFilePath);
            if (!fontFile.Exists){
                return 0;
            }
            try{
                int retVal = AddFontResource(fontFilePath);
                bool posted = PostMessage(HWND_BROADCAST, WM.FONTCHANGE, IntPtr.Zero, IntPtr.Zero);
                return retVal;
            }
            catch{
                return 0;
            }
        }

        public static int RemoveFont(string fontFileName) {
            try{
                int retVal = RemoveFontResource(fontFileName);
                bool posted = PostMessage(HWND_BROADCAST, WM.FONTCHANGE, IntPtr.Zero, IntPtr.Zero);
                return retVal;
            }
            catch {
                return 0;
            }
        }

        public enum WM : uint
        {
            FONTCHANGE = 0x001D
        }

    }
}
'@
Add-Type $CSharpCode

# ////////////////////////////////////////////////////////////////////////////////////////
# ////////////////////////////////////////////////////////////////////////////////////////
function Get-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Collections.Hashtable])]
    param
    (
        [ValidateSet("Present","Absent")]
        [string]
        $Ensure = 'Present',

        [parameter(Mandatory = $true)]
        [ValidatePattern('\.(ttf|ttc|otf|fon)$')]
        [string]
        $FontFile,

        [pscredential]
        $Credential
    )
    $GetRes = @{
        Ensure = $Ensure
        FontFile = $FontFile
        Credential = $Credential
    }
    $FontFolder = Join-Path $Env:windir '\Fonts'
    $Filename = Split-Path $FontFile -Leaf

    if(Test-Path (Join-Path $FontFolder $Filename) -PathType Leaf){
        $GetRes.Ensure = 'Present'
    }
    else{
        $GetRes.Ensure = 'Absent'
    }

    $GetRes
} # end of Get-TargetResource



# ////////////////////////////////////////////////////////////////////////////////////////
# ////////////////////////////////////////////////////////////////////////////////////////
function Test-TargetResource
{
    [CmdletBinding()]
    [OutputType([System.Boolean])]
    param
    (
        [ValidateSet("Present","Absent")]
        [string]
        $Ensure = 'Present',

        [parameter(Mandatory = $true)]
        [ValidatePattern('\.(ttf|ttc|otf|fon)$')]
        [string]
        $FontFile,

        [pscredential]
        $Credential
    )

    return ((Get-TargetResource @PSBoundParameters).Ensure -eq $Ensure)
} # end of Test-TargetResource

# ////////////////////////////////////////////////////////////////////////////////////////
# ////////////////////////////////////////////////////////////////////////////////////////
function Set-TargetResource
{
    [CmdletBinding()]
    param
    (
        [ValidateSet("Present","Absent")]
        [string]
        $Ensure = 'Present',

        [parameter(Mandatory = $true)]
        [ValidatePattern('\.(ttf|ttc|otf|fon)$')]
        [string]
        $FontFile,

        [pscredential]
        $Credential
    )
    $ErrorActionPreference = 'Stop'

    if($Ensure -eq 'Absent'){
    #アンインストール
        $Filename = Split-Path $FontFile -Leaf
        Write-Verbose ('Uninstalling font...')
        UnInstall-Font $Filename -ErrorAction Stop
    }
    elseif($Ensure -eq 'Present'){
    #インストール
        $private:tmpFolder = $Env:TEMP
        $TempFont = (Get-RemoteFile -Path $FontFile -DestinationFolder $tmpFolder -Credential $Credential -Force -PassThru -ErrorAction Stop)   # TEMPフォルダに一度コピー
        if(Test-Path $TempFont){
            try{
                Write-Verbose ('Installing font...')
                Install-Font $TempFont.Fullname -ErrorAction Stop
                Write-Verbose ('Font Installed')
            }
            catch{
                Write-Error $_.Exception.Message
            }
            finally{
                if(Test-Path $TempFont){
                    Remove-Item $TempFont -Force -ErrorAction SilentlyContinue
                }
            }
        }
        else{
            Write-Error 'Failed to copy the font file'
        }
    }

} # end of Set-TargetResource

# ////////////////////////////////////////////////////////////////////////////////////////
# ////////////////////////////////////////////////////////////////////////////////////////
function Install-Font {
    # C:\Windows\Fontsにファイルコピーしただけではインストールされない
    # shell.application経由でコピーする
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory,Position=0)]
        [ValidateScript({Test-Path $_})]
        $Path
    )
    $FullPath = Resolve-Path $Path -ErrorAction Stop
    $local:FONTS = 0x14;
    $local:CopyOptions = 4 + 16 + 1024  #GUIを非表示にするオプション
    $ObjShell = New-Object -ComObject Shell.Application
    $FontsFolder = $ObjShell.Namespace($FONTS)
    try{
        $FontsFolder.CopyHere($FullPath.Path, $CopyOptions)
        [void][FontResource.AddRemoveFonts]::PostFontChangedMessage()
    }
    catch{
        Write-Error $_.Exception.Message
    }
}

# ////////////////////////////////////////////////////////////////////////////////////////
# ////////////////////////////////////////////////////////////////////////////////////////
function UnInstall-Font {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory,Position=0)]
        $Font
    )
    $RegFont = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"
    $FontFolder = Join-Path $Env:windir '\Fonts'
    $Filename = Split-Path $Font -Leaf
    $FontPath = Join-Path $FontFolder $Filename

    if(-not (Test-Path $FontPath -PathType Leaf)){
        Write-Error 'Font file not found'
        return
    }

    if(-not [FontResource.AddRemoveFonts]::RemoveFont($FontPath)){
        Write-Error 'Remove font failed'
    }
    else{
        #Write-Verbose 'Remove font succeeded'
        Start-Sleep -Seconds 2  # ちょっと待たないと失敗することがあるような、ないような...
        if($Value = (Get-ItemProperty $RegFont).PsObject.Properties | where {$_.value -eq $Filename}){
            Remove-ItemProperty -Path $RegFont -Name $Value.Name
        }
        if(Test-Path $FontPath){
            Remove-Item $FontPath -Force -ErrorAction Stop
        }
        Write-Verbose 'Remove font succeeded'
    }
}

# ////////////////////////////////////////////////////////////////////////////////////////
# ////////////////////////////////////////////////////////////////////////////////////////
function Get-RemoteFile {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, Position=0)]
        [Alias("Uri")]
        [Alias("SourcePath")]
        [System.Uri[]] $Path, # ダウンロードするファイルパス（URI）

        [Parameter(Mandatory=$true, Position=1)]
        [string]$DestinationFolder, # ダウンロード先フォルダ

        [Parameter()]
        [AllowNull()]
        [pscredential]$Credential,  # 資格情報

        [Parameter()]
        [int]$TimeoutSec = 0,

        [Parameter()]
        [switch]$Force,

        [Parameter()]
        [switch]$PassThru
    )
    begin{
        if(-not (Test-Path $DestinationFolder -PathType Container)){
            Write-Verbose ('DestinationFolder Folder "{0}" is not exist. Will create it.' -f $DestinationFolder)
            New-Item $DestinationFolder -ItemType Directory -Force -ErrorAction Stop
        }
    }

    Process{
        foreach($private:tempPath in $Path){
            try{
                $private:OutFile = ''
                $private:valid = $true
                $private:tmpDriveName = [Guid]::NewGuid()

                if($tempPath.IsLoopback -eq $null){
                    $valid = $false
                    throw ("{0} is not valid uri." -f $tempPath)
                }

                # ファイルの場所によって処理分岐(ローカル or 共有フォルダ or Web)
                if($tempPath.IsLoopback -or $tempPath.IsUnc){ # ローカル or 共有フォルダ
                    # 資格情報を使う場合は一度ドライブをマップする必要あり
                    if($PSBoundParameters.Credential){
                        New-PSDrive -Name $tmpDriveName -PSProvider FileSystem -Root (Split-Path $tempPath.LocalPath) -Credential $Credential -ErrorAction Stop | Out-Null
                    }
                    # ローカルにコピーする
                    $OutFile = Join-Path $DestinationFolder ([System.IO.Path]::GetFileName($tempPath.LocalPath))
                    if(Test-Path $OutFile -PathType Leaf){
                        if($tempPath.LocalPath -eq $OutFile){
                            if($PassThru){
                                if(Test-Path $OutFile){
                                    Get-Item $OutFile
                                }
                            }
                            continue
                        }
                        elseif($Force){
                            Write-Warning ('"{0}" will be overwritten.'-f $OutFile)
                        }
                        else{
                            $valid = $false
                            throw ("'{0}' is exist. If you want to replace existing file, Use 'Force' switch." -f $OutFile)
                        }
                    }

                    Write-Verbose ("Copy file from '{0}' to '{1}'" -f $tempPath.LocalPath, $DestinationFolder)
                    Copy-Item -Path $tempPath.LocalPath -Destination $DestinationFolder -ErrorAction Stop -Force:$Force
                }
                elseif($tempPath.Scheme -match 'http|https|ftp'){
                    # WebからDL
                    $OutFile = Join-Path $DestinationFolder ([System.IO.Path]::GetFileName($tempPath.AbsoluteUri))
                    if(Test-Path $OutFile -PathType Leaf){
                        if($Force){
                            Write-Warning ('"{0}" will be overwritten.'-f $OutFile)
                        }
                        else{
                            $valid = $false
                            throw ("'{0}' is exist. If you want to replace existing file, Use 'Force' switch." -f $OutFile)
                        }
                    }

                    Write-Verbose ("Download file from '{0}' to '{1}'" -f $tempPath.AbsoluteUri, $OutFile)
                    $private:origVerbose = $VerbosePreference; $VerbosePreference = 'SilentlyContinue'
                    Invoke-WebRequest -Uri $tempPath.AbsoluteUri -OutFile $OutFile -Credential $Credential -TimeoutSec $TimeoutSec -ErrorAction stop
                    $VerbosePreference = $origVerbose
                }
                else{
                    $valid = $false
                    throw ("{0} is not valid uri." -f $tempPath)
                }

                if($valid -and $OutFile -and $PassThru){
                    if(Test-Path $OutFile){
                        Get-Item $OutFile
                    }
                }
            }
            catch [Exception]{
                Write-Error $_.Exception.Message
            }
            finally{
                if(Get-PSDrive | where {$_.Name -eq $tmpDriveName}){
                    Remove-PSDrive -Name $tmpDriveName -Force
                }
            }
        }
    }
}

# ////////////////////////////////////////////////////////////////////////////////////////
# ////////////////////////////////////////////////////////////////////////////////////////
Export-ModuleMember -Function *-TargetResource
