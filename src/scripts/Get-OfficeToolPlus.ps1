# Enable TLSv1.2 for compatibility with older clients.
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12
# Do not display progress for WebRequest.
$ProgressPreference = 'SilentlyContinue'
$Host.UI.RawUI.WindowTitle = "Office Tool Plus | Downloader"

# Localization
$CurrentLang = (Get-WinUserLanguageList)[0].LanguageTag.Replace("en_AU", "en_GB2")
$SupportedLanguages = @("en_AU", "en_GB")
# Fallback to default language if not supported.
if ($SupportedLanguages -notcontains $CurrentLang) {
    $CurrentLang = $SupportedLanguages[0].Replace("en_AU", "en_GB")
}
$AllLanguages = @{
    "ChooseOP"            = [PSCustomObject]@{
        en_AU      = "1"
        en_GB      = "2"
    }
    "GoBack"              = [PSCustomObject]@{
        en_AU      = "Back"
        en_GB      = "Back"
    }
    "HomeSelOP"           = [PSCustomObject]@{
        en_AU      = "  Select an option:"
        en_GB      = "  Select an option:"
    }
    "HomeOP1"             = [PSCustomObject]@{
        en_AU      = "Download now"
        en_GB      = "Download now"
    }
    "HomeOP2"             = [PSCustomObject]@{
        en_AU      = "Select a edition to download"
        en_GB      = "Select a edition to download"
    }
    "Exit"                = [PSCustomObject]@{
        en_AU      = "Exit"
        en_GB      = "Exit"
    }
    "OSInfo"              = [PSCustomObject]@{
        en_AU      = "OS info:"
        en_GB      = "OS info:"
    }
    "SelToDownload"       = [PSCustomObject]@{
        en_AU      = "  Select the edition to download:"
        en_GB      = "  Select the edition to download:"
    }
    "DownSelx64Runtime"   = [PSCustomObject]@{
        en_AU      = "64-bit with runtime"
        en_GB      = "64-bit with runtime"
    }
    "DownSelx86Runtime"   = [PSCustomObject]@{
        en_AU      = "32-bit with runtime"
        en_GB      = "32-bit with runtime"
    }
    "DownSelArm64Runtime" = [PSCustomObject]@{
        en_AU      = "ARM64 with runtime"
        en_GB      = "ARM64 with runtime"
    }
    "DownSelx64"          = [PSCustomObject]@{
        en_AU      = "64-bit"
        en_GB      = "64-bit"
    }
    "DownSelx86"          = [PSCustomObject]@{
        en_AU      = "32-bit"
        en_GB      = "32-bit"
    }
    "DownSelArm64"        = [PSCustomObject]@{
        en_AU      = "ARM64"
        en_GB      = "ARM64"
    }
    "DownNormal"          = [PSCustomObject]@{
        en_AU      = "  The {0} edition of Office Tool Plus will be downloaded."
        en_GB      = "  The {0} edition of Office Tool Plus will be downloaded."
    }
    "DownRuntime"         = [PSCustomObject]@{
        en_AU      = "  The {0} edition of Office Tool Plus with runtime will be downloaded."
        en_GB      = "  The {0} edition of Office Tool Plus with runtime will be downloaded."
    }
    "SelLocation"         = [PSCustomObject]@{
        en_AU      = "  Select the save location for Office Tool Plus:"
        en_GB      = "  Select the save location for Office Tool Plus:"
    }
    "LocationDesktop"     = [PSCustomObject]@{
        en_AU      = "Desktop"
        en_GB      = "Desktop"
    }
    "LocationCustom"      = [PSCustomObject]@{
        en_AU      = "Select a folder"
        en_GB      = "Select a folder"
    }
    "SelLocationTip"      = [PSCustomObject]@{
        en_AU      = "  If you don't see the window to select the folder, it may be behind the window."
        en_GB      = "  If you don't see the window to select the folder, it may be behind the window." 
    }
    "Downloading"         = [PSCustomObject]@{
        en_AU      = "  Downloading Office Tool Plus, please wait."
        en_GB      = "  Downloading Office Tool Plus, please wait."
    }
    "Extracting"          = [PSCustomObject]@{
        en_AU      = "  Extracting files, please wait."
        en_GB      = "  Extracting files, please wait."
    }
    "ErrorDownloading"    = [PSCustomObject]@{
        en_AU      = "  An error occurred while downloading the file."
        en_GB      = "  An error occurred while downloading the file."
    }
    "RetryDownload"       = [PSCustomObject]@{
        en_AU      = "  Do you want to retry? (Y/N)"
        en_GB      = "  Do you want to retry? (Y/N)"
    }
    "RedownloadTip"       = [PSCustomObject]@{
        en_AU      = "  Please download Office Tool Plus from https://www.officetool.plus/ or try again."
        en_GB      = "  Plwase download Office Tool Plus from https://www.officetool.plus/ 下载 Office Tool Plus，或者再试一遍。"
    }
    "DownloadSuccess"     = [PSCustomObject]@{
        en_AU      = "  Office Tool Plus was extracted to {0}"
        en_GB      = "  Office Tool Plus was extracted to {0}"
    }
    "StartProgram"        = [PSCustomObject]@{
        en_AU      = "Start program"
        en_GB      = "Start program"
    }
    "QuickInstallation"   = [PSCustomObject]@{
        en_AU      = "Quick installation"
        en_GB      = "Quick installation"
    }
    "ManuallyForMore"     = [PSCustomObject]@{
        en_AU      = "Tips: for more install options please launch application and configure as you want."
        en_GB      = "Tips: for more install options please launch application and configure as you want." 
    }
    "InvokeCommand"       = [PSCustomObject]@{
        en_AU      = "Invoke commands"
        en_GB      = "Invoke commands"
    }
    "EnterCommand"        = [PSCustomObject]@{
        en_AU      = "Enter a command to execute, or enter nothing to go back"
        en_GB      = "Enter a command to execute, or enter nothing to go back"
    }
    "PressToContinue"     = [PSCustomObject]@{
        en_AU      = "Press Enter to continue."
        en_GB      = "Press Enter to continue."
    }
}

function Get-LString {
    param([string]$Key)

    return $AllLanguages[$Key].$CurrentLang
}

function Invoke-Command {
    param([string]$Path)

    Clear-Host
    Write-Host "=========================== Office Tool Plus ==========================="
    Write-Host
    $UserInput = Read-Host "  $(Get-LString -Key "EnterCommand")"
    if ($null -eq $UserInput -or $UserInput -eq "") {
        Invoke-Launch-App -Path $Path
    }
    else {
        Clear-Host
        Start-Process -FilePath "$Path\Office Tool\Office Tool Plus.Console.exe" -ArgumentList $UserInput -Wait -NoNewWindow -WorkingDirectory "$Path\Office Tool"
        Write-Host
        Get-LString "PressToContinue"
        Read-Host
        Invoke-Command -Path $Path
    }
}

function Invoke-Quick-Install {
    param([string]$Path)

    $DownloadURL = "https://server-win.ccin.top/office-tool-plus/api/xml_generator/?module=1"
    $Channel = "Current"

    Clear-Host
    Write-Host "=========================== Office Tool Plus ==========================="
    Write-Host
    Write-Host "  $(Get-LString "QuickInstallation")"
    Write-Host
    Write-Host "    1: Microsoft 365"
    Write-Host "    2: Microsoft 365 & Visio"
    Write-Host "    3: Office 2024 Pro Plus Volume"
    Write-Host "    4: Office 2024 Pro Plus Volume & Visio 2024 Pro Volume"
    Write-Host "    5: $(Get-LString "GoBack")"
    Write-Host
    Write-Host "  $(Get-LString "ManuallyForMore")"
    Write-Host
    $UserChoice = Read-Host "  $(Get-LString -Key "ChooseOP")"
    switch ($UserChoice) {
        "1" {
            $DownloadURL += "&products=O365ProPlusRetail_MatchOS&O365ProPlusRetail.exclapps=Access,Bing,Groove,Lync,OneDrive,OneNote,Outlook,Publisher,Teams"
        }
        "2" {
            $DownloadURL += "&products=O365ProPlusRetail_MatchOS|VisioProRetail_MatchOS&O365ProPlusRetail.exclapps=Access,Bing,Groove,Lync,OneDrive,OneNote,Outlook,Publisher,Teams&VisioProRetail.exclapps=Groove,OneDrive"
        }
        "3" {
            $DownloadURL += "&products=ProPlus2024Volume_MatchOS&ProPlus2024Volume.exclapps=Access,Lync,OneDrive,OneNote,Outlook"
            $Channel = "PerpetualVL2024"
        }
        "4" {
            $DownloadURL += "&products=ProPlus2024Volume_MatchOS|VisioPro2024Volume_MatchOS&ProPlus2024Volume.exclapps=Access,Lync,OneDrive,OneNote,Outlook&VisioPro2024Volume.exclapps=OneDrive"
            $Channel = "PerpetualVL2024"
        }
        "5" { Invoke-Launch-App -Path $Path }
        default { Invoke-Quick-Install -Path $Path }
    }
    $DownloadURL += "&channel=$Channel"

    $FileName = "$env:TEMP\Config.xml"
    $DownloadSuccess = $false
    do {
        try {
            Invoke-WebRequest -Uri $DownloadURL -UseBasicParsing -OutFile $FileName -ErrorAction Stop
            $DownloadSuccess = $true
        }
        catch {
            Get-LString "ErrorDownloading"
            $UserChoice = Read-Host $(Get-LString "RetryDownload")
            if ($UserChoice -ne "Y") {
                Get-LString "RedownloadTip"
                Read-Host
                Exit
            }
        }
    } while (-not $DownloadSuccess)
    Start-Process -FilePath "$Path\Office Tool\Office Tool Plus.exe" -ArgumentList "`"$FileName`""
}

function Invoke-Launch-App {
    param([string]$Path)

    Clear-Host
    Write-Host "=========================== Office Tool Plus ==========================="
    Write-Host
    Write-Host $([string]::Format($(Get-LString "DownloadSuccess"), "$Path"))
    Write-Host
    Write-Host "    1: $(Get-LString "StartProgram")"
    Write-Host "    2: $(Get-LString "QuickInstallation")"
    Write-Host "    3: $(Get-LString "InvokeCommand")"
    Write-Host "    4: $(Get-LString "Exit")"
    Write-Host
    $UserChoice = Read-Host "  $(Get-LString -Key "ChooseOP")"
    switch ($UserChoice) {
        "1" { Start-Process "$Path\Office Tool\Office Tool Plus.exe" }
        "2" { Invoke-Quick-Install -Path $Path }
        "3" { Invoke-Command -Path $Path }
        "4" { Exit }
        default { Invoke-Launch-App -Path $Path }
    }
    Exit
}

function Get-OTP {
    param([string]$DownloadURL, [string]$SavePath)

    $FileName = "$env:TEMP\Office Tool Plus.zip"
    $DownloadSuccess = $false
    do {
        try {
            Get-LString "Downloading"
            Invoke-WebRequest -Uri $DownloadURL -UseBasicParsing -OutFile $FileName -ErrorAction Stop
            Get-LString "Extracting"
            Expand-Archive -LiteralPath $FileName -DestinationPath $SavePath -Force
            $DownloadSuccess = $true
        }
        catch {
            Get-LString "ErrorDownloading"
            $UserChoice = Read-Host $(Get-LString "RetryDownload")
            if ($UserChoice -ne "Y") {
                Get-LString "RedownloadTip"
                Read-Host
                Exit
            }
        }
        finally {
            if (Test-Path $FileName) {
                $item = Get-Item -LiteralPath $FileName
                $item.Delete()
            }
        }
    } while (-not $DownloadSuccess)
    Invoke-Launch-App -Path $SavePath
}

function Get-RuntimeVersion {
    try {
        $DotnetInfo = dotnet --list-runtimes | Select-String -Pattern "Microsoft.WindowsDesktop.App 8"
        $IsX86Version = $DotnetInfo | Select-String -Pattern "(x86)"
        # If x86 version of runtime is installed on system, ignore it. Because we will download x64 version of OTP by default.
        if ($null -ne $IsX86Version) {
            return $false
        }
        if ($null -ne $DotnetInfo) {
            return $true
        }
    }
    catch {
        return $false
    }
    return $false
}

function Get-FolderName {
    [CmdletBinding()]
    [OutputType([string])]
    Param
    (
        # InitialDirectory help description
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = "Initial Directory for browsing",
            Position = 0)]
        [String]$SelectedPath,

        # Description help description
        [Parameter(
            Mandatory = $false,
            ValueFromPipelineByPropertyName = $true,
            HelpMessage = "Message Box Title")]
        [String]$Description = "Select a folder",

        # ShowNewFolderButton help description
        [Parameter(
            Mandatory = $false,
            HelpMessage = "Show New Folder Button when used")]
        [Switch]$ShowNewFolderButton
    )

    # Load Assembly
    Add-Type -AssemblyName System.Windows.Forms

    # Open Class
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog

    # Define Title
    $FolderBrowser.Description = $Description

    # Define Initial Directory
    if (-Not [String]::IsNullOrWhiteSpace($SelectedPath)) {
        $FolderBrowser.SelectedPath = $SelectedPath
    }

    if ($FolderBrowser.ShowDialog() -eq "OK") {
        $Folder += $FolderBrowser.SelectedPath
    }
    return $Folder
}

function Get-SelecFolderPage {
    param([string]$Type, [string]$Architecture)

    $DesktopPath = [Environment]::GetFolderPath("Desktop")
    Clear-Host
    Write-Host "=========================== Office Tool Plus ==========================="
    Write-Host
    Get-LString "SelLocation"
    Write-Host
    Write-Host "    1: $(Get-LString "LocationDesktop")"
    Write-Host "    2: $(Get-LString "LocationCustom")"
    Write-Host "    3: $(Get-LString "GoBack")"
    Write-Host
    $UserChoice = Read-Host "  $(Get-LString -Key "ChooseOP")"

    switch ($UserChoice) {
        "1" { $UserSpecifiedPath = $DesktopPath }
        "2" {
            Get-LString "SelLocationTip"
            $UserSpecifiedPath = Get-FolderName -SelectedPath $DesktopPath
            if ($null -eq $UserSpecifiedPath) {
                Get-SelecFolderPage -Type $Type -Architecture $Architecture
            }
        }
        3 { Get-Homepage }
        default { Get-SelecFolderPage -Type $Type -Architecture $Architecture }
    }
    Write-Host
    Get-OTP -DownloadURL "https://www.officetool.plus/redirect/download.php?type=$Type&arch=$Architecture" -SavePath $UserSpecifiedPath
}

function Get-DownloadPage {
    # Collect system information
    $OsVersion = [System.Environment]::OSVersion.VersionString
    $Arch = (Get-CimInstance Win32_OperatingSystem).OSArchitecture
    $Arch = if ($Arch -Match "ARM64") { "ARM64" } elseif ($Arch -Match "64") { "x64" } else { "x86" }

    Clear-Host
    Write-Host "=========================== Office Tool Plus ==========================="
    Write-Host
    Write-Host "  $(Get-LString "OSInfo") $OsVersion $Arch"
    Write-Host
    if (Get-RuntimeVersion -eq $true) {
        $Type = "normal"
        Write-Host $([string]::Format($(Get-LString "DownNormal"), $Arch))
    }
    else {
        $Type = "runtime"
        Write-Host $([string]::Format($(Get-LString "DownRuntime"), $Arch))
    }
    Start-Sleep -Seconds 3
    Get-SelecFolderPage -Type $Type -Architecture $Arch
}

function Get-SelectEditionPage {
    # Collect system information
    $OsVersion = [System.Environment]::OSVersion.VersionString
    $Arch = (Get-CimInstance Win32_OperatingSystem).OSArchitecture
    $Arch = if ($Arch -Match "ARM64") { "ARM64" } elseif ($Arch -Match "64") { "x64" } else { "x86" }
    $Type = "runtime"

    Clear-Host
    Write-Host "=========================== Office Tool Plus ==========================="
    Write-Host
    Write-Host "  $(Get-LString "OSInfo") $OsVersion $Arch"
    Write-Host
    Get-LString "SelToDownload"
    Write-Host
    Write-Host "    1: $(Get-LString "DownSelx64Runtime")"
    Write-Host "    2: $(Get-LString "DownSelx86Runtime")"
    Write-Host "    3: $(Get-LString "DownSelArm64Runtime")"
    Write-Host "    4: $(Get-LString "DownSelx64")"
    Write-Host "    5: $(Get-LString "DownSelx86")"
    Write-Host "    6: $(Get-LString "DownSelArm64")"
    Write-Host
    Write-Host "    7: $(Get-LString "GoBack")"
    Write-Host
    $UserChoice = Read-Host "  $(Get-LString -Key "ChooseOP")"
    switch ($UserChoice) {
        "1" { $Arch = "x64" }
        "2" { $Arch = "x86" }
        "3" { $Arch = "arm64" }
        "4" {
            $Arch = "x64"
            $Type = "normal"
        }
        "5" {
            $Arch = "x86"
            $Type = "normal"
        }
        "6" {
            $Arch = "arm64"
            $Type = "normal"
        }
        "7" { Get-Homepage }
        default { Get-SelectEditionPage }
    }
    Get-SelecFolderPage -Type $Type -Architecture $Arch
}

function Set-Language {
    Clear-Host
    Write-Host "=========================== Office Tool Plus ==========================="
    Write-Host
    Write-Host "  Please choose a language:"
    Write-Host
    Write-Host "    1: English (Australia)"
    Write-Host "    2: English (Great Britain)"
    Write-Host
    $UserChoice = Read-Host "  Please enter a number"
    switch ($UserChoice) {
        "1" { $script:CurrentLang = "en_AU" }
        "2" { $script:CurrentLang = "en_GB" }
        default { Set-Language }
    }
    Get-Homepage
}

function Get-Homepage {
    Clear-Host
    Write-Host "=========================== Office Tool Plus ==========================="
    Write-Host
    Get-LString -Key "HomeSelOP"
    Write-Host
    Write-Host "    1: $(Get-LString "HomeOP1")"
    Write-Host "    2: $(Get-LString "HomeOP2")"
    Write-Host "    3: Choose language"
    Write-Host "    4: $(Get-LString "Exit")"
    Write-Host
    $UserChoice = Read-Host "  $(Get-LString -Key "ChooseOP")"
    switch ($UserChoice) {
        "1" { Get-DownloadPage }
        "2" { Get-SelectEditionPage }
        "3" { Set-Language }
        "4" { Exit }
        default { Get-Homepage }
    }
}

Get-Homepage