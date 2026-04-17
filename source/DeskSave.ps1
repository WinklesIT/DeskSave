# ============================================
# DeskSave - User Data Backup Tool - V1
#
# Copyright (C) 2026 Joshua Winkles
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program. If not, see <https://www.gnu.org/licenses/>.
# ============================================
# - Full 2026-era dark UI overhaul
# - IT Managed badge in sidebar + header
# - Colored dot indicators per PC in Recent Backups
# - Icon-forward action tiles (backup/restore)
# - Relative time display on stat cards ("2 days ago")
# - Split numeric+unit storage display (38 GB)
# - "View all" link on dashboard
# - Email-style user display in sidebar footer
# - Deeper blue-tinted dark palette
# - Added: Network Printer Backup/Restore via Registry
# - Added: Firefox bookmark backup/restore (places.sqlite, all profiles)
# - Added: Added cbRFileAssociations reg.exe import FileAssociations.reg
# - Added: Wallpaper Backup/Restore
#   * Detects custom image wallpapers and backs up the file
#   * Saves wallpaper style/fit settings from registry
#   * Handles Windows Spotlight gracefully (saves metadata, skips file copy)
#   * Restore re-applies wallpaper image via SystemParametersInfo + registry style
# ============================================

# ---------------------------
# Ensure STA (required for WPF)
# ---------------------------
if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -STA -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# ---------------------------
# Debug logging toggle
# ---------------------------
$global:EnableBackupToolLogging = $true

# ---------------------------
# Debug logging
# ---------------------------
$global:BackupToolLog        = $null
$global:BackupToolTranscript = $null

function Initialize-BackupToolLog {
    if (-not $global:EnableBackupToolLogging) { return }

    $computer = $env:COMPUTERNAME
    $date     = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

    $logRoot = Join-Path $env:TEMP "BackupToolLogs"
    New-Item -ItemType Directory -Path $logRoot -Force | Out-Null

    $global:BackupToolLog        = Join-Path $logRoot "BackupTool_${computer}_$date.log"
    $global:BackupToolTranscript = Join-Path $logRoot "BackupTool_${computer}_$date.transcript.txt"

    Add-Content -Path $global:BackupToolLog -Value ("===== Backup Tool Log Started: {0} =====" -f (Get-Date)) -Encoding UTF8
    Add-Content -Path $global:BackupToolLog -Value ("Computer: {0}" -f $computer) -Encoding UTF8
    Add-Content -Path $global:BackupToolLog -Value ("User: {0}" -f $env:USERNAME) -Encoding UTF8
    Add-Content -Path $global:BackupToolLog -Value ("PowerShell: {0}" -f $PSVersionTable.PSVersion) -Encoding UTF8
    Add-Content -Path $global:BackupToolLog -Value ("ApartmentState: {0}" -f [System.Threading.Thread]::CurrentThread.ApartmentState) -Encoding UTF8
    Add-Content -Path $global:BackupToolLog -Value ("Script: {0}" -f $PSCommandPath) -Encoding UTF8
    Add-Content -Path $global:BackupToolLog -Value "" -Encoding UTF8

    try { Start-Transcript -Path $global:BackupToolTranscript -Append | Out-Null } catch {}
}

function Write-BackupToolLog {
    param([string]$Message)
    if ($global:BackupToolLog) {
        $line = ("[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $Message)
        Add-Content -Path $global:BackupToolLog -Value $line -Encoding UTF8
    }
}

function Dump-Exception {
    param([System.Exception]$ex)

    Write-BackupToolLog "===== EXCEPTION ====="
    Write-BackupToolLog ("Message: " + $ex.Message)
    Write-BackupToolLog ("Type: " + $ex.GetType().FullName)
    if ($ex.StackTrace) { Write-BackupToolLog ("StackTrace:`n" + $ex.StackTrace) }

    $inner = $ex.InnerException
    $lvl = 1
    while ($inner) {
        Write-BackupToolLog "----- INNER EXCEPTION (Level $lvl) -----"
        Write-BackupToolLog ("Message: " + $inner.Message)
        Write-BackupToolLog ("Type: " + $inner.GetType().FullName)
        if ($inner.StackTrace) { Write-BackupToolLog ("StackTrace:`n" + $inner.StackTrace) }
        $inner = $inner.InnerException
        $lvl++
    }

    Write-BackupToolLog "===== POWERSHELL `$Error DUMP ====="
    $Error | ForEach-Object { Write-BackupToolLog ("ERROR: " + $_.ToString()) }
}

Initialize-BackupToolLog
Write-BackupToolLog "Script started."

try {
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName System.Windows.Forms

    # ---- Load/save backup root from per-user config file ----
    $global:ConfigFile = Join-Path $env:APPDATA "BackupTool\config.json"
    $global:BackupRootPath = ""

    function Save-BackupRootConfig {
        param([string]$Path)
        $dir = Split-Path $global:ConfigFile -Parent
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        @{ BackupRoot = $Path } | ConvertTo-Json | Set-Content -Path $global:ConfigFile -Encoding UTF8
        $global:BackupRootPath = $Path
        Write-BackupToolLog "Saved backup root to config: $Path"
    }

    function Load-BackupRootConfig {
        if (Test-Path $global:ConfigFile) {
            try {
                $cfg = Get-Content $global:ConfigFile -Raw | ConvertFrom-Json
                if ($cfg.BackupRoot -and (Test-Path $cfg.BackupRoot)) {
                    $global:BackupRootPath = $cfg.BackupRoot
                    Write-BackupToolLog "Loaded backup root from config: $($global:BackupRootPath)"
                    return $true
                }
            } catch {}
        }
        return $false
    }

    # Try to detect common drive letters used for home drives (U, H, N, Z, etc.)
    $global:DetectedNetworkDrive = $null
    $commonHomeDrives = @("U","H","N","Z","V","W","X","Y")
    foreach ($letter in $commonHomeDrives) {
        $testPath = "${letter}:\"
        if (Test-Path $testPath) {
            $global:DetectedNetworkDrive = $letter
            Write-BackupToolLog "Detected potential home/network drive: ${letter}:\"
            break
        }
    }

    $global:ConfigLoaded = Load-BackupRootConfig

    # ---------------------------
    # Separate shared state for Backup vs Restore
    # ---------------------------
    $syncBackup = [hashtable]::Synchronized(@{
        Cancel   = $false
        RoboPids = New-Object System.Collections.Generic.List[int]
    })

    $syncRestore = [hashtable]::Synchronized(@{
        Cancel   = $false
        RoboPids = New-Object System.Collections.Generic.List[int]
    })

    # ---------------------------
    # Keep backup/restore objects alive (prevents DispatcherTimer GC)
    # ---------------------------
    $script:BackupTimer  = $null
    $script:BackupPs     = $null
    $script:BackupRs     = $null
    $script:BackupHandle = $null
    
    $script:RestoreTimer  = $null
    $script:RestorePs     = $null
    $script:RestoreRs     = $null
    $script:RestoreHandle = $null

    # ---------------------------
    # Browser EXE detection
    # ---------------------------
    function Get-AppPathExe {
        param([string]$ExeName)
        $regPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\$ExeName",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\$ExeName"
        )
        foreach ($rp in $regPaths) {
            try {
                $p = (Get-ItemProperty -Path $rp -ErrorAction Stop)."(default)"
                if ($p -and (Test-Path $p)) { return $p }
            } catch {}
        }
        return $null
    }

    function Get-ChromeExe {
        $exe = Get-AppPathExe "chrome.exe"
        if ($exe) { return $exe }

        $candidates = @(
            "$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
            "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
        )
        foreach ($c in $candidates) { if (Test-Path $c) { return $c } }

        try { return (Get-Command chrome.exe -ErrorAction Stop).Source } catch {}
        return $null
    }

    function Get-EdgeExe {
        $exe = Get-AppPathExe "msedge.exe"
        if ($exe) { return $exe }

        $candidates = @(
            "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe",
            "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
        )
        foreach ($c in $candidates) { if (Test-Path $c) { return $c } }

        try { return (Get-Command msedge.exe -ErrorAction Stop).Source } catch {}
        return $null
    }

    function Get-FirefoxExe {
        $exe = Get-AppPathExe "firefox.exe"
        if ($exe) { return $exe }

        $candidates = @(
            "$env:ProgramFiles\Mozilla Firefox\firefox.exe",
            "${env:ProgramFiles(x86)}\Mozilla Firefox\firefox.exe"
        )
        foreach ($c in $candidates) { if (Test-Path $c) { return $c } }

        try { return (Get-Command firefox.exe -ErrorAction Stop).Source } catch {}
        return $null
    }

    function Get-FirefoxProfilesRoot {
        return [IO.Path]::Combine($env:APPDATA, "Mozilla", "Firefox", "Profiles")
    }

    function Get-FirefoxInstallRoot {
        return [IO.Path]::Combine($env:APPDATA, "Mozilla", "Firefox")
    }

    function Get-LatestBackupFolderForThisPc {
        $backupRoot = $global:BackupRootPath
        if (-not $backupRoot -or -not (Test-Path $backupRoot)) { return $null }

        $pcFolders = Get-ChildItem -Path $backupRoot -Directory -ErrorAction SilentlyContinue
        if (-not $pcFolders) { return $null }

        $allBackups = @()
        foreach ($pcFolder in $pcFolders) {
            $timestampFolders = Get-ChildItem -Path $pcFolder.FullName -Directory -ErrorAction SilentlyContinue
            if ($timestampFolders) {
                $allBackups += $timestampFolders
            }
        }

        $latest = $allBackups | 
                  Sort-Object LastWriteTime -Descending | 
                  Select-Object -First 1

        if ($latest) { return $latest.FullName }
        return $null
    }

    function Get-FolderSize {
        param([string]$Path)
        if (-not (Test-Path $Path)) { return 0 }
        try {
            $size = (Get-ChildItem -Path $Path -Recurse -File -ErrorAction SilentlyContinue | 
                     Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
            if ($null -eq $size) { return 0 }
            return $size
        } catch {
            return 0
        }
    }

    function Format-FileSize {
        param([long]$Bytes)
        if ($Bytes -ge 1TB) { return "{0:N2} TB" -f ($Bytes / 1TB) }
        if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
        if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
        if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
        return "$Bytes bytes"
    }

    function Get-BackupContents {
        param([string]$BackupPath)
        
        $contents = @{
            Desktop = $false
            Documents = $false
            Pictures = $false
            Downloads = $false
            Music = $false
            Videos = $false
            ChromeBookmarks = $false
            EdgeBookmarks = $false
            FirefoxBookmarks = $false
            LegacyFavorites = $false
            Printers = $false
            Wallpaper = $false
            FileAssociations = $false
            TotalSize = 0
            Note = ""
        }

        if (-not (Test-Path $BackupPath)) { return $contents }

        $contents.Desktop = Test-Path (Join-Path $BackupPath "Desktop")
        $contents.Documents = Test-Path (Join-Path $BackupPath "Documents")
        $contents.Pictures = Test-Path (Join-Path $BackupPath "Pictures")
        $contents.Downloads = Test-Path (Join-Path $BackupPath "Downloads")
        $contents.Music = Test-Path (Join-Path $BackupPath "Music")
        $contents.Videos = Test-Path (Join-Path $BackupPath "Videos")
        $contents.Printers = Test-Path (Join-Path $BackupPath "NetworkPrinters.reg")
        $contents.Wallpaper = Test-Path (Join-Path $BackupPath "Wallpaper\_wallpaper_info.json")
        $contents.FileAssociations = Test-Path (Join-Path $BackupPath "FileAssociations.reg")
        
        $browserBackup = Join-Path $BackupPath "Browser_Backup"
        if (Test-Path $browserBackup) {
            $contents.ChromeBookmarks = Test-Path (Join-Path $browserBackup "Chrome")
            $contents.EdgeBookmarks = Test-Path (Join-Path $browserBackup "Edge")
            $contents.FirefoxBookmarks = Test-Path (Join-Path $browserBackup "Firefox")
            $contents.LegacyFavorites = Test-Path (Join-Path $browserBackup "Legacy_Favorites_Folder")
        }

        $noteFile = Join-Path $BackupPath "_backup_note.txt"
        if (Test-Path $noteFile) {
            $contents.Note = (Get-Content $noteFile -Raw -ErrorAction SilentlyContinue) -replace "`r`n","`n"
        }

        try {
            $contents.TotalSize = Get-FolderSize $BackupPath
        } catch {}

        return $contents
    }

    function Get-AllBackups {
        $backupRoot = $global:BackupRootPath
        if (-not $backupRoot -or -not (Test-Path $backupRoot)) { return @() }

        $allBackups = @()
        $pcFolders = Get-ChildItem -Path $backupRoot -Directory -ErrorAction SilentlyContinue
        
        foreach ($pcFolder in $pcFolders) {
            $timestampFolders = Get-ChildItem -Path $pcFolder.FullName -Directory -ErrorAction SilentlyContinue
            foreach ($ts in $timestampFolders) {
                $contents = Get-BackupContents $ts.FullName
                
                $contentsList = @()
                if ($contents.Desktop) { $contentsList += "Desktop" }
                if ($contents.Documents) { $contentsList += "Documents" }
                if ($contents.Pictures) { $contentsList += "Pictures" }
                if ($contents.Downloads) { $contentsList += "Downloads" }
                if ($contents.Music) { $contentsList += "Music" }
                if ($contents.Videos) { $contentsList += "Videos" }
                if ($contents.ChromeBookmarks) { $contentsList += "Chrome" }
                if ($contents.EdgeBookmarks) { $contentsList += "Edge" }
                if ($contents.FirefoxBookmarks) { $contentsList += "Firefox" }
                if ($contents.Printers) { $contentsList += "Printers" }
                if ($contents.Wallpaper) { $contentsList += "Wallpaper" }
                if ($contents.FileAssociations) { $contentsList += "File Assoc" }
                
                $contentsDisplay = if ($contentsList.Count -gt 0) { 
                    $contentsList -join ", " 
                } else { 
                    "Empty" 
                }
                
                $allBackups += [PSCustomObject]@{
                    PCName = $pcFolder.Name
                    Timestamp = $ts.Name
                    FullPath = $ts.FullName
                    LastWriteTime = $ts.LastWriteTime
                    Size = $contents.TotalSize
                    SizeFormatted = Format-FileSize $contents.TotalSize
                    Contents = $contents
                    ContentsDisplay = $contentsDisplay
                    Note = $contents.Note
                }
            }
        }

        return @($allBackups | Sort-Object LastWriteTime -Descending)
    }

    # ---------------------------
    # V5 Modern UI with Sidebar Navigation
    # ---------------------------
    [xml]$Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="DeskSave" Height="720" Width="980"
        WindowStartupLocation="CenterScreen" ResizeMode="CanResize"
        MinHeight="600" MinWidth="800"
        FontSize="13"
        Background="#FF0C0C10">

    <Window.Resources>
        <SolidColorBrush x:Key="WindowBg"      Color="#FF0C0C10"/>
        <SolidColorBrush x:Key="SidebarBg"     Color="#FF08080C"/>
        <SolidColorBrush x:Key="PanelBg"       Color="#FF161622"/>
        <SolidColorBrush x:Key="CardBg"        Color="#FF161622"/>
        <SolidColorBrush x:Key="RowBg"         Color="#FF0C0C10"/>
        <SolidColorBrush x:Key="BorderColor"   Color="#FF2C2C40"/>
        <SolidColorBrush x:Key="AccentColor"   Color="#FF5C6CF5"/>
        <SolidColorBrush x:Key="AccentBg"      Color="#FF181D3D"/>
        <SolidColorBrush x:Key="AccentHover"   Color="#FF6B7AF8"/>
        <SolidColorBrush x:Key="TextPrimary"   Color="#FFF8F8FC"/>
        <SolidColorBrush x:Key="TextSecondary" Color="#FFCCCCDE"/>
        <SolidColorBrush x:Key="TextMuted"     Color="#FF8888AA"/>
        <SolidColorBrush x:Key="NavActive"     Color="#FF1A1A26"/>
        <SolidColorBrush x:Key="GreenDot"      Color="#FF2ECC71"/>
        <SolidColorBrush x:Key="BlueDot"       Color="#FF5C6CF5"/>
        <SolidColorBrush x:Key="OrangeDot"     Color="#FFE8924A"/>
        <SolidColorBrush x:Key="BadgeBg"       Color="#FF1E1E30"/>
        <SolidColorBrush x:Key="RowHoverBg"    Color="#FF1C1C2C"/>

        <Style TargetType="Button" x:Key="NavButton">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{StaticResource TextSecondary}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Height" Value="40"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="HorizontalContentAlignment" Value="Left"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="NavBorder" Background="{TemplateBinding Background}"
                                CornerRadius="6" Margin="4,1">
                            <ContentPresenter VerticalAlignment="Center" Margin="10,0,0,0"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="NavBorder" Property="Background" Value="#FF16161F"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="Button" x:Key="TileButton">
            <Setter Property="Background" Value="{StaticResource CardBg}"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
            <Setter Property="BorderThickness" Value="0.5"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderColor}"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="TileBorder" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="10">
                            <ContentPresenter HorizontalAlignment="Stretch" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="TileBorder" Property="Background" Value="#FF181825"/>
                                <Setter TargetName="TileBorder" Property="BorderBrush" Value="{StaticResource AccentColor}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
            <Setter Property="Margin" Value="0,3"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>

        <Style TargetType="Button" x:Key="ActionButton">
            <Setter Property="Background" Value="#FF1A1A26"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
            <Setter Property="BorderThickness" Value="0.5"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderColor}"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="ABorder" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="5" Padding="10,6">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="ABorder" Property="Background" Value="#FF22223A"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.35"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="Button" x:Key="PrimaryButton">
            <Setter Property="Background" Value="{StaticResource AccentColor}"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="Medium"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="PBorder" Background="{TemplateBinding Background}"
                                CornerRadius="5" Padding="12,7">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="PBorder" Property="Background" Value="{StaticResource AccentHover}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.35"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="Button" x:Key="LinkButton">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{StaticResource AccentColor}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <TextBlock x:Name="LTxt" Text="{TemplateBinding Content}"
                                   Foreground="{TemplateBinding Foreground}"
                                   FontSize="{TemplateBinding FontSize}"/>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="LTxt" Property="TextDecorations" Value="Underline"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#FF0E0E11"/>
            <Setter Property="Foreground" Value="{StaticResource TextPrimary}"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderColor}"/>
            <Setter Property="BorderThickness" Value="0.5"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="CaretBrush" Value="{StaticResource AccentColor}"/>
        </Style>

        <Style TargetType="ProgressBar">
            <Setter Property="Background" Value="#FF1A1A26"/>
            <Setter Property="Foreground" Value="{StaticResource AccentColor}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Height" Value="4"/>
        </Style>
    </Window.Resources>

    <Grid Background="{StaticResource WindowBg}">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="200"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>

        <Border Grid.Column="0" Background="{StaticResource SidebarBg}"
                BorderBrush="{StaticResource BorderColor}" BorderThickness="0,0,0.5,0">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <StackPanel Grid.Row="0" Margin="16,18,16,14">
                    <Grid Margin="0,0,0,6">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Image x:Name="imgSidebarLogo" Grid.Column="0"
                               Width="36" Height="36"
                               HorizontalAlignment="Left" Margin="0,0,10,0"
                               VerticalAlignment="Center"
                               RenderOptions.BitmapScalingMode="HighQuality"/>
                        <StackPanel Grid.Column="1" VerticalAlignment="Center">
                            <TextBlock Text="DeskSave" FontSize="15" FontWeight="Medium"
                                       Foreground="{StaticResource TextPrimary}"/>
                            <TextBlock Text="v1.0" FontSize="11" Foreground="{StaticResource TextMuted}" Margin="0,1,0,0"/>
                        </StackPanel>
                    </Grid>
                </StackPanel>

                <Border Grid.Row="0" Height="0.5" Background="{StaticResource BorderColor}"
                        VerticalAlignment="Bottom" Margin="0,0,0,0"/>

                <StackPanel Grid.Row="1" Margin="8,10,8,0">
                    <Button x:Name="navHome" Style="{StaticResource NavButton}">
                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                            <Viewbox Width="14" Height="14" Margin="0,0,9,0">
                                <Path Data="M1,1 L6.5,1 L6.5,6.5 L1,6.5 Z M8.5,1 L14,1 L14,6.5 L8.5,6.5 Z M1,8.5 L6.5,8.5 L6.5,14 L1,14 Z M8.5,8.5 L14,8.5 L14,14 L8.5,14 Z"
                                      Stroke="{Binding Foreground, RelativeSource={RelativeSource AncestorType=Button}}"
                                      StrokeThickness="1.2" Fill="Transparent"
                                      StrokeLineJoin="Round"/>
                            </Viewbox>
                            <TextBlock Text="Dashboard" VerticalAlignment="Center"/>
                        </StackPanel>
                    </Button>
                    <Button x:Name="navBackup" Style="{StaticResource NavButton}">
                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                            <Viewbox Width="14" Height="14" Margin="0,0,9,0">
                                <Path Data="M7.5,1 L7.5,10 M4.5,7 L7.5,10.5 L10.5,7 M2,12 L2,13.5 Q2,14 2.5,14 L12.5,14 Q13,14 13,13.5 L13,12"
                                      Stroke="{Binding Foreground, RelativeSource={RelativeSource AncestorType=Button}}"
                                      StrokeThickness="1.2" Fill="Transparent"
                                      StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
                            </Viewbox>
                            <TextBlock Text="New backup" VerticalAlignment="Center"/>
                        </StackPanel>
                    </Button>
                    <Button x:Name="navRestore" Style="{StaticResource NavButton}">
                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                            <Viewbox Width="14" Height="14" Margin="0,0,9,0">
                                <Path Data="M7.5,13 L7.5,4 M4.5,7 L7.5,3.5 L10.5,7 M2,12 L2,13.5 Q2,14 2.5,14 L12.5,14 Q13,14 13,13.5 L13,12"
                                      Stroke="{Binding Foreground, RelativeSource={RelativeSource AncestorType=Button}}"
                                      StrokeThickness="1.2" Fill="Transparent"
                                      StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
                            </Viewbox>
                            <TextBlock Text="Restore" VerticalAlignment="Center"/>
                        </StackPanel>
                    </Button>
                    <Button x:Name="navHistory" Style="{StaticResource NavButton}">
                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                            <Viewbox Width="14" Height="14" Margin="0,0,9,0">
                                <Path Data="M7.5,1.5 A6,6 0 1 1 1.5,7.5 M7.5,4.5 L7.5,7.5 L9.5,9"
                                      Stroke="{Binding Foreground, RelativeSource={RelativeSource AncestorType=Button}}"
                                      StrokeThickness="1.2" Fill="Transparent"
                                      StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
                            </Viewbox>
                            <TextBlock Text="History" VerticalAlignment="Center"/>
                        </StackPanel>
                    </Button>
                    <Button x:Name="navAbout" Style="{StaticResource NavButton}">
                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                            <Viewbox Width="14" Height="14" Margin="0,0,9,0">
                                <Path Data="M7.5,1.5 A6,6 0 1 0 7.5,13.5 A6,6 0 1 0 7.5,1.5 M7.5,6.5 L7.5,7 M7.5,8 L7.5,11"
                                      Stroke="{Binding Foreground, RelativeSource={RelativeSource AncestorType=Button}}"
                                      StrokeThickness="1.2" Fill="Transparent"
                                      StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
                            </Viewbox>
                            <TextBlock Text="About" VerticalAlignment="Center"/>
                        </StackPanel>
                    </Button>
                </StackPanel>

                <Border Grid.Row="2" BorderBrush="{StaticResource BorderColor}" BorderThickness="0,0.5,0,0"
                        Padding="14,12">
                    <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                        <Border Width="30" Height="30" CornerRadius="15"
                                Background="{StaticResource AccentColor}" Margin="0,0,10,0">
                            <TextBlock x:Name="txtUserInitials" Text="?" FontSize="15" FontWeight="SemiBold"
                                       Foreground="White"
                                       HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <StackPanel VerticalAlignment="Center">
                            <TextBlock x:Name="txtSidebarUser" Text="" FontSize="13" FontWeight="Medium"
                                       Foreground="{StaticResource TextPrimary}"/>
                            <TextBlock x:Name="txtSidebarMachine" Text="" FontSize="11"
                                       Foreground="{StaticResource TextMuted}"/>
                        </StackPanel>
                    </StackPanel>
                </Border>
            </Grid>
        </Border>

        <Border Grid.Column="1" Background="{StaticResource WindowBg}" Padding="24">
            <Grid>
                <Grid x:Name="viewHome" Visibility="Visible">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>

                    <Grid Grid.Row="0" Margin="0,0,0,20">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Text="Dashboard" FontSize="24" FontWeight="SemiBold"
                                   Foreground="{StaticResource TextPrimary}" VerticalAlignment="Center"/>
                        <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                            <Ellipse x:Name="ellDriveStatus" Width="7" Height="7" Fill="{StaticResource GreenDot}" Margin="0,0,6,0"
                                     VerticalAlignment="Center"/>
                            <TextBlock x:Name="txtDriveStatus" Text="No location set" FontSize="13"
                                       Foreground="{StaticResource TextSecondary}" VerticalAlignment="Center"
                                       Margin="0,0,10,0"/>
                            <Button x:Name="btnChangeLocation" Content="Change Location" Height="26"
                                    Style="{StaticResource ActionButton}" FontSize="11" Padding="8,3"
                                    VerticalAlignment="Center"/>
                        </StackPanel>
                    </Grid>

                    <Grid Grid.Row="1" Margin="0,0,0,16">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="12"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="12"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Border Grid.Column="0" Background="{StaticResource CardBg}"
                                BorderBrush="{StaticResource BorderColor}" BorderThickness="0.5"
                                CornerRadius="10" Padding="18,16">
                            <StackPanel>
                                <TextBlock Text="Total backups" FontSize="15" Foreground="{StaticResource TextSecondary}"
                                           Margin="0,0,0,6"/>
                                <TextBlock x:Name="txtStatTotalBackups" Text="—" FontSize="28" FontWeight="SemiBold"
                                           Foreground="{StaticResource TextPrimary}"/>
                                <TextBlock x:Name="txtStatMachineCount" Text="" FontSize="12"
                                           Foreground="{StaticResource TextMuted}" Margin="0,3,0,0"/>
                            </StackPanel>
                        </Border>
                        <Border Grid.Column="2" Background="{StaticResource CardBg}"
                                BorderBrush="{StaticResource BorderColor}" BorderThickness="0.5"
                                CornerRadius="10" Padding="18,16">
                            <StackPanel>
                                <TextBlock Text="Last backup" FontSize="15" Foreground="{StaticResource TextSecondary}"
                                           Margin="0,0,0,6"/>
                                <TextBlock x:Name="txtStatLastBackup" Text="—" FontSize="18" FontWeight="SemiBold"
                                           Foreground="{StaticResource TextPrimary}" TextWrapping="Wrap"/>
                                <TextBlock x:Name="txtStatLastMachine" Text="" FontSize="12"
                                           Foreground="{StaticResource TextMuted}" Margin="0,3,0,0"/>
                            </StackPanel>
                        </Border>
                        <Border Grid.Column="4" Background="{StaticResource CardBg}"
                                BorderBrush="{StaticResource BorderColor}" BorderThickness="0.5"
                                CornerRadius="10" Padding="18,16">
                            <StackPanel>
                                <TextBlock Text="Total stored" FontSize="15" Foreground="{StaticResource TextSecondary}"
                                           Margin="0,0,0,6"/>
                                <TextBlock x:Name="txtStatStorage" Text="—" FontSize="28" FontWeight="SemiBold"
                                           Foreground="{StaticResource TextPrimary}"/>
                                <TextBlock x:Name="txtStatStorageLabel" Text="in backup location" FontSize="12"
                                           Foreground="{StaticResource TextMuted}" Margin="0,3,0,0"/>
                            </StackPanel>
                        </Border>
                    </Grid>

                    <Grid Grid.Row="2" Margin="0,0,0,16">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="12"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <Button x:Name="btnHomeBackup" Grid.Column="0" Height="80"
                                Style="{StaticResource TileButton}">
                            <Grid Margin="18,0">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Border Grid.Column="0" Width="36" Height="36" CornerRadius="9"
                                        Background="{StaticResource AccentBg}" Margin="0,0,14,0">
                                    <Viewbox Width="16" Height="16">
                                        <Path Data="M8,2 L8,11 M5,8 L8,11.5 L11,8 M3,14 L13,14"
                                              Stroke="{StaticResource AccentColor}" StrokeThickness="1.4"
                                              StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
                                    </Viewbox>
                                </Border>
                                <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                    <TextBlock Text="New backup" FontSize="15" FontWeight="Medium"
                                               Foreground="{StaticResource TextPrimary}"/>
                                    <TextBlock Text="Back up files to U:\ or custom location" FontSize="15"
                                               Foreground="{StaticResource TextSecondary}"
                                               Margin="0,3,0,0"/>
                                </StackPanel>
                            </Grid>
                        </Button>
                        <Button x:Name="btnHomeRestore" Grid.Column="2" Height="80"
                                Style="{StaticResource TileButton}">
                            <Grid Margin="18,0">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Border Grid.Column="0" Width="36" Height="36" CornerRadius="9"
                                        Background="{StaticResource AccentBg}" Margin="0,0,14,0">
                                    <Viewbox Width="16" Height="16">
                                        <Path Data="M8,12 L8,3 M5,6 L8,2.5 L11,6 M3,14 L13,14"
                                              Stroke="{StaticResource AccentColor}" StrokeThickness="1.4"
                                              StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
                                    </Viewbox>
                                </Border>
                                <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                    <TextBlock Text="Restore files" FontSize="15" FontWeight="Medium"
                                               Foreground="{StaticResource TextPrimary}"/>
                                    <TextBlock Text="From any backup point" FontSize="15"
                                               Foreground="{StaticResource TextSecondary}"
                                               Margin="0,3,0,0"/>
                                </StackPanel>
                            </Grid>
                        </Button>
                    </Grid>

                    <Border Grid.Row="3" Background="{StaticResource CardBg}"
                            BorderBrush="{StaticResource BorderColor}" BorderThickness="0.5"
                            CornerRadius="10" Padding="18">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <Grid Grid.Row="0" Margin="0,0,0,14">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Text="RECENT BACKUPS" FontSize="11" FontWeight="SemiBold"
                                           Foreground="{StaticResource TextSecondary}"
                                           VerticalAlignment="Center"/>
                                <Button x:Name="btnViewAll" Grid.Column="1" Content="View all"
                                        Style="{StaticResource LinkButton}"
                                        VerticalAlignment="Center"/>
                            </Grid>
                            <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" MaxHeight="340" Padding="0,0,12,0">
                                <ItemsControl x:Name="icRecentBackups">
                                    <ItemsControl.ItemTemplate>
                                        <DataTemplate>
                                            <Border Margin="0,0,0,8"
                                                    Background="#FF0C0C10"
                                                    BorderBrush="{StaticResource BorderColor}"
                                                    BorderThickness="0.5"
                                                    CornerRadius="8"
                                                    Padding="16,14"
                                                    Cursor="Hand">
                                                <Grid>
                                                    <Grid.ColumnDefinitions>
                                                        <ColumnDefinition Width="Auto"/>
                                                        <ColumnDefinition Width="*"/>
                                                        <ColumnDefinition Width="Auto"/>
                                                    </Grid.ColumnDefinitions>
                                                    <Ellipse Grid.Column="0" Width="9" Height="9"
                                                             Margin="0,5,14,0" VerticalAlignment="Top">
                                                        <Ellipse.Fill>
                                                            <SolidColorBrush Color="{Binding DotColor}"/>
                                                        </Ellipse.Fill>
                                                    </Ellipse>
                                                    <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                                        <TextBlock Text="{Binding PCName}" FontSize="18" FontWeight="SemiBold"
                                                                   Foreground="{StaticResource TextPrimary}"/>
                                                        <TextBlock Text="{Binding DateDisplay}" FontSize="13"
                                                                   Foreground="{StaticResource TextSecondary}" Margin="0,4,0,6"/>
                                                        <ItemsControl ItemsSource="{Binding ContentTags}">
                                                            <ItemsControl.ItemsPanel>
                                                                <ItemsPanelTemplate>
                                                                    <WrapPanel/>
                                                                </ItemsPanelTemplate>
                                                            </ItemsControl.ItemsPanel>
                                                            <ItemsControl.ItemTemplate>
                                                                <DataTemplate>
                                                                    <Border Background="#FF1A1A2E" BorderBrush="#FF2C2C40"
                                                                            BorderThickness="0.5" CornerRadius="4"
                                                                            Padding="6,2" Margin="0,0,4,4">
                                                                        <TextBlock Text="{Binding}" FontSize="11"
                                                                                   Foreground="{StaticResource TextSecondary}"/>
                                                                    </Border>
                                                                </DataTemplate>
                                                            </ItemsControl.ItemTemplate>
                                                        </ItemsControl>
                                                    </StackPanel>
                                                    <TextBlock Grid.Column="2" Text="{Binding SizeFormatted}"
                                                               FontSize="15" Foreground="{StaticResource TextSecondary}"
                                                               VerticalAlignment="Top" Margin="12,2,0,0"/>
                                                </Grid>
                                            </Border>
                                        </DataTemplate>
                                    </ItemsControl.ItemTemplate>
                                </ItemsControl>
                            </ScrollViewer>
                        </Grid>
                    </Border>
                </Grid>

                <ScrollViewer x:Name="viewBackup" VerticalScrollBarVisibility="Auto" Visibility="Collapsed" Padding="0,0,12,0">
                    <StackPanel Margin="0,0,0,16">
                        <Grid Margin="0,0,0,24">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel>
                                <TextBlock Text="New Backup" FontSize="24" FontWeight="SemiBold"
                                           Foreground="{StaticResource TextPrimary}"/>
                                <TextBlock Text="Choose what to back up and where" FontSize="13"
                                           Foreground="{StaticResource TextSecondary}" Margin="0,4,0,0"/>
                            </StackPanel>
                            <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                                <Button x:Name="btnBackupSelectAll" Content="Select All" Width="90" Height="30"
                                        Style="{StaticResource ActionButton}" Margin="0,0,8,0"/>
                                <Button x:Name="btnBackupSelectNone" Content="Select None" Width="90" Height="30"
                                        Style="{StaticResource ActionButton}"/>
                            </StackPanel>
                        </Grid>

                        <TextBlock Text="BACKUP DESTINATION" FontSize="11" FontWeight="SemiBold"
                                   Foreground="{StaticResource TextMuted}" Margin="0,0,0,10"/>
                        <Border Background="{StaticResource CardBg}" BorderBrush="{StaticResource BorderColor}"
                                BorderThickness="0.5" CornerRadius="10" Padding="16,14" Margin="0,0,0,16">
                            <StackPanel>
                                <TextBlock x:Name="txtBackupDestLabel" FontSize="13" TextWrapping="Wrap"
                                           Foreground="{StaticResource TextSecondary}" Margin="0,0,0,10"/>
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBox x:Name="tbBackupDest" Grid.Column="0" Height="34" Margin="0,0,8,0"
                                             FontSize="13" IsReadOnly="True"
                                             Background="#FF0E0E11" BorderBrush="{StaticResource BorderColor}"
                                             BorderThickness="0.5"/>
                                    <Button x:Name="btnBrowseBackupDest" Grid.Column="1" Content="Browse..."
                                            Width="90" Height="34" Style="{StaticResource ActionButton}"/>
                                </Grid>
                            </StackPanel>
                        </Border>

                        <TextBlock Text="USER FOLDERS" FontSize="11" FontWeight="SemiBold"
                                   Foreground="{StaticResource TextMuted}" Margin="0,0,0,10"/>
                        <Border Background="{StaticResource CardBg}" BorderBrush="{StaticResource BorderColor}"
                                BorderThickness="0.5" CornerRadius="10" Margin="0,0,0,16">
                            <StackPanel>
                                <Border Padding="16,14" BorderBrush="{StaticResource BorderColor}" BorderThickness="0,0,0,0.5">
                                    <Grid>
                                        <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                        <Border Width="32" Height="32" CornerRadius="8" Background="#FF1E1E30" Margin="0,0,14,0" VerticalAlignment="Center">
                                            <Viewbox Width="15" Height="15">
                                                <Path Data="M2,3 L14,3 L14,11 L2,11 Z M6,11 L6,13 M10,11 L10,13 M4,13 L12,13"
                                                      Stroke="{StaticResource AccentColor}" StrokeThickness="1.3" Fill="Transparent"
                                                      StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
                                            </Viewbox>
                                        </Border>
                                        <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                            <TextBlock Text="Desktop" FontSize="15" FontWeight="Medium" Foreground="{StaticResource TextPrimary}"/>
                                            <TextBlock Text="Files and shortcuts on your desktop" FontSize="15" Foreground="{StaticResource TextSecondary}" Margin="0,2,0,0"/>
                                        </StackPanel>
                                        <CheckBox x:Name="cbDesktop" Grid.Column="2" IsChecked="True" VerticalAlignment="Center"/>
                                    </Grid>
                                </Border>
                                <Border Padding="16,14" BorderBrush="{StaticResource BorderColor}" BorderThickness="0,0,0,0.5">
                                    <Grid>
                                        <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                        <Border Width="32" Height="32" CornerRadius="8" Background="#FF1E1E30" Margin="0,0,14,0" VerticalAlignment="Center">
                                            <Viewbox Width="15" Height="15">
                                                <Path Data="M4,2 L10,2 L14,6 L14,14 L4,14 Z M10,2 L10,6 L14,6 M6,9 L12,9 M6,12 L10,12"
                                                      Stroke="{StaticResource AccentColor}" StrokeThickness="1.3" Fill="Transparent"
                                                      StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
                                            </Viewbox>
                                        </Border>
                                        <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                            <TextBlock Text="Documents" FontSize="15" FontWeight="Medium" Foreground="{StaticResource TextPrimary}"/>
                                            <TextBlock Text="Word docs, PDFs and other documents" FontSize="15" Foreground="{StaticResource TextSecondary}" Margin="0,2,0,0"/>
                                        </StackPanel>
                                        <CheckBox x:Name="cbDocuments" Grid.Column="2" IsChecked="True" VerticalAlignment="Center"/>
                                    </Grid>
                                </Border>
                                <Border Padding="16,14" BorderBrush="{StaticResource BorderColor}" BorderThickness="0,0,0,0.5">
                                    <Grid>
                                        <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                        <Border Width="32" Height="32" CornerRadius="8" Background="#FF1E1E30" Margin="0,0,14,0" VerticalAlignment="Center">
                                            <Viewbox Width="15" Height="15">
                                                <Path Data="M1,3 L15,3 L15,13 L1,13 Z M1,9 L5,6 L9,10 L11,8 L15,13 M11,5.5 A0.5,0.5 0 1 1 11,6 Z"
                                                      Stroke="{StaticResource AccentColor}" StrokeThickness="1.3" Fill="Transparent"
                                                      StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
                                            </Viewbox>
                                        </Border>
                                        <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                            <TextBlock Text="Pictures" FontSize="15" FontWeight="Medium" Foreground="{StaticResource TextPrimary}"/>
                                            <TextBlock Text="Photos and images" FontSize="15" Foreground="{StaticResource TextSecondary}" Margin="0,2,0,0"/>
                                        </StackPanel>
                                        <CheckBox x:Name="cbPictures" Grid.Column="2" IsChecked="True" VerticalAlignment="Center"/>
                                    </Grid>
                                </Border>
                                <Border Padding="16,14" BorderBrush="{StaticResource BorderColor}" BorderThickness="0,0,0,0.5">
                                    <Grid>
                                        <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                        <Border Width="32" Height="32" CornerRadius="8" Background="#FF1E1E30" Margin="0,0,14,0" VerticalAlignment="Center">
                                            <Viewbox Width="15" Height="15">
                                                <Path Data="M8,2 L8,11 M5,8 L8,11.5 L11,8 M3,14 L13,14"
                                                      Stroke="{StaticResource AccentColor}" StrokeThickness="1.3" Fill="Transparent"
                                                      StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
                                            </Viewbox>
                                        </Border>
                                        <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                            <TextBlock Text="Downloads" FontSize="15" FontWeight="Medium" Foreground="{StaticResource TextPrimary}"/>
                                            <TextBlock Text="Files saved from the internet" FontSize="15" Foreground="{StaticResource TextSecondary}" Margin="0,2,0,0"/>
                                        </StackPanel>
                                        <CheckBox x:Name="cbDownloads" Grid.Column="2" IsChecked="True" VerticalAlignment="Center"/>
                                    </Grid>
                                </Border>
                                <Border Padding="16,14" BorderBrush="{StaticResource BorderColor}" BorderThickness="0,0,0,0.5">
                                    <Grid>
                                        <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                        <Border Width="32" Height="32" CornerRadius="8" Background="#FF1E1E30" Margin="0,0,14,0" VerticalAlignment="Center">
                                            <Viewbox Width="15" Height="15">
                                                <Path Data="M6,13 A2,2 0 1 1 6,13.1 M6,13 L6,3 L13,1 L13,11 M13,11 A2,2 0 1 1 13,11.1"
                                                      Stroke="{StaticResource AccentColor}" StrokeThickness="1.3" Fill="Transparent"
                                                      StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
                                            </Viewbox>
                                        </Border>
                                        <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                            <TextBlock Text="Music" FontSize="15" FontWeight="Medium" Foreground="{StaticResource TextPrimary}"/>
                                            <TextBlock Text="Audio files and playlists" FontSize="15" Foreground="{StaticResource TextSecondary}" Margin="0,2,0,0"/>
                                        </StackPanel>
                                        <CheckBox x:Name="cbMusic" Grid.Column="2" IsChecked="False" VerticalAlignment="Center"/>
                                    </Grid>
                                </Border>
                                <Border Padding="16,14">
                                    <Grid>
                                        <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                        <Border Width="32" Height="32" CornerRadius="8" Background="#FF1E1E30" Margin="0,0,14,0" VerticalAlignment="Center">
                                            <Viewbox Width="15" Height="15">
                                                <Path Data="M2,3 L14,3 L14,13 L2,13 Z M6,6 L11,8 L6,10 Z"
                                                      Stroke="{StaticResource AccentColor}" StrokeThickness="1.3" Fill="Transparent"
                                                      StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
                                            </Viewbox>
                                        </Border>
                                        <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                            <TextBlock Text="Videos" FontSize="15" FontWeight="Medium" Foreground="{StaticResource TextPrimary}"/>
                                            <TextBlock Text="Video files and recordings" FontSize="15" Foreground="{StaticResource TextSecondary}" Margin="0,2,0,0"/>
                                        </StackPanel>
                                        <CheckBox x:Name="cbVideos" Grid.Column="2" IsChecked="False" VerticalAlignment="Center"/>
                                    </Grid>
                                </Border>
                            </StackPanel>
                        </Border>

                        <TextBlock Text="BROWSER DATA" FontSize="11" FontWeight="SemiBold"
                                   Foreground="{StaticResource TextMuted}" Margin="0,0,0,10"/>
                        <Border Background="{StaticResource CardBg}" BorderBrush="{StaticResource BorderColor}"
                                BorderThickness="0.5" CornerRadius="10" Margin="0,0,0,16">
                            <StackPanel>
                                <Border Padding="16,14" BorderBrush="{StaticResource BorderColor}" BorderThickness="0,0,0,0.5">
                                    <Grid>
                                        <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                        <Border Width="32" Height="32" CornerRadius="8" Background="#FF1E2A1E" Margin="0,0,14,0" VerticalAlignment="Center">
                                            <Viewbox Width="15" Height="15">
                                                <Path Data="M8,8 m-5,0 a5,5 0 1 0 10,0 a5,5 0 1 0 -10,0 M8,5 L8,8 L11,8"
                                                      Stroke="#FF2ECC71" StrokeThickness="1.3" Fill="Transparent"
                                                      StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
                                            </Viewbox>
                                        </Border>
                                        <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                            <TextBlock Text="Chrome bookmarks" FontSize="15" FontWeight="Medium" Foreground="{StaticResource TextPrimary}"/>
                                            <TextBlock Text="All Chrome profiles" FontSize="15" Foreground="{StaticResource TextSecondary}" Margin="0,2,0,0"/>
                                        </StackPanel>
                                        <CheckBox x:Name="cbChromeBookmarks" Grid.Column="2" IsChecked="True" VerticalAlignment="Center"/>
                                    </Grid>
                                </Border>
                                <Border Padding="16,14" BorderBrush="{StaticResource BorderColor}" BorderThickness="0,0,0,0.5">
                                    <Grid>
                                        <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                        <Border Width="32" Height="32" CornerRadius="8" Background="#FF1A1E2A" Margin="0,0,14,0" VerticalAlignment="Center">
                                            <Viewbox Width="15" Height="15">
                                                <Path Data="M13,7 Q13,3 8,3 Q3,3 3,8 Q3,13 8,13 Q11,13 12.5,11 M13,7 Q10,9 7,8"
                                                      Stroke="#FF5C8CF5" StrokeThickness="1.3" Fill="Transparent"
                                                      StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
                                            </Viewbox>
                                        </Border>
                                        <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                            <TextBlock Text="Edge favorites" FontSize="15" FontWeight="Medium" Foreground="{StaticResource TextPrimary}"/>
                                            <TextBlock Text="All Edge profiles and bookmarks" FontSize="15" Foreground="{StaticResource TextSecondary}" Margin="0,2,0,0"/>
                                        </StackPanel>
                                        <CheckBox x:Name="cbEdgeFavorites" Grid.Column="2" IsChecked="True" VerticalAlignment="Center"/>
                                    </Grid>
                                </Border>
                                <Border Padding="16,14" BorderBrush="{StaticResource BorderColor}" BorderThickness="0,0.5,0,0">
                                    <Grid>
                                        <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                        <Border Width="32" Height="32" CornerRadius="8" Background="#FF1E1A2A" Margin="0,0,14,0" VerticalAlignment="Center">
                                            <Viewbox Width="15" Height="15">
                                                <Path Data="M8,2 Q4,2 3,6 Q2,10 5,12 Q7,14 8,12 Q9,14 11,12 Q14,10 13,6 Q12,2 8,2 M8,2 L8,12 M5,5 Q6.5,6 8,5 Q9.5,6 11,5"
                                                      Stroke="#FFFF9500" StrokeThickness="1.3" Fill="Transparent"
                                                      StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
                                            </Viewbox>
                                        </Border>
                                        <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                            <TextBlock Text="Firefox bookmarks" FontSize="15" FontWeight="Medium" Foreground="{StaticResource TextPrimary}"/>
                                            <TextBlock Text="All Firefox profiles (places.sqlite)" FontSize="15" Foreground="{StaticResource TextSecondary}" Margin="0,2,0,0"/>
                                        </StackPanel>
                                        <CheckBox x:Name="cbFirefoxBookmarks" Grid.Column="2" IsChecked="True" VerticalAlignment="Center"/>
                                    </Grid>
                                </Border>
                            </StackPanel>
                        </Border>

                        <TextBlock Text="SYSTEM &amp; DEVICES" FontSize="11" FontWeight="SemiBold"
                                   Foreground="{StaticResource TextMuted}" Margin="0,0,0,10"/>
                        <Border Background="{StaticResource CardBg}" BorderBrush="{StaticResource BorderColor}"
                                BorderThickness="0.5" CornerRadius="10" Margin="0,0,0,16">
                            <StackPanel>
                                <Border Padding="16,14" BorderBrush="{StaticResource BorderColor}" BorderThickness="0,0,0,0.5">
                                    <Grid>
                                        <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                        <Border Width="32" Height="32" CornerRadius="8" Background="#FF1E1E2A" Margin="0,0,14,0" VerticalAlignment="Center">
                                            <Viewbox Width="15" Height="15">
                                                <Path Data="M2,4 L14,4 L14,10 L2,10 Z M4,2 L12,2 L12,4 L4,4 Z M3,10 L13,10 L13,14 L3,14 Z M4,11 L12,11 M4,12 L12,12 M4,13 L8,13"
                                                      Stroke="{StaticResource AccentColor}" StrokeThickness="1.3" Fill="Transparent"
                                                      StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
                                            </Viewbox>
                                        </Border>
                                        <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                            <TextBlock Text="Network Printers" FontSize="15" FontWeight="Medium" Foreground="{StaticResource TextPrimary}"/>
                                            <TextBlock Text="Export mapped network printer connections" FontSize="15" Foreground="{StaticResource TextSecondary}" Margin="0,2,0,0"/>
                                        </StackPanel>
                                        <CheckBox x:Name="cbPrinters" Grid.Column="2" IsChecked="True" VerticalAlignment="Center"/>
                                    </Grid>
                                </Border>
                                <Border Padding="16,14">
                                    <Grid>
                                        <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                        <Border Width="32" Height="32" CornerRadius="8" Background="#FF1A1E2A" Margin="0,0,14,0" VerticalAlignment="Center">
                                            <Viewbox Width="15" Height="15">
                                                <Path Data="M1,3 L15,3 L15,13 L1,13 Z M1,9 L5,6 L9,10 L11,8 L15,13 M3,3 L3,13 M13,3 L13,13"
                                                      Stroke="#FF9B8CF5" StrokeThickness="1.3" Fill="Transparent"
                                                      StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
                                            </Viewbox>
                                        </Border>
                                        <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                            <TextBlock Text="Wallpaper" FontSize="15" FontWeight="Medium" Foreground="{StaticResource TextPrimary}"/>
                                            <TextBlock Text="Desktop wallpaper image and display settings" FontSize="15" Foreground="{StaticResource TextSecondary}" Margin="0,2,0,0"/>
                                        </StackPanel>
                                        <CheckBox x:Name="cbWallpaper" Grid.Column="2" IsChecked="True" VerticalAlignment="Center"/>
                                    </Grid>
                                </Border>
                                <Border Padding="16,14" BorderBrush="{StaticResource BorderColor}" BorderThickness="0,0,0,0.5">
                                    <Grid>
                                        <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                        <Border Width="32" Height="32" CornerRadius="8" Background="#FF1A2A1E" Margin="0,0,14,0" VerticalAlignment="Center">
                                            <Viewbox Width="15" Height="15">
                                                <Path Data="M2,2 L10,2 L14,6 L14,14 L2,14 Z M10,2 L10,6 L14,6 M5,8 L11,8 M5,10 L9,10"
                                                      Stroke="#FF4EC994" StrokeThickness="1.3" Fill="Transparent"
                                                      StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
                                            </Viewbox>
                                        </Border>
                                        <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                            <TextBlock Text="File Associations" FontSize="15" FontWeight="Medium" Foreground="{StaticResource TextPrimary}"/>
                                            <TextBlock Text="Default app assignments for file types and protocols" FontSize="15" Foreground="{StaticResource TextSecondary}" Margin="0,2,0,0"/>
                                        </StackPanel>
                                        <CheckBox x:Name="cbFileAssociations" Grid.Column="2" IsChecked="True" VerticalAlignment="Center"/>
                                    </Grid>
                                </Border>
                            </StackPanel>
                        </Border>

                        <TextBlock Text="OPTIONS" FontSize="11" FontWeight="SemiBold"
                                   Foreground="{StaticResource TextMuted}" Margin="0,0,0,10"/>
                        <Border Background="{StaticResource CardBg}" BorderBrush="{StaticResource BorderColor}"
                                BorderThickness="0.5" CornerRadius="10" Margin="0,0,0,16">
                            <StackPanel>
                                <Border Padding="16,14" BorderBrush="{StaticResource BorderColor}" BorderThickness="0,0,0,0.5">
                                    <Grid>
                                        <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                        <Border Width="32" Height="32" CornerRadius="8" Background="#FF2A1E1E" Margin="0,0,14,0" VerticalAlignment="Center">
                                            <Viewbox Width="15" Height="15">
                                                <Path Data="M5,8 L5,6 Q5,3 8,3 Q11,3 11,6 L11,8 M3,8 L13,8 L13,14 L3,14 Z M8,10 L8,12"
                                                      Stroke="#FFE8924A" StrokeThickness="1.3" Fill="Transparent"
                                                      StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
                                            </Viewbox>
                                        </Border>
                                        <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                            <TextBlock Text="Password helper pages" FontSize="15" FontWeight="Medium" Foreground="{StaticResource TextPrimary}"/>
                                            <TextBlock Text="Create export helper pages for Chrome and Edge" FontSize="15" Foreground="{StaticResource TextSecondary}" Margin="0,2,0,0"/>
                                        </StackPanel>
                                        <CheckBox x:Name="cbPasswordHelpers" Grid.Column="2" IsChecked="True" VerticalAlignment="Center"/>
                                    </Grid>
                                </Border>
                                <Border Padding="16,14" BorderBrush="{StaticResource BorderColor}" BorderThickness="0,0,0,0.5">
                                    <Grid>
                                        <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                        <Border Width="32" Height="32" CornerRadius="8" Background="#FF1E1E2A" Margin="0,0,14,0" VerticalAlignment="Center">
                                            <Viewbox Width="15" Height="15">
                                                <Path Data="M3,8 Q3,3 8,3 Q13,3 13,8 Q13,13 8,13 M8,6 L8,10 M6,8 L10,8"
                                                      Stroke="{StaticResource AccentColor}" StrokeThickness="1.3" Fill="Transparent"
                                                      StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
                                            </Viewbox>
                                        </Border>
                                        <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                            <TextBlock Text="Open helper pages automatically" FontSize="15" FontWeight="Medium" Foreground="{StaticResource TextPrimary}"/>
                                            <TextBlock Text="Launch browser pages when backup starts" FontSize="15" Foreground="{StaticResource TextSecondary}" Margin="0,2,0,0"/>
                                        </StackPanel>
                                        <CheckBox x:Name="cbOpenHelperPages" Grid.Column="2" IsChecked="True" VerticalAlignment="Center"/>
                                    </Grid>
                                </Border>
                                <Border Padding="16,14">
                                    <Grid>
                                        <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                        <Border Width="32" Height="32" CornerRadius="8" Background="#FF1E1E2A" Margin="0,0,14,0" VerticalAlignment="Center">
                                            <Viewbox Width="15" Height="15">
                                                <Path Data="M2,5 L6,5 L7,3 L14,3 L14,12 L2,12 Z"
                                                      Stroke="{StaticResource AccentColor}" StrokeThickness="1.3" Fill="Transparent"
                                                      StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
                                            </Viewbox>
                                        </Border>
                                        <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                            <TextBlock Text="Open backup folder when done" FontSize="15" FontWeight="Medium" Foreground="{StaticResource TextPrimary}"/>
                                            <TextBlock Text="Show the backup location in File Explorer" FontSize="15" Foreground="{StaticResource TextSecondary}" Margin="0,2,0,0"/>
                                        </StackPanel>
                                        <CheckBox x:Name="cbOpenBackupFolder" Grid.Column="2" IsChecked="True" VerticalAlignment="Center"/>
                                    </Grid>
                                </Border>
                            </StackPanel>
                        </Border>

                        <TextBlock Text="BACKUP NOTE" FontSize="11" FontWeight="SemiBold"
                                   Foreground="{StaticResource TextMuted}" Margin="0,0,0,10"/>
                        <Border Background="{StaticResource CardBg}" BorderBrush="{StaticResource BorderColor}"
                                BorderThickness="0.5" CornerRadius="10" Padding="16,12" Margin="0,0,0,16">
                            <StackPanel>
                                <TextBlock Text="Optional note for this backup" FontSize="15"
                                           Foreground="{StaticResource TextSecondary}" Margin="0,0,0,8"/>
                                <TextBox x:Name="tbBackupNote" Height="36" TextWrapping="Wrap"
                                         AcceptsReturn="False" FontSize="13"
                                         Background="#FF0E0E11" BorderBrush="{StaticResource BorderColor}"
                                         BorderThickness="0.5"/>
                            </StackPanel>
                        </Border>

                        <Border Background="{StaticResource CardBg}" BorderBrush="{StaticResource BorderColor}"
                                BorderThickness="0.5" CornerRadius="10" Padding="16,14" Margin="0,0,0,16">
                            <StackPanel>
                                <Grid Margin="0,0,0,10">
                                    <TextBlock x:Name="txtStatus" Text="Ready to back up." FontSize="14" FontWeight="Medium"
                                               Foreground="{StaticResource TextPrimary}"/>
                                </Grid>
                                <ProgressBar x:Name="pbProgress" Minimum="0" Maximum="100" Value="0" Margin="0,0,0,10"/>
                                <TextBox x:Name="txtActivity" Height="65" IsReadOnly="True"
                                         TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                                         FontFamily="Consolas" FontSize="11"
                                         Background="#FF090910" Foreground="#FFAAAACC" BorderThickness="0"/>
                            </StackPanel>
                        </Border>

                        <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                            <Button x:Name="btnCancel" Content="Cancel" Width="90" Height="34"
                                    Style="{StaticResource ActionButton}" Margin="0,0,10,0"/>
                            <Button x:Name="btnStart" Content="Start Backup" Width="130" Height="34"
                                    Style="{StaticResource PrimaryButton}"/>
                        </StackPanel>
                    </StackPanel>
                </ScrollViewer>

                <Grid x:Name="viewRestore" Visibility="Collapsed">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <Grid Grid.Row="0" Margin="0,0,0,20">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel>
                            <TextBlock Text="Restore from Backup" FontSize="24" FontWeight="SemiBold"
                                       Foreground="{StaticResource TextPrimary}"/>
                            <TextBlock Text="Select a backup point and choose what to restore" FontSize="13"
                                       Foreground="{StaticResource TextSecondary}" Margin="0,4,0,0"/>
                        </StackPanel>
                        <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                            <Button x:Name="btnRestoreSelectAll" Content="Select All" Width="90" Height="30"
                                    Style="{StaticResource ActionButton}" Margin="0,0,8,0"/>
                            <Button x:Name="btnRestoreSelectNone" Content="Select None" Width="90" Height="30"
                                    Style="{StaticResource ActionButton}"/>
                        </StackPanel>
                    </Grid>

                    <Border Grid.Row="1" Background="{StaticResource CardBg}" BorderBrush="{StaticResource BorderColor}"
                            BorderThickness="0.5" CornerRadius="10" Padding="16,14" Margin="0,0,0,14">
                        <StackPanel>
                            <TextBlock Text="BACKUP LOCATION" FontSize="11" FontWeight="SemiBold"
                                       Foreground="{StaticResource TextMuted}" Margin="0,0,0,12"/>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBox x:Name="tbRestoreSource" Grid.Column="0" Height="34" Margin="0,0,8,0"
                                         FontSize="13"/>
                                <Button x:Name="btnFindLatestBackup" Grid.Column="1" Content="Find Latest"
                                        Width="100" Height="34" Style="{StaticResource ActionButton}" Margin="0,0,8,0"/>
                                <Button x:Name="btnBrowseBackup" Grid.Column="2" Content="Browse..."
                                        Width="90" Height="34" Style="{StaticResource ActionButton}"/>
                            </Grid>
                        </StackPanel>
                    </Border>

                    <Border Grid.Row="2" x:Name="statsPanel" Visibility="Collapsed"
                            Background="{StaticResource AccentBg}"
                            BorderBrush="{StaticResource AccentColor}"
                            BorderThickness="0.5" CornerRadius="10" Padding="16,12" Margin="0,0,0,14">
                        <StackPanel>
                            <TextBlock Text="BACKUP INFORMATION" FontWeight="SemiBold" FontSize="11"
                                       Foreground="{StaticResource AccentColor}" Margin="0,0,0,6"/>
                            <TextBlock x:Name="txtBackupInfo" FontSize="13" TextWrapping="Wrap"
                                       Foreground="{StaticResource TextSecondary}"/>
                        </StackPanel>
                    </Border>

                    <ScrollViewer Grid.Row="3" VerticalScrollBarVisibility="Auto" Padding="0,0,12,0">
                        <StackPanel>
                            <TextBlock Text="WHAT TO RESTORE" FontSize="11" FontWeight="SemiBold"
                                       Foreground="{StaticResource TextMuted}" Margin="0,0,0,10"/>
                            <Border Background="{StaticResource CardBg}" BorderBrush="{StaticResource BorderColor}"
                                    BorderThickness="0.5" CornerRadius="10" Margin="0,0,0,14">
                                <StackPanel>
                                    <Border Padding="16,14" BorderBrush="{StaticResource BorderColor}" BorderThickness="0,0,0,0.5">
                                        <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                            <Border Width="32" Height="32" CornerRadius="8" Background="#FF1E1E30" Margin="0,0,14,0" VerticalAlignment="Center">
                                                <Viewbox Width="15" Height="15"><Path Data="M2,3 L14,3 L14,11 L2,11 Z M6,11 L6,13 M10,11 L10,13 M4,13 L12,13" Stroke="{StaticResource AccentColor}" StrokeThickness="1.3" Fill="Transparent" StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/></Viewbox>
                                            </Border>
                                            <StackPanel Grid.Column="1" VerticalAlignment="Center"><TextBlock Text="Desktop" FontSize="15" FontWeight="Medium" Foreground="{StaticResource TextPrimary}"/><TextBlock Text="Desktop files and shortcuts" FontSize="15" Foreground="{StaticResource TextSecondary}" Margin="0,2,0,0"/></StackPanel>
                                            <CheckBox x:Name="cbRDesktop" Grid.Column="2" IsChecked="True" VerticalAlignment="Center"/>
                                        </Grid>
                                    </Border>
                                    <Border Padding="16,14" BorderBrush="{StaticResource BorderColor}" BorderThickness="0,0,0,0.5">
                                        <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                            <Border Width="32" Height="32" CornerRadius="8" Background="#FF1E1E30" Margin="0,0,14,0" VerticalAlignment="Center">
                                                <Viewbox Width="15" Height="15"><Path Data="M4,2 L10,2 L14,6 L14,14 L4,14 Z M10,2 L10,6 L14,6 M6,9 L12,9 M6,12 L10,12" Stroke="{StaticResource AccentColor}" StrokeThickness="1.3" Fill="Transparent" StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/></Viewbox>
                                            </Border>
                                            <StackPanel Grid.Column="1" VerticalAlignment="Center"><TextBlock Text="Documents" FontSize="15" FontWeight="Medium" Foreground="{StaticResource TextPrimary}"/><TextBlock Text="Word docs, PDFs and other files" FontSize="15" Foreground="{StaticResource TextSecondary}" Margin="0,2,0,0"/></StackPanel>
                                            <CheckBox x:Name="cbRDocuments" Grid.Column="2" IsChecked="True" VerticalAlignment="Center"/>
                                        </Grid>
                                    </Border>
                                    <Border Padding="16,14" BorderBrush="{StaticResource BorderColor}" BorderThickness="0,0,0,0.5">
                                        <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                            <Border Width="32" Height="32" CornerRadius="8" Background="#FF1E1E30" Margin="0,0,14,0" VerticalAlignment="Center">
                                                <Viewbox Width="15" Height="15"><Path Data="M1,3 L15,3 L15,13 L1,13 Z M1,9 L5,6 L9,10 L11,8 L15,13" Stroke="{StaticResource AccentColor}" StrokeThickness="1.3" Fill="Transparent" StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/></Viewbox>
                                            </Border>
                                            <StackPanel Grid.Column="1" VerticalAlignment="Center"><TextBlock Text="Pictures" FontSize="15" FontWeight="Medium" Foreground="{StaticResource TextPrimary}"/><TextBlock Text="Photos and images" FontSize="15" Foreground="{StaticResource TextSecondary}" Margin="0,2,0,0"/></StackPanel>
                                            <CheckBox x:Name="cbRPictures" Grid.Column="2" IsChecked="True" VerticalAlignment="Center"/>
                                        </Grid>
                                    </Border>
                                    <Border Padding="16,14" BorderBrush="{StaticResource BorderColor}" BorderThickness="0,0,0,0.5">
                                        <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                            <Border Width="32" Height="32" CornerRadius="8" Background="#FF1E1E30" Margin="0,0,14,0" VerticalAlignment="Center">
                                                <Viewbox Width="15" Height="15"><Path Data="M8,2 L8,11 M5,8 L8,11.5 L11,8 M3,14 L13,14" Stroke="{StaticResource AccentColor}" StrokeThickness="1.3" Fill="Transparent" StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/></Viewbox>
                                            </Border>
                                            <StackPanel Grid.Column="1" VerticalAlignment="Center"><TextBlock Text="Downloads" FontSize="15" FontWeight="Medium" Foreground="{StaticResource TextPrimary}"/><TextBlock Text="Files saved from the internet" FontSize="15" Foreground="{StaticResource TextSecondary}" Margin="0,2,0,0"/></StackPanel>
                                            <CheckBox x:Name="cbRDownloads" Grid.Column="2" IsChecked="True" VerticalAlignment="Center"/>
                                        </Grid>
                                    </Border>
                                    <Border Padding="16,14" BorderBrush="{StaticResource BorderColor}" BorderThickness="0,0,0,0.5">
                                        <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                            <Border Width="32" Height="32" CornerRadius="8" Background="#FF1E1E30" Margin="0,0,14,0" VerticalAlignment="Center">
                                                <Viewbox Width="15" Height="15"><Path Data="M6,13 A2,2 0 1 1 6,13.1 M6,13 L6,3 L13,1 L13,11 M13,11 A2,2 0 1 1 13,11.1" Stroke="{StaticResource AccentColor}" StrokeThickness="1.3" Fill="Transparent" StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/></Viewbox>
                                            </Border>
                                            <StackPanel Grid.Column="1" VerticalAlignment="Center"><TextBlock Text="Music" FontSize="15" FontWeight="Medium" Foreground="{StaticResource TextPrimary}"/><TextBlock Text="Audio files and playlists" FontSize="15" Foreground="{StaticResource TextSecondary}" Margin="0,2,0,0"/></StackPanel>
                                            <CheckBox x:Name="cbRMusic" Grid.Column="2" IsChecked="False" VerticalAlignment="Center"/>
                                        </Grid>
                                    </Border>
                                    <Border Padding="16,14" BorderBrush="{StaticResource BorderColor}" BorderThickness="0,0,0,0.5">
                                        <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                            <Border Width="32" Height="32" CornerRadius="8" Background="#FF1E1E30" Margin="0,0,14,0" VerticalAlignment="Center">
                                                <Viewbox Width="15" Height="15"><Path Data="M2,3 L14,3 L14,13 L2,13 Z M6,6 L11,8 L6,10 Z" Stroke="{StaticResource AccentColor}" StrokeThickness="1.3" Fill="Transparent" StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/></Viewbox>
                                            </Border>
                                            <StackPanel Grid.Column="1" VerticalAlignment="Center"><TextBlock Text="Videos" FontSize="15" FontWeight="Medium" Foreground="{StaticResource TextPrimary}"/><TextBlock Text="Video files and recordings" FontSize="15" Foreground="{StaticResource TextSecondary}" Margin="0,2,0,0"/></StackPanel>
                                            <CheckBox x:Name="cbRVideos" Grid.Column="2" IsChecked="False" VerticalAlignment="Center"/>
                                        </Grid>
                                    </Border>
                                    <Border Padding="16,14" BorderBrush="{StaticResource BorderColor}" BorderThickness="0,0,0,0.5">
                                        <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                            <Border Width="32" Height="32" CornerRadius="8" Background="#FF1E2A1E" Margin="0,0,14,0" VerticalAlignment="Center">
                                                <Viewbox Width="15" Height="15"><Path Data="M8,8 m-5,0 a5,5 0 1 0 10,0 a5,5 0 1 0 -10,0 M8,5 L8,8 L11,8" Stroke="#FF2ECC71" StrokeThickness="1.3" Fill="Transparent" StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/></Viewbox>
                                            </Border>
                                            <StackPanel Grid.Column="1" VerticalAlignment="Center"><TextBlock Text="Chrome bookmarks" FontSize="15" FontWeight="Medium" Foreground="{StaticResource TextPrimary}"/><TextBlock Text="All Chrome profiles" FontSize="15" Foreground="{StaticResource TextSecondary}" Margin="0,2,0,0"/></StackPanel>
                                            <CheckBox x:Name="cbRChromeBookmarks" Grid.Column="2" IsChecked="True" VerticalAlignment="Center"/>
                                        </Grid>
                                    </Border>
                                    <Border Padding="16,14" BorderBrush="{StaticResource BorderColor}" BorderThickness="0,0,0,0.5">
                                        <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                            <Border Width="32" Height="32" CornerRadius="8" Background="#FF1A1E2A" Margin="0,0,14,0" VerticalAlignment="Center">
                                                <Viewbox Width="15" Height="15"><Path Data="M13,7 Q13,3 8,3 Q3,3 3,8 Q3,13 8,13 Q11,13 12.5,11 M13,7 Q10,9 7,8" Stroke="#FF5C8CF5" StrokeThickness="1.3" Fill="Transparent" StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/></Viewbox>
                                            </Border>
                                            <StackPanel Grid.Column="1" VerticalAlignment="Center"><TextBlock Text="Edge favorites" FontSize="15" FontWeight="Medium" Foreground="{StaticResource TextPrimary}"/><TextBlock Text="All Edge profiles and bookmarks" FontSize="15" Foreground="{StaticResource TextSecondary}" Margin="0,2,0,0"/></StackPanel>
                                            <CheckBox x:Name="cbREdgeBookmarks" Grid.Column="2" IsChecked="True" VerticalAlignment="Center"/>
                                        </Grid>
                                    </Border>
                                    <Border Padding="16,14" BorderBrush="{StaticResource BorderColor}" BorderThickness="0,0,0,0.5">
                                        <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                            <Border Width="32" Height="32" CornerRadius="8" Background="#FF1E1A2A" Margin="0,0,14,0" VerticalAlignment="Center">
                                                <Viewbox Width="15" Height="15"><Path Data="M8,2 Q4,2 3,6 Q2,10 5,12 Q7,14 8,12 Q9,14 11,12 Q14,10 13,6 Q12,2 8,2 M8,2 L8,12 M5,5 Q6.5,6 8,5 Q9.5,6 11,5" Stroke="#FFFF9500" StrokeThickness="1.3" Fill="Transparent" StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/></Viewbox>
                                            </Border>
                                            <StackPanel Grid.Column="1" VerticalAlignment="Center"><TextBlock Text="Firefox bookmarks" FontSize="15" FontWeight="Medium" Foreground="{StaticResource TextPrimary}"/><TextBlock Text="All Firefox profiles (places.sqlite)" FontSize="15" Foreground="{StaticResource TextSecondary}" Margin="0,2,0,0"/></StackPanel>
                                            <CheckBox x:Name="cbRFirefoxBookmarks" Grid.Column="2" IsChecked="True" VerticalAlignment="Center"/>
                                        </Grid>
                                    </Border>
                                    <Border Padding="16,14">
                                        <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                            <Border Width="32" Height="32" CornerRadius="8" Background="#FF2A1E2A" Margin="0,0,14,0" VerticalAlignment="Center">
                                                <Viewbox Width="15" Height="15"><Path Data="M3,3 L13,3 L13,13 L3,13 Z M6,3 L6,13 M3,8 L13,8" Stroke="#FF9B59B6" StrokeThickness="1.3" Fill="Transparent" StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/></Viewbox>
                                            </Border>
                                            <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                                <TextBlock Text="Legacy Favorites folder" FontSize="15" FontWeight="Medium" Foreground="{StaticResource TextPrimary}"/>
                                                <TextBlock Text="Internet Explorer / legacy favorites (if present)" FontSize="15" Foreground="{StaticResource TextSecondary}" Margin="0,2,0,0"/>
                                            </StackPanel>
                                            <CheckBox x:Name="cbRLegacyFavorites" Grid.Column="2" IsChecked="False" VerticalAlignment="Center"/>
                                        </Grid>
                                    </Border>
                                </StackPanel>
                            </Border>

                            <TextBlock Text="SYSTEM &amp; DEVICES" FontSize="11" FontWeight="SemiBold"
                                       Foreground="{StaticResource TextMuted}" Margin="0,0,0,10"/>
                            <Border Background="{StaticResource CardBg}" BorderBrush="{StaticResource BorderColor}"
                                    BorderThickness="0.5" CornerRadius="10" Margin="0,0,0,14">
                                <StackPanel>
                                    <Border Padding="16,14" BorderBrush="{StaticResource BorderColor}" BorderThickness="0,0,0,0.5">
                                        <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                            <Border Width="32" Height="32" CornerRadius="8" Background="#FF1E1E2A" Margin="0,0,14,0" VerticalAlignment="Center">
                                                <Viewbox Width="15" Height="15"><Path Data="M2,4 L14,4 L14,10 L2,10 Z M4,2 L12,2 L12,4 L4,4 Z M3,10 L13,10 L13,14 L3,14 Z M4,11 L12,11 M4,12 L12,12 M4,13 L8,13" Stroke="{StaticResource AccentColor}" StrokeThickness="1.3" Fill="Transparent" StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/></Viewbox>
                                            </Border>
                                            <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                                <TextBlock Text="Network Printers" FontSize="15" FontWeight="Medium" Foreground="{StaticResource TextPrimary}"/>
                                                <TextBlock Text="Restore mapped network printer connections" FontSize="15" Foreground="{StaticResource TextSecondary}" Margin="0,2,0,0"/>
                                            </StackPanel>
                                            <CheckBox x:Name="cbRPrinters" Grid.Column="2" IsChecked="False" VerticalAlignment="Center"/>
                                        </Grid>
                                    </Border>
                                    <Border Padding="16,14">
                                        <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                            <Border Width="32" Height="32" CornerRadius="8" Background="#FF1A1E2A" Margin="0,0,14,0" VerticalAlignment="Center">
                                                <Viewbox Width="15" Height="15"><Path Data="M1,3 L15,3 L15,13 L1,13 Z M1,9 L5,6 L9,10 L11,8 L15,13 M3,3 L3,13 M13,3 L13,13" Stroke="#FF9B8CF5" StrokeThickness="1.3" Fill="Transparent" StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/></Viewbox>
                                            </Border>
                                            <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                                <TextBlock Text="Wallpaper" FontSize="15" FontWeight="Medium" Foreground="{StaticResource TextPrimary}"/>
                                                <TextBlock Text="Restore desktop wallpaper and display settings" FontSize="15" Foreground="{StaticResource TextSecondary}" Margin="0,2,0,0"/>
                                            </StackPanel>
                                            <CheckBox x:Name="cbRWallpaper" Grid.Column="2" IsChecked="True" VerticalAlignment="Center"/>
                                        </Grid>
                                    </Border>
                                    <Border Padding="16,14">
                                        <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                            <Border Width="32" Height="32" CornerRadius="8" Background="#FF1A2A1E" Margin="0,0,14,0" VerticalAlignment="Center">
                                                <Viewbox Width="15" Height="15"><Path Data="M2,2 L10,2 L14,6 L14,14 L2,14 Z M10,2 L10,6 L14,6 M5,8 L11,8 M5,10 L9,10" Stroke="#FF4EC994" StrokeThickness="1.3" Fill="Transparent" StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/></Viewbox>
                                            </Border>
                                            <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                                <TextBlock Text="File Associations" FontSize="15" FontWeight="Medium" Foreground="{StaticResource TextPrimary}"/>
                                                <TextBlock Text="Restore default app assignments for file types and protocols" FontSize="15" Foreground="{StaticResource TextSecondary}" Margin="0,2,0,0"/>
                                            </StackPanel>
                                            <CheckBox x:Name="cbRFileAssociations" Grid.Column="2" IsChecked="False" VerticalAlignment="Center"/>
                                        </Grid>
                                    </Border>
                                </StackPanel>
                            </Border>

                            <Border Background="#FF1A1310" BorderBrush="#FF3A2810" BorderThickness="0.5"
                                    CornerRadius="8" Padding="14,10" Margin="0,0,0,14">
                                <StackPanel Orientation="Horizontal">
                                    <Viewbox Width="14" Height="14" Margin="0,0,10,0" VerticalAlignment="Top">
                                        <Path Data="M8,2 L15,14 L1,14 Z M8,7 L8,10 M8,12 L8,12.5"
                                              Stroke="#FFE8924A" StrokeThickness="1.3" Fill="Transparent"
                                              StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
                                    </Viewbox>
                                    <TextBlock FontSize="15" TextWrapping="Wrap" Foreground="#FFE8924A">
                                        Close Chrome, Edge, and Firefox before restoring bookmarks to avoid file lock errors.
                                    </TextBlock>
                                </StackPanel>
                            </Border>

                            <Border Background="{StaticResource CardBg}" BorderBrush="{StaticResource BorderColor}"
                                    BorderThickness="0.5" CornerRadius="10" Padding="16,14" Margin="0,0,0,4">
                                <StackPanel>
                                    <TextBlock x:Name="txtRestoreStatus" Text="Ready to restore." FontSize="14" FontWeight="Medium"
                                               Foreground="{StaticResource TextPrimary}" Margin="0,0,0,10"/>
                                    <ProgressBar x:Name="pbRestoreProgress" Minimum="0" Maximum="100" Value="0" Margin="0,0,0,10"/>
                                    <TextBox x:Name="txtRestoreActivity" Height="90" IsReadOnly="True"
                                             TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"
                                             FontFamily="Consolas" FontSize="11"
                                             Background="#FF090910" Foreground="#FFAAAACC" BorderThickness="0"/>
                                </StackPanel>
                            </Border>
                        </StackPanel>
                    </ScrollViewer>

                    <StackPanel Grid.Row="4" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,12,0,0">
                        <Button x:Name="btnCancelRestore" Content="Cancel" Width="90" Height="34"
                                Style="{StaticResource ActionButton}" Margin="0,0,10,0"/>
                        <Button x:Name="btnStartRestore" Content="Start Restore" Width="140" Height="34"
                                Style="{StaticResource PrimaryButton}"/>
                    </StackPanel>
                </Grid>

                <Grid x:Name="viewHistory" Visibility="Collapsed">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <Grid Grid.Row="0" Margin="0,0,0,20">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel>
                            <TextBlock Text="Backup History" FontSize="24" FontWeight="SemiBold"
                                       Foreground="{StaticResource TextPrimary}"/>
                            <TextBlock x:Name="txtHistoryCount" Text="Loading..."
                                       FontSize="13" Foreground="{StaticResource TextSecondary}" Margin="0,4,0,0"/>
                        </StackPanel>
                        <Button x:Name="btnRefreshHistory" Grid.Column="1" Content="↻  Refresh"
                                Width="100" Height="34" Style="{StaticResource ActionButton}"
                                VerticalAlignment="Center"/>
                    </Grid>

                    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" Margin="0,0,0,14" Padding="0,0,12,0">
                        <ItemsControl x:Name="lvBackupHistory">
                            <ItemsControl.ItemTemplate>
                                <DataTemplate>
                                    <Border Margin="0,0,0,8"
                                            Background="{StaticResource CardBg}"
                                            BorderBrush="{StaticResource BorderColor}"
                                            BorderThickness="0.5" CornerRadius="10"
                                            Cursor="Hand">
                                        <Grid>
                                            <Grid.ColumnDefinitions>
                                                <ColumnDefinition Width="4"/>
                                                <ColumnDefinition Width="*"/>
                                            </Grid.ColumnDefinitions>
                                            <Border Grid.Column="0" CornerRadius="10,0,0,10">
                                                <Border.Background>
                                                    <SolidColorBrush Color="{Binding DotColor}"/>
                                                </Border.Background>
                                            </Border>
                                            <Grid Grid.Column="1" Margin="16,14,16,14">
                                                <Grid.ColumnDefinitions>
                                                    <ColumnDefinition Width="Auto"/>
                                                    <ColumnDefinition Width="*"/>
                                                    <ColumnDefinition Width="Auto"/>
                                                    <ColumnDefinition Width="Auto"/>
                                                </Grid.ColumnDefinitions>

                                                <Border Grid.Column="0" Width="36" Height="36" CornerRadius="8"
                                                        Background="{StaticResource AccentBg}" Margin="0,0,14,0"
                                                        VerticalAlignment="Center">
                                                    <Viewbox Width="16" Height="16">
                                                        <Path Data="M2,3 L14,3 L14,11 L2,11 Z M4,11 L4,13 L12,13 L12,11 M6,13 L6,14 M10,13 L10,14 M0,14 L16,14"
                                                              Stroke="{StaticResource AccentColor}" StrokeThickness="1.2"
                                                              Fill="Transparent" StrokeLineJoin="Round"
                                                              StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
                                                    </Viewbox>
                                                </Border>

                                                <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                                    <TextBlock Text="{Binding PCName}" FontSize="16" FontWeight="SemiBold"
                                                               Foreground="{StaticResource TextPrimary}"/>
                                                    <TextBlock Text="{Binding Timestamp}" FontSize="15"
                                                               Foreground="{StaticResource TextSecondary}" Margin="0,2,0,4"/>
                                                    <ItemsControl ItemsSource="{Binding ContentTags}">
                                                        <ItemsControl.ItemsPanel>
                                                            <ItemsPanelTemplate>
                                                                <WrapPanel/>
                                                            </ItemsPanelTemplate>
                                                        </ItemsControl.ItemsPanel>
                                                        <ItemsControl.ItemTemplate>
                                                            <DataTemplate>
                                                                <Border Background="#FF1A1A2E" BorderBrush="#FF2C2C40"
                                                                        BorderThickness="0.5" CornerRadius="4"
                                                                        Padding="6,2" Margin="0,0,4,4">
                                                                    <TextBlock Text="{Binding}" FontSize="11"
                                                                               Foreground="{StaticResource TextSecondary}"/>
                                                                </Border>
                                                            </DataTemplate>
                                                        </ItemsControl.ItemTemplate>
                                                    </ItemsControl>
                                                </StackPanel>

                                                <Border Grid.Column="2" Background="#FF0C0C18"
                                                        BorderBrush="{StaticResource BorderColor}"
                                                        BorderThickness="0.5" CornerRadius="6"
                                                        Padding="10,6" Margin="12,0" VerticalAlignment="Center">
                                                    <TextBlock Text="{Binding SizeFormatted}" FontSize="14"
                                                               FontWeight="Medium" Foreground="{StaticResource TextPrimary}"/>
                                                </Border>

                                                <Button Grid.Column="3" Content="Restore →"
                                                        Tag="{Binding}"
                                                        Width="90" Height="30"
                                                        Style="{StaticResource PrimaryButton}"
                                                        VerticalAlignment="Center"/>
                                            </Grid>
                                        </Grid>
                                    </Border>
                                </DataTemplate>
                            </ItemsControl.ItemTemplate>
                        </ItemsControl>
                    </ScrollViewer>

                    <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
                        <Button x:Name="btnUseSelectedBackup" Content="Use Selected for Restore"
                                Width="190" Height="34" IsEnabled="False"
                                Style="{StaticResource PrimaryButton}"/>
                    </StackPanel>
                </Grid>

                <!-- About View -->
                <ScrollViewer x:Name="viewAbout" VerticalScrollBarVisibility="Auto" Visibility="Collapsed" Padding="0,0,12,0">
                    <StackPanel MaxWidth="560" HorizontalAlignment="Left">

                        <!-- Header -->
                        <TextBlock Text="About" FontSize="24" FontWeight="SemiBold"
                                   Foreground="{StaticResource TextPrimary}" Margin="0,0,0,24"/>

                        <!-- App identity card -->
                        <Border Background="{StaticResource CardBg}"
                                BorderBrush="{StaticResource BorderColor}" BorderThickness="0.5"
                                CornerRadius="10" Padding="20,18" Margin="0,0,0,12">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Border Grid.Column="0" Width="52" Height="52" CornerRadius="12"
                                        Background="{StaticResource AccentBg}" Margin="0,0,16,0"
                                        VerticalAlignment="Center">
                                    <Viewbox Width="26" Height="26">
                                        <Path Data="M8,2 L8,10 M5,7 L8,10.5 L11,7 M3,13 L13,13 M3,3 L13,3"
                                              Stroke="{StaticResource AccentColor}" StrokeThickness="1.3"
                                              StrokeLineJoin="Round" StrokeStartLineCap="Round" StrokeEndLineCap="Round"/>
                                    </Viewbox>
                                </Border>
                                <StackPanel Grid.Column="1" VerticalAlignment="Center">
                                    <TextBlock Text="DeskSave" FontSize="20" FontWeight="SemiBold"
                                               Foreground="{StaticResource TextPrimary}"/>
                                    <TextBlock Text="Windows PC backup &amp; restore utility"
                                               FontSize="13" Foreground="{StaticResource TextSecondary}" Margin="0,3,0,0"/>
                                    <StackPanel Orientation="Horizontal" Margin="0,8,0,0">
                                        <Border Background="{StaticResource AccentBg}" CornerRadius="4" Padding="7,3" Margin="0,0,6,0">
                                            <StackPanel Orientation="Horizontal">
                                                <Ellipse Width="6" Height="6" Fill="{StaticResource GreenDot}" Margin="0,0,5,0" VerticalAlignment="Center"/>
                                                <TextBlock Text="Version 1.0" FontSize="11" Foreground="{StaticResource AccentColor}" FontWeight="Medium"/>
                                            </StackPanel>
                                        </Border>
                                        <Border Background="{StaticResource BadgeBg}" CornerRadius="4" Padding="7,3" Margin="0,0,6,0">
                                            <TextBlock Text="Windows 10 / 11" FontSize="11" Foreground="{StaticResource TextSecondary}"/>
                                        </Border>
                                        <Border Background="{StaticResource BadgeBg}" CornerRadius="4" Padding="7,3">
                                            <TextBlock Text="PowerShell 5.1" FontSize="11" Foreground="{StaticResource TextSecondary}"/>
                                        </Border>
                                    </StackPanel>
                                </StackPanel>
                            </Grid>
                        </Border>

                        <!-- Build info cards row -->
                        <Grid Margin="0,0,0,12">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="10"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Border Grid.Column="0" Background="{StaticResource CardBg}"
                                    BorderBrush="{StaticResource BorderColor}" BorderThickness="0.5"
                                    CornerRadius="8" Padding="16,12">
                                <StackPanel>
                                    <TextBlock Text="RELEASE" FontSize="10" FontWeight="Medium"
                                               Foreground="{StaticResource TextMuted}" Margin="0,0,0,5"/>
                                    <TextBlock Text="2026" FontSize="18" FontWeight="SemiBold"
                                               Foreground="{StaticResource TextPrimary}"/>
                                </StackPanel>
                            </Border>
                            <Border Grid.Column="2" Background="{StaticResource CardBg}"
                                    BorderBrush="{StaticResource BorderColor}" BorderThickness="0.5"
                                    CornerRadius="8" Padding="16,12">
                                <StackPanel>
                                    <TextBlock Text="BUILD" FontSize="10" FontWeight="Medium"
                                               Foreground="{StaticResource TextMuted}" Margin="0,0,0,5"/>
                                    <TextBlock Text="v1.0.0" FontSize="18" FontWeight="SemiBold"
                                               Foreground="{StaticResource AccentColor}"/>
                                </StackPanel>
                            </Border>
                        </Grid>

                        <!-- Executable / Log location row -->
                        <Grid Margin="0,0,0,12">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="10"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Border Grid.Column="0" Background="{StaticResource CardBg}"
                                    BorderBrush="{StaticResource BorderColor}" BorderThickness="0.5"
                                    CornerRadius="8" Padding="16,12">
                                <StackPanel>
                                    <TextBlock Text="EXECUTABLE" FontSize="10" FontWeight="Medium"
                                               Foreground="{StaticResource TextMuted}" Margin="0,0,0,5"/>
                                    <TextBlock Text="DeskSave.exe" FontSize="13"
                                               Foreground="{StaticResource AccentColor}" FontFamily="Consolas"/>
                                </StackPanel>
                            </Border>
                            <Border Grid.Column="2" Background="{StaticResource CardBg}"
                                    BorderBrush="{StaticResource BorderColor}" BorderThickness="0.5"
                                    CornerRadius="8" Padding="16,12">
                                <StackPanel>
                                    <TextBlock Text="LOG LOCATION" FontSize="10" FontWeight="Medium"
                                               Foreground="{StaticResource TextMuted}" Margin="0,0,0,5"/>
                                    <TextBlock Text="%TEMP%\BackupToolLogs" FontSize="12"
                                               Foreground="{StaticResource AccentColor}" FontFamily="Consolas" TextWrapping="Wrap"/>
                                </StackPanel>
                            </Border>
                        </Grid>

                        <!-- Separator -->
                        <Border Height="0.5" Background="{StaticResource BorderColor}" Margin="0,4,0,12"/>

                        <!-- Author card -->
                        <Border Background="{StaticResource CardBg}"
                                BorderBrush="{StaticResource BorderColor}" BorderThickness="0.5"
                                CornerRadius="10" Padding="18,14" Margin="0,0,0,16">
                            <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                                <Border Width="40" Height="40" CornerRadius="20"
                                        Background="{StaticResource AccentColor}" Margin="0,0,14,0">
                                    <TextBlock Text="JW" FontSize="14" FontWeight="SemiBold" Foreground="White"
                                               HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                </Border>
                                <StackPanel VerticalAlignment="Center">
                                    <TextBlock Text="Joshua Winkles" FontSize="14" FontWeight="SemiBold"
                                               Foreground="{StaticResource TextPrimary}"/>
                                    <StackPanel Orientation="Horizontal" Margin="0,2,0,0">
                                        <TextBlock Text="Creator &amp; original author" FontSize="12"
                                                   Foreground="{StaticResource TextSecondary}"/>
                                        <TextBlock Text="  ·  Copyright © 2026" FontSize="12"
                                                   Foreground="{StaticResource TextMuted}"/>
                                    </StackPanel>
			<StackPanel Orientation="Horizontal" Margin="0,6,0,0">
    <TextBlock FontSize="13"
               FontWeight="SemiBold"
               Foreground="{StaticResource TextMuted}">

        <Hyperlink x:Name="KoFiLink"
                   NavigateUri="https://ko-fi.com/winklesit"
                   TextDecorations="None">

            <Hyperlink.Style>
                <Style TargetType="Hyperlink">
                    <Setter Property="Foreground" Value="#FF4DA3FF"/>
                    <Style.Triggers>
                        <Trigger Property="IsMouseOver" Value="True">
                            <Setter Property="TextDecorations" Value="Underline"/>
                            <Setter Property="Foreground" Value="#66B3FF"/>
                        </Trigger>
                    </Style.Triggers>
                </Style>
            </Hyperlink.Style>

            ☕ Support me on Ko-fi →
        </Hyperlink>

    </TextBlock>
</StackPanel>
                                </StackPanel>
                            </StackPanel>
                        </Border>

                        <!-- Feature list -->
                        <StackPanel Margin="0,0,0,16">
                            <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                                <TextBlock Text="✓" FontSize="13" Foreground="{StaticResource AccentColor}" Margin="0,0,10,0" VerticalAlignment="Center"/>
                                <TextBlock Text="Backs up Desktop, Documents, Pictures, Downloads, Music &amp; Videos"
                                           FontSize="13" Foreground="{StaticResource TextSecondary}" TextWrapping="Wrap"/>
                            </StackPanel>
                            <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                                <TextBlock Text="✓" FontSize="13" Foreground="{StaticResource AccentColor}" Margin="0,0,10,0" VerticalAlignment="Center"/>
                                <TextBlock Text="Chrome, Edge &amp; Firefox bookmark backup — all profiles"
                                           FontSize="13" Foreground="{StaticResource TextSecondary}"/>
                            </StackPanel>
                            <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                                <TextBlock Text="✓" FontSize="13" Foreground="{StaticResource AccentColor}" Margin="0,0,10,0" VerticalAlignment="Center"/>
                                <TextBlock Text="Network printer connections, wallpaper &amp; file associations"
                                           FontSize="13" Foreground="{StaticResource TextSecondary}"/>
                            </StackPanel>
                            <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                                <TextBlock Text="✓" FontSize="13" Foreground="{StaticResource AccentColor}" Margin="0,0,10,0" VerticalAlignment="Center"/>
                                <TextBlock Text="Timestamped, multi-PC backup history in one location"
                                           FontSize="13" Foreground="{StaticResource TextSecondary}"/>
                            </StackPanel>
                            <StackPanel Orientation="Horizontal">
                                <TextBlock Text="✓" FontSize="13" Foreground="{StaticResource AccentColor}" Margin="0,0,10,0" VerticalAlignment="Center"/>
                                <TextBlock Text="No installation required — portable executable"
                                           FontSize="13" Foreground="{StaticResource TextSecondary}"/>
                            </StackPanel>
                        </StackPanel>

                        <!-- Separator -->
                        <Border Height="0.5" Background="{StaticResource BorderColor}" Margin="0,0,0,12"/>

                        <!-- License footer -->
                        <Grid Margin="0,0,0,20">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel Grid.Column="0" VerticalAlignment="Center">
                                <TextBlock FontSize="12" Foreground="{StaticResource TextSecondary}" TextWrapping="Wrap">
                                    <Run Text="Released under the "/>
                                    <Run Text="GNU General Public License v3" Foreground="{StaticResource AccentColor}"/>
                                </TextBlock>
                                <TextBlock Text="Free to use, share, and modify under the terms of GPL-3.0"
                                           FontSize="12" Foreground="{StaticResource TextMuted}" Margin="0,3,0,0"/>
                            </StackPanel>
                            <Border Grid.Column="1" Background="{StaticResource BadgeBg}"
                                    BorderBrush="{StaticResource BorderColor}" BorderThickness="0.5"
                                    CornerRadius="5" Padding="10,5" VerticalAlignment="Center">
                                <TextBlock Text="GPL-3.0" FontSize="12" Foreground="{StaticResource TextSecondary}"/>
                            </Border>
                        </Grid>

                        <!-- Close button -->
                        <Button x:Name="btnAboutClose" Content="Close"
                                Height="36" Style="{StaticResource ActionButton}"
                                HorizontalAlignment="Stretch" Margin="0,0,0,8"/>

                    </StackPanel>
                </ScrollViewer>

            </Grid>
        </Border>
    </Grid>
</Window>
"@

    $Reader = New-Object System.Xml.XmlNodeReader $Xaml
    $Window = [Windows.Markup.XamlReader]::Load($Reader)

    # ---------------------------
    # Embed application icon (Base64-encoded .ico)
    # ---------------------------
    try {
        $icoBase64 = @"
AAABAAEAEBAAAAAAIADxAAAAFgAAAIlQTkcNChoKAAAADUlIRFIAAAAQAAAAEAgGAAAAH/P/YQAAALhJ
REFUeJy1kjEOgkAQRR+ERDtKD0CFJQmJvbfgHB7Dc3ALexMSSrfiAJZ0WmE1OouzLA2/miXz3i67A1sm
PxynWE+2BOZFDTABjM9HYvV6H2egl3HofrWSJfOjanhXXr/1211MUWaBS5E+EaWrqIVsJ9D/b62jArm0
0DoqWJsUoGx7xqHznkjvqneXvrLtATVIp9trAnBNBYSHSUCA+3mfeJMoEolrKvKi/gODgpDIAqU2BZZE
QzofXX9PXZR0He0AAAAASUVORK5CYII=
"@
        $icoBytes  = [Convert]::FromBase64String(($icoBase64 -replace '\s',''))
        $icoStream = New-Object System.IO.MemoryStream(,$icoBytes)
        $bitmap    = New-Object System.Windows.Media.Imaging.BitmapImage
        $bitmap.BeginInit()
        $bitmap.StreamSource = $icoStream
        $bitmap.CacheOption  = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $bitmap.EndInit()
        $bitmap.Freeze()
        $Window.Icon = $bitmap
    } catch {
        Write-BackupToolLog "Icon load failed: $($_.Exception.Message)"
    }

    # ---------------------------
    # Load sidebar logo from embedded Base64 PNG
    # ---------------------------
    try {
        $pngBase64 = "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"
        $pngBytes  = [Convert]::FromBase64String(($pngBase64 -replace '\s',''))
        $pngStream = New-Object System.IO.MemoryStream(,$pngBytes)
        $logoBitmap = New-Object System.Windows.Media.Imaging.BitmapImage
        $logoBitmap.BeginInit()
        $logoBitmap.StreamSource = $pngStream
        $logoBitmap.CacheOption  = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $logoBitmap.EndInit()
        $logoBitmap.Freeze()
        $imgSidebarLogo = $Window.FindName("imgSidebarLogo")
        if ($imgSidebarLogo) { $imgSidebarLogo.Source = $logoBitmap }
    } catch {
        Write-BackupToolLog "Sidebar logo load failed: $($_.Exception.Message)"
    }

    # Map all controls
    $navHome = $Window.FindName("navHome")
    $navBackup = $Window.FindName("navBackup")
    $navRestore = $Window.FindName("navRestore")
    $navHistory = $Window.FindName("navHistory")
    $navAbout = $Window.FindName("navAbout")
    
    $viewHome = $Window.FindName("viewHome")
    $viewBackup = $Window.FindName("viewBackup")
    $viewRestore = $Window.FindName("viewRestore")
    $viewHistory = $Window.FindName("viewHistory")
    $viewAbout = $Window.FindName("viewAbout")
    
    # Dashboard controls
    $btnHomeBackup   = $Window.FindName("btnHomeBackup")
    $btnHomeRestore  = $Window.FindName("btnHomeRestore")
    $btnViewAll      = $Window.FindName("btnViewAll")
    $icRecentBackups = $Window.FindName("icRecentBackups")
    
    # Backup controls
    $cbDesktop = $Window.FindName("cbDesktop")
    $cbDocuments = $Window.FindName("cbDocuments")
    $cbPictures = $Window.FindName("cbPictures")
    $cbDownloads = $Window.FindName("cbDownloads")
    $cbMusic = $Window.FindName("cbMusic")
    $cbVideos = $Window.FindName("cbVideos")
    $cbChromeBookmarks = $Window.FindName("cbChromeBookmarks")
    $cbEdgeFavorites = $Window.FindName("cbEdgeFavorites")
    $cbFirefoxBookmarks = $Window.FindName("cbFirefoxBookmarks")
    $cbPrinters = $Window.FindName("cbPrinters")
    $cbWallpaper = $Window.FindName("cbWallpaper")
    $cbFileAssociations = $Window.FindName("cbFileAssociations")
    $cbPasswordHelpers = $Window.FindName("cbPasswordHelpers")
    $cbOpenHelperPages = $Window.FindName("cbOpenHelperPages")
    $cbOpenBackupFolder = $Window.FindName("cbOpenBackupFolder")
    $btnBackupSelectAll  = $Window.FindName("btnBackupSelectAll")
    $btnBackupSelectNone = $Window.FindName("btnBackupSelectNone")
    $tbBackupNote = $Window.FindName("tbBackupNote")
    $txtStatus = $Window.FindName("txtStatus")
    $pbProgress = $Window.FindName("pbProgress")
    $txtActivity = $Window.FindName("txtActivity")
    $btnStart = $Window.FindName("btnStart")
    $btnCancel = $Window.FindName("btnCancel")
    
    # Restore controls
    $tbRestoreSource = $Window.FindName("tbRestoreSource")
    $btnFindLatestBackup = $Window.FindName("btnFindLatestBackup")
    $btnBrowseBackup = $Window.FindName("btnBrowseBackup")
    $statsPanel = $Window.FindName("statsPanel")
    $txtBackupInfo = $Window.FindName("txtBackupInfo")
    $cbRDesktop = $Window.FindName("cbRDesktop")
    $cbRDocuments = $Window.FindName("cbRDocuments")
    $cbRPictures = $Window.FindName("cbRPictures")
    $cbRDownloads = $Window.FindName("cbRDownloads")
    $cbRMusic = $Window.FindName("cbRMusic")
    $cbRVideos = $Window.FindName("cbRVideos")
    $cbRChromeBookmarks = $Window.FindName("cbRChromeBookmarks")
    $cbREdgeBookmarks = $Window.FindName("cbREdgeBookmarks")
    $cbRFirefoxBookmarks = $Window.FindName("cbRFirefoxBookmarks")
    $cbRLegacyFavorites = $Window.FindName("cbRLegacyFavorites")
    $cbRPrinters = $Window.FindName("cbRPrinters")
    $cbRWallpaper = $Window.FindName("cbRWallpaper")
    $cbRFileAssociations = $Window.FindName("cbRFileAssociations")
    $btnRestoreSelectAll  = $Window.FindName("btnRestoreSelectAll")
    $btnRestoreSelectNone = $Window.FindName("btnRestoreSelectNone")
    $txtRestoreStatus = $Window.FindName("txtRestoreStatus")
    $pbRestoreProgress = $Window.FindName("pbRestoreProgress")
    $txtRestoreActivity = $Window.FindName("txtRestoreActivity")
    $btnStartRestore = $Window.FindName("btnStartRestore")
    $btnCancelRestore = $Window.FindName("btnCancelRestore")
    
    # History controls
    $btnRefreshHistory = $Window.FindName("btnRefreshHistory")
    $txtHistoryCount = $Window.FindName("txtHistoryCount")
    $lvBackupHistory = $Window.FindName("lvBackupHistory")
    $btnUseSelectedBackup = $Window.FindName("btnUseSelectedBackup")

    # Drive status indicator (dashboard header)
    $ellDriveStatus  = $Window.FindName("ellDriveStatus")
    $txtDriveStatus  = $Window.FindName("txtDriveStatus")
    $txtStatStorageLabel = $Window.FindName("txtStatStorageLabel")

    # Backup destination controls
    $tbBackupDest        = $Window.FindName("tbBackupDest")
    $btnBrowseBackupDest = $Window.FindName("btnBrowseBackupDest")
    $txtBackupDestLabel  = $Window.FindName("txtBackupDestLabel")
	
	# Hyperlink to Support me page
	$hyperlink = $window.FindName("KoFiLink")

$hyperlink.Add_RequestNavigate({
    param($sender, $e)

    Start-Process $e.Uri.AbsoluteUri
    $e.Handled = $true
})

    # ---------------------------
    # Helper: Update backup location indicator on dashboard
    # ---------------------------
    function Update-LocationIndicator {
        param([string]$RootPath)
        if ($RootPath) {
            $ellDriveStatus.Fill = [System.Windows.Media.SolidColorBrush]([System.Windows.Media.Color]::FromRgb(0x2E,0xCC,0x71))
            $shortPath = if ($RootPath.Length -gt 35) { "..." + $RootPath.Substring($RootPath.Length - 32) } else { $RootPath }
            $txtDriveStatus.Text = $shortPath
            $txtStatStorageLabel.Text = "in backup location"
        } else {
            $ellDriveStatus.Fill = [System.Windows.Media.SolidColorBrush]([System.Windows.Media.Color]::FromRgb(0xE8,0x92,0x4A))
            $txtDriveStatus.Text = "No location set"
            $txtStatStorageLabel.Text = "location not set"
        }
    }

    # ---------------------------
    # Helper: Prompt user to pick a backup root folder
    # ---------------------------
    function Select-BackupRoot {
        param([string]$Title = "Choose your backup root folder")
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = $Title
        $dlg.ShowNewFolderButton = $true
        # Suggest detected network drive or Desktop as starting point
        if ($global:DetectedNetworkDrive) {
            $dlg.SelectedPath = "$($global:DetectedNetworkDrive):\"
        } else {
            $dlg.SelectedPath = [Environment]::GetFolderPath("Desktop")
        }
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $chosen = Join-Path $dlg.SelectedPath "_PCBackups"
            Save-BackupRootConfig $chosen
            Update-LocationIndicator $chosen
            $tbBackupDest.Text = Join-Path $chosen $env:COMPUTERNAME
            $txtBackupDestLabel.Text = "Backing up to: $chosen"
            & $RefreshHistoryData
            return $true
        }
        return $false
    }

    # ---------------------------
    # Initialize backup location indicator + destination field
    # ---------------------------
    $btnChangeLocation = $Window.FindName("btnChangeLocation")

    if ($global:ConfigLoaded -and $global:BackupRootPath) {
        # Saved config found — use it
        Update-LocationIndicator $global:BackupRootPath
        $tbBackupDest.Text = Join-Path $global:BackupRootPath $env:COMPUTERNAME
        $txtBackupDestLabel.Text = "Backing up to: $($global:BackupRootPath). Browse to change."
    } elseif ($global:DetectedNetworkDrive) {
        # Auto-suggest detected network drive
        $suggested = "$($global:DetectedNetworkDrive):\_PCBackups"
        $ellDriveStatus.Fill = [System.Windows.Media.SolidColorBrush]([System.Windows.Media.Color]::FromRgb(0x5C,0x6C,0xF5))
        $txtDriveStatus.Text = "$($global:DetectedNetworkDrive):\ detected"
        $tbBackupDest.Text   = Join-Path $suggested $env:COMPUTERNAME
        $txtBackupDestLabel.Text = "$($global:DetectedNetworkDrive):\ drive detected. Using it as your default backup location. Browse to change, or click 'Change Location' on the dashboard."
        Save-BackupRootConfig $suggested
        Update-LocationIndicator $suggested
    } else {
        # No drive and no config — prompt on first run
        $ellDriveStatus.Fill = [System.Windows.Media.SolidColorBrush]([System.Windows.Media.Color]::FromRgb(0xE8,0x92,0x4A))
        $txtDriveStatus.Text = "No location set"
        $tbBackupDest.Text   = ""
        $txtBackupDestLabel.Text = "No backup location set. Click 'Browse' below or 'Change Location' on the Dashboard to choose where backups are saved (e.g. a USB drive, NAS, or network share)."
    }

    $btnChangeLocation.Add_Click({ Select-BackupRoot "Choose your backup root folder" | Out-Null })

    $btnBrowseBackupDest.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $dlg.Description = "Select where to save this backup"
        $dlg.ShowNewFolderButton = $true
        if ($global:BackupRootPath -and (Test-Path $global:BackupRootPath)) {
            $dlg.SelectedPath = $global:BackupRootPath
        } elseif ($global:DetectedNetworkDrive) {
            $dlg.SelectedPath = "$($global:DetectedNetworkDrive):\"
        }
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $newRoot = Join-Path $dlg.SelectedPath "_PCBackups"
            Save-BackupRootConfig $newRoot
            Update-LocationIndicator $newRoot
            $tbBackupDest.Text = Join-Path $newRoot $env:COMPUTERNAME
            $txtBackupDestLabel.Text = "Backing up to: $newRoot"
            & $RefreshHistoryData
        }
    })

    # ---------------------------
    # Populate sidebar user/machine info
    # ---------------------------
    $txtUserInitials   = $Window.FindName("txtUserInitials")
    $txtSidebarUser    = $Window.FindName("txtSidebarUser")
    $txtSidebarMachine = $Window.FindName("txtSidebarMachine")

    $sidebarUsername = $env:USERNAME
    $sidebarMachine  = $env:COMPUTERNAME

    $nameParts = $sidebarUsername -split '[._\-\s]' | Where-Object { $_ }
    $initials  = ($nameParts | Select-Object -First 2 | ForEach-Object { $_[0].ToString().ToUpper() }) -join ""

    $txtUserInitials.Text   = $initials
    $txtSidebarUser.Text    = $sidebarUsername.ToLower()
    $txtSidebarMachine.Text = $sidebarMachine

    # ---------------------------
    # Populate dashboard stat cards on load
    # ---------------------------
    $txtStatLastBackup  = $Window.FindName("txtStatLastBackup")
    $txtStatLastMachine = $Window.FindName("txtStatLastMachine")
    $txtStatTotalBackups = $Window.FindName("txtStatTotalBackups")
    $txtStatMachineCount = $Window.FindName("txtStatMachineCount")
    $txtStatStorage     = $Window.FindName("txtStatStorage")

    # ---------------------------
    # Navigation Functions
    # ---------------------------
    function Switch-View {
        param([string]$ViewName, [System.Windows.Controls.Button]$NavButton)
        
        $viewHome.Visibility = "Collapsed"
        $viewBackup.Visibility = "Collapsed"
        $viewRestore.Visibility = "Collapsed"
        $viewHistory.Visibility = "Collapsed"
        $viewAbout.Visibility = "Collapsed"
        
        $navHome.Background = "Transparent"
        $navBackup.Background = "Transparent"
        $navRestore.Background = "Transparent"
        $navHistory.Background = "Transparent"
        $navAbout.Background = "Transparent"
        
        switch ($ViewName) {
            "viewHome" { $viewHome.Visibility = "Visible" }
            "viewBackup" { $viewBackup.Visibility = "Visible" }
            "viewRestore" { $viewRestore.Visibility = "Visible" }
            "viewHistory" { $viewHistory.Visibility = "Visible" }
            "viewAbout" { $viewAbout.Visibility = "Visible" }
        }
        
        if ($NavButton) {
            $NavButton.Background = "#FF1A1A26"
        }
    }

    $navHome.Add_Click({ Switch-View "viewHome" $navHome })
    $navBackup.Add_Click({ Switch-View "viewBackup" $navBackup })
    $navRestore.Add_Click({ Switch-View "viewRestore" $navRestore })
    $navHistory.Add_Click({ 
        Switch-View "viewHistory" $navHistory
        & $RefreshHistoryData
    })
    $navAbout.Add_Click({ Switch-View "viewAbout" $navAbout })

    # About close button
    $btnAboutClose = $Window.FindName("btnAboutClose")
    if ($btnAboutClose) {
        $btnAboutClose.Add_Click({ Switch-View "viewHome" $navHome })
    }

    # Dashboard tile navigation
    $btnHomeBackup.Add_Click({ Switch-View "viewBackup" $navBackup })
    $btnHomeRestore.Add_Click({ Switch-View "viewRestore" $navRestore })
    $btnViewAll.Add_Click({
        Switch-View "viewHistory" $navHistory
        & $RefreshHistoryData
    })

    # Start on dashboard
    Switch-View "viewHome" $navHome

    # ---------------------------
    # History Refresh Function
    # ---------------------------

    # Dot color palette
    $dotColors = @("#5C6CF5", "#2ECC71", "#E8924A", "#E74C3C", "#9B59B6", "#1ABC9C")

    function Parse-BackupTimestamp {
        param([string]$ts)
        $formats = @(
            "yyyy-MM-dd_HH-mm-ss",
            "yyyy-MM-dd_hh-mm-tt",
            "yyyy-MM-dd_hh-mm-sstt",
            "yyyy-MM-dd HH:mm:ss"
        )
        $normalized = $ts -replace '_(\d{2})-(\d{2})-(AM|PM)$', ' $1:$2 $3'
        $formats2 = @("yyyy-MM-dd HH:mm tt", "yyyy-MM-dd h:mm tt")
        foreach ($fmt in ($formats + $formats2)) {
            try { return [datetime]::ParseExact($ts, $fmt, [System.Globalization.CultureInfo]::InvariantCulture) } catch {}
            try { return [datetime]::ParseExact($normalized, $fmt, [System.Globalization.CultureInfo]::InvariantCulture) } catch {}
        }
        try { return [datetime]::Parse($ts) } catch {}
        return $null
    }

    function Get-RelativeTime {
        param([datetime]$dt)
        $diff = (Get-Date) - $dt
        if ($diff.TotalMinutes -lt 2)  { return "Just now" }
        if ($diff.TotalMinutes -lt 60) { return "$([int]$diff.TotalMinutes) minutes ago" }
        if ($diff.TotalHours   -lt 24) { return "$([int]$diff.TotalHours) hours ago" }
        if ($diff.TotalDays    -lt 2)  { return "Yesterday" }
        if ($diff.TotalDays    -lt 7)  { return "$([int]$diff.TotalDays) days ago" }
        if ($diff.TotalDays    -lt 30) { return "$([int]($diff.TotalDays/7)) weeks ago" }
        return $dt.ToString("MMM d, yyyy")
    }

    $RefreshHistoryData = {
        $backups = @(Get-AllBackups)
        $txtHistoryCount.Text = "$($backups.Count) backup(s) found"

        $pcColorMap = @{}
        $colorIdx = 0
        foreach ($b in $backups) {
            if (-not $pcColorMap.ContainsKey($b.PCName)) {
                $pcColorMap[$b.PCName] = $dotColors[$colorIdx % $dotColors.Count]
                $colorIdx++
            }
        }

        $historyItems = $backups | ForEach-Object {
            $wpfColor = if ($pcColorMap.ContainsKey($_.PCName)) { $pcColorMap[$_.PCName] } else { "#5C6CF5" }

            $tags = @()
            $c = $_.Contents
            if ($c.Desktop)          { $tags += "Desktop" }
            if ($c.Documents)        { $tags += "Documents" }
            if ($c.Pictures)         { $tags += "Pictures" }
            if ($c.Downloads)        { $tags += "Downloads" }
            if ($c.Music)            { $tags += "Music" }
            if ($c.Videos)           { $tags += "Videos" }
            if ($c.ChromeBookmarks)  { $tags += "Chrome" }
            if ($c.EdgeBookmarks)    { $tags += "Edge" }
            if ($c.FirefoxBookmarks) { $tags += "Firefox" }
            if ($c.Printers)         { $tags += "Printers" }
            if ($c.Wallpaper)        { $tags += "Wallpaper" }
            if ($c.FileAssociations) { $tags += "File Assoc" }
            if ($tags.Count -eq 0)   { $tags += "Empty" }

            [PSCustomObject]@{
                PCName        = $_.PCName
                Timestamp     = $_.Timestamp
                SizeFormatted = $_.SizeFormatted
                DotColor      = $wpfColor
                ContentTags   = [string[]]$tags
                FullPath      = $_.FullPath
                Contents      = $_.Contents
                Note          = $_.Note
            }
        }

        $lvBackupHistory.ItemsSource = [object[]]@($historyItems)

        $recentItems = @($historyItems | Select-Object -First 5) | ForEach-Object {
            $relTime = $_.Timestamp 
            $formats = @("yyyy-MM-dd_HH-mm-ss","yyyy-MM-dd_hh-mm-tt","yyyy-MM-dd_hh-mm-ss-tt","yyyy-MM-dd_HH-mm")
            foreach ($fmt in $formats) {
                try {
                    $parsed = [datetime]::ParseExact($_.Timestamp, $fmt, [System.Globalization.CultureInfo]::InvariantCulture)
                    $relTime = (Get-RelativeTime $parsed) + " — " + $parsed.ToString("MMM d, h:mm tt")
                    break
                } catch {}
            }
            if ($relTime -eq $_.Timestamp) {
                try {
                    $cleaned = $_.Timestamp -replace '-([AP]M)$', ' $1' -replace '_', ' '
                    $parsed = [datetime]::Parse($cleaned, [System.Globalization.CultureInfo]::InvariantCulture)
                    $relTime = (Get-RelativeTime $parsed) + " — " + $parsed.ToString("MMM d, h:mm tt")
                } catch {}
            }

            [PSCustomObject]@{
                PCName        = $_.PCName
                DateDisplay   = $relTime
                SizeFormatted = $_.SizeFormatted
                DotColor      = $_.DotColor
                FullPath      = $_.FullPath
                Contents      = $_.Contents
                ContentTags   = $_.ContentTags
                Timestamp     = $_.Timestamp
                Note          = $_.Note
            }
        }

        $icRecentBackups.ItemsSource = [object[]]@($recentItems)

        $machineCount = ($backups | Select-Object -ExpandProperty PCName -Unique).Count

        if ($backups.Count -gt 0) {
            $latest = $backups[0]
            $relTime = $latest.Timestamp
            $formats = @("yyyy-MM-dd_HH-mm-ss","yyyy-MM-dd_hh-mm-tt","yyyy-MM-dd_hh-mm-ss-tt","yyyy-MM-dd_HH-mm")
            foreach ($fmt in $formats) {
                try {
                    $parsed = [datetime]::ParseExact($latest.Timestamp, $fmt, [System.Globalization.CultureInfo]::InvariantCulture)
                    $relTime = Get-RelativeTime $parsed
                    break
                } catch {}
            }
            if ($relTime -eq $latest.Timestamp) {
                try {
                    $cleaned = $latest.Timestamp -replace '-([AP]M)$', ' $1' -replace '_', ' '
                    $parsed = [datetime]::Parse($cleaned, [System.Globalization.CultureInfo]::InvariantCulture)
                    $relTime = Get-RelativeTime $parsed
                } catch {}
            }

            $txtStatLastBackup.Text  = $relTime
            $txtStatLastMachine.Text = $latest.PCName

            $totalBytes = ($backups | Measure-Object -Property Size -Sum).Sum
            $txtStatStorage.Text = Format-FileSize $totalBytes
        } else {
            $txtStatLastBackup.Text  = "No backups yet"
            $txtStatLastMachine.Text = ""
            $txtStatStorage.Text     = "0 GB"
        }
        $txtStatTotalBackups.Text = "$($backups.Count)"
        $txtStatMachineCount.Text = if ($machineCount -eq 1) { "Across 1 machine" } elseif ($machineCount -gt 1) { "Across $machineCount machines" } else { "" }
    }

    $btnRefreshHistory.Add_Click($RefreshHistoryData)
    & $RefreshHistoryData

    # ---------------------------
    # Wire "Restore →" buttons
    # ---------------------------
    $lvBackupHistory.AddHandler(
        [System.Windows.Controls.Button]::ClickEvent,
        [System.Windows.RoutedEventHandler]{
            param($sender, $e)
            $btn = $e.OriginalSource
            if ($btn -is [System.Windows.Controls.Button] -and $btn.Content -eq "Restore →") {
                $item = $btn.Tag
                if ($item -and $item.FullPath) {
                    $tbRestoreSource.Text = $item.FullPath
                    $statsPanel.Visibility = "Visible"
                    $txtBackupInfo.Text = "PC: $($item.PCName)`nDate: $($item.Timestamp)`nSize: $($item.SizeFormatted)`nNote: $($item.Note)"
                    $cbRDesktop.IsChecked        = $item.Contents.Desktop
                    $cbRDocuments.IsChecked      = $item.Contents.Documents
                    $cbRPictures.IsChecked       = $item.Contents.Pictures
                    $cbRDownloads.IsChecked      = $item.Contents.Downloads
                    $cbRMusic.IsChecked          = $item.Contents.Music
                    $cbRVideos.IsChecked         = $item.Contents.Videos
                    $cbRChromeBookmarks.IsChecked = $item.Contents.ChromeBookmarks
                    $cbREdgeBookmarks.IsChecked  = $item.Contents.EdgeBookmarks
                    $cbRFirefoxBookmarks.IsChecked = $item.Contents.FirefoxBookmarks
                    $cbRLegacyFavorites.IsChecked = $item.Contents.LegacyFavorites
                    $cbRPrinters.IsChecked       = $item.Contents.Printers
                    $cbRFileAssociations.IsChecked = $item.Contents.FileAssociations
                    Switch-View "viewRestore" $navRestore
                }
            }
        }
    )

    # ---------------------------
    # Wire click on recent backup cards
    # ---------------------------
    $icRecentBackups.AddHandler(
        [System.Windows.UIElement]::MouseLeftButtonUpEvent,
        [System.Windows.Input.MouseButtonEventHandler]{
            param($sender, $e)
            $el = $e.OriginalSource
            while ($el -and -not ($el -is [System.Windows.FrameworkElement] -and $el.DataContext -and $el.DataContext.FullPath)) {
                $el = [System.Windows.Media.VisualTreeHelper]::GetParent($el)
            }
            if ($el -and $el.DataContext -and $el.DataContext.FullPath) {
                $item = $el.DataContext
                $tbRestoreSource.Text = $item.FullPath
                $statsPanel.Visibility = "Visible"
                $txtBackupInfo.Text = "PC: $($item.PCName)`nDate: $($item.Timestamp)`nSize: $($item.SizeFormatted)`nNote: $($item.Note)"
                $cbRDesktop.IsChecked         = $item.Contents.Desktop
                $cbRDocuments.IsChecked       = $item.Contents.Documents
                $cbRPictures.IsChecked        = $item.Contents.Pictures
                $cbRDownloads.IsChecked       = $item.Contents.Downloads
                $cbRMusic.IsChecked           = $item.Contents.Music
                $cbRVideos.IsChecked          = $item.Contents.Videos
                $cbRChromeBookmarks.IsChecked = $item.Contents.ChromeBookmarks
                $cbREdgeBookmarks.IsChecked   = $item.Contents.EdgeBookmarks
                $cbRFirefoxBookmarks.IsChecked = $item.Contents.FirefoxBookmarks
                $cbRLegacyFavorites.IsChecked = $item.Contents.LegacyFavorites
                $cbRPrinters.IsChecked        = $item.Contents.Printers
                $cbRFileAssociations.IsChecked = $item.Contents.FileAssociations
                Switch-View "viewRestore" $navRestore
            }
        },
        $true
    )

    # ---------------------------
    # UI update helpers (thread-safe)
    # ---------------------------
    $ui = [hashtable]::Synchronized(@{})

    $ui.SetStatus = {
        param([string]$text)
        try {
            if ($Window.Dispatcher.HasShutdownStarted -or $Window.Dispatcher.HasShutdownFinished) { return }
            $Window.Dispatcher.Invoke([action]{ $txtStatus.Text = $text })
        } catch {}
    }

    $ui.SetProgress = {
        param([int]$pct)
        try {
            if ($pct -lt 0) { $pct = 0 }
            if ($pct -gt 100) { $pct = 100 }
            if ($Window.Dispatcher.HasShutdownStarted -or $Window.Dispatcher.HasShutdownFinished) { return }
            $Window.Dispatcher.Invoke([action]{ $pbProgress.Value = $pct })
        } catch {}
    }

    $ui.AppendActivity = {
        param([string]$line)
        if (-not $line) { return }
        try {
            if ($Window.Dispatcher.HasShutdownStarted -or $Window.Dispatcher.HasShutdownFinished) { return }
            $Window.Dispatcher.Invoke([action]{
                $txtActivity.AppendText($line + [Environment]::NewLine)
                $txtActivity.ScrollToEnd()
            })
        } catch {}
    }

    $ui.RestoreSetStatus = {
        param([string]$text)
        try {
            if ($Window.Dispatcher.HasShutdownStarted -or $Window.Dispatcher.HasShutdownFinished) { return }
            $Window.Dispatcher.Invoke([action]{ $txtRestoreStatus.Text = $text })
        } catch {}
    }

    $ui.RestoreSetProgress = {
        param([int]$pct)
        try {
            if ($pct -lt 0) { $pct = 0 }
            if ($pct -gt 100) { $pct = 100 }
            if ($Window.Dispatcher.HasShutdownStarted -or $Window.Dispatcher.HasShutdownFinished) { return }
            $Window.Dispatcher.Invoke([action]{ $pbRestoreProgress.Value = $pct })
        } catch {}
    }

    $ui.RestoreAppend = {
        param([string]$line)
        if (-not $line) { return }
        try {
            if ($Window.Dispatcher.HasShutdownStarted -or $Window.Dispatcher.HasShutdownFinished) { return }
            $Window.Dispatcher.Invoke([action]{
                $txtRestoreActivity.AppendText($line + [Environment]::NewLine)
                $txtRestoreActivity.ScrollToEnd()
            })
        } catch {}
    }


    # ---------------------------
    # Backup Select All / Select None
    # ---------------------------
    $btnBackupSelectAll.Add_Click({
        $cbDesktop.IsChecked = $true
        $cbDocuments.IsChecked = $true
        $cbPictures.IsChecked = $true
        $cbDownloads.IsChecked = $true
        $cbMusic.IsChecked = $true
        $cbVideos.IsChecked = $true
        $cbChromeBookmarks.IsChecked = $true
        $cbEdgeFavorites.IsChecked = $true
        $cbFirefoxBookmarks.IsChecked = $true
        $cbPrinters.IsChecked = $true
        $cbWallpaper.IsChecked = $true
        $cbFileAssociations.IsChecked = $true
        $cbPasswordHelpers.IsChecked = $true
    })

    $btnBackupSelectNone.Add_Click({
        $cbDesktop.IsChecked = $false
        $cbDocuments.IsChecked = $false
        $cbPictures.IsChecked = $false
        $cbDownloads.IsChecked = $false
        $cbMusic.IsChecked = $false
        $cbVideos.IsChecked = $false
        $cbChromeBookmarks.IsChecked = $false
        $cbEdgeFavorites.IsChecked = $false
        $cbFirefoxBookmarks.IsChecked = $false
        $cbPrinters.IsChecked = $false
        $cbWallpaper.IsChecked = $false
        $cbFileAssociations.IsChecked = $false
        $cbPasswordHelpers.IsChecked = $false
    })

    # ---------------------------
    # Restore Select All / Select None
    # ---------------------------
    $btnRestoreSelectAll.Add_Click({
        $cbRDesktop.IsChecked = $true
        $cbRDocuments.IsChecked = $true
        $cbRPictures.IsChecked = $true
        $cbRDownloads.IsChecked = $true
        $cbRMusic.IsChecked = $true
        $cbRVideos.IsChecked = $true
        $cbRChromeBookmarks.IsChecked = $true
        $cbREdgeBookmarks.IsChecked = $true
        $cbRFirefoxBookmarks.IsChecked = $true
        $cbRLegacyFavorites.IsChecked = $true
        $cbRPrinters.IsChecked = $true
        $cbRWallpaper.IsChecked = $true
        $cbRFileAssociations.IsChecked = $true
    })

    $btnRestoreSelectNone.Add_Click({
        $cbRDesktop.IsChecked = $false
        $cbRDocuments.IsChecked = $false
        $cbRPictures.IsChecked = $false
        $cbRDownloads.IsChecked = $false
        $cbRMusic.IsChecked = $false
        $cbRVideos.IsChecked = $false
        $cbRChromeBookmarks.IsChecked = $false
        $cbREdgeBookmarks.IsChecked = $false
        $cbRFirefoxBookmarks.IsChecked = $false
        $cbRLegacyFavorites.IsChecked = $false
        $cbRPrinters.IsChecked = $false
        $cbRWallpaper.IsChecked = $false
        $cbRFileAssociations.IsChecked = $false
    })

    # ---------------------------
    # Cancel Backup
    # ---------------------------
    $btnCancel.Add_Click({
        $syncBackup.Cancel = $true
        foreach ($pid in @($syncBackup.RoboPids)) {
            try {
                Write-BackupToolLog "Killing backup robocopy PID: $pid"
                Stop-Process -Id $pid -Force -ErrorAction SilentlyContinue
            } catch {}
        }
    })

    # ---------------------------
    # Start Backup
    # ---------------------------
    $btnStart.Add_Click({
        $txtActivity.Clear()
        $btnStart.IsEnabled = $false
        $btnCancel.IsEnabled = $true

        $syncBackup.Cancel = $false
        $syncBackup.RoboPids.Clear()

        $BackupRoot = $tbBackupDest.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($BackupRoot)) {
            [System.Windows.MessageBox]::Show("Please choose a backup destination before starting.","No Destination","OK","Warning") | Out-Null
            $btnStart.IsEnabled = $true
            $btnCancel.IsEnabled = $true
            return
        }
        $Computer = $env:COMPUTERNAME
        $Date = Get-Date -Format "yyyy-MM-dd_hh-mm-tt"
        if ($BackupRoot -like "*\$Computer") {
            $BackupPath = Join-Path $BackupRoot $Date
        } else {
            $BackupPath = Join-Path $BackupRoot "$Computer\$Date"
        }

        $cfg = [pscustomobject]@{
            BackupPath = $BackupPath
            UserProfile = $env:USERPROFILE
            LogFile = $(if ($global:EnableBackupToolLogging) { Join-Path $BackupPath "BackupLog.txt" } else { $null })
            DoDesktop = [bool]$cbDesktop.IsChecked
            DoDocuments = [bool]$cbDocuments.IsChecked
            DoPictures = [bool]$cbPictures.IsChecked
            DoDownloads = [bool]$cbDownloads.IsChecked
            DoMusic = [bool]$cbMusic.IsChecked
            DoVideos = [bool]$cbVideos.IsChecked
            DoChromeBookmarks = [bool]$cbChromeBookmarks.IsChecked
            DoEdgeBookmarks = [bool]$cbEdgeFavorites.IsChecked
            DoFirefoxBookmarks = [bool]$cbFirefoxBookmarks.IsChecked
            DoPrinters = [bool]$cbPrinters.IsChecked
            DoWallpaper = [bool]$cbWallpaper.IsChecked
            DoFileAssociations = ($cbFileAssociations.IsChecked -eq $true)
            DoPasswordHelpers = [bool]$cbPasswordHelpers.IsChecked
            OpenHelperPages = [bool]$cbOpenHelperPages.IsChecked
            OpenBackupFolder = [bool]$cbOpenBackupFolder.IsChecked
            BackupNote = $tbBackupNote.Text
            ChromeExe = Get-ChromeExe
            EdgeExe = Get-EdgeExe
            FirefoxExe = Get-FirefoxExe
            FirefoxProfilesRoot = Get-FirefoxProfilesRoot
            LocalAppData = $env:LOCALAPPDATA
        }

        $anySelected = $cfg.DoDesktop -or $cfg.DoDocuments -or $cfg.DoPictures -or $cfg.DoDownloads -or 
                       $cfg.DoMusic -or $cfg.DoVideos -or $cfg.DoChromeBookmarks -or $cfg.DoEdgeBookmarks -or
                       $cfg.DoFirefoxBookmarks -or $cfg.DoPrinters -or $cfg.DoPasswordHelpers -or $cfg.DoWallpaper -or
                       $cfg.DoFileAssociations

        if (-not $anySelected) {
            [System.Windows.MessageBox]::Show("Nothing selected to back up.","No Selection","OK","Information") | Out-Null
            $btnStart.IsEnabled = $true
            $btnCancel.IsEnabled = $true
            return
        }

        New-Item -ItemType Directory -Path $cfg.BackupPath -Force | Out-Null

        & $ui.SetProgress 0
        & $ui.SetStatus "Starting..."
        & $ui.AppendActivity ("Backup destination: {0}" -f $cfg.BackupPath)

        Write-BackupToolLog "Backup Start clicked. BackupPath: $($cfg.BackupPath)"

        $worker = {
            param($cfg, $ui, $sync, $mainLogPath)

            function LogMain {
                param([string]$msg)
                try {
                    if (-not $mainLogPath) { return }
                    $line = ("[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $msg)
                    Add-Content -Path $mainLogPath -Value $line -Encoding UTF8
                } catch {}
            }

            function Copy-BookmarksAllProfiles {
                param([string]$UserDataRoot, [string]$DestRoot, [string]$BrowserName)
                if (-not (Test-Path $UserDataRoot)) {
                    LogMain "$BrowserName user data folder not found: $UserDataRoot"
                    $ui.AppendActivity.Invoke("$BrowserName user data folder not found.")
                    return
                }

                New-Item -ItemType Directory -Path $DestRoot -Force | Out-Null
                $profiles = Get-ChildItem -Path $UserDataRoot -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq "Default" -or $_.Name -like "Profile *" }

                if (-not $profiles) {
                    LogMain "$BrowserName profiles not found under: $UserDataRoot"
                    $ui.AppendActivity.Invoke("${BrowserName}: no profiles found.")
                }

                foreach ($p in $profiles) {
                    if ($sync.Cancel) { throw "Cancelled" }
                    $srcBookmarks = Join-Path $p.FullName "Bookmarks"
                    $srcBak = Join-Path $p.FullName "Bookmarks.bak"
                    $destProfile = Join-Path $DestRoot $p.Name

                    if ((Test-Path $srcBookmarks) -or (Test-Path $srcBak)) {
                        LogMain "Backing up $BrowserName bookmarks for profile: $($p.Name)"
                        $ui.AppendActivity.Invoke("${BrowserName}: $($p.Name) bookmarks...")
                        New-Item -ItemType Directory -Path $destProfile -Force | Out-Null
                        if (Test-Path $srcBookmarks) { Copy-Item $srcBookmarks -Destination $destProfile -Force -ErrorAction SilentlyContinue }
                        if (Test-Path $srcBak) { Copy-Item $srcBak -Destination $destProfile -Force -ErrorAction SilentlyContinue }
                    } else {
                        LogMain "No bookmark files for ${BrowserName}: profile: $($p.Name)"
                    }
                }
            }

            function New-PasswordHelperHtml {
                param([string]$Path, [string]$TargetUrl, [string]$Title)
                $html = @"
<!doctype html>
<html><head><meta charset="utf-8"><title>$Title</title><style>body { font-family: Segoe UI, Arial; margin: 24px; } code { background:#f3f3f3; padding:8px 10px; border-radius:6px; display:inline-block; } button { padding:8px 12px; border-radius:6px; border:1px solid #bbb; cursor:pointer; margin-top:10px; } .ok { color:#0a6; margin-left:10px; font-weight:600; } .hint { opacity:.8; margin-top:12px; }</style></head><body><h2>$Title</h2><p><b>Important:</b> Browsers block opening internal <code>chrome://</code> and <code>edge://</code> URLs from webpages.</p><p>Copy this URL and paste it into the browser address bar:</p><div><code id="u">$TargetUrl</code></div><div><button onclick="copyUrl()">Copy to Clipboard</button><span id="msg" class="ok"></span></div><p class="hint">If Copy to Clipboard is blocked, highlight the URL and press <b>Ctrl+C</b>, then paste into the address bar.</p><script>function copyUrl(){const t=document.getElementById('u').textContent;navigator.clipboard.writeText(t).then(()=>{document.getElementById('msg').textContent="Copied!";}).catch(()=>{document.getElementById('msg').textContent="Copy blocked. Please Ctrl+C the URL.";});}</script></body></html>
"@
                $folder = Split-Path $Path -Parent
                if (-not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
                Set-Content -Path $Path -Value $html -Encoding UTF8
            }

            function Invoke-RoboCopyCancellable {
                param([string]$Source, [string]$Destination, [string]$LogFile = $null)
                LogMain "Robocopy: '$Source' -> '$Destination'"

                if (-not (Test-Path $Source)) {
                    LogMain "Source not found: $Source"
                    $ui.AppendActivity.Invoke("Source not found: $Source")
                    return
                }

                New-Item -ItemType Directory -Path $Destination -Force | Out-Null

                $args = @("`"$Source`"", "`"$Destination`"", "/E","/Z","/R:2","/W:2","/XJ","/COPY:DAT","/DCOPY:T", "/NFL","/NDL","/NP","/NJH","/NJS")
                if ($LogFile) { $args += "/LOG+:`"$LogFile`"" }

                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName = "robocopy.exe"
                $psi.Arguments = ($args -join " ")
                $psi.UseShellExecute = $false
                $psi.RedirectStandardOutput = $false
                $psi.RedirectStandardError = $false
                $psi.CreateNoWindow = $true

                $p = New-Object System.Diagnostics.Process
                $p.StartInfo = $psi
                
                try {
                    [void]$p.Start()
                    [void]$sync.RoboPids.Add($p.Id)

                    while (-not $p.HasExited) {
                        if ($sync.Cancel) {
                            LogMain "Cancel requested: killing robocopy PID $($p.Id)"
                            try { $p.Kill() } catch {}
                            try { $p.WaitForExit(3000) | Out-Null } catch {}
                            throw "Cancelled"
                        }
                        Start-Sleep -Milliseconds 200
                    }

                    LogMain "Robocopy exit code: $($p.ExitCode)"
                    if ($p.ExitCode -ge 8) {
                        if ($LogFile) { throw "Robocopy reported failure (ExitCode $($p.ExitCode)). See $LogFile" } 
                        else { throw "Robocopy reported failure (ExitCode $($p.ExitCode))." }
                    }
                }
                finally {
                    try { [void]$sync.RoboPids.Remove($p.Id) } catch {}
                    try { $p.Dispose() } catch {}
                }
            }

            try {
                $tasks = New-Object System.Collections.Generic.List[string]
                if ($cfg.DoDesktop) { $tasks.Add("Desktop") }
                if ($cfg.DoDocuments) { $tasks.Add("Documents") }
                if ($cfg.DoPictures) { $tasks.Add("Pictures") }
                if ($cfg.DoDownloads) { $tasks.Add("Downloads") }
                if ($cfg.DoMusic) { $tasks.Add("Music") }
                if ($cfg.DoVideos) { $tasks.Add("Videos") }
                if ($cfg.DoChromeBookmarks) { $tasks.Add("Chrome Bookmarks") }
                if ($cfg.DoEdgeBookmarks) { $tasks.Add("Edge Bookmarks") }
                if ($cfg.DoFirefoxBookmarks) { $tasks.Add("Firefox Bookmarks") }
                if ($cfg.DoPrinters) { $tasks.Add("Network Printers") }
                if ($cfg.DoPasswordHelpers) { $tasks.Add("Password Helpers") }
                if ($cfg.DoWallpaper) { $tasks.Add("Wallpaper") }
                if ($cfg.DoFileAssociations) { $tasks.Add("File Associations") }

                $total = $tasks.Count
                $done = 0

                function Do-Folder {
                    param([string]$FolderName)
                    $src = Join-Path $cfg.UserProfile $FolderName
                    $dst = Join-Path $cfg.BackupPath $FolderName
                    $ui.SetStatus.Invoke("Backing up $FolderName...")
                    $ui.AppendActivity.Invoke("Copying $FolderName...")
                    Invoke-RoboCopyCancellable -Source $src -Destination $dst -LogFile $cfg.LogFile
                }

                if ($cfg.DoDesktop) { if ($sync.Cancel) { throw "Cancelled" }; $done++; $ui.SetProgress.Invoke([int][math]::Round(($done/$total)*100)); Do-Folder "Desktop" }
                if ($cfg.DoDocuments) { if ($sync.Cancel) { throw "Cancelled" }; $done++; $ui.SetProgress.Invoke([int][math]::Round(($done/$total)*100)); Do-Folder "Documents" }
                if ($cfg.DoPictures) { if ($sync.Cancel) { throw "Cancelled" }; $done++; $ui.SetProgress.Invoke([int][math]::Round(($done/$total)*100)); Do-Folder "Pictures" }
                if ($cfg.DoDownloads) { if ($sync.Cancel) { throw "Cancelled" }; $done++; $ui.SetProgress.Invoke([int][math]::Round(($done/$total)*100)); Do-Folder "Downloads" }
                if ($cfg.DoMusic) { if ($sync.Cancel) { throw "Cancelled" }; $done++; $ui.SetProgress.Invoke([int][math]::Round(($done/$total)*100)); Do-Folder "Music" }
                if ($cfg.DoVideos) { if ($sync.Cancel) { throw "Cancelled" }; $done++; $ui.SetProgress.Invoke([int][math]::Round(($done/$total)*100)); Do-Folder "Videos" }

                $browserRoot = Join-Path $cfg.BackupPath "Browser_Backup"
                New-Item -ItemType Directory -Path $browserRoot -Force | Out-Null

                if ($cfg.DoChromeBookmarks) {
                    if ($sync.Cancel) { throw "Cancelled" }
                    $done++; $ui.SetProgress.Invoke([int][math]::Round(($done/$total)*100))
                    $ui.SetStatus.Invoke("Backing up Chrome bookmarks...")
                    $chromeUserData = [IO.Path]::Combine($cfg.LocalAppData, "Google", "Chrome", "User Data")
                    Copy-BookmarksAllProfiles -UserDataRoot $chromeUserData -DestRoot (Join-Path $browserRoot "Chrome") -BrowserName "Chrome"
                }

                if ($cfg.DoEdgeBookmarks) {
                    if ($sync.Cancel) { throw "Cancelled" }
                    $done++; $ui.SetProgress.Invoke([int][math]::Round(($done/$total)*100))
                    $ui.SetStatus.Invoke("Backing up Edge bookmarks...")
                    $edgeUserData = [IO.Path]::Combine($cfg.LocalAppData, "Microsoft", "Edge", "User Data")
                    Copy-BookmarksAllProfiles -UserDataRoot $edgeUserData -DestRoot (Join-Path $browserRoot "Edge") -BrowserName "Edge"

                    $legacyFav = Join-Path $cfg.UserProfile "Favorites"
                    if (Test-Path $legacyFav) {
                        $ui.AppendActivity.Invoke("Copying legacy Favorites folder...")
                        Invoke-RoboCopyCancellable -Source $legacyFav -Destination (Join-Path $browserRoot "Legacy_Favorites_Folder") -LogFile $cfg.LogFile
                    }
                }

                if ($cfg.DoFirefoxBookmarks) {
                    if ($sync.Cancel) { throw "Cancelled" }
                    $done++; $ui.SetProgress.Invoke([int][math]::Round(($done/$total)*100))
                    $ui.SetStatus.Invoke("Backing up Firefox bookmarks...")

                    $ffProfilesRoot = $cfg.FirefoxProfilesRoot
                    if (Test-Path $ffProfilesRoot) {
                        $ffDestRoot = Join-Path $browserRoot "Firefox"
                        New-Item -ItemType Directory -Path $ffDestRoot -Force | Out-Null

                        $ffProfiles = Get-ChildItem -Path $ffProfilesRoot -Directory -ErrorAction SilentlyContinue
                        $ffBacked = 0
                        foreach ($ffp in $ffProfiles) {
                            if ($sync.Cancel) { throw "Cancelled" }
                            $placesFile = Join-Path $ffp.FullName "places.sqlite"
                            $faviconsFile = Join-Path $ffp.FullName "favicons.sqlite"
                            if (Test-Path $placesFile) {
                                $ffProfileDest = Join-Path $ffDestRoot $ffp.Name
                                New-Item -ItemType Directory -Path $ffProfileDest -Force | Out-Null
                                Copy-Item -Path $placesFile -Destination $ffProfileDest -Force -ErrorAction SilentlyContinue
                                if (Test-Path $faviconsFile) {
                                    Copy-Item -Path $faviconsFile -Destination $ffProfileDest -Force -ErrorAction SilentlyContinue
                                }
                                $ui.AppendActivity.Invoke("Firefox: backed up profile '$($ffp.Name)'")
                                $ffBacked++
                            }
                        }

                        # Also back up profiles.ini so restore knows which profile is default
                        $profilesIni = Join-Path ([IO.Path]::Combine($env:APPDATA, "Mozilla", "Firefox")) "profiles.ini"
                        if (Test-Path $profilesIni) {
                            Copy-Item -Path $profilesIni -Destination $ffDestRoot -Force -ErrorAction SilentlyContinue
                            $ui.AppendActivity.Invoke("Firefox: backed up profiles.ini")
                        }

                        if ($ffBacked -eq 0) {
                            $ui.AppendActivity.Invoke("Firefox: no profiles with places.sqlite found under $ffProfilesRoot")
                        }
                    } else {
                        $ui.AppendActivity.Invoke("Firefox: profiles folder not found ($ffProfilesRoot). Is Firefox installed?")
                    }
                }

                if ($cfg.DoPrinters) {
                    if ($sync.Cancel) { throw "Cancelled" }
                    $done++; $ui.SetProgress.Invoke([int][math]::Round(($done/$total)*100))
                    $ui.SetStatus.Invoke("Backing up Network Printers...")
                    $ui.AppendActivity.Invoke("Exporting network printer registry keys...")
        
                    $printersKey = "HKCU:\Printers\Connections"
                    if (Test-Path $printersKey) {
                        $printers = Get-ChildItem -Path $printersKey -ErrorAction SilentlyContinue
                        if ($printers) {
                            foreach ($p in $printers) {
                                # Registry keys look like: ,,ServerName,PrinterName
                                # We replace commas with slashes to show the standard UNC path
                                $pName = $p.PSChildName -replace ',', '\'
                                $ui.AppendActivity.Invoke(" - Found: $pName")
                            }
                            $regFile = Join-Path $cfg.BackupPath "NetworkPrinters.reg"
                            $process = Start-Process -FilePath "reg.exe" -ArgumentList "export `"HKCU\Printers\Connections`" `"$regFile`" /y" -Wait -PassThru -WindowStyle Hidden
                            if ($process.ExitCode -eq 0) {
                                $ui.AppendActivity.Invoke("Successfully exported printers.")
                            } else {
                                $ui.AppendActivity.Invoke("Warning: Failed to export printers via reg.exe.")
                            }
                        } else {
                            $ui.AppendActivity.Invoke("No network printers found in registry.")
                        }
                    } else {
                        $ui.AppendActivity.Invoke("Printer connections key not found.")
                    }
                }

                if ($cfg.DoPasswordHelpers) {
                    if ($sync.Cancel) { throw "Cancelled" }
                    $done++; $ui.SetProgress.Invoke([int][math]::Round(($done/$total)*100))
                    $ui.SetStatus.Invoke("Creating password helper pages...")

                    $pwFolder = Join-Path $cfg.BackupPath "BrowserPasswordExports"
                    New-Item -ItemType Directory -Path $pwFolder -Force | Out-Null

                    $chromeHtml = Join-Path $pwFolder "Open_Chrome_Passwords.html"
                    $edgeHtml = Join-Path $pwFolder "Open_Edge_Passwords.html"

                    New-PasswordHelperHtml -Path $chromeHtml -TargetUrl "chrome://password-manager/passwords" -Title "Chrome Passwords"
                    New-PasswordHelperHtml -Path $edgeHtml -TargetUrl "edge://settings/autofill/passwords" -Title "Edge Passwords"

                    $ui.AppendActivity.Invoke("Created password helper pages.")

                    if ($cfg.OpenHelperPages) {
                        $ui.SetStatus.Invoke("Opening helper pages...")
                        if ($cfg.ChromeExe -and (Test-Path $cfg.ChromeExe)) { Start-Process -FilePath $cfg.ChromeExe -ArgumentList "`"$chromeHtml`"" | Out-Null }
                        else { Start-Process -FilePath $chromeHtml | Out-Null }

                        if ($cfg.EdgeExe -and (Test-Path $cfg.EdgeExe)) { Start-Process -FilePath $cfg.EdgeExe -ArgumentList "`"$edgeHtml`"" | Out-Null }
                        else { Start-Process -FilePath $edgeHtml | Out-Null }
                    }
                }

                if ($cfg.DoWallpaper) {
                    if ($sync.Cancel) { throw "Cancelled" }
                    $done++; $ui.SetProgress.Invoke([int][math]::Round(($done/$total)*100))
                    $ui.SetStatus.Invoke("Backing up wallpaper...")
                    LogMain "Starting wallpaper backup."

                    try {
                        $wpDest = Join-Path $cfg.BackupPath "Wallpaper"
                        New-Item -ItemType Directory -Path $wpDest -Force | Out-Null

                        # Read current wallpaper path from registry
                        $wpRegPath  = "HKCU:\Control Panel\Desktop"
                        $wpFilePath = (Get-ItemProperty -Path $wpRegPath -Name "WallPaper"    -ErrorAction SilentlyContinue).WallPaper
                        $wpStyle    = (Get-ItemProperty -Path $wpRegPath -Name "WallpaperStyle" -ErrorAction SilentlyContinue).WallpaperStyle
                        $wpTile     = (Get-ItemProperty -Path $wpRegPath -Name "TileWallpaper"  -ErrorAction SilentlyContinue).TileWallpaper

                        # Detect Windows Spotlight — its path lives under Packages\Microsoft.Windows.ContentDeliveryManager
                        $isSpotlight = $false
                        if ($wpFilePath -and $wpFilePath -like "*ContentDeliveryManager*") {
                            $isSpotlight = $true
                        }
                        # Also check the Personalization registry key for Spotlight
                        $slideShowSrc = (Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Lock Screen\Creative" -ErrorAction SilentlyContinue)
                        $bgType = (Get-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "WallpaperType" -ErrorAction SilentlyContinue).WallpaperType

                        $info = [ordered]@{
                            SourceType    = if ($isSpotlight) { "Spotlight" } elseif ([string]::IsNullOrEmpty($wpFilePath)) { "SolidColor" } else { "Image" }
                            OriginalPath  = $wpFilePath
                            WallpaperStyle = $wpStyle
                            TileWallpaper  = $wpTile
                            BackedUpAt     = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
                            BackedUpFile   = ""
                        }

                        if ($info.SourceType -eq "Image" -and (Test-Path $wpFilePath)) {
                            $ext       = [IO.Path]::GetExtension($wpFilePath)
                            $destFile  = Join-Path $wpDest ("wallpaper" + $ext)
                            Copy-Item -Path $wpFilePath -Destination $destFile -Force -ErrorAction Stop
                            $info.BackedUpFile = [IO.Path]::GetFileName($destFile)
                            $ui.AppendActivity.Invoke("Wallpaper: backed up image '$([IO.Path]::GetFileName($wpFilePath))' (style=$wpStyle).")
                            LogMain "Wallpaper image copied from '$wpFilePath' to '$destFile'."
                        } elseif ($info.SourceType -eq "Spotlight") {
                            # Try to copy the most recent Spotlight asset as a best-effort fallback
                            $spotlightAssets = [IO.Path]::Combine($env:LOCALAPPDATA,
                                "Packages", "Microsoft.Windows.ContentDeliveryManager_cw5n1h2txyewy",
                                "LocalState", "Assets")
                            if (Test-Path $spotlightAssets) {
                                $bestAsset = Get-ChildItem -Path $spotlightAssets -File -ErrorAction SilentlyContinue |
                                             Where-Object { $_.Length -gt 100KB } |
                                             Sort-Object Length -Descending |
                                             Select-Object -First 1
                                if ($bestAsset) {
                                    $destFile = Join-Path $wpDest "wallpaper_spotlight.jpg"
                                    Copy-Item -Path $bestAsset.FullName -Destination $destFile -Force -ErrorAction SilentlyContinue
                                    $info.BackedUpFile = "wallpaper_spotlight.jpg"
                                    $ui.AppendActivity.Invoke("Wallpaper: Windows Spotlight detected — saved current Spotlight image as fallback.")
                                    LogMain "Spotlight best asset copied from '$($bestAsset.FullName)'."
                                } else {
                                    $ui.AppendActivity.Invoke("Wallpaper: Windows Spotlight active — no Spotlight asset found to copy.")
                                    LogMain "Spotlight active but no large assets found."
                                }
                            } else {
                                $ui.AppendActivity.Invoke("Wallpaper: Windows Spotlight active — Spotlight assets folder not found.")
                                LogMain "Spotlight assets folder not found at: $spotlightAssets"
                            }
                        } else {
                            $ui.AppendActivity.Invoke("Wallpaper: no image file set (solid color or unsupported type). Settings saved.")
                            LogMain "Wallpaper type: $($info.SourceType) — no file to copy."
                        }

                        # Always save the metadata JSON so restore knows what was here
                        $jsonPath = Join-Path $wpDest "_wallpaper_info.json"
                        $info | ConvertTo-Json | Set-Content -Path $jsonPath -Encoding UTF8
                        LogMain "Wallpaper metadata saved to: $jsonPath"

                    } catch {
                        $ui.AppendActivity.Invoke("Wallpaper: WARNING - $($_.Exception.Message)")
                        LogMain "Wallpaper backup error: $($_.Exception.Message)"
                    }
                }

                if ($cfg.DoFileAssociations) {
                    if ($sync.Cancel) { throw "Cancelled" }
                    $done++; $ui.SetProgress.Invoke([int][math]::Round(($done/$total)*100))
                    $ui.SetStatus.Invoke("Backing up File Associations...")
                    LogMain "Starting File Associations backup."

                    try {
                        $faRegKey = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts"
                        $faRegFile = Join-Path $cfg.BackupPath "FileAssociations.reg"
                        $process = Start-Process -FilePath "reg.exe" `
                            -ArgumentList "export `"$faRegKey`" `"$faRegFile`" /y" `
                            -Wait -PassThru -WindowStyle Hidden
                        if ($process.ExitCode -eq 0) {
                            $ui.AppendActivity.Invoke("File Associations: exported successfully.")
                            LogMain "File Associations exported to: $faRegFile"
                        } else {
                            $ui.AppendActivity.Invoke("File Associations: WARNING — reg.exe exited with code $($process.ExitCode).")
                            LogMain "File Associations export failed with exit code: $($process.ExitCode)"
                        }
                    } catch {
                        $ui.AppendActivity.Invoke("File Associations: ERROR — $($_.Exception.Message)")
                        LogMain "File Associations backup error: $($_.Exception.Message)"
                    }
                }

                if ($cfg.BackupNote -and $cfg.BackupNote.Trim()) {
                    $noteFile = Join-Path $cfg.BackupPath "_backup_note.txt"
                    try {
                        Set-Content -Path $noteFile -Value $cfg.BackupNote -Encoding UTF8
                        LogMain "Saved backup note to: $noteFile"
                    } catch { LogMain "Failed to save note: $($_.Exception.Message)" }
                }

                $ui.SetProgress.Invoke(100)
                $ui.SetStatus.Invoke("Backup complete.")
                $ui.AppendActivity.Invoke("Backup completed successfully.")

                if ($cfg.OpenBackupFolder) { Start-Process -FilePath $cfg.BackupPath | Out-Null }
            }
            catch {
                if ($_.Exception -and $_.Exception.Message -eq "Cancelled") {
                    LogMain "Backup cancelled."; $ui.SetStatus.Invoke("Cancelled."); $ui.AppendActivity.Invoke("Cancelled.")
                } elseif ($sync.Cancel) {
                    LogMain "Backup cancelled (flag)."; $ui.SetStatus.Invoke("Cancelled."); $ui.AppendActivity.Invoke("Cancelled.")
                } else {
                    LogMain ("Backup failed: " + $_.Exception.Message)
                    $ui.SetStatus.Invoke("Backup failed. See log.")
                    $ui.AppendActivity.Invoke("ERROR: " + $_.Exception.Message)
                }
            }
        }

        $ps = [PowerShell]::Create()
        $rs = [RunspaceFactory]::CreateRunspace()
        $rs.ApartmentState = "STA"
        $rs.ThreadOptions = "ReuseThread"
        $rs.Open()
        $ps.Runspace = $rs

        $ps.AddScript($worker).AddArgument($cfg).AddArgument($ui).AddArgument($syncBackup).AddArgument($global:BackupToolLog) | Out-Null
        $handleBackup = $ps.BeginInvoke()

        $timerBackup = New-Object System.Windows.Threading.DispatcherTimer
        $timerBackup.Interval = [TimeSpan]::FromMilliseconds(250)

        $timerBackup.Add_Tick({
            $timer = $script:BackupTimer
            $handle = $script:BackupHandle
            $psInstance = $script:BackupPs
            $rsInstance = $script:BackupRs
            
            if ($handle -and $handle.IsCompleted) {
                try { if ($timer) { $timer.Stop() } } catch {}
                try {
                    if ($btnStart) { $btnStart.IsEnabled = $true }
                    if ($btnCancel) { $btnCancel.IsEnabled = $true }
                } catch {}

                try { if ($psInstance) { $psInstance.Dispose() } } catch {}
                try { if ($rsInstance) { $rsInstance.Close(); $rsInstance.Dispose() } } catch {}
                
                $script:BackupTimer  = $null
                $script:BackupPs     = $null
                $script:BackupRs     = $null
                $script:BackupHandle = $null
            }
        })

        $script:BackupTimer  = $timerBackup
        $script:BackupPs     = $ps
        $script:BackupRs     = $rs
        $script:BackupHandle = $handleBackup

        $timerBackup.Start()
    })

    # ---------------------------
    # Restore Functions
    # ---------------------------
    $btnFindLatestBackup.Add_Click({
        if (-not $global:BackupRootPath -or -not (Test-Path $global:BackupRootPath)) {
            $result = [System.Windows.MessageBox]::Show(
                "No backup location is configured yet.`n`nWould you like to choose your backup root folder now?",
                "No Location Set", "YesNo", "Question")
            if ($result -eq "Yes") { Select-BackupRoot | Out-Null }
            return
        }
        $backups = @(Get-AllBackups)
        if ($backups) {
            $latest = $backups[0]
            $tbRestoreSource.Text = $latest.FullPath
            $txtRestoreActivity.AppendText("Selected: $($latest.FullPath)`n")
            
            $cbRDesktop.IsChecked = $latest.Contents.Desktop
            $cbRDocuments.IsChecked = $latest.Contents.Documents
            $cbRPictures.IsChecked = $latest.Contents.Pictures
            $cbRPrinters.IsChecked = $latest.Contents.Printers
            $cbRWallpaper.IsChecked = $latest.Contents.Wallpaper
            $cbRFileAssociations.IsChecked = $latest.Contents.FileAssociations
            
            $statsPanel.Visibility = "Visible"
            $txtBackupInfo.Text = "PC: $($latest.PCName)`nDate: $($latest.Timestamp)`nSize: $($latest.SizeFormatted)`nNote: $($latest.Note)"
        } else {
            [System.Windows.MessageBox]::Show("No backups found.","No Backups","OK","Information")
        }
    })

    $btnBrowseBackup.Add_Click({
        $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
        if ($dlg.ShowDialog() -eq "OK") { $tbRestoreSource.Text = $dlg.SelectedPath }
    })

    $btnCancelRestore.Add_Click({
        $syncRestore.Cancel = $true
        foreach ($pid in @($syncRestore.RoboPids)) {
            try {
                Write-BackupToolLog "Killing restore robocopy PID: $pid"
                Stop-Process -Id $pid -Force -ErrorAction SilentlyContinue
            } catch {}
        }
    })

    $btnStartRestore.Add_Click({
        $src = $tbRestoreSource.Text
        if (-not (Test-Path $src)) {
            [System.Windows.MessageBox]::Show("Invalid backup path. Please select a valid backup folder.","Invalid Path","OK","Warning")
            return
        }

        $txtRestoreActivity.Clear()
        $btnStartRestore.IsEnabled = $false
        $btnCancelRestore.IsEnabled = $true
        $syncRestore.Cancel = $false
        $syncRestore.RoboPids.Clear()

        $cfgR = [pscustomobject]@{
            RestoreSource = $src
            UserProfile = $env:USERPROFILE
            LocalAppData = $env:LOCALAPPDATA
            LogFile = Join-Path $src "RestoreLog.txt"
            DoDesktop = $cbRDesktop.IsChecked
            DoDocuments = $cbRDocuments.IsChecked
            DoPictures = $cbRPictures.IsChecked
            DoDownloads = $cbRDownloads.IsChecked
            DoMusic = $cbRMusic.IsChecked
            DoVideos = $cbRVideos.IsChecked
            DoChromeBookmarks = $cbRChromeBookmarks.IsChecked
            DoEdgeBookmarks = $cbREdgeBookmarks.IsChecked
            DoFirefoxBookmarks = $cbRFirefoxBookmarks.IsChecked
            DoLegacyFavorites = $cbRLegacyFavorites.IsChecked
            DoPrinters = $cbRPrinters.IsChecked
            DoWallpaper = [bool]$cbRWallpaper.IsChecked
            DoFileAssociations = [bool]$cbRFileAssociations.IsChecked
        }

        $restoreWorker = {
            param($cfgR, $ui, $syncR, $mainLogPath)

            function LogMain {
                param([string]$msg)
                try {
                    if (-not $mainLogPath) { return }
                    $line = ("[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $msg)
                    Add-Content -Path $mainLogPath -Value $line -Encoding UTF8
                } catch {}
            }

            function Invoke-RoboCopyCancellable {
                param([string]$Source, [string]$Destination, [string]$LogFile = $null)
                LogMain "Robocopy: '$Source' -> '$Destination'"

                if (-not (Test-Path $Source)) {
                    LogMain "Source not found: $Source"
                    $ui.RestoreAppend.Invoke("Source not found: $Source")
                    return
                }

                New-Item -ItemType Directory -Path $Destination -Force | Out-Null

                $args = @("`"$Source`"", "`"$Destination`"", "/E","/XO","/R:0","/W:0","/NFL","/NDL","/NP","/NJH","/NJS")
                if ($LogFile) { $args += "/LOG+:`"$LogFile`"" }

                $psi = New-Object System.Diagnostics.ProcessStartInfo
                $psi.FileName = "robocopy.exe"
                $psi.Arguments = ($args -join " ")
                $psi.UseShellExecute = $false
                $psi.RedirectStandardOutput = $false
                $psi.RedirectStandardError = $false
                $psi.CreateNoWindow = $true

                $p = New-Object System.Diagnostics.Process
                $p.StartInfo = $psi

                try {
                    [void]$p.Start()
                    [void]$syncR.RoboPids.Add($p.Id)

                    while (-not $p.HasExited) {
                        if ($syncR.Cancel) {
                            LogMain "Cancel requested: killing robocopy PID $($p.Id)"
                            try { $p.Kill() } catch {}
                            try { $p.WaitForExit(3000) | Out-Null } catch {}
                            throw "Cancelled"
                        }
                        Start-Sleep -Milliseconds 200
                    }

                    LogMain "Robocopy exit code: $($p.ExitCode)"
                    if ($p.ExitCode -ge 8) {
                        if ($LogFile) { throw "Robocopy reported failure (ExitCode $($p.ExitCode)). See $LogFile" } 
                        else { throw "Robocopy reported failure (ExitCode $($p.ExitCode))." }
                    }
                }
                finally {
                    try { [void]$syncR.RoboPids.Remove($p.Id) } catch {}
                    try { $p.Dispose() } catch {}
                }
            }

            function Restore-BookmarksAllProfiles {
                param([string]$BackupBrowserRoot, [string]$UserDataRoot, [string]$BrowserName)
                if ($syncR.Cancel) { throw "Cancelled" }

                if (-not (Test-Path $BackupBrowserRoot)) {
                    $ui.RestoreAppend.Invoke("${BrowserName}: no bookmark backup folder found: $BackupBrowserRoot")
                    return
                }
                if (-not (Test-Path $UserDataRoot)) {
                    $ui.RestoreAppend.Invoke("${BrowserName}: user data folder not found: $UserDataRoot")
                    return
                }

                $procName = if ($BrowserName -eq "Chrome") { "chrome" } else { "msedge" }
                $running = Get-Process -Name $procName -ErrorAction SilentlyContinue
                if ($running) {
                    $ui.RestoreAppend.Invoke("${BrowserName}: WARNING - $procName is currently running. Bookmarks may be locked. Close the browser and run restore again for best results.")
                }

                $profiles = Get-ChildItem -Path $BackupBrowserRoot -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq "Default" -or $_.Name -like "Profile *" }
                if (-not $profiles) {
                    $ui.RestoreAppend.Invoke("${BrowserName}: no profile backups found under: $BackupBrowserRoot")
                    return
                }

                $bestBookmarksSource = $null
                $bestBookmarksBak = $null

                foreach ($p in $profiles) {
                    if ($syncR.Cancel) { throw "Cancelled" }

                    $destProfile = Join-Path $UserDataRoot $p.Name
                    New-Item -ItemType Directory -Path $destProfile -Force | Out-Null

                    $srcBookmarks = Join-Path $p.FullName "Bookmarks"
                    $srcBak = Join-Path $p.FullName "Bookmarks.bak"
                    $dstBookmarks = Join-Path $destProfile "Bookmarks"
                    $dstBak = Join-Path $destProfile "Bookmarks.bak"

                    $copiedAny = $false

                    if (Test-Path $srcBookmarks) {
                        try {
                            Copy-Item -Path $srcBookmarks -Destination $dstBookmarks -Force -ErrorAction Stop
                            if (Test-Path $dstBookmarks) { $copiedAny = $true }
                            if (-not $bestBookmarksSource -or $p.Name -eq "Default") { $bestBookmarksSource = $srcBookmarks }
                        } catch { $ui.RestoreAppend.Invoke("${BrowserName}: FAILED to copy Bookmarks for '$($p.Name)'.") }
                    }

                    if (Test-Path $srcBak) {
                        try {
                            Copy-Item -Path $srcBak -Destination $dstBak -Force -ErrorAction Stop
                            if (Test-Path $dstBak) { $copiedAny = $true }
                            if (-not $bestBookmarksBak -or $p.Name -eq "Default") { $bestBookmarksBak = $srcBak }
                        } catch { $ui.RestoreAppend.Invoke("${BrowserName}: FAILED to copy Bookmarks.bak for '$($p.Name)'.") }
                    }

                    if ($copiedAny) {
                        $ui.RestoreAppend.Invoke("${BrowserName}: restored bookmarks for profile '$($p.Name)' -> $destProfile")
                    }
                }

                if ($BrowserName -eq "Chrome" -and $bestBookmarksSource) {
                    $targets = @((Join-Path $UserDataRoot "Default"), ([IO.Path]::Combine($UserDataRoot, "Profile 1")))
                    foreach ($t in $targets) {
                        try {
                            New-Item -ItemType Directory -Path $t -Force | Out-Null
                            Copy-Item -Path $bestBookmarksSource -Destination (Join-Path $t "Bookmarks") -Force -ErrorAction Stop
                            if ($bestBookmarksBak) { Copy-Item -Path $bestBookmarksBak -Destination (Join-Path $t "Bookmarks.bak") -Force -ErrorAction SilentlyContinue }
                            $ui.RestoreAppend.Invoke("Chrome: also copied bookmarks into '$t'")
                        } catch {}
                    }
                }
            }

            try {
                $tasks = New-Object System.Collections.Generic.List[string]

                if ($cfgR.DoDesktop) { $tasks.Add("Desktop") }
                if ($cfgR.DoDocuments) { $tasks.Add("Documents") }
                if ($cfgR.DoPictures) { $tasks.Add("Pictures") }
                if ($cfgR.DoDownloads) { $tasks.Add("Downloads") }
                if ($cfgR.DoMusic) { $tasks.Add("Music") }
                if ($cfgR.DoVideos) { $tasks.Add("Videos") }
                if ($cfgR.DoChromeBookmarks) { $tasks.Add("Chrome Bookmarks") }
                if ($cfgR.DoEdgeBookmarks) { $tasks.Add("Edge Bookmarks") }
                if ($cfgR.DoFirefoxBookmarks) { $tasks.Add("Firefox Bookmarks") }
                if ($cfgR.DoLegacyFavorites) { $tasks.Add("Legacy Favorites") }
                if ($cfgR.DoPrinters) { $tasks.Add("Printers") }
                if ($cfgR.DoWallpaper) { $tasks.Add("Wallpaper") }
                if ($cfgR.DoFileAssociations) { $tasks.Add("File Associations") }

                $total = $tasks.Count
                $done = 0

                $ui.RestoreSetProgress.Invoke(0)
                $ui.RestoreSetStatus.Invoke("Starting...")
                $ui.RestoreAppend.Invoke("Tasks: " + ($tasks -join ", "))

                function Step {
                    param([string]$Status)
                    $script:done++
                    $pct = [int][math]::Round(($script:done / $total) * 100)
                    $ui.RestoreSetProgress.Invoke($pct)
                    if ($Status) { $ui.RestoreSetStatus.Invoke($Status) }
                }

                $script:done = 0

                if ($cfgR.DoDesktop) { Step "Restoring Desktop..."; Invoke-RoboCopyCancellable -Source (Join-Path $cfgR.RestoreSource "Desktop") -Destination (Join-Path $cfgR.UserProfile "Desktop") -LogFile $cfgR.LogFile }
                if ($cfgR.DoDocuments) { Step "Restoring Documents..."; Invoke-RoboCopyCancellable -Source (Join-Path $cfgR.RestoreSource "Documents") -Destination (Join-Path $cfgR.UserProfile "Documents") -LogFile $cfgR.LogFile }
                if ($cfgR.DoPictures) { Step "Restoring Pictures..."; Invoke-RoboCopyCancellable -Source (Join-Path $cfgR.RestoreSource "Pictures") -Destination (Join-Path $cfgR.UserProfile "Pictures") -LogFile $cfgR.LogFile }
                if ($cfgR.DoDownloads) { Step "Restoring Downloads..."; Invoke-RoboCopyCancellable -Source (Join-Path $cfgR.RestoreSource "Downloads") -Destination (Join-Path $cfgR.UserProfile "Downloads") -LogFile $cfgR.LogFile }
                if ($cfgR.DoMusic) { Step "Restoring Music..."; Invoke-RoboCopyCancellable -Source (Join-Path $cfgR.RestoreSource "Music") -Destination (Join-Path $cfgR.UserProfile "Music") -LogFile $cfgR.LogFile }
                if ($cfgR.DoVideos) { Step "Restoring Videos..."; Invoke-RoboCopyCancellable -Source (Join-Path $cfgR.RestoreSource "Videos") -Destination (Join-Path $cfgR.UserProfile "Videos") -LogFile $cfgR.LogFile }

                $browserBackupRoot = Join-Path $cfgR.RestoreSource "Browser_Backup"

                if ($cfgR.DoChromeBookmarks) {
                    Step "Restoring Chrome bookmarks..."
                    $chromeUserData = [IO.Path]::Combine($cfgR.LocalAppData, "Google", "Chrome", "User Data")
                    Restore-BookmarksAllProfiles -BackupBrowserRoot (Join-Path $browserBackupRoot "Chrome") -UserDataRoot $chromeUserData -BrowserName "Chrome"
                }

                if ($cfgR.DoEdgeBookmarks) {
                    Step "Restoring Edge bookmarks..."
                    $edgeUserData = [IO.Path]::Combine($cfgR.LocalAppData, "Microsoft", "Edge", "User Data")
                    Restore-BookmarksAllProfiles -BackupBrowserRoot (Join-Path $browserBackupRoot "Edge") -UserDataRoot $edgeUserData -BrowserName "Edge"
                }

                if ($cfgR.DoFirefoxBookmarks) {
                    Step "Restoring Firefox bookmarks..."
                    $ffBackupRoot = Join-Path $browserBackupRoot "Firefox"

                    if (-not (Test-Path $ffBackupRoot)) {
                        $ui.RestoreAppend.Invoke("Firefox: no backup folder found at $ffBackupRoot")
                    } else {
                        # Warn if Firefox is running
                        $ffRunning = Get-Process -Name "firefox" -ErrorAction SilentlyContinue
                        if ($ffRunning) {
                            $ui.RestoreAppend.Invoke("Firefox: WARNING - Firefox is currently running. Close it before restoring for best results.")
                        }

                        $ffProfilesRoot = [IO.Path]::Combine($env:APPDATA, "Mozilla", "Firefox", "Profiles")
                        $ffBackedProfiles = Get-ChildItem -Path $ffBackupRoot -Directory -ErrorAction SilentlyContinue

                        if (-not $ffBackedProfiles) {
                            $ui.RestoreAppend.Invoke("Firefox: no profile folders found in backup.")
                        } else {
                            foreach ($bp in $ffBackedProfiles) {
                                if ($syncR.Cancel) { throw "Cancelled" }

                                $srcPlaces   = Join-Path $bp.FullName "places.sqlite"
                                $srcFavicons = Join-Path $bp.FullName "favicons.sqlite"

                                if (-not (Test-Path $srcPlaces)) { continue }

                                # Try to match to an existing live profile by folder name first
                                $destProfileDir = Join-Path $ffProfilesRoot $bp.Name

                                if (-not (Test-Path $destProfileDir)) {
                                    # Profile folder name changed (new install) — find any live profile that has places.sqlite
                                    $liveProfiles = Get-ChildItem -Path $ffProfilesRoot -Directory -ErrorAction SilentlyContinue |
                                                    Where-Object { Test-Path (Join-Path $_.FullName "places.sqlite") }
                                    if ($liveProfiles) {
                                        # Prefer a profile whose name contains "default"
                                        $match = $liveProfiles | Where-Object { $_.Name -like "*default*" } | Select-Object -First 1
                                        if (-not $match) { $match = $liveProfiles | Select-Object -First 1 }
                                        $destProfileDir = $match.FullName
                                        $ui.RestoreAppend.Invoke("Firefox: profile '$($bp.Name)' not found — restoring into '$($match.Name)' instead.")
                                    } else {
                                        # No live profiles at all — create the folder and restore anyway
                                        New-Item -ItemType Directory -Path $destProfileDir -Force | Out-Null
                                        $ui.RestoreAppend.Invoke("Firefox: created new profile folder '$($bp.Name)' for restore.")
                                    }
                                }

                                try {
                                    Copy-Item -Path $srcPlaces -Destination $destProfileDir -Force -ErrorAction Stop
                                    $ui.RestoreAppend.Invoke("Firefox: restored places.sqlite -> $destProfileDir")
                                } catch {
                                    $ui.RestoreAppend.Invoke("Firefox: FAILED to restore places.sqlite for '$($bp.Name)': $($_.Exception.Message)")
                                }

                                if (Test-Path $srcFavicons) {
                                    try {
                                        Copy-Item -Path $srcFavicons -Destination $destProfileDir -Force -ErrorAction SilentlyContinue
                                        $ui.RestoreAppend.Invoke("Firefox: restored favicons.sqlite -> $destProfileDir")
                                    } catch {}
                                }
                            }
                        }
                    }
                }

                if ($cfgR.DoLegacyFavorites) {
                    Step "Restoring legacy Favorites..."
                    $legacySrc = Join-Path $browserBackupRoot "Legacy_Favorites_Folder"
                    $legacyDst = Join-Path $cfgR.UserProfile "Favorites"
                    Invoke-RoboCopyCancellable -Source $legacySrc -Destination $legacyDst -LogFile $cfgR.LogFile
                }

                if ($cfgR.DoPrinters) {
                    Step "Restoring network printers..."
                    $regFile = Join-Path $cfgR.RestoreSource "NetworkPrinters.reg"
                    if (Test-Path $regFile) {
                        $ui.RestoreAppend.Invoke("Importing network printers registry file...")
                        $process = Start-Process -FilePath "reg.exe" -ArgumentList "import `"$regFile`"" -Wait -PassThru -WindowStyle Hidden
                        if ($process.ExitCode -eq 0) {
                            $ui.RestoreAppend.Invoke("Successfully imported printers.")
                            $ui.RestoreAppend.Invoke("Note: You may need to log out/in to see them appear.")
                        } else {
                            $ui.RestoreAppend.Invoke("Warning: Failed to import printers via reg.exe.")
                        }
                    } else {
                        $ui.RestoreAppend.Invoke("No NetworkPrinters.reg found in backup.")
                    }
                }

                if ($cfgR.DoWallpaper) {
                    Step "Restoring wallpaper..."
                    $wpBackupDir = Join-Path $cfgR.RestoreSource "Wallpaper"
                    $jsonPath    = Join-Path $wpBackupDir "_wallpaper_info.json"

                    if (-not (Test-Path $jsonPath)) {
                        $ui.RestoreAppend.Invoke("Wallpaper: no wallpaper backup found in this backup set.")
                        LogMain "Wallpaper restore skipped — no _wallpaper_info.json found."
                    } else {
                        try {
                            $info = Get-Content $jsonPath -Raw -ErrorAction Stop | ConvertFrom-Json

                            if ($info.SourceType -eq "Spotlight") {
                                # If a fallback image was saved, apply it; otherwise just note it
                                if ($info.BackedUpFile -and (Test-Path (Join-Path $wpBackupDir $info.BackedUpFile))) {
                                    $imgPath = Join-Path $wpBackupDir $info.BackedUpFile

                                    # Copy to a stable location in the user's AppData so the path won't disappear
                                    $stableDir  = Join-Path $env:APPDATA "RestoredWallpaper"
                                    New-Item -ItemType Directory -Path $stableDir -Force | Out-Null
                                    $stablePath = Join-Path $stableDir ([IO.Path]::GetFileName($imgPath))
                                    Copy-Item -Path $imgPath -Destination $stablePath -Force -ErrorAction Stop

                                    # Apply via SystemParametersInfo
                                    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class WallpaperSetter {
    [DllImport("user32.dll", CharSet=CharSet.Auto)]
    public static extern bool SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
"@ -ErrorAction SilentlyContinue
                                    [WallpaperSetter]::SystemParametersInfo(0x0014, 0, $stablePath, 3) | Out-Null

                                    # Write registry style values
                                    $wpStyle = if ($info.WallpaperStyle) { $info.WallpaperStyle } else { "10" }
                                    $wpTile  = if ($info.TileWallpaper)  { $info.TileWallpaper  } else { "0"  }
                                    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "WallpaperStyle" -Value $wpStyle -ErrorAction SilentlyContinue
                                    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "TileWallpaper"  -Value $wpTile  -ErrorAction SilentlyContinue

                                    $ui.RestoreAppend.Invoke("Wallpaper: restored Spotlight fallback image. Note: Windows Spotlight rotation will not be re-enabled automatically.")
                                    LogMain "Wallpaper Spotlight fallback applied from: $stablePath"
                                } else {
                                    $ui.RestoreAppend.Invoke("Wallpaper: this backup was taken while Windows Spotlight was active — no image was saved to restore. Re-enable Spotlight manually via Settings > Personalization > Background.")
                                    LogMain "Wallpaper: Spotlight was active and no fallback image was captured."
                                }

                            } elseif ($info.SourceType -eq "Image" -and $info.BackedUpFile) {
                                $imgPath = Join-Path $wpBackupDir $info.BackedUpFile
                                if (-not (Test-Path $imgPath)) {
                                    $ui.RestoreAppend.Invoke("Wallpaper: backed-up image file not found: $imgPath")
                                    LogMain "Wallpaper image missing at: $imgPath"
                                } else {
                                    # Copy image to a stable AppData location
                                    $stableDir  = Join-Path $env:APPDATA "RestoredWallpaper"
                                    New-Item -ItemType Directory -Path $stableDir -Force | Out-Null
                                    $stablePath = Join-Path $stableDir $info.BackedUpFile
                                    Copy-Item -Path $imgPath -Destination $stablePath -Force -ErrorAction Stop

                                    # Apply via SystemParametersInfo (SPI_SETDESKWALLPAPER = 0x0014)
                                    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class WallpaperSetter {
    [DllImport("user32.dll", CharSet=CharSet.Auto)]
    public static extern bool SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
}
"@ -ErrorAction SilentlyContinue
                                    [WallpaperSetter]::SystemParametersInfo(0x0014, 0, $stablePath, 3) | Out-Null

                                    # Restore style and tile settings from backup
                                    $wpStyle = if ($info.WallpaperStyle) { $info.WallpaperStyle } else { "10" }
                                    $wpTile  = if ($info.TileWallpaper)  { $info.TileWallpaper  } else { "0"  }
                                    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "WallpaperStyle" -Value $wpStyle -ErrorAction SilentlyContinue
                                    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "TileWallpaper"  -Value $wpTile  -ErrorAction SilentlyContinue
                                    Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "WallPaper"       -Value $stablePath -ErrorAction SilentlyContinue

                                    $ui.RestoreAppend.Invoke("Wallpaper: restored '$($info.BackedUpFile)' (style=$wpStyle, tile=$wpTile).")
                                    LogMain "Wallpaper restored to: $stablePath (style=$wpStyle)"
                                }

                            } else {
                                # Solid color or unknown — nothing to apply, just inform
                                $ui.RestoreAppend.Invoke("Wallpaper: backup indicates no image was set (solid color or unsupported). No wallpaper applied.")
                                LogMain "Wallpaper restore: source type was '$($info.SourceType)' — nothing to apply."
                            }

                        } catch {
                            $ui.RestoreAppend.Invoke("Wallpaper: ERROR during restore — $($_.Exception.Message)")
                            LogMain "Wallpaper restore error: $($_.Exception.Message)"
                        }
                    }
                }

                if ($cfgR.DoFileAssociations) {
                    Step "Restoring File Associations..."
                    $faRegFile = Join-Path $cfgR.RestoreSource "FileAssociations.reg"
                    if (-not (Test-Path $faRegFile)) {
                        $ui.RestoreAppend.Invoke("File Associations: no FileAssociations.reg found in this backup set.")
                        LogMain "File Associations restore skipped — FileAssociations.reg not found."
                    } else {
                        try {
                            $ui.RestoreAppend.Invoke("File Associations: importing registry file...")
                            $process = Start-Process -FilePath "reg.exe" `
                                -ArgumentList "import `"$faRegFile`"" `
                                -Wait -PassThru -WindowStyle Hidden
                            if ($process.ExitCode -eq 0) {
                                $ui.RestoreAppend.Invoke("File Associations: restored successfully.")
                                $ui.RestoreAppend.Invoke("Note: You may need to sign out and back in for all changes to take effect.")
                                LogMain "File Associations imported from: $faRegFile"
                            } else {
                                $ui.RestoreAppend.Invoke("File Associations: WARNING — reg.exe exited with code $($process.ExitCode). Some associations may not have been restored.")
                                LogMain "File Associations import failed with exit code: $($process.ExitCode)"
                            }
                        } catch {
                            $ui.RestoreAppend.Invoke("File Associations: ERROR — $($_.Exception.Message)")
                            LogMain "File Associations restore error: $($_.Exception.Message)"
                        }
                    }
                }

                $ui.RestoreSetProgress.Invoke(100)
                $ui.RestoreSetStatus.Invoke("Restore complete.")
                $ui.RestoreAppend.Invoke("Restore completed successfully.")

                return [pscustomobject]@{ Result="Success"; Message="OK" }
            }
            catch {
                if ($_.Exception -and $_.Exception.Message -eq "Cancelled") {
                    $ui.RestoreSetStatus.Invoke("Cancelled."); $ui.RestoreAppend.Invoke("Cancelled.")
                    return [pscustomobject]@{ Result="Cancelled"; Message="Cancelled" }
                }

                if ($syncR.Cancel) {
                    $ui.RestoreSetStatus.Invoke("Cancelled."); $ui.RestoreAppend.Invoke("Cancelled.")
                    return [pscustomobject]@{ Result="Cancelled"; Message="Cancelled" }
                }

                $msg = $_.Exception.Message
                $ui.RestoreSetStatus.Invoke("Restore failed.")
                $ui.RestoreAppend.Invoke("ERROR: " + $msg)
                return [pscustomobject]@{ Result="Error"; Message=$msg }
            }
        }

        $psR = [PowerShell]::Create()
        $rsR = [RunspaceFactory]::CreateRunspace()
        $rsR.ApartmentState = "STA"
        $rsR.ThreadOptions = "ReuseThread"
        $rsR.Open()
        $psR.Runspace = $rsR

        $psR.AddScript($restoreWorker).AddArgument($cfgR).AddArgument($ui).AddArgument($syncRestore).AddArgument($global:BackupToolLog) | Out-Null
        $handleR = $psR.BeginInvoke()

        $timer = New-Object System.Windows.Threading.DispatcherTimer
        $timer.Interval = [TimeSpan]::FromMilliseconds(250)

        $timer.Add_Tick({
            if ($handleR -and $handleR.IsCompleted) {
                $timer.Stop()
                try { $res = $psR.EndInvoke($handleR) | Select-Object -Last 1 }
                catch { $txtRestoreStatus.Text = "Restore failed."; $txtRestoreActivity.AppendText("ERROR: $($_.Exception.Message)`r`n") }
                finally {
                    try { $psR.Dispose() } catch {}
                    try { $rsR.Close(); $rsR.Dispose() } catch {}
                    $btnStartRestore.IsEnabled = $true; $btnCancelRestore.IsEnabled = $true
                    $script:RestoreTimer = $null; $script:RestorePs = $null; $script:RestoreRs = $null; $script:RestoreHandle = $null
                }
            }
        })

        $script:RestoreTimer = $timer; $script:RestorePs = $psR; $script:RestoreRs = $rsR; $script:RestoreHandle = $handleR
        $timer.Start()
    })

    # ---------------------------
    # Window Closing - Cleanup
    # ---------------------------
    $Window.Add_Closing({
        Write-BackupToolLog "Window closing - cleaning up resources..."
        $syncBackup.Cancel = $true
        $syncRestore.Cancel = $true
        
        foreach ($pid in @($syncBackup.RoboPids)) {
            try { Write-BackupToolLog "Killing backup robocopy PID: $pid"; Stop-Process -Id $pid -Force -ErrorAction SilentlyContinue } catch {}
        }
        foreach ($pid in @($syncRestore.RoboPids)) {
            try { Write-BackupToolLog "Killing restore robocopy PID: $pid"; Stop-Process -Id $pid -Force -ErrorAction SilentlyContinue } catch {}
        }
        
        Start-Sleep -Milliseconds 200
        Write-BackupToolLog "Window closed - cleanup complete."
        [System.Environment]::Exit(0)
    })

    $Window.ShowDialog() | Out-Null
}
catch {
    Write-BackupToolLog "Fatal startup exception."
    Dump-Exception $_.Exception
    try { [System.Windows.MessageBox]::Show("A fatal error occurred while starting the tool.`n`nLog saved to:`n$global:BackupToolLog","Backup Tool Error","OK","Error") | Out-Null } catch {}
}
finally {
    Write-BackupToolLog "Script ending."
    try { Stop-Transcript | Out-Null } catch {}
}
