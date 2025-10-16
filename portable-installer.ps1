# PDF Merge Tool - Portable Installer
# No Python installation required - everything runs from a single folder
# Run: irm https://raw.githubusercontent.com/g-hazard/pdftool/main/portable-installer.ps1 | iex

param(
    [string]$InstallPath = "$env:LOCALAPPDATA\PDFMergeTool"
)

$ErrorActionPreference = "Stop"

function Write-Step {
    param([string]$Message)
    Write-Host "[*] $Message" -ForegroundColor Cyan
}

function Write-Success {
    param([string]$Message)
    Write-Host "[+] $Message" -ForegroundColor Green
}

function Write-Error {
    param([string]$Message)
    Write-Host "[!] $Message" -ForegroundColor Red
}

Write-Host ""
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  PDF Merge Tool - Portable Edition" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check if already installed
if (Test-Path $InstallPath) {
    Write-Host "Installation found at: $InstallPath" -ForegroundColor Yellow
    $response = Read-Host "Reinstall? (y/N)"
    if ($response -notlike "y*") {
        Write-Host "Installation cancelled." -ForegroundColor Yellow
        exit 0
    }
    Write-Step "Removing existing installation..."
    Remove-Item -Path $InstallPath -Recurse -Force
}

# Create installation directory
Write-Step "Creating installation directory..."
New-Item -ItemType Directory -Path $InstallPath -Force | Out-Null

# Download portable Python (embeddable package)
Write-Step "Downloading portable Python 3.11.9 (~25MB)..."
$pythonVersion = "3.11.9"
$pythonUrl = "https://www.python.org/ftp/python/$pythonVersion/python-$pythonVersion-embed-amd64.zip"
$pythonZip = Join-Path $InstallPath "python.zip"

try {
    Invoke-WebRequest -Uri $pythonUrl -OutFile $pythonZip -UseBasicParsing
    Write-Success "Python downloaded"
} catch {
    Write-Error "Failed to download Python: $_"
    exit 1
}

# Extract Python
Write-Step "Extracting Python..."
Expand-Archive -Path $pythonZip -DestinationPath $InstallPath -Force
Remove-Item $pythonZip

# Create pythonw.exe (embedded package only has python.exe)
$pythonExe = Join-Path $InstallPath "python.exe"
$pythonwExe = Join-Path $InstallPath "pythonw.exe"
if ((Test-Path $pythonExe) -and (-not (Test-Path $pythonwExe))) {
    Copy-Item -Path $pythonExe -Destination $pythonwExe
}

# Enable site-packages for pip
Write-Step "Configuring Python..."
$pthFile = Get-ChildItem -Path $InstallPath -Filter "*._pth" | Select-Object -First 1
if ($pthFile) {
    $content = Get-Content $pthFile.FullName
    $content = $content -replace "#import site", "import site"
    $content | Set-Content $pthFile.FullName
}

# Download and install pip
Write-Step "Installing pip..."
$getPipUrl = "https://bootstrap.pypa.io/get-pip.py"
$getPipPath = Join-Path $InstallPath "get-pip.py"
Invoke-WebRequest -Uri $getPipUrl -OutFile $getPipPath -UseBasicParsing

$pythonExe = Join-Path $InstallPath "python.exe"
& $pythonExe $getPipPath --no-warn-script-location 2>&1 | Out-Null
Remove-Item $getPipPath

# Install required packages
Write-Step "Installing PDF libraries (pypdf, winotify, filelock, pywin32)..."
& $pythonExe -m pip install --no-warn-script-location pypdf winotify filelock pywin32 2>&1 | Out-Null
Write-Success "Libraries installed"

# Download scripts from GitHub
Write-Step "Downloading PDF merge scripts..."
$baseUrl = "https://raw.githubusercontent.com/g-hazard/pdftool/main"
$scripts = @(
    "merge_pdfs.py",
    "merge_pdf_handler.py"
)

foreach ($script in $scripts) {
    $url = "$baseUrl/$script"
    $destination = Join-Path $InstallPath $script
    try {
        Invoke-WebRequest -Uri $url -OutFile $destination -UseBasicParsing
        Write-Host "  Downloaded: $script" -ForegroundColor Gray
    } catch {
        Write-Error "Failed to download $script : $_"
        exit 1
    }
}

# Create wrapper script for context menu (VBS for completely silent execution)
Write-Step "Creating launcher..."
$wrapperScript = Join-Path $InstallPath "merge-pdf.vbs"
$wrapperContent = @"
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
strScriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
strPython = strScriptPath & "\pythonw.exe"
strHandler = strScriptPath & "\merge_pdf_handler.py"
strArgs = WScript.Arguments(0)
objShell.Run Chr(34) & strPython & Chr(34) & " " & Chr(34) & strHandler & Chr(34) & " " & Chr(34) & strArgs & Chr(34), 0, False
"@
$wrapperContent | Set-Content $wrapperScript

# Register context menu
Write-Step "Registering context menu..."
$registryBase = "HKCU:\Software\Classes\SystemFileAssociations\.pdf\shell\Merge PDF"
New-Item -Path $registryBase -Force | Out-Null
Set-ItemProperty -Path $registryBase -Name "MUIVerb" -Value "Merge PDF"
Set-ItemProperty -Path $registryBase -Name "MultiSelectModel" -Value "Player"

$iconSource = Join-Path $env:SystemRoot "System32\imageres.dll"
if (Test-Path $iconSource) {
    Set-ItemProperty -Path $registryBase -Name "Icon" -Value "$iconSource,15"
}

$commandKey = Join-Path $registryBase "command"
New-Item -Path $commandKey -Force | Out-Null
$commandValue = "wscript.exe `"$wrapperScript`" `"%1`""
Set-ItemProperty -Path $commandKey -Name "(Default)" -Value $commandValue

Write-Success "Context menu registered"

# Create uninstaller
Write-Step "Creating uninstaller..."
$uninstallerPath = Join-Path $InstallPath "uninstall.ps1"
$uninstallerContent = @"
# PDF Merge Tool - Uninstaller
Write-Host "Uninstalling PDF Merge Tool..." -ForegroundColor Yellow

# Remove context menu
Remove-Item -Path "HKCU:\Software\Classes\SystemFileAssociations\.pdf\shell\Merge PDF" -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "Context menu removed" -ForegroundColor Green

# Remove installation directory
`$installPath = "$InstallPath"
Write-Host "Removing files from: `$installPath" -ForegroundColor Yellow
Start-Sleep -Seconds 2

# Self-delete: move uninstaller to temp and schedule deletion
`$tempUninstaller = Join-Path `$env:TEMP "pdftool-uninstall.ps1"
Move-Item -Path `$PSCommandPath -Destination `$tempUninstaller -Force

# Remove installation directory
Remove-Item -Path `$installPath -Recurse -Force

Write-Host ""
Write-Host "PDF Merge Tool has been completely removed." -ForegroundColor Green
Write-Host "Press any key to exit..." -ForegroundColor Gray
`$null = `$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
"@
$uninstallerContent | Set-Content $uninstallerPath

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "  Installation Complete!" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Installation Path:" -ForegroundColor Cyan
Write-Host "  $InstallPath" -ForegroundColor White
Write-Host ""
Write-Host "Size: ~45MB (portable Python + libraries)" -ForegroundColor Gray
Write-Host ""
Write-Host "Usage:" -ForegroundColor Cyan
Write-Host "  1. Select PDF files in File Explorer" -ForegroundColor White
Write-Host "  2. Right-click → 'Merge PDF'" -ForegroundColor White
Write-Host "  3. Choose output location" -ForegroundColor White
Write-Host ""
Write-Host "To Uninstall:" -ForegroundColor Cyan
Write-Host "  Run: powershell -File `"$InstallPath\uninstall.ps1`"" -ForegroundColor White
Write-Host "  Or simply delete: $InstallPath" -ForegroundColor White
Write-Host ""
Write-Host "No system-wide Python installation required!" -ForegroundColor Green
Write-Host ""

