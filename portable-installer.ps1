# PDF Merge Tool - Portable Installer
# No Python installation required - everything runs from a single folder
# Run: irm https://raw.githubusercontent.com/g-hazard/pdftool/main/portable-installer.ps1 | iex

param(
    [string]$InstallPath = "$env:LOCALAPPDATA\PDFMergeTool"
)

$ErrorActionPreference = "Stop"

function Show-Progress {
    param(
        [int]$Percent,
        [string]$Status
    )
    
    $barLength = 40
    $filled = [math]::Floor($barLength * $Percent / 100)
    $empty = $barLength - $filled
    
    $bar = "[" + ("█" * $filled) + ("░" * $empty) + "]"
    
    Write-Host "`r$bar $Percent% - $Status" -NoNewline -ForegroundColor Cyan
}

Clear-Host
Write-Host ""
Write-Host "  PDF Merge Tool - Installer" -ForegroundColor Cyan
Write-Host ""

# Check if already installed
if (Test-Path $InstallPath) {
    Remove-Item -Path $InstallPath -Recurse -Force -ErrorAction SilentlyContinue
}

try {
    # Step 1: Create directory
    Show-Progress -Percent 5 -Status "Preparing installation..."
    New-Item -ItemType Directory -Path $InstallPath -Force | Out-Null
    Start-Sleep -Milliseconds 300

    # Step 2: Download Python
    Show-Progress -Percent 10 -Status "Downloading Python (~25MB)..."
    $pythonVersion = "3.11.9"
    $pythonUrl = "https://www.python.org/ftp/python/$pythonVersion/python-$pythonVersion-embed-amd64.zip"
    $pythonZip = Join-Path $InstallPath "python.zip"
    Invoke-WebRequest -Uri $pythonUrl -OutFile $pythonZip -UseBasicParsing
    Show-Progress -Percent 30 -Status "Python downloaded"
    Start-Sleep -Milliseconds 300

    # Step 3: Extract Python
    Show-Progress -Percent 35 -Status "Extracting Python..."
    Expand-Archive -Path $pythonZip -DestinationPath $InstallPath -Force
    Remove-Item $pythonZip
    
    # Create pythonw.exe
    $pythonExe = Join-Path $InstallPath "python.exe"
    $pythonwExe = Join-Path $InstallPath "pythonw.exe"
    if ((Test-Path $pythonExe) -and (-not (Test-Path $pythonwExe))) {
        Copy-Item -Path $pythonExe -Destination $pythonwExe
    }
    Show-Progress -Percent 40 -Status "Configuring Python..."
    Start-Sleep -Milliseconds 300
    
    # Enable site-packages
    $pthFile = Get-ChildItem -Path $InstallPath -Filter "*._pth" | Select-Object -First 1
    if ($pthFile) {
        $content = Get-Content $pthFile.FullName
        $content = $content -replace "#import site", "import site"
        $content | Set-Content $pthFile.FullName
    }
    
    # Step 4: Install pip
    Show-Progress -Percent 45 -Status "Installing pip..."
    $getPipUrl = "https://bootstrap.pypa.io/get-pip.py"
    $getPipPath = Join-Path $InstallPath "get-pip.py"
    Invoke-WebRequest -Uri $getPipUrl -OutFile $getPipPath -UseBasicParsing
    & $pythonExe $getPipPath --no-warn-script-location 2>&1 | Out-Null
    Remove-Item $getPipPath
    Show-Progress -Percent 55 -Status "Installing libraries..."
    Start-Sleep -Milliseconds 300
    
    # Step 5: Install packages
    & $pythonExe -m pip install --no-warn-script-location pypdf winotify filelock pywin32 2>&1 | Out-Null
    Show-Progress -Percent 70 -Status "Libraries installed"
    Start-Sleep -Milliseconds 300

    # Step 6: Download scripts
    Show-Progress -Percent 75 -Status "Downloading scripts..."
    $baseUrl = "https://raw.githubusercontent.com/g-hazard/pdftool/main"
    $scripts = @("merge_pdfs.py", "merge_pdf_handler.py")
    
    foreach ($script in $scripts) {
        $url = "$baseUrl/$script"
        $destination = Join-Path $InstallPath $script
        Invoke-WebRequest -Uri $url -OutFile $destination -UseBasicParsing
    }
    Show-Progress -Percent 85 -Status "Creating launcher..."
    Start-Sleep -Milliseconds 300
    
    # Step 7: Create wrapper
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
    
    # Step 8: Register context menu
    Show-Progress -Percent 90 -Status "Registering context menu..."
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
    Start-Sleep -Milliseconds 300

    # Step 9: Create uninstaller
    Show-Progress -Percent 95 -Status "Finalizing..."
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
    
    # Complete
    Show-Progress -Percent 100 -Status "Installation complete!"
    Write-Host ""
    Write-Host ""
    Write-Host "  ✓ Installation successful!" -ForegroundColor Green
    Write-Host ""
    Write-Host "  Right-click PDF files → 'Merge PDF'" -ForegroundColor White
    Write-Host "  Output: merged.pdf in same folder" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Window will close in 5 seconds..." -ForegroundColor DarkGray
    Start-Sleep -Seconds 5
    
} catch {
    Write-Host ""
    Write-Host ""
    Write-Host "  ✗ Installation failed: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "  Press any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

