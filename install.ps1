# PDF Merge Tool - Quick Installer
# Run: irm https://raw.githubusercontent.com/YOUR_USERNAME/pdftool/main/install.ps1 | iex

Write-Host "PDF Merge Tool - Installer" -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan
Write-Host ""

# Check if running as administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    Write-Host "Note: For best results, run as Administrator (optional)" -ForegroundColor Yellow
    Write-Host ""
}

# Create temporary directory
$tempDir = Join-Path $env:TEMP "pdftool_install"
if (Test-Path $tempDir) {
    Remove-Item $tempDir -Recurse -Force
}
New-Item -ItemType Directory -Path $tempDir | Out-Null

Write-Host "Downloading PDF Merge Tool..." -ForegroundColor Green

# Base URL for raw GitHub content (update with your repository)
$baseUrl = "https://raw.githubusercontent.com/YOUR_USERNAME/pdftool/main"

# List of files to download
$files = @(
    "merge_pdfs.py",
    "merge_pdf_handler.py",
    "register_context_menu.ps1",
    "unregister_context_menu.ps1",
    "setup.ps1"
)

# Download each file
foreach ($file in $files) {
    try {
        $url = "$baseUrl/$file"
        $destination = Join-Path $tempDir $file
        Invoke-WebRequest -Uri $url -OutFile $destination -UseBasicParsing
        Write-Host "  Downloaded: $file" -ForegroundColor Gray
    } catch {
        Write-Host "  Failed to download: $file" -ForegroundColor Red
        Write-Host "  Error: $_" -ForegroundColor Red
        exit 1
    }
}

Write-Host ""
Write-Host "Running setup..." -ForegroundColor Green
Write-Host ""

# Run the setup script
Push-Location $tempDir
try {
    & powershell -ExecutionPolicy Bypass -File "$tempDir\setup.ps1"
    $exitCode = $LASTEXITCODE
    
    if ($exitCode -eq 0) {
        Write-Host ""
        Write-Host "Installation completed successfully!" -ForegroundColor Green
        Write-Host ""
        Write-Host "To use:" -ForegroundColor Cyan
        Write-Host "  1. Right-click on PDF files in File Explorer" -ForegroundColor White
        Write-Host "  2. Select 'Merge PDF' to combine multiple files" -ForegroundColor White
        Write-Host ""
        Write-Host "To uninstall:" -ForegroundColor Cyan
        Write-Host "  Run: $tempDir\unregister_context_menu.ps1" -ForegroundColor White
    } else {
        Write-Host ""
        Write-Host "Installation failed. Please check the error messages above." -ForegroundColor Red
    }
} catch {
    Write-Host "Error running setup: $_" -ForegroundColor Red
    exit 1
} finally {
    Pop-Location
}

Write-Host ""
Write-Host "Setup files are in: $tempDir" -ForegroundColor Gray
Write-Host "You can delete this directory after installation is complete." -ForegroundColor Gray

