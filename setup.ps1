param(
    [string]$VerbName = "Merge PDF",
    [string]$PythonVersion = "3.12.3",
    [switch]$SkipRegister
)

Set-StrictMode -Version 3
$ErrorActionPreference = "Stop"

# Check if running as administrator
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

function Write-Note {
    param([string]$Message)
    Write-Host "[*] $Message"
}

function Invoke-PythonCommand {
    param(
        [string]$Launcher,
        [string[]]$Arguments,
        [string]$ErrorMessage
    )

    $display = "$(Split-Path -Path $Launcher -Leaf) $($Arguments -join ' ')"
    Write-Note "Running $display"

    & $Launcher @Arguments
    $exitCode = $LASTEXITCODE
    if ($exitCode -ne 0) {
        throw "$ErrorMessage (exit code $exitCode)"
    }
}

function Resolve-CommandPath {
    param([System.Management.Automation.CommandInfo]$Command)

    if ($null -ne $Command) {
        if ($Command.Source) { return $Command.Source }
        if ($Command.Definition) { return $Command.Definition }
    }
    return $null
}

function Get-PythonLaunchers {
    $cli = $null
    $context = $null

    $pyCommand = Get-Command -Name "py" -ErrorAction SilentlyContinue
    if ($pyCommand) {
        $cli = Resolve-CommandPath -Command $pyCommand
    }

    $pywCommand = Get-Command -Name "pyw" -ErrorAction SilentlyContinue
    if ($pywCommand) {
        $context = Resolve-CommandPath -Command $pywCommand
    }

    if (-not $cli) {
        $pyHint = Join-Path -Path $env:SystemRoot -ChildPath "py.exe"
        if (Test-Path -Path $pyHint) {
            $cli = $pyHint
        }
    }

    if (-not $context) {
        $pywHint = Join-Path -Path $env:SystemRoot -ChildPath "pyw.exe"
        if (Test-Path -Path $pywHint) {
            $context = $pywHint
        }
    }

    if (-not $cli) {
        $pythonCommand = Get-Command -Name "python" -ErrorAction SilentlyContinue
        if ($pythonCommand) {
            $cli = Resolve-CommandPath -Command $pythonCommand
        }
    }

    if (-not $context) {
        $context = $cli
    }

    return [PSCustomObject]@{
        Cli     = $cli
        Context = $context
    }
}

function Install-Python {
    param([string]$Version)

    Write-Note "Attempting to install Python $Version."

    $winget = Get-Command -Name "winget" -ErrorAction SilentlyContinue
    if ($winget) {
        $wingetPath = Resolve-CommandPath -Command $winget
        if (-not $wingetPath) {
            throw "winget command resolution failed."
        }

        $parts = $Version.Split('.')
        if ($parts.Length -ge 2) {
            $majorMinor = "$($parts[0]).$($parts[1])"
        }
        else {
            $majorMinor = $parts[0]
        }
        $packageId = "Python.Python.$majorMinor"

        Write-Note "Using winget to install $packageId."
        $wingetArgs = @(
            "install",
            "--id", $packageId,
            "--exact",
            "--accept-source-agreements",
            "--accept-package-agreements",
            "--silent"
        )

        $process = Start-Process -FilePath $wingetPath -ArgumentList $wingetArgs -Wait -PassThru
        if ($process.ExitCode -eq 0) {
            return
        }
        Write-Warning "winget install exited with code $($process.ExitCode); falling back to direct installer."
    }
    else {
        Write-Note "winget not available."
    }

    $uri = "https://www.python.org/ftp/python/$Version/python-$Version-amd64.exe"
    $installerPath = Join-Path -Path $env:TEMP -ChildPath "python-$Version-amd64.exe"
    Write-Note "Downloading $uri"
    Invoke-WebRequest -Uri $uri -OutFile $installerPath
    try {
        $arguments = "/quiet InstallAllUsers=0 PrependPath=1 Include_launcher=1 Include_test=0 Include_pip=1"
        Write-Note "Running installer quietly."
        $process = Start-Process -FilePath $installerPath -ArgumentList $arguments -Wait -PassThru
        if ($process.ExitCode -ne 0) {
            throw "Python installer exited with code $($process.ExitCode)."
        }
    }
    finally {
        if (Test-Path -Path $installerPath) {
            Remove-Item -Path $installerPath -Force
        }
    }
}

function Refresh-EnvironmentPath {
    Write-Note "Refreshing environment PATH..."
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path", "User")
}

function Ensure-Python {
    param([string]$Version)

    $launchers = Get-PythonLaunchers
    if ($launchers.Cli -and $launchers.Context) {
        Write-Note "Python is already installed."
        return $launchers
    }

    Write-Note "Python not found. Installing Python $Version..."
    Install-Python -Version $Version
    
    # Refresh PATH to pick up newly installed Python
    Refresh-EnvironmentPath
    Start-Sleep -Seconds 2
    
    $launchers = Get-PythonLaunchers
    if (-not $launchers.Cli) {
        Write-Host ""
        Write-Host "Python installation completed, but Python launcher not found in PATH." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Please do ONE of the following:" -ForegroundColor Cyan
        Write-Host "  1. Close this PowerShell window and run the installer again" -ForegroundColor White
        Write-Host "  2. Or manually install Python from: https://www.python.org/downloads/" -ForegroundColor White
        Write-Host "     Make sure to check 'Add Python to PATH' during installation" -ForegroundColor White
        Write-Host ""
        throw "Python installation requires a shell restart. Please rerun the installer."
    }
    if (-not $launchers.Context) {
        Write-Warning "Python windowed launcher (pyw) not found. Context menu will use '$($launchers.Cli)'."
        $launchers = [PSCustomObject]@{ Cli = $launchers.Cli; Context = $launchers.Cli }
    }
    return $launchers
}

function Ensure-PdfLibraries {
    param(
        [string]$CliLauncher
    )

    Write-Note "Ensuring ensurepip is available."
    Invoke-PythonCommand -Launcher $CliLauncher -Arguments @("-m", "ensurepip", "--upgrade") -ErrorMessage "Failed to run ensurepip."

    Write-Note "Upgrading pip."
    Invoke-PythonCommand -Launcher $CliLauncher -Arguments @("-m", "pip", "install", "--upgrade", "pip") -ErrorMessage "Failed to upgrade pip."

    Write-Note "Installing required packages (pypdf, winotify, filelock)."
    Invoke-PythonCommand -Launcher $CliLauncher -Arguments @("-m", "pip", "install", "--upgrade", "pypdf", "winotify", "filelock") -ErrorMessage "Failed to install required packages."
}

function Register-ContextMenu {
    param(
        [string]$Verb,
        [string]$PythonLauncher
    )

    $registerScript = Join-Path -Path $PSScriptRoot -ChildPath "register_context_menu.ps1"
    if (-not (Test-Path -Path $registerScript)) {
        throw "Could not find register_context_menu.ps1 at $registerScript."
    }

    Write-Note "Registering context menu entry '$Verb'."
    & $registerScript -VerbName $Verb -PythonLauncher $PythonLauncher
}

Write-Note "Starting PDF merge tool setup."

if (-not $isAdmin) {
    Write-Host ""
    Write-Host "NOTE: Not running as Administrator" -ForegroundColor Yellow
    Write-Host "If Python installation fails, please:" -ForegroundColor Yellow
    Write-Host "  1. Right-click PowerShell" -ForegroundColor White
    Write-Host "  2. Select 'Run as Administrator'" -ForegroundColor White
    Write-Host "  3. Run the installer again" -ForegroundColor White
    Write-Host ""
}

$launchers = Ensure-Python -Version $PythonVersion

Ensure-PdfLibraries -CliLauncher $launchers.Cli

if (-not $SkipRegister.IsPresent) {
    Register-ContextMenu -Verb $VerbName -PythonLauncher $launchers.Context
}
else {
    Write-Note "SkipRegister specified; context menu not updated."
}

Write-Note "Setup complete."
