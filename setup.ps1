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

function Test-RealPython {
    param([string]$Path)
    
    if (-not $Path -or -not (Test-Path $Path)) {
        return $false
    }
    
    # Test if this is actually Python and not the Microsoft Store stub
    try {
        $output = & $Path --version 2>&1
        $exitCode = $LASTEXITCODE
        # Microsoft Store stub exits with 9009 and outputs error message
        if ($exitCode -ne 0 -or $output -like "*Microsoft Store*") {
            return $false
        }
        return $true
    } catch {
        return $false
    }
}

function Get-PythonLaunchers {
    $cli = $null
    $context = $null

    # Try py launcher first (most reliable on Windows)
    $pyHint = Join-Path -Path $env:SystemRoot -ChildPath "py.exe"
    if ((Test-Path -Path $pyHint) -and (Test-RealPython -Path $pyHint)) {
        $cli = $pyHint
    }

    # Try pyw launcher
    $pywHint = Join-Path -Path $env:SystemRoot -ChildPath "pyw.exe"
    if ((Test-Path -Path $pywHint) -and (Test-RealPython -Path $pywHint)) {
        $context = $pywHint
    }

    # If py not found in System32, try PATH
    if (-not $cli) {
        $pyCommand = Get-Command -Name "py" -ErrorAction SilentlyContinue
        if ($pyCommand) {
            $pyPath = Resolve-CommandPath -Command $pyCommand
            if (Test-RealPython -Path $pyPath) {
                $cli = $pyPath
            }
        }
    }

    # Try python.exe as last resort (avoid Microsoft Store stub)
    if (-not $cli) {
        $pythonCommand = Get-Command -Name "python" -ErrorAction SilentlyContinue
        if ($pythonCommand) {
            $pythonPath = Resolve-CommandPath -Command $pythonCommand
            # Extra check: Microsoft Store stub is in WindowsApps folder
            if ($pythonPath -notlike "*WindowsApps*" -and (Test-RealPython -Path $pythonPath)) {
                $cli = $pythonPath
            }
        }
    }

    if (-not $context) {
        $pywCommand = Get-Command -Name "pyw" -ErrorAction SilentlyContinue
        if ($pywCommand) {
            $pywPath = Resolve-CommandPath -Command $pywCommand
            if (Test-RealPython -Path $pywPath) {
                $context = $pywPath
            }
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
        Write-Host "This might be due to the Windows Store Python alias interfering." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Please do ONE of the following:" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Option 1 - Disable Windows Store Python alias:" -ForegroundColor White
        Write-Host "  1. Open Settings > Apps > Advanced app settings > App execution aliases" -ForegroundColor Gray
        Write-Host "  2. Turn OFF both 'python.exe' and 'python3.exe' aliases" -ForegroundColor Gray
        Write-Host "  3. Close PowerShell and run the installer again" -ForegroundColor Gray
        Write-Host ""
        Write-Host "Option 2 - Manually install Python:" -ForegroundColor White
        Write-Host "  1. Go to: https://www.python.org/downloads/" -ForegroundColor Gray
        Write-Host "  2. Download and run the installer" -ForegroundColor Gray
        Write-Host "  3. CHECK 'Add Python to PATH' during installation" -ForegroundColor Gray
        Write-Host "  4. Run the installer again" -ForegroundColor Gray
        Write-Host ""
        throw "Python installation requires configuration. Please follow the steps above."
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

    Write-Note "Installing required packages (pypdf, winotify, filelock, pywin32)."
    Invoke-PythonCommand -Launcher $CliLauncher -Arguments @("-m", "pip", "install", "--upgrade", "pypdf", "winotify", "filelock", "pywin32") -ErrorMessage "Failed to install required packages."
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
