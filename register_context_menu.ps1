param(
    [string]$VerbName = "Merge PDF",
    [string]$PythonLauncher = ""
)

$handlerScript = Join-Path -Path $PSScriptRoot -ChildPath "merge_pdf_handler.py"
if (-not (Test-Path -Path $handlerScript)) {
    throw "Cannot find merge_pdf_handler.py at $handlerScript"
}

if ([string]::IsNullOrWhiteSpace($PythonLauncher)) {
    $candidateLaunchers = @(
        (Join-Path -Path $env:SystemRoot -ChildPath "pyw.exe"),
        (Join-Path -Path $env:SystemRoot -ChildPath "py.exe")
    )
    $PythonLauncher = $candidateLaunchers | Where-Object { Test-Path -Path $_ } | Select-Object -First 1
}

if (-not $PythonLauncher) {
    throw "Unable to locate pyw.exe or py.exe. Specify -PythonLauncher with the full path."
}

$pythonLauncherPath = (Resolve-Path -Path $PythonLauncher).ProviderPath
$resolvedHandlerPath = (Resolve-Path -Path $handlerScript).ProviderPath
$command = ('"{0}" "{1}" "%1"' -f $pythonLauncherPath, $resolvedHandlerPath)

$registryBase = "HKCU:\Software\Classes\SystemFileAssociations\.pdf\shell\$VerbName"
New-Item -Path $registryBase -Force | Out-Null

Set-ItemProperty -Path $registryBase -Name "MUIVerb" -Value $VerbName
Set-ItemProperty -Path $registryBase -Name "MultiSelectModel" -Value "Player"

$iconSource = Join-Path -Path $env:SystemRoot -ChildPath "System32\imageres.dll"
if (Test-Path -Path $iconSource) {
    Set-ItemProperty -Path $registryBase -Name "Icon" -Value "$iconSource,15"
}

$commandKey = Join-Path -Path $registryBase -ChildPath "command"
New-Item -Path $commandKey -Force | Out-Null
$commandRegistryPath = "HKEY_CURRENT_USER\Software\Classes\SystemFileAssociations\.pdf\shell\$VerbName\command"
[Microsoft.Win32.Registry]::SetValue(
    $commandRegistryPath,
    "",
    $command,
    [Microsoft.Win32.RegistryValueKind]::String
)

Write-Output "Context menu entry '$VerbName' registered for PDF files."
Write-Output "Command: $command"
