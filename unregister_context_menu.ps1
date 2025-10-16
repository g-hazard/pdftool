param(
    [string]$VerbName = "Merge PDF"
)

$registryBase = "HKCU:\Software\Classes\SystemFileAssociations\.pdf\shell\$VerbName"

if (Test-Path -Path $registryBase) {
    Remove-Item -Path $registryBase -Recurse -Force
    Write-Output "Context menu entry '$VerbName' removed."
} else {
    Write-Output "No context menu entry named '$VerbName' was found."
}
