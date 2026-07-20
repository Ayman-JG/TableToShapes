# Removes the TableToShapes COM add-in registration for the current user.

$ErrorActionPreference = 'Stop'

# Remove per-user COM registration (CLSID + ProgId under HKCU\Software\Classes).
$clsid = '{6F9B0C64-3C1A-4E6E-9B7D-2D3E8A11F0AB}'
$paths = @(
    "HKCU:\Software\Classes\CLSID\$clsid",
    'HKCU:\Software\Classes\TableToShapes.AddIn.Connect',
    'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\TableToShapes.AddIn.Connect'
)

foreach ($path in $paths) {
    if (Test-Path $path) {
        Remove-Item -Path $path -Recurse -Force
        Write-Host "Removed $path"
    }
}

Write-Host 'Add-in unregistered.'
