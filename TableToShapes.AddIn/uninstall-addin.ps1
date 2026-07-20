# Removes the TableToShapes COM add-in registration for the current user.

$ErrorActionPreference = 'Stop'

$dll = Join-Path $PSScriptRoot 'bin\Debug\net48\TableToShapes.AddIn.dll'
$regasm = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe'

if (Test-Path $dll) {
    & $regasm $dll /unregister
}

$key = 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\TableToShapes.AddIn.Connect'
if (Test-Path $key) {
    Remove-Item -Path $key -Recurse -Force
}

Write-Host 'Add-in unregistered.'
