# Registers the TableToShapes COM add-in for the CURRENT USER (no admin needed).
# Run from a Developer PowerShell after building the solution (Debug|net48).
#
# Two things must be registered:
#  1. The managed class as a COM server (RegAsm /codebase) so PowerPoint can create it.
#  2. The add-in entry under HKCU\...\PowerPoint\Addins so PowerPoint knows to load it.

$ErrorActionPreference = 'Stop'

$dll = Join-Path $PSScriptRoot 'bin\Debug\net48\TableToShapes.AddIn.dll'
if (-not (Test-Path $dll)) { throw "Build first: $dll not found." }

# 64-bit Office needs the 64-bit RegAsm; adjust to Framework\ for 32-bit Office.
$regasm = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe'

& $regasm $dll /codebase /register
if ($LASTEXITCODE -ne 0) { throw "RegAsm failed with exit code $LASTEXITCODE" }

$key = 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\TableToShapes.AddIn.Connect'
New-Item -Path $key -Force | Out-Null
Set-ItemProperty -Path $key -Name 'FriendlyName' -Value 'Table to Shapes'
Set-ItemProperty -Path $key -Name 'Description'  -Value 'Converts a selected table into a visually identical group of shapes.'
Set-ItemProperty -Path $key -Name 'LoadBehavior' -Value 3 -Type DWord   # 3 = load at startup

Write-Host 'Add-in registered. Restart PowerPoint; look for "Table to Shapes" on the Home tab.'
