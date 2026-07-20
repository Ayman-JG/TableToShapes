# Registers the TableToShapes COM add-in for the CURRENT USER (no admin needed).
# Run from PowerShell after building the solution (Debug|net48).
#
# RegAsm /register writes to HKCR (needs admin), so instead we use RegAsm /regfile
# to capture the entries and import them into HKCU\Software\Classes - the per-user
# COM hive, no elevation required. Then we add the PowerPoint Addins entry.

$ErrorActionPreference = 'Stop'

$dll = Join-Path $PSScriptRoot 'bin\Debug\net48\TableToShapes.AddIn.dll'
if (-not (Test-Path $dll)) { throw "Build first: $dll not found." }

# 64-bit Office needs the 64-bit RegAsm; change Framework64 to Framework for 32-bit Office.
$regasm = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe'

# 1. Generate the registration entries instead of writing HKCR directly.
$regFile = Join-Path $env:TEMP 'TableToShapes.AddIn.reg'
& $regasm $dll /codebase /regfile:$regFile
if ($LASTEXITCODE -ne 0) { throw "RegAsm failed with exit code $LASTEXITCODE" }

# 2. Retarget HKEY_CLASSES_ROOT -> HKCU\Software\Classes (per-user COM).
$content = Get-Content $regFile -Raw
$content = $content -replace '\[HKEY_CLASSES_ROOT\\', '[HKEY_CURRENT_USER\Software\Classes\'
Set-Content $regFile $content -Encoding Unicode

# 3. Import silently.
reg.exe import $regFile 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) { throw "reg import failed with exit code $LASTEXITCODE" }
Remove-Item $regFile

# 4. Tell PowerPoint to load the add-in at startup.
$key = 'HKCU:\Software\Microsoft\Office\PowerPoint\Addins\TableToShapes.AddIn.Connect'
New-Item -Path $key -Force | Out-Null
Set-ItemProperty -Path $key -Name 'FriendlyName' -Value 'Table to Shapes'
Set-ItemProperty -Path $key -Name 'Description'  -Value 'Converts a selected table into a visually identical group of shapes.'
Set-ItemProperty -Path $key -Name 'LoadBehavior' -Value 3 -Type DWord   # 3 = load at startup

Write-Host 'Add-in registered for current user. Restart PowerPoint; look for "Table to Shapes" on the Home tab.'
