# Registers the TableToShapes COM add-in for the CURRENT USER (no admin needed).
# Run from PowerShell after building the solution (Debug|net48).
#
# Writes the COM class registration directly to HKCU\Software\Classes (the per-user
# COM hive) instead of using RegAsm: RegAsm /register needs admin (writes HKCR), and
# its /regfile output proved unreliable when imported. A .NET COM class only needs
# a handful of well-known values; we set them explicitly.

$ErrorActionPreference = 'Stop'

$dll = Join-Path $PSScriptRoot 'bin\Debug\net48\TableToShapes.AddIn.dll'
if (-not (Test-Path $dll)) { throw "Build first: $dll not found." }

$clsid    = '{6F9B0C64-3C1A-4E6E-9B7D-2D3E8A11F0AB}'   # must match [Guid] on Connect
$progId   = 'TableToShapes.AddIn.Connect'                # must match [ProgId] on Connect
$assembly = 'TableToShapes.AddIn, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'
$codeBase = ([Uri]$dll).AbsoluteUri

# --- 1. COM class registration (per-user) ---
$clsidKey  = "HKCU:\Software\Classes\CLSID\$clsid"
$inprocKey = "$clsidKey\InprocServer32"

New-Item -Path $inprocKey -Force | Out-Null
New-Item -Path "$clsidKey\ProgId" -Force | Out-Null

Set-ItemProperty -Path $clsidKey -Name '(Default)' -Value $progId
Set-ItemProperty -Path "$clsidKey\ProgId" -Name '(Default)' -Value $progId

# mscoree.dll is the .NET Framework COM shim that loads the managed class.
Set-ItemProperty -Path $inprocKey -Name '(Default)'      -Value 'mscoree.dll'
Set-ItemProperty -Path $inprocKey -Name 'ThreadingModel' -Value 'Both'
Set-ItemProperty -Path $inprocKey -Name 'Class'          -Value 'TableToShapes.AddIn.Connect'
Set-ItemProperty -Path $inprocKey -Name 'Assembly'       -Value $assembly
Set-ItemProperty -Path $inprocKey -Name 'RuntimeVersion' -Value 'v4.0.30319'
Set-ItemProperty -Path $inprocKey -Name 'CodeBase'       -Value $codeBase

# ProgId -> CLSID lookup (PowerPoint resolves the add-in by ProgId).
$progIdKey = "HKCU:\Software\Classes\$progId"
New-Item -Path "$progIdKey\CLSID" -Force | Out-Null
Set-ItemProperty -Path $progIdKey -Name '(Default)' -Value $progId
Set-ItemProperty -Path "$progIdKey\CLSID" -Name '(Default)' -Value $clsid

# --- 2. Tell PowerPoint to load the add-in at startup ---
$addinKey = "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\$progId"
New-Item -Path $addinKey -Force | Out-Null
Set-ItemProperty -Path $addinKey -Name 'FriendlyName' -Value 'Table to Shapes'
Set-ItemProperty -Path $addinKey -Name 'Description'  -Value 'Converts a selected table into a visually identical group of shapes.'
Set-ItemProperty -Path $addinKey -Name 'LoadBehavior' -Value 3 -Type DWord   # 3 = load at startup

Write-Host "Registered $progId -> $codeBase"
Write-Host 'Restart PowerPoint; look for "Table to Shapes" on the Home tab.'
