# sign-dev.ps1 - DEVELOPMENT / TEST code signing only.
#
# Creates (once) a self-signed code-signing certificate in the current user's store and signs
# the three add-in assemblies with it. A self-signed certificate is trusted ONLY on machines
# where you import it (use -Trust below, or import it by hand). Do NOT use this for distributing
# to clients - production signing needs a real code-signing certificate (internal PKI, a public
# CA token, or a cloud service). See the README "Signing" note.
#
# Typical dev flow:   dotnet build   ->   .\sign-dev.ps1 -Trust   ->   .\install-addin.ps1

[CmdletBinding()]
param(
    [string]$Configuration = 'Debug',
    [string]$Subject       = 'CN=TableToShapes Dev Code Signing',
    [string]$TimestampUrl  = 'http://timestamp.digicert.com',
    [switch]$NoTimestamp,                 # skip timestamping (allows fully offline signing)
    [switch]$Trust                        # import the cert into LocalMachine Root + TrustedPublisher (needs admin)
)

$ErrorActionPreference = 'Stop'

# 1. Find or create the self-signed code-signing certificate.
$cert = Get-ChildItem Cert:\CurrentUser\My |
        Where-Object { $_.Subject -eq $Subject -and $_.HasPrivateKey -and $_.NotAfter -gt (Get-Date) } |
        Select-Object -First 1

if (-not $cert) {
    Write-Host "Creating self-signed code-signing certificate: $Subject"
    $cert = New-SelfSignedCertificate `
        -Type CodeSigningCert `
        -Subject $Subject `
        -CertStoreLocation Cert:\CurrentUser\My `
        -NotAfter (Get-Date).AddYears(3)
}
Write-Host "Using certificate thumbprint $($cert.Thumbprint)"

# 2. Locate signtool.exe (Windows 10/11 SDK).
$signtool = Get-ChildItem "${env:ProgramFiles(x86)}\Windows Kits\10\bin" -Recurse -Filter signtool.exe -ErrorAction SilentlyContinue |
            Where-Object { $_.FullName -match '\\x64\\' } |
            Sort-Object FullName -Descending | Select-Object -First 1
if (-not $signtool) { throw "signtool.exe not found. Install the Windows 10/11 SDK." }

# 3. Resolve the built assemblies.
$binDir = Join-Path $PSScriptRoot "bin\$Configuration\net48"
$targets = @(
    'TableToShapes.AddIn.dll',
    'TableToShapes.Interop.dll',
    'TableToShapes.Core.dll'
) | ForEach-Object { Join-Path $binDir $_ }
foreach ($t in $targets) { if (-not (Test-Path $t)) { throw "Build first: $t not found." } }

# 4. (Optional) trust the cert on THIS machine so Windows/Office accept the signature.
if ($Trust) {
    Write-Host "Importing certificate into LocalMachine Root + TrustedPublisher (requires admin)."
    foreach ($storeName in 'Root', 'TrustedPublisher') {
        $store = Get-Item "Cert:\LocalMachine\$storeName"
        $store.Open('ReadWrite'); $store.Add($cert); $store.Close()
    }
}

# 5. Sign.
$signArgs = @('sign', '/sha1', $cert.Thumbprint, '/fd', 'SHA256')
if (-not $NoTimestamp) { $signArgs += @('/tr', $TimestampUrl, '/td', 'SHA256') }
$signArgs += $targets

& $signtool.FullName @signArgs
if ($LASTEXITCODE -ne 0) { throw "signtool failed with exit code $LASTEXITCODE" }

# 6. Verify (with the default policy). Reports an untrusted chain unless the cert was trusted.
& $signtool.FullName verify /pa /v $targets[0]
if ($LASTEXITCODE -ne 0) {
    Write-Warning "verify /pa reports an untrusted chain - expected for a self-signed cert unless you ran with -Trust (or imported it manually)."
}

Write-Host "Signed: $($targets -join ', ')"
