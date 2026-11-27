[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$TargetPath,

    [Parameter()]
    [string]$CertificateBase64 = $env:WINDOWS_CERTIFICATE_BASE64,

    [Parameter()]
    [string]$CertificatePassword = $env:WINDOWS_CERTIFICATE_PASSWORD,

    [Parameter()]
    [string]$TimestampUrl = $(if ($env:WINDOWS_TIMESTAMP_URL) { $env:WINDOWS_TIMESTAMP_URL } else { "http://timestamp.digicert.com" })
)

$ErrorActionPreference = "Stop"
$resolvedTarget = (Resolve-Path -LiteralPath $TargetPath).Path
if ([IO.Path]::GetFileName($resolvedTarget) -ne "PGT.exe") {
    throw "Signing target must be the application PGT.exe, not a third-party binary."
}

$signTool = Get-Command signtool.exe -ErrorAction SilentlyContinue
if ($null -eq $signTool) {
    throw "signtool.exe was not found. Install the Windows SDK or add SignTool to PATH."
}

if (-not $PSCmdlet.ShouldProcess($resolvedTarget, "Sign and verify the application executable")) {
    Write-Host "WhatIf: would sign only $resolvedTarget and verify its Authenticode signature."
    return
}

if ([string]::IsNullOrWhiteSpace($CertificateBase64) -or [string]::IsNullOrWhiteSpace($CertificatePassword)) {
    throw "Both WINDOWS_CERTIFICATE_BASE64 and WINDOWS_CERTIFICATE_PASSWORD are required to sign."
}

$pfxPath = Join-Path ([IO.Path]::GetTempPath()) ("pgt-signing-" + [guid]::NewGuid().ToString("N") + ".pfx")
$thumbprint = $null
$certificate = $null
$securePassword = $null

try {
    [IO.File]::WriteAllBytes($pfxPath, [Convert]::FromBase64String($CertificateBase64))
    $securePassword = ConvertTo-SecureString $CertificatePassword -AsPlainText -Force
    $certificate = Import-PfxCertificate -FilePath $pfxPath -CertStoreLocation "Cert:\CurrentUser\My" -Password $securePassword
    $thumbprint = $certificate.Thumbprint
    if ([string]::IsNullOrWhiteSpace($thumbprint)) {
        throw "The imported certificate has no thumbprint."
    }

    & $signTool.Source sign /fd SHA256 /tr $TimestampUrl /td SHA256 /sha1 $thumbprint $resolvedTarget
    if ($LASTEXITCODE -ne 0) {
        throw "SignTool signing failed with exit code $LASTEXITCODE."
    }

    & $signTool.Source verify /pa /tw $resolvedTarget
    if ($LASTEXITCODE -ne 0) {
        throw "SignTool verification failed with exit code $LASTEXITCODE."
    }
    Write-Host "PGT.exe was signed and its Authenticode signature was verified."
}
finally {
    if ($thumbprint) {
        Remove-Item -LiteralPath "Cert:\CurrentUser\My\$thumbprint" -Force -ErrorAction SilentlyContinue
    }
    if ($certificate) {
        $certificate.Dispose()
    }
    if ($securePassword) {
        $securePassword.Dispose()
    }
    Remove-Item -LiteralPath $pfxPath -Force -ErrorAction SilentlyContinue
}
