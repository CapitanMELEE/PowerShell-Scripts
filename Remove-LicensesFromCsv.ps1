<#
.SYNOPSIS
  Remove a specific Microsoft 365 SKU from all users listed in a CSV.

.DESCRIPTION
  CSV must have at least one column for the user (UserPrincipalName/UPN/userPrincipalName).
  CSV should also include SkuId (GUID) OR supply -SkuIdOverride.
  Optional columns: DisplayName, SkuPartNumber.

  Works in Windows PowerShell 5.1 and PowerShell 7+.

.EXAMPLE
  .\Remove-LicensesFromCsv.ps1 -CsvPath .\Direct-ENTERPRISEPACK-Assignments.csv -WhatIf -Verbose

.EXAMPLE
  .\Remove-LicensesFromCsv.ps1 -CsvPath .\users.csv -SkuIdOverride 6fd2c87f-b296-42f0-b197-1e91e994b900
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
  [Parameter(Mandatory)]
  [string]$CsvPath,

  # If your CSV lacks SkuId or has multiple SkuIds, force the GUID here
  [Guid]$SkuIdOverride,

  # Stop on first failure
  [switch]$StopOnError
)

# ---- Helpers ----
function Get-ColVal {
  param(
    [Parameter(Mandatory)] $Row,
    [Parameter(Mandatory)][string[]] $Names
  )
  foreach ($n in $Names) {
    if ($Row.PSObject.Properties.Name -contains $n) {
      $v = $Row.$n
      if ($null -ne $v -and ("" + $v).Trim().Length -gt 0) { return $v }
    }
  }
  return $null
}

# ---- Pre-flight ----
if (-not (Test-Path -LiteralPath $CsvPath)) { throw "CSV not found at '$CsvPath'." }

try {
  Import-Module Microsoft.Graph.Users -ErrorAction Stop | Out-Null
} catch {
  throw "Microsoft Graph PowerShell SDK not found. Install with: Install-Module Microsoft.Graph -Scope CurrentUser"
}

$requiredScopes = @('User.ReadWrite.All','Organization.Read.All')
if (-not (Get-MgContext)) {
  Write-Verbose "Connecting to Microsoft Graph..."
  Connect-MgGraph -Scopes $requiredScopes | Out-Null
} else {
  $ctx = Get-MgContext
  $missing = @()
  if ($ctx -and $ctx.Scopes) {
    $missing = $requiredScopes | Where-Object { $_ -notin $ctx.Scopes }
  } else {
    $missing = $requiredScopes
  }
  if ($missing.Count -gt 0) {
    Write-Verbose ("Reconnecting to Graph to add missing scopes: {0}" -f ($missing -join ', '))
    Connect-MgGraph -Scopes $requiredScopes | Out-Null
  }
}

# ---- Read + normalize CSV ----
$rows = Import-Csv -LiteralPath $CsvPath
if (-not $rows -or $rows.Count -eq 0) { throw "CSV '$CsvPath' is empty." }

$items = foreach ($r in $rows) {
  $upn          = Get-ColVal -Row $r -Names @('UserPrincipalName','UPN','userPrincipalName')
  $skuIdStr     = Get-ColVal -Row $r -Names @('SkuId','skuId','SKUId')
  $displayName  = Get-ColVal -Row $r -Names @('DisplayName','displayName')
  $skuPart      = Get-ColVal -Row $r -Names @('SkuPartNumber','skuPartNumber','Sku')

  [pscustomobject]@{
    UserPrincipalName = $upn
    SkuId             = $skuIdStr
    SkuPartNumber     = $skuPart
    DisplayName       = $displayName
  }
}

if ( ($items | Where-Object { -not $_.UserPrincipalName }).Count -gt 0 ) {
  throw "CSV must contain a 'UserPrincipalName' (or UPN/userPrincipalName) column with values."
}

# ---- Figure out which SKU to remove ----
$skuGuid = $null
if ($PSBoundParameters.ContainsKey('SkuIdOverride')) {
  $skuGuid = [Guid]$SkuIdOverride
} else {
  $distinctSkus = ($items | Where-Object { $_.SkuId } | Select-Object -ExpandProperty SkuId -Unique)
  if ($null -eq $distinctSkus) { $distinctSkus = @() }
  if ($distinctSkus.Count -eq 0) {
    throw "No 'SkuId' in CSV and no -SkuIdOverride provided."
  } elseif ($distinctSkus.Count -gt 1) {
    Write-Warning ("Multiple SkuId values found in CSV: {0}" -f ($distinctSkus -join ', '))
    throw "Provide -SkuIdOverride to choose which SKU to remove."
  } else {
    try { $skuGuid = [Guid]$distinctSkus[0] }
    catch { throw "SKU in CSV is not a valid GUID: '$($distinctSkus[0])'." }
  }
}

Write-Host ("`n>> Will attempt to REMOVE SKU: {0} from {1} user(s)." -f $skuGuid, $items.Count) -ForegroundColor Yellow
Write-Host ">> Tip: add -WhatIf to preview, -Verbose for details." -ForegroundColor DarkGray

# ---- Work lists ----
$success = New-Object System.Collections.Generic.List[object]
$failed  = New-Object System.Collections.Generic.List[object]

$i = 0
foreach ($it in $items) {
  $i++
  $upn = $it.UserPrincipalName
  $label = if ($it.DisplayName) { ("{0} <{1}>" -f $it.DisplayName, $upn) } else { $upn }

  if ($PSCmdlet.ShouldProcess(($label), ("Remove license {0}" -f $skuGuid))) {

    $maxAttempts = 4
    $delaySec = 2
    $attempt = 0
    $done = $false

    while (-not $done -and $attempt -lt $maxAttempts) {
      $attempt++
      try {
        # Remove only; AddLicenses must be supplied but can be empty
        Set-MgUserLicense -UserId $upn -RemoveLicenses @([Guid]$skuGuid) -AddLicenses @() -ErrorAction Stop

        Write-Verbose ("[{0}/{1}] Removed {2} from {3}" -f $i, $items.Count, $skuGuid, $upn)
        $success.Add([pscustomobject]@{ UserPrincipalName=$upn; SkuId=$skuGuid; Status='Removed'; Attempt=$attempt })
        $done = $true

      } catch {
        $msg = $_.Exception.Message
        $status = $null
        try { $status = $_.Exception.ResponseStatusCode } catch {}

        $isThrottle = ($msg -match 'Too Many Requests|throttl') -or ($status -eq 429)
        $isConflict = ($msg -match 'conflict') -or ($status -eq 409)

        if ($isThrottle -and $attempt -lt $maxAttempts) {
          Write-Warning ("Throttle on {0} (attempt {1}/{2}). Sleeping {3}s…" -f $upn, $attempt, $maxAttempts, $delaySec)
          Start-Sleep -Seconds $delaySec
          $delaySec = [Math]::Min($delaySec * 2, 30)
        } elseif ($isConflict -and $attempt -eq 1) {
          # Treat as idempotent no-op if license not currently held
          Write-Verbose ("[{0}] {1} conflict/no-op." -f $i, $upn)
          $success.Add([pscustomobject]@{ UserPrincipalName=$upn; SkuId=$skuGuid; Status='NoChange'; Attempt=$attempt })
          $done = $true
        } else {
          Write-Error ("[{0}] Failed for {1}: {2}" -f $i, $upn, $msg)
          $failed.Add([pscustomobject]@{ UserPrincipalName=$upn; SkuId=$skuGuid; Status='Failed'; Error=$msg })
          if ($StopOnError) { throw }
          $done = $true
        }
      }
    }
  }
}

# ---- Summary ----
Write-Host ""
Write-Host "================= SUMMARY =================" -ForegroundColor Cyan
Write-Host ("Removed : {0}" -f ($success | Where-Object { $_.Status -eq 'Removed' }).Count)
Write-Host ("No change: {0}" -f ($success | Where-Object { $_.Status -eq 'NoChange' }).Count)
Write-Host ("Failed  : {0}" -f $failed.Count)
Write-Host "===========================================" -ForegroundColor Cyan

# CSV outputs
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$okPath = "Removed-$($skuGuid)-$timestamp.csv"
$errPath = "Failed-$($skuGuid)-$timestamp.csv"

if ($success.Count -gt 0) {
  $success | Export-Csv -NoTypeInformation -Encoding UTF8 $okPath
  Write-Host ("Wrote successes to '{0}'" -f $okPath)
}
if ($failed.Count -gt 0) {
  $failed | Export-Csv -NoTypeInformation -Encoding UTF8 $errPath
  Write-Host ("Wrote failures to '{0}'" -f $errPath)
}

