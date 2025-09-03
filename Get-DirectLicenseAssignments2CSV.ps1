<# 
Requires: Microsoft Graph PowerShell SDK
Install-Module Microsoft.Graph -Scope CurrentUser
#>

# ---- Config ----
$SkuPartNumber = " M365EDU_A3_STUUSEBNFT"   # e.g., ENTERPRISEPACK (M365 E3), SPE_E5, O365_BUSINESS, M365_E5, etc.
$Scopes = @(
  "User.Read.All",
  "Organization.Read.All"
)

# ---- Connect ----
if (-not (Get-MgContext)) {
  Connect-MgGraph -Scopes $Scopes | Out-Null
}

# ---- Resolve SKU to GUID ----
$sku = Get-MgSubscribedSku | Where-Object { $_.SkuPartNumber -eq $SkuPartNumber }
if (-not $sku) {
  throw "Could not find a subscribed SKU with SkuPartNumber '$SkuPartNumber'. Run: Get-MgSubscribedSku | Select SkuId,SkuPartNumber"
}
$SkuId = $sku.SkuId

# ---- Pull users with license assignment states ----
# LicenseAssignmentStates includes the source:
#  - AssignedByGroup = GUID of the group if assigned via group
#  - AssignedByGroup = $null if assigned directly
$users = Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,LicenseAssignmentStates

$directAssignments =
  foreach ($u in $users) {
    if ($u.LicenseAssignmentStates) {
      $match = $u.LicenseAssignmentStates |
        Where-Object { $_.SkuId -eq $SkuId -and [string]::IsNullOrEmpty($_.AssignedByGroup) }
      if ($match) {
        [pscustomobject]@{
          DisplayName        = $u.DisplayName
          UserPrincipalName  = $u.UserPrincipalName
          SkuPartNumber      = $SkuPartNumber
          SkuId              = $SkuId
        }
      }
    }
  }

# ---- Output ----
$directAssignments | Sort-Object DisplayName | Format-Table -AutoSize

# Optional: export to CSV
# $directAssignments | Sort-Object DisplayName | Export-Csv ".\Direct-$($SkuPartNumber)-Assignments.csv" -NoTypeInformation -Encoding UTF8

