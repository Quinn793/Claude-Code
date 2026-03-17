<#
.SYNOPSIS
    SharePoint Online Site Audit Script
    Exports all SPO sites with size, owners and members to CSV.

.DESCRIPTION
    Connects to SharePoint Online and retrieves all site collections (excluding
    Team Channel sites and Redirect sites). For each site it reports:
      - URL
      - Title
      - Template
      - Storage used (GB)
      - Site Owners (Site Admins)
      - Site Members (non-admin users)

.PARAMETER TenantAdminUrl
    The SharePoint admin center URL, e.g. https://contoso-admin.sharepoint.com

.PARAMETER OutputPath
    Full path of the output CSV file. Defaults to C:\Temp\SPO_SiteAudit.csv

.PARAMETER SkipTemplates
    Additional template prefixes to exclude (comma-separated).

.EXAMPLE
    .\SharePoint-SiteAudit.ps1 -TenantAdminUrl "https://contoso-admin.sharepoint.com"

.EXAMPLE
    .\SharePoint-SiteAudit.ps1 `
        -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
        -OutputPath "D:\Reports\SPO_Audit_$(Get-Date -f yyyyMMdd).csv"

.NOTES
    Required module: Microsoft.Online.SharePoint.PowerShell (PnP.PowerShell also works)
    Required role  : SharePoint Administrator or Global Administrator
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$TenantAdminUrl,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "C:\Temp\SPO_SiteAudit_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",

    [Parameter(Mandatory = $false)]
    [string[]]$SkipTemplates = @()
)

#region ── Helper functions ──────────────────────────────────────────────────

function Ensure-Module {
    param([string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        Write-Host "Module '$Name' not found. Installing from PSGallery..." -ForegroundColor Yellow
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $Name -ErrorAction Stop
}

function Connect-ToSPO {
    param([string]$AdminUrl)
    Write-Host "Connecting to SharePoint Online: $AdminUrl" -ForegroundColor Cyan
    try {
        Connect-SPOService -Url $AdminUrl -ErrorAction Stop
        Write-Host "Connected successfully." -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to connect to SPO: $_"
        exit 1
    }
}

function Get-SiteUsers {
    <#
    .SYNOPSIS
        Retrieves users for a single site, with error handling.
    #>
    param([string]$SiteUrl)
    try {
        Get-SPOUser -Site $SiteUrl -Limit All -ErrorAction Stop
    }
    catch {
        Write-Warning "Could not retrieve users for $SiteUrl : $_"
        return $null
    }
}

#endregion

#region ── Main ──────────────────────────────────────────────────────────────

# Ensure output directory exists
$outputDir = Split-Path -Path $OutputPath -Parent
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    Write-Host "Created output directory: $outputDir" -ForegroundColor Yellow
}

# Load the SPO module
Ensure-Module -Name "Microsoft.Online.SharePoint.PowerShell"

# Connect
Connect-ToSPO -AdminUrl $TenantAdminUrl

# Build the list of excluded template prefixes
$defaultSkip = @("TEAMCHANNEL", "RedirectSite", "SRCHCEN", "SPSMSITEHOST")
$allSkip     = ($defaultSkip + $SkipTemplates) | Select-Object -Unique

Write-Host "`nRetrieving all site collections..." -ForegroundColor Cyan
$allSites = Get-SPOSite -Limit All -IncludePersonalSite $false -ErrorAction Stop

# Filter out unwanted templates
$sites = $allSites | Where-Object {
    $template = $_.Template
    -not ($allSkip | Where-Object { $template -like "$_*" })
}

Write-Host "Found $($sites.Count) site(s) to process (after filtering)." -ForegroundColor Cyan

$totalSites  = $sites.Count
$currentIdx  = 0
$result      = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($site in $sites) {
    $currentIdx++
    $percent = [math]::Round(($currentIdx / $totalSites) * 100, 0)
    Write-Progress -Activity "Auditing SharePoint Sites" `
                   -Status "[$currentIdx / $totalSites] $($site.Url)" `
                   -PercentComplete $percent

    $sizeGB  = [math]::Round($site.StorageUsageCurrent / 1024, 2)
    $users   = Get-SiteUsers -SiteUrl $site.Url

    if ($users) {
        $owners  = ($users | Where-Object { $_.IsSiteAdmin }  | ForEach-Object { $_.LoginName }) -join "; "
        $members = ($users | Where-Object { -not $_.IsSiteAdmin } | ForEach-Object { $_.LoginName }) -join "; "
    }
    else {
        $owners  = "ERROR - could not retrieve"
        $members = "ERROR - could not retrieve"
    }

    $result.Add([PSCustomObject]@{
        Url              = $site.Url
        Title            = $site.Title
        Template         = $site.Template
        SizeGB           = $sizeGB
        StorageLimitGB   = [math]::Round($site.StorageQuota / 1024, 2)
        StorageUsedPct   = if ($site.StorageQuota -gt 0) {
                               [math]::Round(($site.StorageUsageCurrent / $site.StorageQuota) * 100, 1)
                           } else { 0 }
        LastContentModified = $site.LastContentModifiedDate
        Status           = $site.Status
        LockState        = $site.LockState
        SharingCapability = $site.SharingCapability
        Owners           = $owners
        Members          = $members
    })
}

Write-Progress -Activity "Auditing SharePoint Sites" -Completed

# Export to CSV
$result | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host "`n===== Audit Complete =====" -ForegroundColor Green
Write-Host "Sites processed : $totalSites"
Write-Host "Output file     : $OutputPath" -ForegroundColor Cyan

# Summary: top 10 largest sites
Write-Host "`nTop 10 largest sites (GB):" -ForegroundColor Yellow
$result | Sort-Object SizeGB -Descending | Select-Object -First 10 |
    Format-Table -AutoSize Title, SizeGB, StorageUsedPct, Owners

Disconnect-SPOService -ErrorAction SilentlyContinue

#endregion
