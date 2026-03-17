Connect-SPOService -Url "https://JOUW-TENANT-admin.sharepoint.com"

Get-SPOSite -Limit All | ? { $_.Template -notlike "TEAMCHANNEL*" -and $_.Template -notlike "RedirectSite*" } | % {
    $u = Get-SPOUser -Site $_.Url -Limit All
    [PSCustomObject]@{
        Url     = $_.Url
        Title   = $_.Title
        SizeGB  = [math]::Round($_.StorageUsageCurrent/1024,2)
        Owners  = ($u | ? { $_.IsSiteAdmin }  | % { $_.LoginName }) -join "; "
        Members = ($u | ? { !$_.IsSiteAdmin } | % { $_.LoginName }) -join "; "
    }
} | Export-Csv C:\Temp\SPO_Audit.csv -NoTypeInformation -Encoding UTF8
