<#!
.SYNOPSIS
    HarjTech SharePoint Online Permission Audit Tool
.DESCRIPTION
    This PnP PowerShell script audits a SharePoint Online site and produces a CSV report summarising
    all files and folders with their permission inheritance status.  It counts how many files and
    folders exist, identifies those with broken inheritance (unique permissions), and detects
    documents or folders exposed to "Everyone" in the organisation.  It also summarises the
    number of files by type and lists document libraries that still inherit parent site permissions.

    Built and branded by HarjTech – your partner for secure collaboration and Microsoft 365
    optimisation.  Visit https://www.harjtech.com for consulting and support.

.PARAMETER SiteUrl
    The URL of the SharePoint Online site you want to assess.

.NOTES
    Author: HarjTech Solutions (contact@harjtech.com)
    Copyright © 2025 HarjTech. All rights reserved.

    This script requires the PnP.PowerShell module and SharePoint Online administrative
    permissions to access site and list information.

#>
param(
    [Parameter(Mandatory=$true, HelpMessage="Enter the URL of the SharePoint Online site to scan.")]
    [string]$SiteUrl
)

# Ensure PnP PowerShell module is available
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host "PnP.PowerShell module not found. Installing..." -ForegroundColor Yellow
    Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force
}

# Connect to SharePoint Online
Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive

# Prepare report file
$timestamp = (Get-Date).ToString('yyyyMMdd-HHmmss')
$reportFile = "SPPermissionsReport-$timestamp.csv"
$summaryFile = "SPPermissionsSummary-$timestamp.txt"

# Initialise counters and collections
$totalFiles = 0
$totalFolders = 0
$filesWithUniquePerm = 0
$foldersWithUniquePerm = 0
$fileTypeCounts = @{}
$reportData = @()
$docLibsNoCustomPerm = @()

# Retrieve document libraries (template 101) that are not hidden
$lists = Get-PnPList | Where-Object { $_.BaseTemplate -eq 101 -and -not $_.Hidden }

foreach ($list in $lists) {
    # Identify libraries that inherit permissions (no custom permissions)
    if (-not $list.HasUniqueRoleAssignments) {
        $docLibsNoCustomPerm += $list.Title
    }

    Write-Host "Scanning library: $($list.Title)" -ForegroundColor Green

    # Retrieve items with relevant properties. FileSystemObjectType: 0=File, 1=Folder【879391428515485†L318-L320】.
    $items = Get-PnPListItem -List $list -PageSize 500 -Fields "FileSystemObjectType", "FileLeafRef", "FileRef", "HasUniqueRoleAssignments" -ScriptBlock {
        param($batchItems)
        return $batchItems
    }

    foreach ($item in $items) {
        $type = $item.FileSystemObjectType
        $path = $item.FieldValues["FileRef"]
        $name = $item.FieldValues["FileLeafRef"]
        $hasUnique = $item.HasUniqueRoleAssignments
        $exposedToEveryone = $false
        $ext = ""

        if ($type -eq 0) { # File
            $totalFiles++
            if ($hasUnique) { $filesWithUniquePerm++ }
            $ext = [System.IO.Path]::GetExtension($name)
            if ([string]::IsNullOrEmpty($ext)) { $ext = "NoExtension" } else { $ext = $ext.TrimStart('.') }
            if ($fileTypeCounts.ContainsKey($ext)) {
                $fileTypeCounts[$ext] += 1
            } else {
                $fileTypeCounts[$ext] = 1
            }
        }
        elseif ($type -eq 1) { # Folder
            $totalFolders++
            if ($hasUnique) { $foldersWithUniquePerm++ }
        }

        # Determine if item is shared with "Everyone" or "Everyone except external users"
        try {
            $roleAssignments = Get-PnPProperty -ClientObject $item -Property RoleAssignments
            foreach ($ra in $roleAssignments) {
                $member = Get-PnPProperty -ClientObject $ra -Property Member
                if ($null -ne $member -and $member.Title -match "Everyone") {
                    $exposedToEveryone = $true
                    break
                }
            }
        } catch {
            # Ignore errors retrieving role assignments (insufficient rights may cause failure)
        }

        # Record information
        $reportData += [PSCustomObject]@{
            Library             = $list.Title
            ItemPath            = $path
            ItemName            = $name
            #ItemType            = (if ($type -eq 0) { "File" } elseif ($type -eq 1) { "Folder" } else { "Other" })
            ItemType = $type -eq 0 ? "File" : ($type -eq 1 ? "Folder" : "Other")
            UniquePermissions   = $hasUnique
            ExposedToEveryone   = $exposedToEveryone
            #FileType            = (if ($type -eq 0) { $ext } else { "" })
            FileType = $type -eq 0 ? $ext : ""
        }
    }
}

# Export detailed report to CSV
$reportData | Export-Csv -Path $reportFile -NoTypeInformation

# Build summary information and export to text file
$summaryLines = @()
$summaryLines += "SharePoint Online Permission Audit Summary"
$summaryLines += "Site: $SiteUrl"
$summaryLines += "Date: $(Get-Date)"
$summaryLines += "--------------------------------------------------"
$summaryLines += "Total files: $totalFiles"
$summaryLines += "Total folders: $totalFolders"
$summaryLines += "Files with unique permissions: $filesWithUniquePerm"
$summaryLines += "Folders with unique permissions: $foldersWithUniquePerm"
$summaryLines += ""
$summaryLines += "File type counts:"
foreach ($kvp in ($fileTypeCounts.GetEnumerator() | Sort-Object Name)) {
    $summaryLines += " - $($kvp.Name): $($kvp.Value)"
}
$summaryLines += ""
$summaryLines += "Document libraries inheriting permissions (no custom permissions):"
if ($docLibsNoCustomPerm.Count -gt 0) {
    foreach ($lib in $docLibsNoCustomPerm) {
        $summaryLines += " - $lib"
    }
} else {
    $summaryLines += " None"
}
$summaryLines += ""
$summaryLines += "Detailed report saved to: $reportFile"
$summaryLines | Out-File -FilePath $summaryFile -Encoding UTF8

Write-Host "Audit completed successfully." -ForegroundColor Cyan
Write-Host "Detailed report: $reportFile" -ForegroundColor Yellow
Write-Host "Summary report: $summaryFile" -ForegroundColor Yellow

# Marketing message
Write-Host "\nThis tool was built by HarjTech, a Microsoft solutions consultancy." -ForegroundColor Magenta
Write-Host "For advanced SharePoint governance, automation and analytics solutions, visit https://www.harjtech.com or contact us at info@harjtech.com." -ForegroundColor Magenta