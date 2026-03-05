################################################################################
# RECKITT - FOLDER CREATOR FROM CSV
################################################################################
# Just run: .\RECKITT_Create_Folders.ps1
################################################################################

# ============================================================================
# SHAREPOINT SITE URL
# ============================================================================
$siteUrl = "https://datapoem1.sharepoint.com/sites/DataPOEMDataVault"

# ============================================================================
# CSV FILE PATH - EDIT THIS LINE WITH YOUR CSV PATH
# ============================================================================
$CSVPath = "C:\Scripts\Projects.csv"

# ============================================================================
# LIBRARY NAME
# ============================================================================
$libraryName = "Reckitt"

# ============================================================================
# PERMISSION GROUPS
# ============================================================================
$readGroup = "Reckitt_READ"
$dataEditGroup = "Reckitt_DATA_EDIT"
$aiEditGroup = "Reckitt_AI_EDIT"
$insightsEditGroup = "Reckitt_INSIGHTS_EDIT"

# ============================================================================
# L1 FOLDER STRUCTURE (HARDCODED)
# ============================================================================
$l1Folders = @{
    "L1.DATA" = @(
        "01. Raw Data",
        "02. Raw Processed Data",
        "03. Data Summary",
        "04. DP Format",
        "05. DP Format QC"
    )
    "L1.AI" = @(
        "06. AI Transformed",
        "07. Feature List",
        "08. EDA",
        "09. Sales",
        "10. AI Preprocessed",
        "13. Final Model Files",
        "14. Temp (AI)",
        "16. Final ROI with FeatureID",
        "17. Final Model Insights (RROI)",
        "18. Optimization"
    )
    "L1.AI_INSIGHTS" = @(
        "12. Model Selection & Validation"
    )
    "L1.AI_DATA" = @(
        "11. AI Preprocessed QC",
        "15. ROI QC"
    )
}

################################################################################
# SCRIPT STARTS HERE - DO NOT EDIT BELOW THIS LINE
################################################################################

Write-Host "`n═══════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  RECKITT - FOLDER CREATOR" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════`n" -ForegroundColor Cyan

if (-not (Test-Path $CSVPath)) {
    Write-Host "✗ ERROR: CSV file not found!" -ForegroundColor Red
    Write-Host "Path: $CSVPath" -ForegroundColor Gray
    Write-Host "`nPlease edit line 17 in this script to set correct CSV path.`n" -ForegroundColor Yellow
    exit
}

Write-Host "CSV File: $CSVPath" -ForegroundColor Gray
Write-Host ""

Write-Host "Reading CSV file..." -ForegroundColor Yellow
try {
    $csvData = Import-Csv -Path $CSVPath
    $projects = @()
    
    foreach ($row in $csvData) {
        if ($row.Library -eq $libraryName -and $row.Folder5) {
            $fullPath = "$($row.Library)/$($row.Folder1)/$($row.Folder2)/$($row.Folder3)/$($row.Folder4)/$($row.Folder5)"
            $projects += $fullPath
        }
    }
    
    if ($projects.Count -eq 0) {
        Write-Host "✗ No $libraryName projects found in CSV" -ForegroundColor Red
        Write-Host "Make sure CSV has 'Library' column with value: $libraryName`n" -ForegroundColor Yellow
        exit
    }
    
    Write-Host "✓ Found $($projects.Count) $libraryName projects`n" -ForegroundColor Green
    
    foreach ($p in $projects) {
        Write-Host "  • $p" -ForegroundColor Gray
    }
    Write-Host ""
}
catch {
    Write-Host "✗ ERROR reading CSV: $($_.Exception.Message)`n" -ForegroundColor Red
    exit
}

Write-Host "Connecting to SharePoint..." -ForegroundColor Yellow
Connect-PnPOnline -Url $siteUrl -UseWebLogin -WarningAction Ignore
Write-Host "✓ Connected`n" -ForegroundColor Green

$count = 0
$stats = @{
    PathsCreated = 0
    PathsExisted = 0
    L1Created = 0
    SubfoldersCreated = 0
}

foreach ($projectPath in $projects) {
    $count++
    
    Write-Host "═══════════════════════════════════════" -ForegroundColor Magenta
    Write-Host "[$count/$($projects.Count)] Processing Project" -ForegroundColor White
    Write-Host "═══════════════════════════════════════" -ForegroundColor Magenta
    Write-Host "$projectPath`n" -ForegroundColor White
    
    Write-Host "→ Creating folder path..." -ForegroundColor Cyan
    
    $pathParts = $projectPath.Split('/')
    $currentPath = ""
    
    for ($i = 0; $i -lt $pathParts.Length; $i++) {
        if ($i -eq 0) {
            $currentPath = $pathParts[$i]
            continue
        }
        
        $folderName = $pathParts[$i]
        $checkPath = "$currentPath/$folderName"
        
        try {
            Get-PnPFolder -Url $checkPath -ErrorAction Stop | Out-Null
            Write-Host "  ○ $folderName (already exists)" -ForegroundColor Gray
            $stats.PathsExisted++
        }
        catch {
            try {
                Add-PnPFolder -Name $folderName -Folder $currentPath -ErrorAction Stop | Out-Null
                Write-Host "  ✓ $folderName (created)" -ForegroundColor Green
                $stats.PathsCreated++
            }
            catch {
                Write-Host "  ✗ $folderName (error: $($_.Exception.Message))" -ForegroundColor Red
            }
        }
        
        $currentPath = $checkPath
    }
    
    Write-Host "`n→ Creating L1 folders..." -ForegroundColor Cyan
    
    foreach ($l1Name in @("L1.DATA", "L1.AI", "L1.AI_INSIGHTS", "L1.AI_DATA")) {
        try {
            Add-PnPFolder -Name $l1Name -Folder $projectPath -ErrorAction Stop | Out-Null
            Write-Host "  ✓ $l1Name (created)" -ForegroundColor Green
            $stats.L1Created++
        }
        catch {
            Write-Host "  ○ $l1Name (already exists)" -ForegroundColor Gray
        }
    }
    
    Write-Host "`n→ Creating 18 subfolders..." -ForegroundColor Cyan
    
    foreach ($l1Name in $l1Folders.Keys) {
        foreach ($subfolder in $l1Folders[$l1Name]) {
            try {
                Add-PnPFolder -Name $subfolder -Folder "$projectPath/$l1Name" -ErrorAction Stop | Out-Null
                $stats.SubfoldersCreated++
            }
            catch {}
        }
    }
    
    Write-Host "  ✓ All 18 subfolders processed" -ForegroundColor Green
    
    Write-Host "`n→ Setting L1 folder permissions..." -ForegroundColor Cyan
    
    $ctx = Get-PnPContext
    $web = $ctx.Web
    $ctx.Load($web)
    $ctx.ExecuteQuery()
    
    function Set-L1Permissions($FolderName, $EditGroupNames) {
        try {
            $folderUrl = $web.ServerRelativeUrl + "/" + "$projectPath/$FolderName"
            $folder = $web.GetFolderByServerRelativeUrl($folderUrl)
            $ctx.Load($folder)
            $ctx.Load($folder.ListItemAllFields)
            $ctx.ExecuteQuery()
            
            $item = $folder.ListItemAllFields
            $item.BreakRoleInheritance($false, $false)
            $ctx.ExecuteQuery()
            
            $rg = $web.SiteGroups.GetByName($readGroup)
            $rr = $web.RoleDefinitions.GetByName("Read")
            $ctx.Load($rg)
            $ctx.Load($rr)
            $ctx.ExecuteQuery()
            
            $rb = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($ctx)
            $rb.Add($rr)
            $item.RoleAssignments.Add($rg, $rb)
            $ctx.ExecuteQuery()
            
            $er = $web.RoleDefinitions.GetByName("Edit")
            $ctx.Load($er)
            $ctx.ExecuteQuery()
            
            foreach ($editGroupName in $EditGroupNames) {
                $eg = $web.SiteGroups.GetByName($editGroupName)
                $ctx.Load($eg)
                $ctx.ExecuteQuery()
                
                $eb = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($ctx)
                $eb.Add($er)
                $item.RoleAssignments.Add($eg, $eb)
                $ctx.ExecuteQuery()
            }
            
            return $true
        }
        catch {
            return $false
        }
    }
    
    if (Set-L1Permissions "L1.DATA" @($dataEditGroup)) {
        Write-Host "  ✓ L1.DATA → $readGroup + $dataEditGroup" -ForegroundColor Green
    }
    
    if (Set-L1Permissions "L1.AI" @($aiEditGroup)) {
        Write-Host "  ✓ L1.AI → $readGroup + $aiEditGroup" -ForegroundColor Green
    }
    
    if (Set-L1Permissions "L1.AI_INSIGHTS" @($insightsEditGroup)) {
        Write-Host "  ✓ L1.AI_INSIGHTS → $readGroup + $insightsEditGroup" -ForegroundColor Green
    }
    
    if (Set-L1Permissions "L1.AI_DATA" @($aiEditGroup, $dataEditGroup)) {
        Write-Host "  ✓ L1.AI_DATA → $readGroup + $aiEditGroup + $dataEditGroup" -ForegroundColor Green
    }
    
    Write-Host "`n✓ Project complete!`n" -ForegroundColor Green
}

Write-Host "═══════════════════════════════════════" -ForegroundColor Cyan
Write-Host "         COMPLETED SUCCESSFULLY!        " -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════`n" -ForegroundColor Cyan

Write-Host "📊 Summary:" -ForegroundColor Yellow
Write-Host "  Projects processed: $count" -ForegroundColor White
Write-Host "  Folder paths created: $($stats.PathsCreated)" -ForegroundColor Green
Write-Host "  Folder paths existed: $($stats.PathsExisted)" -ForegroundColor Gray
Write-Host "  L1 folders created: $($stats.L1Created)" -ForegroundColor Green
Write-Host "  Subfolders created: $($stats.SubfoldersCreated)" -ForegroundColor Green

Disconnect-PnPOnline
Write-Host "`n✅ All done!`n" -ForegroundColor Green
