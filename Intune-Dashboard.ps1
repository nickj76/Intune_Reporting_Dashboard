<#
Intune Dashboard (PowerShell + Microsoft Graph)
A custom-built Intune reporting dashboard written in PowerShell that uses Microsoft Graph to generate a modern, interactive HTML overview of all Intune‑managed devices.
This script was fully rewritten and extended to provide a single‑page, executive‑friendly dashboard while still offering deep, audit‑ready device inventory tables for administrators.

.Version 

3.3 - 14-04-2026
- Added "Windows Feature Update Status" section showing counts of devices on current vs outdated feature updates (24H2/25H2 vs 23H2/Win10).
- Added "Windows Quality Update Status" section comparing each device's OS version against the latest known build for their version family to identify devices missing recent quality updates. 
- Updated dashboard layout to include Windows update sections in the Overview page and added new sidebar navigation links for easier access to Windows version breakdown and update status pages.
- Expanded device inactivity section to identify devices that haven't synced in over 30 days, 60 days, or 90 days+, which may indicate decommissioned or lost devices still enrolled in Intune.

3.2 - 14-04-2026
- Added "Windows OS Version Breakdown" section with interactive charts showing distribution of Windows 10 vs 11 versions. Included export buttons to download CSVs for each Windows version group. 
- Updated Chart.js library to latest version for improved visuals and performance.
- Added "Android OS Version Breakdown" section with charts and exports for Android 10-16 versions.

3.1 - 13-04-2026
- Added export to JSON option for all tables, allowing administrators to download raw device data in addition to Excel/CSV/PDF. 
- Updated export buttons to include JSON format and adjusted styling for better visibility.
- Fixed issue with linux devices not showing in the overview cards.
- Added Windows & Android device counts to the overview cards.
- Added per-column search inputs for more granular filtering in large tables. Updated DataTables library to the latest version for improved performance and features.
- Organised dashboard sections into Overview, Security, and Inventory for better navigation. Added new cards for inactive devices and low storage to highlight potential issues.

3.0 - 13-04-2026
- Complete rewrite with new modern dashboard design, improved performance, and additional device insights. Added interactive charts, enhanced filtering, and export capabilities. 
- Updated to use latest Microsoft Graph API endpoints and optimized data retrieval with parallel processing. Refreshed UI with new color scheme, icons, and responsive layout for better user experience.

2.4 - 13-04-2026
- Added Unicode Emoji Icons
- Includes: macOS count, SerialNumber in all tables, Unicode emojis for OS icons, reverted Defender logic, suppress Graph welcome, All Devices section

.NOTES
Requires Powershell 7.0 or later and the PSWriteHTML module. 
#>

if (!(Get-Module -Name PSWriteHTML -ListAvailable)) { Install-Module -Name PSWriteHTML -Force -AllowClobber }
Import-Module PSWriteHTML

$OutputFolder = "C:\temp"; if (!(Test-Path $OutputFolder)) { mkdir -Path $OutputFolder }

# Use Invoke-MgGraphRequest (part of Microsoft.Graph.Authentication, already loaded) to avoid
# version conflicts with Microsoft.Graph.DeviceManagement sub-module.
$deviceUri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$top=999"
$deviceList = [System.Collections.Generic.List[object]]::new()
do {
    $response = Invoke-MgGraphRequest -Uri $deviceUri -Method GET -OutputType PSObject
    foreach ($d in $response.value) { $deviceList.Add($d) }
    $deviceUri = $response.'@odata.nextLink'
} while ($deviceUri)
$AllDevices     = $deviceList.ToArray()
$ManagedDevices = $AllDevices | Where-Object managedDeviceOwnerType -EQ 'company'

$WindowsDevices = $ManagedDevices | Where-Object OperatingSystem -EQ "Windows"
$Win10Devices   = $WindowsDevices | Where-Object { $_.osVersion -match "^10\.0\.1[0-9]{4}" }
$Win11_23H2     = $WindowsDevices | Where-Object { $_.osVersion -match "^10\.0\.22631" }
$Win11_24H2     = $WindowsDevices | Where-Object { $_.osVersion -match "^10\.0\.26100" }
$Win11_25H2     = $WindowsDevices | Where-Object { $_.osVersion -match "^10\.0\.262[0-9]{2}" }

# ── Feature Update analysis ──────────────────────────────────────────────────
$WinFUCurrentCount  = $Win11_24H2.Count + $Win11_25H2.Count   # 24H2 or 25H2 = current
$WinFUOutdated      = @($Win10Devices) + @($Win11_23H2)        # Win10 or 23H2 = needs FU

# ── Quality Update analysis (highest build per version family = 'latest') ────
$LatestWin10Build    = if ($Win10Devices.Count -gt 0)  { ($Win10Devices | Sort-Object { [version]$_.osVersion } -Descending | Select-Object -First 1).osVersion } else { $null }
$LatestWin11_23Build = if ($Win11_23H2.Count   -gt 0)  { ($Win11_23H2   | Sort-Object { [version]$_.osVersion } -Descending | Select-Object -First 1).osVersion } else { $null }
$LatestWin11_24Build = if ($Win11_24H2.Count   -gt 0)  { ($Win11_24H2   | Sort-Object { [version]$_.osVersion } -Descending | Select-Object -First 1).osVersion } else { $null }
$LatestWin11_25Build = if ($Win11_25H2.Count   -gt 0)  { ($Win11_25H2   | Sort-Object { [version]$_.osVersion } -Descending | Select-Object -First 1).osVersion } else { $null }
$LatestBuilds        = @($LatestWin10Build, $LatestWin11_23Build, $LatestWin11_24Build, $LatestWin11_25Build) | Where-Object { $_ }
$WinQUCurrent        = $WindowsDevices | Where-Object { $_.osVersion -in $LatestBuilds }
$WinQUOutdated       = $WindowsDevices | Where-Object { $_.osVersion -notin $LatestBuilds }
$UniqueOsVersionCount = ($WindowsDevices | Group-Object osVersion).Count
$OsVersionDist       = $WindowsDevices | Group-Object osVersion | Sort-Object Count -Descending | Select-Object -First 12
$OsVersionLabels     = @($OsVersionDist | ForEach-Object { $_.Name })  | ConvertTo-Json -Compress -AsArray
$OsVersionCounts     = @($OsVersionDist | ForEach-Object { [int]$_.Count }) | ConvertTo-Json -Compress -AsArray

$MacDevices     = $ManagedDevices | Where-Object OperatingSystem -EQ "macOS"
$iOSDevices     = $ManagedDevices | Where-Object OperatingSystem -EQ "iOS"
$AndroidDevices = $ManagedDevices | Where-Object OperatingSystem -EQ "Android"
$Android10      = $AndroidDevices | Where-Object { $_.osVersion -match "^10(\.|$)" }
$Android11      = $AndroidDevices | Where-Object { $_.osVersion -match "^11(\.|$)" }
$Android12      = $AndroidDevices | Where-Object { $_.osVersion -match "^12(\.|$)" }
$Android13      = $AndroidDevices | Where-Object { $_.osVersion -match "^13(\.|$)" }
$Android14      = $AndroidDevices | Where-Object { $_.osVersion -match "^14(\.|$)" }
$Android15      = $AndroidDevices | Where-Object { $_.osVersion -match "^15(\.|$)" }
$Android16      = $AndroidDevices | Where-Object { $_.osVersion -match "^16(\.|$)" }
$LinuxDevices   = $AllDevices | Where-Object { $_.operatingSystem -like "*linux*" -or $_.operatingSystem -like "*ubuntu*" -or $_.operatingSystem -like "*debian*" -or $_.operatingSystem -like "*fedora*" -or $_.operatingSystem -like "*rhel*" }
$LinuxUbuntu    = $LinuxDevices | Where-Object { $_.operatingSystem -like "*ubuntu*" }
$LinuxDebian    = $LinuxDevices | Where-Object { $_.operatingSystem -like "*debian*" }
$LinuxFedora    = $LinuxDevices | Where-Object { $_.operatingSystem -like "*fedora*" }
$LinuxRHEL      = $LinuxDevices | Where-Object { $_.operatingSystem -like "*rhel*" }
$LinuxOther     = $LinuxDevices | Where-Object { $_.operatingSystem -notlike "*ubuntu*" -and $_.operatingSystem -notlike "*debian*" -and $_.operatingSystem -notlike "*fedora*" -and $_.operatingSystem -notlike "*rhel*" }

$ProtectedBag   = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
$UnprotectedBag = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
$ManagedDevices | ForEach-Object -Parallel {
    if ($_.WindowsProtectionState.RealTimeProtectionEnabled -eq $false) { ($using:UnprotectedBag).Add($_) } else { ($using:ProtectedBag).Add($_) }
} -ThrottleLimit 20
$ProtectedDevices   = $ProtectedBag.ToArray()
$UnprotectedDevices = $UnprotectedBag.ToArray()

$CompliantDevices=$ManagedDevices|Where-Object ComplianceState -EQ "compliant"
$NoncompliantDevices=$ManagedDevices|Where-Object ComplianceState -EQ "noncompliant"
$ThirtyDaysAgo=(Get-Date).AddDays(-30)
$SixtyDaysAgo  =(Get-Date).AddDays(-60)
$NinetyDaysAgo =(Get-Date).AddDays(-90)
$InactiveDevices=$ManagedDevices|Where-Object{$_.LastSyncDateTime -lt $ThirtyDaysAgo}
$Inactive30_60  =$ManagedDevices|Where-Object{$_.LastSyncDateTime -lt $ThirtyDaysAgo  -and $_.LastSyncDateTime -ge $SixtyDaysAgo}
$Inactive60_90  =$ManagedDevices|Where-Object{$_.LastSyncDateTime -lt $SixtyDaysAgo   -and $_.LastSyncDateTime -ge $NinetyDaysAgo}
$Inactive90Plus =$ManagedDevices|Where-Object{$_.LastSyncDateTime -lt $NinetyDaysAgo}
$MinimumFreeSpace=100
$LowStorageDevices=$ManagedDevices|Where-Object{($_.FreeStorageSpaceInBytes/1GB)-lt $MinimumFreeSpace}
$EncryptedDevices=$ManagedDevices|Where-Object IsEncrypted -EQ $true
$UnecryptedDevices=$ManagedDevices|Where-Object IsEncrypted -EQ $false
$AzureRegisteredDevices=$ManagedDevices|Where-Object AzureAdRegistered -EQ $True
$AzureUnregisteredDevices=$ManagedDevices|Where-Object AzureAdRegistered -EQ $False
$CompanyDevices=$ManagedDevices
$PersonalDevices=$AllDevices|Where-Object managedDeviceOwnerType -EQ 'personal'

# ── Helper: convert a PS object array to a JSON-safe array of plain hashtables ──
function ConvertTo-JsonRows {
    param([object[]]$Rows, [string[]]$Props)
    if (-not $Rows -or $Rows.Count -eq 0) { return '[]' }
    $out = @(foreach ($r in $Rows) {
        $h = [ordered]@{}
        foreach ($p in $Props) { $h[$p] = if ($null -ne $r.$p) { "$($r.$p)" } else { "" } }
        $h
    })
    return ($out | ConvertTo-Json -Compress -Depth 3 -AsArray)
}

$allProps      = @("deviceName","serialNumber","userPrincipalName","operatingSystem","manufacturer","model","osVersion","complianceState","isEncrypted","lastSyncDateTime","enrolledDateTime")
$allPropsHdr   = @("Device Name","Serial Number","UPN","OS","Manufacturer","Model","OS Version","Compliance","Encrypted","Last Sync","Enrolled")
$overviewProps = @("deviceName","serialNumber","model","operatingSystem","userPrincipalName")
$overviewHdr   = @("Device Name","Serial Number","Model","OS","UPN")
$storageProps  = $allProps + @("freeSpaceGB")
$storageHdr    = $allPropsHdr + @("Free Space (GB)")

$AllDevicesRows      = $ManagedDevices | Select-Object *,@{n="freeSpaceGB";e={[math]::Round($_.freeStorageSpaceInBytes/1GB,2)}}
$LowStorageRows      = $LowStorageDevices | Select-Object *,@{n="freeSpaceGB";e={[math]::Round($_.freeStorageSpaceInBytes/1GB,2)}}

$jsonOverview    = ConvertTo-JsonRows -Rows $ManagedDevices          -Props $overviewProps
$jsonDefender    = ConvertTo-JsonRows -Rows $UnprotectedDevices       -Props $allProps
$jsonNoncompliant= ConvertTo-JsonRows -Rows $NoncompliantDevices      -Props $allProps
$jsonBitlocker   = ConvertTo-JsonRows -Rows $UnecryptedDevices        -Props $allProps
$jsonAzure       = ConvertTo-JsonRows -Rows $AzureUnregisteredDevices -Props $allProps
$jsonPersonal    = ConvertTo-JsonRows -Rows $PersonalDevices          -Props $allProps
$jsonInactive    = ConvertTo-JsonRows -Rows $InactiveDevices          -Props $allProps
$jsonInactive30  = ConvertTo-JsonRows -Rows $Inactive30_60             -Props $allProps
$jsonInactive60  = ConvertTo-JsonRows -Rows $Inactive60_90             -Props $allProps
$jsonInactive90  = ConvertTo-JsonRows -Rows $Inactive90Plus            -Props $allProps
$jsonStorage     = ConvertTo-JsonRows -Rows $LowStorageRows           -Props $storageProps
$jsonLinux       = ConvertTo-JsonRows -Rows $LinuxDevices              -Props $allProps
$jsonLinuxUbuntu = ConvertTo-JsonRows -Rows $LinuxUbuntu               -Props $allProps
$jsonLinuxDebian = ConvertTo-JsonRows -Rows $LinuxDebian               -Props $allProps
$jsonLinuxFedora = ConvertTo-JsonRows -Rows $LinuxFedora               -Props $allProps
$jsonLinuxRHEL   = ConvertTo-JsonRows -Rows $LinuxRHEL                 -Props $allProps
$jsonLinuxOther  = ConvertTo-JsonRows -Rows $LinuxOther                -Props $allProps
$jsonWindows     = ConvertTo-JsonRows -Rows $WindowsDevices            -Props $allProps
$jsonWin10       = ConvertTo-JsonRows -Rows $Win10Devices              -Props $allProps
$jsonWin11_23H2  = ConvertTo-JsonRows -Rows $Win11_23H2               -Props $allProps
$jsonWin11_24H2  = ConvertTo-JsonRows -Rows $Win11_24H2               -Props $allProps
$jsonWin11_25H2  = ConvertTo-JsonRows -Rows $Win11_25H2               -Props $allProps
$jsonAndroid     = ConvertTo-JsonRows -Rows $AndroidDevices            -Props $allProps
$jsonAndroid10   = ConvertTo-JsonRows -Rows $Android10                   -Props $allProps
$jsonAndroid11   = ConvertTo-JsonRows -Rows $Android11                   -Props $allProps
$jsonAndroid12   = ConvertTo-JsonRows -Rows $Android12                   -Props $allProps
$jsonAndroid13   = ConvertTo-JsonRows -Rows $Android13                   -Props $allProps
$jsonAndroid14   = ConvertTo-JsonRows -Rows $Android14                   -Props $allProps
$jsonAndroid15   = ConvertTo-JsonRows -Rows $Android15                   -Props $allProps
$jsonAndroid16   = ConvertTo-JsonRows -Rows $Android16                   -Props $allProps

$wuProps        = @("deviceName","serialNumber","userPrincipalName","osVersion","complianceState","lastSyncDateTime")
$wuHdr          = @("Device Name","Serial Number","UPN","OS Version","Compliance","Last Sync")
$jsonFUOutdated = ConvertTo-JsonRows -Rows $WinFUOutdated  -Props $wuProps
$jsonQUOutdated = ConvertTo-JsonRows -Rows $WinQUOutdated  -Props $wuProps

$generatedAt = Get-Date -Format "dd MMM yyyy HH:mm"

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>Intune Dashboard</title>
<link rel="stylesheet" href="https://cdn.datatables.net/1.13.8/css/jquery.dataTables.min.css"/>
<link rel="stylesheet" href="https://cdn.datatables.net/buttons/2.4.2/css/buttons.dataTables.min.css"/>
<link rel="stylesheet" href="https://cdn.datatables.net/searchbuilder/1.5.0/css/searchBuilder.dataTables.min.css"/>
<style>
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:'Segoe UI',system-ui,sans-serif;background:#0f1117;color:#e2e8f0;display:flex;min-height:100vh}

  /* ── Sidebar ── */
  nav{width:230px;min-width:230px;background:#161b27;padding:24px 0;display:flex;flex-direction:column;gap:4px;position:sticky;top:0;height:100vh;overflow-y:auto}
  nav .brand{padding:0 20px 24px;font-size:18px;font-weight:700;color:#fff;letter-spacing:.5px;border-bottom:1px solid #2a3042}
  nav .brand span{color:#3b82f6}
  nav a{display:flex;align-items:center;gap:10px;padding:10px 20px;color:#94a3b8;text-decoration:none;font-size:13.5px;border-left:3px solid transparent;transition:all .15s}
  nav a:hover,nav a.active{color:#e2e8f0;background:#1e2535;border-left-color:#3b82f6}
  nav .section-label{padding:16px 20px 6px;font-size:10.5px;text-transform:uppercase;letter-spacing:1.2px;color:#4b5563}

  /* ── Main ── */
  main{flex:1;padding:32px;overflow:auto}
  .page{display:none}.page.active{display:block}
  h1{font-size:22px;font-weight:700;margin-bottom:4px}
  .subtitle{font-size:13px;color:#64748b;margin-bottom:28px}

  /* ── Stat cards ── */
  .cards{display:grid;grid-template-columns:repeat(auto-fill,minmax(190px,1fr));gap:16px;margin-bottom:32px}
  .card{background:#161b27;border:1px solid #2a3042;border-radius:12px;padding:20px;display:flex;flex-direction:column;gap:8px;transition:transform .15s}
  .card:hover{transform:translateY(-2px)}
  .card .icon{font-size:26px}
  .card .num{font-size:32px;font-weight:800;color:#fff}
  .card .label{font-size:12px;color:#64748b;font-weight:500}
  .card.blue .num{color:#3b82f6}.card.purple .num{color:#a855f7}
  .card.orange .num{color:#f97316}.card.green .num{color:#22c55e}
  .card.yellow .num{color:#eab308}.card.red .num{color:#ef4444}
  .card.sky .num{color:#0ea5e9}.card.pink .num{color:#ec4899}
  .card-section{margin-bottom:28px}
  .card-section-title{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:1.2px;color:#4b5563;margin-bottom:12px;padding-bottom:6px;border-bottom:1px solid #2a3042}

  /* ── Tables ── */
  .table-wrap{background:#161b27;border:1px solid #2a3042;border-radius:12px;padding:24px;overflow-x:auto}
  table.dataTable{width:100%!important;border-collapse:collapse;font-size:13px}
  table.dataTable thead tr{background:#1e2535}
  table.dataTable thead th{color:#94a3b8;font-weight:600;padding:10px 14px;border-bottom:1px solid #2a3042;white-space:nowrap}
  table.dataTable tbody tr{border-bottom:1px solid #1e2535;transition:background .1s}
  table.dataTable tbody tr:hover{background:#1e2535}
  table.dataTable tbody td{padding:9px 14px;color:#cbd5e1}
  .dataTables_wrapper .dataTables_filter input{background:#1e2535;border:1px solid #2a3042;color:#e2e8f0;border-radius:6px;padding:5px 10px}
  .dataTables_wrapper .dataTables_length select{background:#1e2535;border:1px solid #2a3042;color:#e2e8f0;border-radius:6px;padding:4px 8px}
  .dataTables_wrapper .dataTables_info,.dataTables_wrapper .dataTables_paginate{color:#64748b;font-size:12px;margin-top:12px}
  .dataTables_wrapper .dataTables_paginate .paginate_button{color:#94a3b8!important;border-radius:6px;padding:4px 10px!important}
  .dataTables_wrapper .dataTables_paginate .paginate_button.current{background:#3b82f6!important;color:#fff!important;border:none!important}
  .dataTables_wrapper .dataTables_paginate .paginate_button:hover{background:#1e2535!important;color:#fff!important}

  .badge{display:inline-block;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:600}
  .badge-green{background:#14532d;color:#4ade80}.badge-red{background:#450a0a;color:#f87171}
  .badge-yellow{background:#422006;color:#fbbf24}.badge-blue{background:#172554;color:#60a5fa}

  /* ── Export buttons ── */
  .dt-buttons{margin-bottom:12px;display:flex;gap:8px;flex-wrap:wrap}
  .dt-button{background:#1e2535!important;border:1px solid #2a3042!important;color:#94a3b8!important;border-radius:6px!important;padding:5px 14px!important;font-size:12px!important;cursor:pointer;transition:all .15s}
  .dt-button:hover{background:#3b82f6!important;border-color:#3b82f6!important;color:#fff!important}
  .dt-button.buttons-excel{border-left-color:#22c55e!important}
  .dt-button.buttons-csv{border-left-color:#eab308!important}
  .dt-button.buttons-pdf{border-left-color:#ef4444!important}
  .dt-toolbar{display:flex;flex-direction:column;gap:8px;margin-bottom:12px}
  .dt-buttons{display:flex;gap:8px;flex-wrap:wrap}
  .dt-searchbuilder{display:block}
  /* ── SearchBuilder dark theme ── */
  .dtsb-searchBuilder{padding:12px 0 8px}
  .dtsb-group{background:#1e2535!important;border:1px solid #2a3042!important;border-radius:8px!important;padding:12px!important}
  .dtsb-logicContainer{display:flex;flex-direction:column;gap:4px;margin-right:10px}
  .dtsb-logic,.dtsb-add{background:#0f1117!important;border:1px solid #2a3042!important;color:#94a3b8!important;border-radius:6px!important;padding:4px 10px!important;font-size:12px!important;cursor:pointer}
  .dtsb-logic:hover,.dtsb-add:hover{background:#3b82f6!important;border-color:#3b82f6!important;color:#fff!important}
  .dtsb-criteria{display:flex;flex-wrap:wrap;align-items:center;gap:6px;margin-bottom:6px}
  .dtsb-data,.dtsb-condition,.dtsb-value{background:#0f1117!important;border:1px solid #2a3042!important;color:#e2e8f0!important;border-radius:6px!important;padding:5px 10px!important;font-size:12px!important}
  .dtsb-data:focus,.dtsb-condition:focus,.dtsb-value:focus{outline:none!important;border-color:#3b82f6!important}
  .dtsb-delete,.dtsb-clearAll{background:#450a0a!important;border:1px solid #7f1d1d!important;color:#f87171!important;border-radius:6px!important;padding:4px 10px!important;font-size:12px!important;cursor:pointer}
  .dtsb-delete:hover,.dtsb-clearAll:hover{background:#ef4444!important;border-color:#ef4444!important;color:#fff!important}
  .dtsb-title{font-size:13px;font-weight:600;color:#94a3b8}
  /* ── Per-column search inputs ── */
  table.dataTable thead tr:last-child th{padding:6px 8px!important;background:#1a2030}
  input.col-search{width:100%;background:#0f1117;border:1px solid #2a3042;color:#e2e8f0;border-radius:5px;padding:4px 8px;font-size:11px;box-sizing:border-box}
  input.col-search:focus{outline:none;border-color:#3b82f6}
  input.col-search::placeholder{color:#4b5563}

  /* ── Windows OS version breakdown ── */
  .card .export-btn{background:#3b82f6;border:none;color:#fff;border-radius:6px;padding:4px 12px;font-size:11px;cursor:pointer;margin-top:6px;align-self:flex-start;transition:background .15s}
  .card .export-btn:hover{background:#2563eb}
  .win-charts{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-bottom:28px}
  .chart-box{background:#161b27;border:1px solid #2a3042;border-radius:12px;padding:24px;position:relative;height:320px}
</style>
</head>
<body>
<nav>
  <div class="brand">Intune <span>Dashboard</span></div>
  <div class="section-label">Overview</div>
  <a href="javascript:void(0)" class="active" onclick="return showPage('overview',this)">📊 Summary</a>
  <a href="javascript:void(0)" onclick="return showPage('all',this)">🖥️ All Devices</a>
  <div class="section-label">Security</div>
  <a href="javascript:void(0)" onclick="return showPage('defender',this)">🛡️ Defender Disabled</a>
  <a href="javascript:void(0)" onclick="return showPage('noncompliant',this)">⚠️ Noncompliant</a>
  <a href="javascript:void(0)" onclick="return showPage('bitlocker',this)">🔓 BitLocker Disabled</a>
  <a href="javascript:void(0)" onclick="return showPage('winupdates',this)">🔄 Windows Updates</a>
  <div class="section-label">Inventory</div>
  <a href="javascript:void(0)" onclick="return showPage('azure',this)">☁️ Azure Unregistered</a>
  <a href="javascript:void(0)" onclick="return showPage('personal',this)">📱 Personal Devices</a>
  <a href="javascript:void(0)" onclick="return showPage('inactive',this)">⏳ Inactive Devices</a>
  <a href="javascript:void(0)" onclick="return showPage('storage',this)">💾 Low Storage</a>
  <a href="javascript:void(0)" onclick="return showPage('linux',this)">🐧 Linux Devices</a>
  <a href="javascript:void(0)" onclick="return showPage('windows',this)">🖥️ Windows Devices</a>
  <a href="javascript:void(0)" onclick="return showPage('android',this)">🤖 Android Devices</a>
</nav>

<main>
  <!-- OVERVIEW -->
  <div id="page-overview" class="page active">
    <h1>Reporting Overview</h1>
    <p class="subtitle">Corporate managed devices &mdash; Generated $generatedAt</p>
    <div class="card-section">
      <div class="card-section-title">Platforms</div>
      <div class="cards">
        <div class="card green"><div class="icon">✅</div><div class="num">$($ManagedDevices.Count)</div><div class="label">Total Corporate</div></div>
        <div class="card blue"><div class="icon">🖥️</div><div class="num">$($WindowsDevices.Count)</div><div class="label">Windows</div></div>
        <div class="card purple"><div class="icon">🍎</div><div class="num">$($MacDevices.Count)</div><div class="label">macOS</div></div>
        <div class="card orange"><div class="icon">🤖</div><div class="num">$($AndroidDevices.Count)</div><div class="label">Android</div></div>
        <div class="card sky"><div class="icon">📱</div><div class="num">$($iOSDevices.Count)</div><div class="label">iOS</div></div>
        <div class="card green"><div class="icon">🐧</div><div class="num">$($LinuxDevices.Count)</div><div class="label">Linux</div></div>
      </div>
    </div>
    <div class="card-section">
      <div class="card-section-title">Security</div>
      <div class="cards">
        <div class="card red"><div class="icon">🛡️</div><div class="num">$($UnprotectedDevices.Count)</div><div class="label">Defender Disabled</div></div>
        <div class="card yellow"><div class="icon">⚠️</div><div class="num">$($NoncompliantDevices.Count)</div><div class="label">Noncompliant</div></div>
        <div class="card red"><div class="icon">🔓</div><div class="num">$($UnecryptedDevices.Count)</div><div class="label">BitLocker Off</div></div>
      </div>
    </div>
    <div class="card-section">
      <div class="card-section-title">Windows Updates</div>
      <div class="cards">
        <div class="card green"><div class="icon">🔄</div><div class="num">$($WinFUCurrentCount)</div><div class="label">On Current Feature Update</div></div>
        <div class="card red"><div class="icon">⬆️</div><div class="num">$($WinFUOutdated.Count)</div><div class="label">Needs Feature Update</div></div>
        <div class="card green"><div class="icon">✅</div><div class="num">$($WinQUCurrent.Count)</div><div class="label">On Latest Quality Update</div></div>
        <div class="card yellow"><div class="icon">🕐</div><div class="num">$($WinQUOutdated.Count)</div><div class="label">Behind on Quality Update</div></div>
      </div>
    </div>
    <div class="card-section">
      <div class="card-section-title">Inventory</div>
      <div class="cards">
        <div class="card purple"><div class="icon">⏳</div><div class="num">$($InactiveDevices.Count)</div><div class="label">Inactive ≥30 days</div></div>
        <div class="card yellow"><div class="icon">💾</div><div class="num">$($LowStorageDevices.Count)</div><div class="label">Low Storage &lt;100GB</div></div>
        <div class="card pink"><div class="icon">📲</div><div class="num">$($PersonalDevices.Count)</div><div class="label">Personal Devices</div></div>
      </div>
    </div>
  </div>

  <!-- ALL DEVICES -->
  <div id="page-all" class="page">
    <h1>All Devices</h1><p class="subtitle">Full corporate device inventory</p>
    <div class="card-section">
      <div class="card-section-title">Platform Overview</div>
      <div class="cards">
        <div class="card green"><div class="icon">✅</div><div class="num">$($ManagedDevices.Count)</div><div class="label">Total Corporate</div></div>
        <div class="card blue"><div class="icon">🖥️</div><div class="num">$($WindowsDevices.Count)</div><div class="label">Windows</div></div>
        <div class="card purple"><div class="icon">🍎</div><div class="num">$($MacDevices.Count)</div><div class="label">macOS</div></div>
        <div class="card orange"><div class="icon">🤖</div><div class="num">$($AndroidDevices.Count)</div><div class="label">Android</div></div>
        <div class="card sky"><div class="icon">📱</div><div class="num">$($iOSDevices.Count)</div><div class="label">iOS</div></div>
        <div class="card green"><div class="icon">🐧</div><div class="num">$($LinuxDevices.Count)</div><div class="label">Linux</div></div>
      </div>
    </div>
    <div class="table-wrap"><table id="tbl-all" class="dataTable"></table></div>
  </div>

  <!-- DEFENDER -->
  <div id="page-defender" class="page">
    <h1>Defender Disabled</h1><p class="subtitle">Devices with real-time protection off</p>
    <div class="table-wrap"><table id="tbl-defender" class="dataTable"></table></div>
  </div>

  <!-- NONCOMPLIANT -->
  <div id="page-noncompliant" class="page">
    <h1>Noncompliant Devices</h1><p class="subtitle">Devices not meeting compliance policy</p>
    <div class="table-wrap"><table id="tbl-noncompliant" class="dataTable"></table></div>
  </div>

  <!-- BITLOCKER -->
  <div id="page-bitlocker" class="page">
    <h1>BitLocker Disabled</h1><p class="subtitle">Devices without drive encryption</p>
    <div class="table-wrap"><table id="tbl-bitlocker" class="dataTable"></table></div>
  </div>

  <!-- AZURE UNREGISTERED -->
  <div id="page-azure" class="page">
    <h1>Azure Unregistered</h1><p class="subtitle">Devices not registered in Entra ID</p>
    <div class="table-wrap"><table id="tbl-azure" class="dataTable"></table></div>
  </div>

  <!-- PERSONAL -->
  <div id="page-personal" class="page">
    <h1>Personal Devices</h1><p class="subtitle">Personally-owned enrolled devices</p>
    <div class="table-wrap"><table id="tbl-personal" class="dataTable"></table></div>
  </div>

  <!-- INACTIVE -->
  <div id="page-inactive" class="page">
    <h1>Inactive Devices</h1><p class="subtitle">Devices with no sync in the last 30+ days</p>
    <div class="card-section">
      <div class="card-section-title">Inactivity Breakdown</div>
      <div class="cards">
        <div class="card red"><div class="icon">⏳</div><div class="num">$($InactiveDevices.Count)</div><div class="label">Total Inactive</div></div>
        <div class="card yellow"><div class="num">$($Inactive30_60.Count)</div><div class="label">30–60 Days</div></div>
        <div class="card orange"><div class="num">$($Inactive60_90.Count)</div><div class="label">60–90 Days</div></div>
        <div class="card red"><div class="num">$($Inactive90Plus.Count)</div><div class="label">90+ Days</div></div>
      </div>
    </div>
    <div class="win-charts">
      <div class="chart-box"><canvas id="inactiveBarChart"></canvas></div>
      <div class="chart-box"><canvas id="inactiveDoughnutChart"></canvas></div>
    </div>
    <div class="table-wrap"><table id="tbl-inactive" class="dataTable"></table></div>
  </div>

  <!-- LOW STORAGE -->
  <div id="page-storage" class="page">
    <h1>Low Storage</h1><p class="subtitle">Devices with less than 100 GB free</p>
    <div class="table-wrap"><table id="tbl-storage" class="dataTable"></table></div>
  </div>

  <!-- LINUX -->
  <div id="page-linux" class="page">
    <h1>Linux Devices</h1><p class="subtitle">OS &amp; Distro View</p>
    <div class="card-section">
      <div class="card-section-title">Distro Breakdown</div>
      <div class="cards">
        <div class="card green"><div class="icon">🐧</div><div class="num">$($LinuxDevices.Count)</div><div class="label">Total Devices</div></div>
        <div class="card orange"><div class="num">$($LinuxUbuntu.Count)</div><div class="label">Ubuntu</div><button class="export-btn" onclick="exportVersionCSV('linuxubuntu')">⬇ Export</button></div>
        <div class="card blue"><div class="num">$($LinuxDebian.Count)</div><div class="label">Debian</div><button class="export-btn" onclick="exportVersionCSV('linuxdebian')">⬇ Export</button></div>
        <div class="card sky"><div class="num">$($LinuxFedora.Count)</div><div class="label">Fedora</div><button class="export-btn" onclick="exportVersionCSV('linuxfedora')">⬇ Export</button></div>
        <div class="card red"><div class="num">$($LinuxRHEL.Count)</div><div class="label">RHEL</div><button class="export-btn" onclick="exportVersionCSV('linuxrhel')">⬇ Export</button></div>
        <div class="card purple"><div class="num">$($LinuxOther.Count)</div><div class="label">Other</div><button class="export-btn" onclick="exportVersionCSV('linuxother')">⬇ Export</button></div>
      </div>
    </div>
    <div class="win-charts">
      <div class="chart-box"><canvas id="linuxBarChart"></canvas></div>
      <div class="chart-box"><canvas id="linuxDoughnutChart"></canvas></div>
    </div>
    <div class="table-wrap"><table id="tbl-linux" class="dataTable"></table></div>
  </div>

  <!-- WINDOWS -->
  <div id="page-windows" class="page">
    <h1>Windows Devices</h1><p class="subtitle">OS &amp; Version View</p>
    <div class="card-section">
      <div class="card-section-title">OS Version Breakdown</div>
      <div class="cards">
        <div class="card blue"><div class="icon">🖥️</div><div class="num">$($WindowsDevices.Count)</div><div class="label">Total Devices</div></div>
        <div class="card purple"><div class="num">$($Win10Devices.Count)</div><div class="label">Windows 10</div><button class="export-btn" onclick="exportVersionCSV('win10')">⬇ Export</button></div>
        <div class="card sky"><div class="num">$($Win11_23H2.Count)</div><div class="label">Windows 11 23H2</div><button class="export-btn" onclick="exportVersionCSV('win11_23h2')">⬇ Export</button></div>
        <div class="card green"><div class="num">$($Win11_24H2.Count)</div><div class="label">Windows 11 24H2</div><button class="export-btn" onclick="exportVersionCSV('win11_24h2')">⬇ Export</button></div>
        <div class="card orange"><div class="num">$($Win11_25H2.Count)</div><div class="label">Windows 11 25H2</div><button class="export-btn" onclick="exportVersionCSV('win11_25h2')">⬇ Export</button></div>
      </div>
    </div>
    <div class="win-charts">
      <div class="chart-box"><canvas id="winBarChart"></canvas></div>
      <div class="chart-box"><canvas id="winDoughnutChart"></canvas></div>
    </div>
    <div class="table-wrap"><table id="tbl-windows" class="dataTable"></table></div>
  </div>

  <!-- ANDROID -->
  <div id="page-android" class="page">
    <h1>Android Devices</h1><p class="subtitle">OS &amp; Version View</p>
    <div class="card-section">
      <div class="card-section-title">OS Version Breakdown</div>
      <div class="cards">
        <div class="card orange"><div class="icon">🤖</div><div class="num">$($AndroidDevices.Count)</div><div class="label">Total Devices</div></div>
        <div class="card blue"><div class="num">$($Android10.Count)</div><div class="label">Android 10</div><button class="export-btn" onclick="exportVersionCSV('android10')">⬇ Export</button></div>
        <div class="card sky"><div class="num">$($Android11.Count)</div><div class="label">Android 11</div><button class="export-btn" onclick="exportVersionCSV('android11')">⬇ Export</button></div>
        <div class="card green"><div class="num">$($Android12.Count)</div><div class="label">Android 12</div><button class="export-btn" onclick="exportVersionCSV('android12')">⬇ Export</button></div>
        <div class="card purple"><div class="num">$($Android13.Count)</div><div class="label">Android 13</div><button class="export-btn" onclick="exportVersionCSV('android13')">⬇ Export</button></div>
        <div class="card yellow"><div class="num">$($Android14.Count)</div><div class="label">Android 14</div><button class="export-btn" onclick="exportVersionCSV('android14')">⬇ Export</button></div>
        <div class="card red"><div class="num">$($Android15.Count)</div><div class="label">Android 15</div><button class="export-btn" onclick="exportVersionCSV('android15')">⬇ Export</button></div>
        <div class="card pink"><div class="num">$($Android16.Count)</div><div class="label">Android 16</div><button class="export-btn" onclick="exportVersionCSV('android16')">⬇ Export</button></div>
      </div>
    </div>
    <div class="win-charts">
      <div class="chart-box"><canvas id="androidBarChart"></canvas></div>
      <div class="chart-box"><canvas id="androidDoughnutChart"></canvas></div>
    </div>
    <div class="table-wrap"><table id="tbl-android" class="dataTable"></table></div>
  </div>

  <!-- WINDOWS UPDATES -->
  <div id="page-winupdates" class="page">
    <h1>Windows Update Status</h1>
    <p class="subtitle">Feature &amp; Quality Update compliance across all Windows devices</p>

    <div class="card-section">
      <div class="card-section-title">Feature Update Status</div>
      <div class="cards">
        <div class="card blue"><div class="icon">🖥️</div><div class="num">$($WindowsDevices.Count)</div><div class="label">Total Windows</div></div>
        <div class="card green"><div class="num">$($WinFUCurrentCount)</div><div class="label">On Current (24H2 / 25H2)</div></div>
        <div class="card red"><div class="num">$($WinFUOutdated.Count)</div><div class="label">Needs Feature Update</div></div>
      </div>
    </div>
    <div class="win-charts">
      <div class="chart-box"><canvas id="fuBarChart"></canvas></div>
      <div class="chart-box"><canvas id="fuDoughnutChart"></canvas></div>
    </div>
    <div class="card-section-title" style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:1.2px;color:#4b5563;margin-bottom:12px;padding-bottom:6px;border-bottom:1px solid #2a3042">Devices Needing Feature Update</div>
    <div class="table-wrap" style="margin-bottom:36px"><table id="tbl-fuoutdated" class="dataTable"></table></div>

    <div class="card-section">
      <div class="card-section-title">Quality Update Status</div>
      <div class="cards">
        <div class="card green"><div class="num">$($WinQUCurrent.Count)</div><div class="label">On Latest Build</div></div>
        <div class="card yellow"><div class="num">$($WinQUOutdated.Count)</div><div class="label">Behind on Quality Update</div></div>
        <div class="card purple"><div class="num">$($UniqueOsVersionCount)</div><div class="label">Unique Build Versions</div></div>
      </div>
    </div>
    <div class="win-charts">
      <div class="chart-box" style="height:420px"><canvas id="quBarChart"></canvas></div>
      <div class="chart-box"><canvas id="quDoughnutChart"></canvas></div>
    </div>
    <div class="card-section-title" style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:1.2px;color:#4b5563;margin-bottom:12px;padding-bottom:6px;border-bottom:1px solid #2a3042">Devices Behind on Quality Update</div>
    <div class="table-wrap"><table id="tbl-quoutdated" class="dataTable"></table></div>
  </div>
</main>

<script src="https://code.jquery.com/jquery-3.7.1.min.js"></script>
<script src="https://cdn.datatables.net/1.13.8/js/jquery.dataTables.min.js"></script>
<script src="https://cdn.datatables.net/buttons/2.4.2/js/dataTables.buttons.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/pdfmake.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdfmake/0.2.7/vfs_fonts.js"></script>
<script src="https://cdn.datatables.net/searchbuilder/1.5.0/js/dataTables.searchBuilder.min.js"></script>
<script src="https://cdn.datatables.net/buttons/2.4.2/js/buttons.html5.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.2/dist/chart.umd.min.js"></script>
<script>
const overviewHeaders = $($overviewHdr | ConvertTo-Json -Compress);
const stdHeaders      = $($allPropsHdr | ConvertTo-Json -Compress);
const storageHeaders  = $($storageHdr  | ConvertTo-Json -Compress);
const wuHeaders       = $($wuHdr       | ConvertTo-Json -Compress);

const data = {
  all:         $jsonOverview,
  defender:    $jsonDefender,
  noncompliant:$jsonNoncompliant,
  bitlocker:   $jsonBitlocker,
  azure:       $jsonAzure,
  personal:    $jsonPersonal,
  inactive:    $jsonInactive,
  storage:     $jsonStorage,
  linux:       $jsonLinux,
  windows:     $jsonWindows,
  android:     $jsonAndroid,
  fuoutdated:  $jsonFUOutdated,
  quoutdated:  $jsonQUOutdated,
  inactive30:  $jsonInactive30,
  inactive60:  $jsonInactive60,
  inactive90:  $jsonInactive90
};

const versionData = {
  win10:      $jsonWin10,
  win11_23h2: $jsonWin11_23H2,
  win11_24h2: $jsonWin11_24H2,
  win11_25h2: $jsonWin11_25H2,
  android10:  $jsonAndroid10,
  android11:  $jsonAndroid11,
  android12:  $jsonAndroid12,
  android13:  $jsonAndroid13,
  android14:  $jsonAndroid14,
  android15:  $jsonAndroid15,
  android16:  $jsonAndroid16,
  linuxubuntu: $jsonLinuxUbuntu,
  linuxdebian: $jsonLinuxDebian,
  linuxfedora: $jsonLinuxFedora,
  linuxrhel:   $jsonLinuxRHEL,
  linuxother:  $jsonLinuxOther
};
const versionNames = {
  win10:      'Windows10',
  win11_23h2: 'Windows11-23H2',
  win11_24h2: 'Windows11-24H2',
  win11_25h2: 'Windows11-25H2',
  android10:  'Android10',
  android11:  'Android11',
  android12:  'Android12',
  android13:  'Android13',
  android14:  'Android14',
  android15:  'Android15',
  android16:  'Android16',
  linuxubuntu: 'Linux-Ubuntu',
  linuxdebian: 'Linux-Debian',
  linuxfedora: 'Linux-Fedora',
  linuxrhel:   'Linux-RHEL',
  linuxother:  'Linux-Other'
};

const tableMap = {
  all:          {id:'tbl-all',         headers:overviewHeaders},
  defender:     {id:'tbl-defender',    headers:stdHeaders},
  noncompliant: {id:'tbl-noncompliant',headers:stdHeaders},
  bitlocker:    {id:'tbl-bitlocker',   headers:stdHeaders},
  azure:        {id:'tbl-azure',       headers:stdHeaders},
  personal:     {id:'tbl-personal',    headers:stdHeaders},
  inactive:     {id:'tbl-inactive',    headers:stdHeaders},
  storage:      {id:'tbl-storage',     headers:storageHeaders},
  linux:        {id:'tbl-linux',       headers:stdHeaders},
  windows:      {id:'tbl-windows',     headers:stdHeaders},
  android:      {id:'tbl-android',     headers:stdHeaders},
  fuoutdated:   {id:'tbl-fuoutdated',  headers:wuHeaders},
  quoutdated:   {id:'tbl-quoutdated',  headers:wuHeaders},
  inactive30:   {id:'tbl-inactive30',  headers:stdHeaders},
  inactive60:   {id:'tbl-inactive60',  headers:stdHeaders},
  inactive90:   {id:'tbl-inactive90',  headers:stdHeaders}
};

const initialized = {};

function buildTable(key) {
  if (initialized[key]) return;
  initialized[key] = true;
  const cfg = tableMap[key];
  const rows = data[key];
  const cols = cfg.headers.map((h,i) => ({ title: h, data: i.toString() }));
  const tableData = rows.map(r => Object.values(r));

  const dt = `$('#' + cfg.id).DataTable({
    data: tableData,
    columns: cols,
    pageLength: 50,
    order: [],
    dom: '<"dt-toolbar"<"dt-buttons"B><"dt-searchbuilder"Q>>frtip',
    buttons: [
      { extend: 'excel', text: '⬇ Excel', className: 'buttons-excel', title: cfg.id },
      { extend: 'csv',   text: '⬇ CSV',   className: 'buttons-csv',   title: cfg.id },
      { extend: 'pdf',   text: '⬇ PDF',   className: 'buttons-pdf',   title: cfg.id, orientation: 'landscape', pageSize: 'A3' }
    ],
    language: { search: 'Filter:' }
  });

  // Inject second header row with per-column search inputs (after DT creates thead)
  const tr = document.createElement('tr');
  cfg.headers.forEach(h => {
    const th = document.createElement('th');
    const inp = document.createElement('input');
    inp.type = 'text';
    inp.placeholder = h;
    inp.className = 'col-search';
    th.appendChild(inp);
    tr.appendChild(th);
  });
  document.getElementById(cfg.id).querySelector('thead').appendChild(tr);

  dt.columns().every(function(i) {
    const inp = document.querySelectorAll('#' + cfg.id + ' thead tr:last-child input')[i];
    if (inp) inp.addEventListener('keyup', () => { this.search(inp.value).draw(); });
  });
}

function showPage(name, el) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('nav a').forEach(a => a.classList.remove('active'));
  document.getElementById('page-' + name).classList.add('active');
  el.classList.add('active');
  if (name === 'winupdates') {
    buildTable('fuoutdated');
    buildTable('quoutdated');
    initUpdateCharts();
  } else if (name === 'inactive') {
    buildTable('inactive');
    initInactiveCharts();
  } else if (name !== 'overview') {
    buildTable(name);
  }
  if (name === 'windows') initWinCharts();
  if (name === 'android') initAndroidCharts();
  if (name === 'linux')   initLinuxCharts();
  return false;
}

let winChartsInited = false;
function initWinCharts() {
  if (winChartsInited) return;
  winChartsInited = true;
  const labels = ['Windows 10', 'Win 11 23H2', 'Win 11 24H2', 'Win 11 25H2'];
  const counts = [$($Win10Devices.Count), $($Win11_23H2.Count), $($Win11_24H2.Count), $($Win11_25H2.Count)];
  const colors = ['#a855f7', '#0ea5e9', '#22c55e', '#f97316'];
  const gridColor = '#2a3042';
  const tickColor = '#94a3b8';

  new Chart(document.getElementById('winBarChart'), {
    type: 'bar',
    data: { labels: labels, datasets: [{ data: counts, backgroundColor: colors, borderRadius: 6, borderSkipped: false }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { color: tickColor }, grid: { color: gridColor } },
        y: { ticks: { color: tickColor }, grid: { color: gridColor }, beginAtZero: true }
      }
    }
  });

  new Chart(document.getElementById('winDoughnutChart'), {
    type: 'doughnut',
    data: { labels: labels, datasets: [{ data: counts, backgroundColor: colors, borderColor: '#161b27', borderWidth: 3 }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: true, position: 'top', labels: { color: tickColor, font: { size: 12 }, usePointStyle: true, padding: 16 } }
      },
      cutout: '55%'
    }
  });
}

let androidChartsInited = false;
function initAndroidCharts() {
  if (androidChartsInited) return;
  androidChartsInited = true;
  const labels = ['Android 10', 'Android 11', 'Android 12', 'Android 13', 'Android 14', 'Android 15', 'Android 16'];
  const counts = [$($Android10.Count), $($Android11.Count), $($Android12.Count), $($Android13.Count), $($Android14.Count), $($Android15.Count), $($Android16.Count)];
  const colors = ['#3b82f6', '#0ea5e9', '#22c55e', '#a855f7', '#eab308', '#ef4444', '#ec4899'];
  const gridColor = '#2a3042';
  const tickColor = '#94a3b8';

  new Chart(document.getElementById('androidBarChart'), {
    type: 'bar',
    data: { labels: labels, datasets: [{ data: counts, backgroundColor: colors, borderRadius: 6, borderSkipped: false }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { color: tickColor }, grid: { color: gridColor } },
        y: { ticks: { color: tickColor }, grid: { color: gridColor }, beginAtZero: true }
      }
    }
  });

  new Chart(document.getElementById('androidDoughnutChart'), {
    type: 'doughnut',
    data: { labels: labels, datasets: [{ data: counts, backgroundColor: colors, borderColor: '#161b27', borderWidth: 3 }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: true, position: 'top', labels: { color: tickColor, font: { size: 12 }, usePointStyle: true, padding: 16 } }
      },
      cutout: '55%'
    }
  });
}

let updateChartsInited = false;
function initUpdateCharts() {
  if (updateChartsInited) return;
  updateChartsInited = true;
  const gridColor = '#2a3042';
  const tickColor = '#94a3b8';

  // Feature Update — version family distribution
  const fuLabels = ['Windows 10', 'Win 11 23H2', 'Win 11 24H2', 'Win 11 25H2'];
  const fuCounts = [$($Win10Devices.Count), $($Win11_23H2.Count), $($Win11_24H2.Count), $($Win11_25H2.Count)];
  const fuColors = ['#a855f7', '#0ea5e9', '#22c55e', '#f97316'];

  new Chart(document.getElementById('fuBarChart'), {
    type: 'bar',
    data: { labels: fuLabels, datasets: [{ data: fuCounts, backgroundColor: fuColors, borderRadius: 6, borderSkipped: false }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, title: { display: true, text: 'Feature Update Distribution', color: tickColor, font: { size: 13 } } },
      scales: {
        x: { ticks: { color: tickColor }, grid: { color: gridColor } },
        y: { ticks: { color: tickColor }, grid: { color: gridColor }, beginAtZero: true }
      }
    }
  });

  new Chart(document.getElementById('fuDoughnutChart'), {
    type: 'doughnut',
    data: { labels: fuLabels, datasets: [{ data: fuCounts, backgroundColor: fuColors, borderColor: '#161b27', borderWidth: 3 }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: true, position: 'top', labels: { color: tickColor, font: { size: 12 }, usePointStyle: true, padding: 16 } },
        title: { display: true, text: 'Feature Update Share', color: tickColor, font: { size: 13 } }
      },
      cutout: '55%'
    }
  });

  // Quality Update — build version distribution (top 12, horizontal bar)
  const quLabels = $OsVersionLabels;
  const quCounts = $OsVersionCounts;

  new Chart(document.getElementById('quBarChart'), {
    type: 'bar',
    data: { labels: quLabels, datasets: [{ data: quCounts, backgroundColor: '#3b82f6', borderRadius: 4 }] },
    options: {
      indexAxis: 'y',
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, title: { display: true, text: 'Build Version Distribution (Top 12)', color: tickColor, font: { size: 13 } } },
      scales: {
        x: { ticks: { color: tickColor }, grid: { color: gridColor }, beginAtZero: true },
        y: { ticks: { color: tickColor, font: { size: 11 } }, grid: { color: gridColor } }
      }
    }
  });

  // Quality Update compliance split
  new Chart(document.getElementById('quDoughnutChart'), {
    type: 'doughnut',
    data: {
      labels: ['On Latest Build', 'Behind on QU'],
      datasets: [{ data: [$($WinQUCurrent.Count), $($WinQUOutdated.Count)], backgroundColor: ['#22c55e', '#ef4444'], borderColor: '#161b27', borderWidth: 3 }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: true, position: 'top', labels: { color: tickColor, font: { size: 12 }, usePointStyle: true, padding: 16 } },
        title: { display: true, text: 'Quality Update Compliance', color: tickColor, font: { size: 13 } }
      },
      cutout: '55%'
    }
  });
}

let linuxChartsInited = false;
function initLinuxCharts() {
  if (linuxChartsInited) return;
  linuxChartsInited = true;
  const labels = ['Ubuntu', 'Debian', 'Fedora', 'RHEL', 'Other'];
  const counts = [$($LinuxUbuntu.Count), $($LinuxDebian.Count), $($LinuxFedora.Count), $($LinuxRHEL.Count), $($LinuxOther.Count)];
  const colors = ['#f97316', '#3b82f6', '#0ea5e9', '#ef4444', '#a855f7'];
  const gridColor = '#2a3042';
  const tickColor = '#94a3b8';

  new Chart(document.getElementById('linuxBarChart'), {
    type: 'bar',
    data: { labels: labels, datasets: [{ data: counts, backgroundColor: colors, borderRadius: 6, borderSkipped: false }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { color: tickColor }, grid: { color: gridColor } },
        y: { ticks: { color: tickColor }, grid: { color: gridColor }, beginAtZero: true }
      }
    }
  });

  new Chart(document.getElementById('linuxDoughnutChart'), {
    type: 'doughnut',
    data: { labels: labels, datasets: [{ data: counts, backgroundColor: colors, borderColor: '#161b27', borderWidth: 3 }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: true, position: 'top', labels: { color: tickColor, font: { size: 12 }, usePointStyle: true, padding: 16 } }
      },
      cutout: '55%'
    }
  });
}

let inactiveChartsInited = false;
function initInactiveCharts() {
  if (inactiveChartsInited) return;
  inactiveChartsInited = true;
  const labels = ['30–60 Days', '60–90 Days', '90+ Days'];
  const counts = [$($Inactive30_60.Count), $($Inactive60_90.Count), $($Inactive90Plus.Count)];
  const colors = ['#eab308', '#f97316', '#ef4444'];
  const gridColor = '#2a3042';
  const tickColor = '#94a3b8';

  new Chart(document.getElementById('inactiveBarChart'), {
    type: 'bar',
    data: { labels: labels, datasets: [{ data: counts, backgroundColor: colors, borderRadius: 6, borderSkipped: false }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: { ticks: { color: tickColor }, grid: { color: gridColor } },
        y: { ticks: { color: tickColor }, grid: { color: gridColor }, beginAtZero: true }
      }
    }
  });

  new Chart(document.getElementById('inactiveDoughnutChart'), {
    type: 'doughnut',
    data: { labels: labels, datasets: [{ data: counts, backgroundColor: colors, borderColor: '#161b27', borderWidth: 3 }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: true, position: 'top', labels: { color: tickColor, font: { size: 12 }, usePointStyle: true, padding: 16 } }
      },
      cutout: '55%'
    }
  });
}

function exportVersionCSV(ver) {
  const rows = versionData[ver];
  if (!rows || rows.length === 0) { alert('No data for this version.'); return; }
  const csvRows = [stdHeaders.join(',')];
  rows.forEach(r => {
    csvRows.push(Object.values(r).map(v => '"' + String(v).replace(/"/g, '""') + '"').join(','));
  });
  const blob = new Blob([csvRows.join('\n')], { type: 'text/csv' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = versionNames[ver] + '-Devices.csv';
  a.click();
  URL.revokeObjectURL(a.href);
}
</script>
</body>
</html>
"@

$htmlPath = "$OutputFolder\Intune-Dashboard.html"
$html | Out-File -FilePath $htmlPath -Encoding utf8 -Force
Start-Process $htmlPath
