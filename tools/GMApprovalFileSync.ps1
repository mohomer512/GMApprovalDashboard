param(
  [string]$SiteUrl = "http://spse26h/GM",
  [string]$RequestsListName = "GM Requests",
  [string]$DocumentsLibraryName = "GM Approval Documents",
  [string]$ShareRootPath = "C:\GMApprovalShare",
  [string]$ShareUncRoot = "\\SPSE26H\GMApprovalShare"
)

$ErrorActionPreference = "Stop"

$pendingPath = Join-Path $ShareRootPath "Pending"
$signedPath = Join-Path $ShareRootPath "Signed"
$archivePath = Join-Path $ShareRootPath "Archive"
$logsPath = Join-Path $ShareRootPath "Logs"
$statePath = Join-Path $logsPath "GMApprovalFileSync.state.json"
$generalLogPath = Join-Path $logsPath "GMApprovalFileSync.log"

foreach ($path in @($pendingPath, $signedPath, $archivePath, $logsPath)) {
  if (-not (Test-Path -LiteralPath $path)) {
    New-Item -Path $path -ItemType Directory -Force | Out-Null
  }
}

Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

function Write-GeneralLog {
  param([string]$Message)

  $line = "[{0}] {1}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $Message
  Add-Content -LiteralPath $generalLogPath -Value $line -Encoding UTF8
}

function Get-SafeFileName {
  param([string]$Value)

  if ([string]::IsNullOrWhiteSpace($Value)) {
    return "UNKNOWN"
  }

  return ($Value -replace '[\\/:*?"<>|]', "_")
}

function Write-RequestLog {
  param(
    [string]$RequestNo,
    [string]$Message
  )

  $safeRequestNo = Get-SafeFileName $RequestNo
  $logPath = Join-Path $logsPath ($safeRequestNo + ".log")
  $line = "[{0}] {1}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $Message
  Add-Content -LiteralPath $logPath -Value $line -Encoding UTF8
}

function Convert-StateToHashtable {
  param($JsonObject)

  $state = @{}

  if ($null -eq $JsonObject) {
    return $state
  }

  foreach ($itemProperty in $JsonObject.PSObject.Properties) {
    $itemState = @{}

    if ($null -ne $itemProperty.Value) {
      foreach ($fieldProperty in $itemProperty.Value.PSObject.Properties) {
        $itemState[$fieldProperty.Name] = [string]$fieldProperty.Value
      }
    }

    $state[$itemProperty.Name] = $itemState
  }

  return $state
}

function Get-State {
  if (-not (Test-Path -LiteralPath $statePath)) {
    return @{}
  }

  $raw = Get-Content -LiteralPath $statePath -Raw

  if ([string]::IsNullOrWhiteSpace($raw)) {
    return @{}
  }

  return Convert-StateToHashtable ($raw | ConvertFrom-Json)
}

function Save-State {
  param([hashtable]$State)

  $State | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $statePath -Encoding UTF8
}

function Get-FieldString {
  param(
    [Microsoft.SharePoint.SPListItem]$Item,
    [string]$FieldName
  )

  try {
    $value = $Item[$FieldName]
  } catch {
    return ""
  }

  if ($null -eq $value) {
    return ""
  }

  return [string]$value
}

function Get-UserFieldString {
  param(
    [Microsoft.SharePoint.SPListItem]$Item,
    [string]$FieldName
  )

  try {
    $value = $Item[$FieldName]
  } catch {
    return ""
  }

  if ($null -eq $value) {
    return ""
  }

  if ($value.LookupValue) {
    return [string]$value.LookupValue
  }

  return [string]$value
}

function Get-ItemSnapshot {
  param([Microsoft.SharePoint.SPListItem]$Item)

  return @{
    Id = [string]$Item.ID
    Title = Get-FieldString $Item "Title"
    RequestNo = Get-FieldString $Item "RequestNo"
    Status = Get-FieldString $Item "Status"
    OfficeManagerComment = Get-FieldString $Item "OfficeManagerComment"
    HODComment = Get-FieldString $Item "HODComment"
    GMComment = Get-FieldString $Item "GMComment"
    SharedFolderPath = Get-FieldString $Item "SharedFolderPath"
    PDFFileUrl = Get-FieldString $Item "PDFFileUrl"
    SignedPDFUrl = Get-FieldString $Item "SignedPDFUrl"
    Modified = ([datetime]$Item["Modified"]).ToString("o")
    ModifiedBy = Get-UserFieldString $Item "Editor"
  }
}

function Write-ChangeLog {
  param(
    [string]$RequestNo,
    [hashtable]$Previous,
    [hashtable]$Current
  )

  if ($null -eq $Previous) {
    Write-RequestLog $RequestNo ("Tracking started. Status={0}; ModifiedBy={1}" -f $Current.Status, $Current.ModifiedBy)
    return
  }

  $fieldsToLog = @(
    "Title",
    "Status",
    "OfficeManagerComment",
    "HODComment",
    "GMComment",
    "SharedFolderPath",
    "PDFFileUrl",
    "SignedPDFUrl"
  )

  foreach ($fieldName in $fieldsToLog) {
    $oldValue = [string]$Previous[$fieldName]
    $newValue = [string]$Current[$fieldName]

    if ($oldValue -ne $newValue) {
      Write-RequestLog $RequestNo ("{0} changed by {1}`r`nOLD: {2}`r`nNEW: {3}" -f $fieldName, $Current.ModifiedBy, $oldValue, $newValue)
    }
  }
}

function Get-RequestPdfFile {
  param(
    [Microsoft.SharePoint.SPList]$Library,
    [string]$RequestNo,
    [int]$RequestItemId
  )

  $safeRequestNo = [System.Security.SecurityElement]::Escape($RequestNo)

  $query = New-Object Microsoft.SharePoint.SPQuery
  $query.ViewAttributes = "Scope='RecursiveAll'"
  $query.RowLimit = 1
  $query.Query = @"
<Where>
  <Or>
    <Eq>
      <FieldRef Name='RequestItemId' />
      <Value Type='Number'>$RequestItemId</Value>
    </Eq>
    <Eq>
      <FieldRef Name='RequestNo' />
      <Value Type='Text'>$safeRequestNo</Value>
    </Eq>
  </Or>
</Where>
<OrderBy>
  <FieldRef Name='ID' Ascending='FALSE' />
</OrderBy>
"@

  $items = $Library.GetItems($query)

  if ($items.Count -gt 0 -and $null -ne $items[0].File) {
    return $items[0].File
  }

  return $null
}

function Copy-SPFileIfMissing {
  param(
    [Microsoft.SharePoint.SPFile]$File,
    [string]$DestinationPath,
    [string]$RequestNo,
    [string]$ActionName
  )

  if ($null -eq $File) {
    Write-RequestLog $RequestNo ("Cannot {0}. Source PDF was not found in SharePoint library." -f $ActionName)
    return
  }

  if (Test-Path -LiteralPath $DestinationPath) {
    return
  }

  $bytes = $File.OpenBinary()
  [System.IO.File]::WriteAllBytes($DestinationPath, $bytes)
  Write-RequestLog $RequestNo ("{0}: {1}" -f $ActionName, $DestinationPath)
}

function Copy-LocalFileIfMissing {
  param(
    [string]$SourcePath,
    [string]$DestinationPath,
    [string]$RequestNo,
    [string]$ActionName
  )

  if (-not (Test-Path -LiteralPath $SourcePath)) {
    Write-RequestLog $RequestNo ("Cannot {0}. Source file was not found: {1}" -f $ActionName, $SourcePath)
    return
  }

  if (Test-Path -LiteralPath $DestinationPath) {
    return
  }

  Copy-Item -LiteralPath $SourcePath -Destination $DestinationPath
  Write-RequestLog $RequestNo ("{0}: {1}" -f $ActionName, $DestinationPath)
}

function Set-SharedFolderPathIfNeeded {
  param(
    [Microsoft.SharePoint.SPListItem]$Item,
    [string]$ExpectedPath,
    [string]$RequestNo
  )

  $currentPath = Get-FieldString $Item "SharedFolderPath"

  if ($currentPath -ne $ExpectedPath) {
    $Item["SharedFolderPath"] = $ExpectedPath
    $Item.Update()
    Write-RequestLog $RequestNo ("SharedFolderPath updated to {0}" -f $ExpectedPath)
  }
}

function Sync-RequestFiles {
  param(
    [Microsoft.SharePoint.SPListItem]$Item,
    [Microsoft.SharePoint.SPList]$Library,
    [hashtable]$Snapshot
  )

  $requestNo = $Snapshot.RequestNo
  $status = $Snapshot.Status

  if ([string]::IsNullOrWhiteSpace($requestNo)) {
    return
  }

  $pdfFile = Get-RequestPdfFile $Library $requestNo $Item.ID
  $pendingFilePath = Join-Path $pendingPath ($requestNo + ".pdf")
  $signedFilePath = Join-Path $signedPath ($requestNo + ".pdf")
  $archiveFilePath = Join-Path $archivePath ($requestNo + ".pdf")

  if ($status -eq "Pending GM Signature") {
    Copy-SPFileIfMissing $pdfFile $pendingFilePath $requestNo "Copied PDF to Pending"
    Set-SharedFolderPathIfNeeded $Item (($ShareUncRoot.TrimEnd("\")) + "\Pending\" + $requestNo + ".pdf") $requestNo
    return
  }

  if ($status -eq "GM Signed Pending Office Manager Confirmation" -or $status -eq "Approved by GM") {
    Copy-LocalFileIfMissing $pendingFilePath $signedFilePath $requestNo "Copied PDF to Signed"
    Set-SharedFolderPathIfNeeded $Item (($ShareUncRoot.TrimEnd("\")) + "\Signed\" + $requestNo + ".pdf") $requestNo
    return
  }

  if ($status -eq "Closed" -or $status -eq "Rejected") {
    if (Test-Path -LiteralPath $signedFilePath) {
      Copy-LocalFileIfMissing $signedFilePath $archiveFilePath $requestNo "Copied PDF to Archive"
    } elseif (Test-Path -LiteralPath $pendingFilePath) {
      Copy-LocalFileIfMissing $pendingFilePath $archiveFilePath $requestNo "Copied PDF to Archive"
    } else {
      Copy-SPFileIfMissing $pdfFile $archiveFilePath $requestNo "Copied PDF to Archive"
    }

    Set-SharedFolderPathIfNeeded $Item (($ShareUncRoot.TrimEnd("\")) + "\Archive\" + $requestNo + ".pdf") $requestNo
  }
}

function Invoke-GMApprovalFileSync {
  $state = Get-State
  $web = $null

  try {
    $web = Get-SPWeb $SiteUrl
    $requests = $web.Lists[$RequestsListName]
    $library = $web.Lists[$DocumentsLibraryName]

    $query = New-Object Microsoft.SharePoint.SPQuery
    $query.ViewAttributes = "Scope='RecursiveAll'"
    $query.Query = "<OrderBy><FieldRef Name='ID' Ascending='FALSE' /></OrderBy>"

    $items = $requests.GetItems($query)

    foreach ($item in $items) {
      $snapshot = Get-ItemSnapshot $item
      $key = [string]$item.ID
      $previous = $null

      if ($state.ContainsKey($key)) {
        $previous = $state[$key]
      }

      Write-ChangeLog $snapshot.RequestNo $previous $snapshot
      Sync-RequestFiles $item $library $snapshot
      $state[$key] = $snapshot
    }

    Save-State $state
    Write-GeneralLog ("Sync completed. Items checked: {0}" -f $items.Count)
  } catch {
    Write-GeneralLog ("ERROR: {0}" -f $_.Exception.Message)
    throw
  } finally {
    if ($null -ne $web) {
      $web.Dispose()
    }
  }
}

Invoke-GMApprovalFileSync

<# 
Suggested scheduled task command, run once as Administrator on the SharePoint server:

$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File C:\GMApprovalDashboard\tools\GMApprovalFileSync.ps1"
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval (New-TimeSpan -Minutes 1) -RepetitionDuration (New-TimeSpan -Days 3650)
Register-ScheduledTask -TaskName "GM Approval File Sync" -Action $action -Trigger $trigger -User "TEST\sp.gmservice" -RunLevel Highest
#>
