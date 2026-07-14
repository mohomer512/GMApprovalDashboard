param(
  [string]$SiteUrl = "http://spse26h/GM",
  [string]$RequestsListName = "GM Requests",
  [string]$DocumentsLibraryName = "GM Approval Documents",
  [string]$ShareRootPath = "C:\GMApprovalShare",
  [string]$ShareUncRoot = "\\SPSE26H\GMApprovalShare",
  [ValidateRange(0, 86400)]
  [int]$PollSeconds = 0,
  [ValidateRange(0, 300)]
  [int]$SignedFileStableSeconds = 5,
  [switch]$SkipExecution
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

function Write-RequestLogBlock {
  param(
    [string]$RequestNo,
    [string[]]$Lines
  )

  $safeRequestNo = Get-SafeFileName $RequestNo
  $logPath = Join-Path $logsPath ($safeRequestNo + ".log")
  Add-Content -LiteralPath $logPath -Value $Lines -Encoding UTF8
}

function Format-AuditValue {
  param([string]$Value)

  if ([string]::IsNullOrWhiteSpace($Value)) {
    return "[blank]"
  }

  return $Value
}

function Add-CommentAuditBlock {
  param(
    [string]$RequestNo,
    [hashtable]$Previous,
    [hashtable]$Current
  )

  if ([string]::IsNullOrWhiteSpace($RequestNo)) {
    return $false
  }

  $safeRequestNo = Get-SafeFileName $RequestNo
  $commentLogPath = Join-Path $logsPath ($safeRequestNo + "-Comments.log")
  $commentFields = @(
    @{ Name = "OfficeManagerComment"; Label = "Office Manager Comment" },
    @{ Name = "HODComment"; Label = "HOD Comment" },
    @{ Name = "GMComment"; Label = "GM Comment" }
  )
  $lines = @()

  if (-not (Test-Path -LiteralPath $commentLogPath)) {
    $lines = @(
      "================================================",
      "Comment audit started",
      ("Audit time  : {0}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")),
      ("Request ID  : {0}" -f $Current.Id),
      ("Request No  : {0}" -f (Format-AuditValue $Current.RequestNo)),
      ("Modified By : {0}" -f (Format-AuditValue $Current.ModifiedBy)),
      "------------------------------------------------"
    )

    foreach ($field in $commentFields) {
      $lines += ("{0} : {1}" -f $field.Label, (Format-AuditValue ([string]$Current[$field.Name])))
      $lines += "."
    }
  } elseif ($null -ne $Previous) {
    $changedCommentFields = @()

    foreach ($field in $commentFields) {
      if ([string]$Previous[$field.Name] -ne [string]$Current[$field.Name]) {
        $changedCommentFields += $field
      }
    }

    if ($changedCommentFields.Count -gt 0) {
      $lines = @(
        "================================================",
        "Comment version changed",
        ("Audit time  : {0}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")),
        ("Request ID  : {0}" -f $Current.Id),
        ("Request No  : {0}" -f (Format-AuditValue $Current.RequestNo)),
        ("Modified By : {0}" -f (Format-AuditValue $Current.ModifiedBy)),
        "------------------------------------------------"
      )

      foreach ($field in $changedCommentFields) {
        $lines += ("{0} : {1}" -f $field.Label, (Format-AuditValue ([string]$Previous[$field.Name])))
        $lines += ("{0} changed to : {1}" -f $field.Label, (Format-AuditValue ([string]$Current[$field.Name])))
        $lines += "."
      }
    }
  }

  if ($lines.Count -eq 0) {
    return $false
  }

  $lines += ""
  Add-Content -LiteralPath $commentLogPath -Value $lines -Encoding UTF8
  return $true
}

function Write-AsciiBytes {
  param(
    [System.IO.Stream]$Stream,
    [string]$Value
  )

  $bytes = [System.Text.Encoding]::ASCII.GetBytes($Value)
  $Stream.Write($bytes, 0, $bytes.Length)
}

function Get-WrappedPdfLines {
  param(
    [System.Drawing.Graphics]$Graphics,
    [System.Drawing.Font]$Font,
    [string[]]$SourceLines,
    [float]$MaximumWidth
  )

  $wrappedLines = New-Object System.Collections.ArrayList

  foreach ($sourceLine in $SourceLines) {
    if ([string]::IsNullOrWhiteSpace($sourceLine)) {
      [void]$wrappedLines.Add("")
      continue
    }

    $words = $sourceLine -split "\s+"
    $currentLine = ""

    foreach ($word in $words) {
      $candidate = if ([string]::IsNullOrEmpty($currentLine)) { $word } else { $currentLine + " " + $word }
      $candidateWidth = $Graphics.MeasureString($candidate, $Font).Width

      if ($candidateWidth -le $MaximumWidth -or [string]::IsNullOrEmpty($currentLine)) {
        $currentLine = $candidate
      } else {
        [void]$wrappedLines.Add($currentLine)
        $currentLine = $word
      }
    }

    if (-not [string]::IsNullOrEmpty($currentLine)) {
      [void]$wrappedLines.Add($currentLine)
    }
  }

  return ,$wrappedLines.ToArray()
}

function New-ImagePdf {
  param(
    [string[]]$Lines,
    [string]$Title,
    [string]$OutputPath
  )

  Add-Type -AssemblyName System.Drawing

  $pageWidth = 1240
  $pageHeight = 1754
  $margin = 85
  $lineHeight = 34
  $bodyTop = 180
  $bodyBottom = $pageHeight - 80
  $linesPerPage = [math]::Floor(($bodyBottom - $bodyTop) / $lineHeight)
  $measureBitmap = New-Object System.Drawing.Bitmap 1, 1
  $measureGraphics = [System.Drawing.Graphics]::FromImage($measureBitmap)
  $bodyFont = New-Object System.Drawing.Font "Arial", 20, ([System.Drawing.FontStyle]::Regular), ([System.Drawing.GraphicsUnit]::Pixel)
  $titleFont = New-Object System.Drawing.Font "Arial", 30, ([System.Drawing.FontStyle]::Bold), ([System.Drawing.GraphicsUnit]::Pixel)
  $metaFont = New-Object System.Drawing.Font "Arial", 16, ([System.Drawing.FontStyle]::Regular), ([System.Drawing.GraphicsUnit]::Pixel)
  $jpegPages = New-Object System.Collections.ArrayList

  try {
    $wrappedLines = Get-WrappedPdfLines $measureGraphics $bodyFont $Lines ($pageWidth - (2 * $margin))

    if ($wrappedLines.Count -eq 0) {
      $wrappedLines = @("No comments have been recorded.")
    }

    $pageCount = [math]::Ceiling($wrappedLines.Count / $linesPerPage)

    for ($pageIndex = 0; $pageIndex -lt $pageCount; $pageIndex++) {
      $bitmap = New-Object System.Drawing.Bitmap $pageWidth, $pageHeight
      $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
      $memoryStream = New-Object System.IO.MemoryStream

      try {
        $graphics.Clear([System.Drawing.Color]::White)
        $graphics.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
        $graphics.DrawString($Title, $titleFont, [System.Drawing.Brushes]::Black, [float]$margin, [float]65)
        $pageLabel = "Generated {0} | Page {1} of {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), ($pageIndex + 1), $pageCount
        $graphics.DrawString($pageLabel, $metaFont, [System.Drawing.Brushes]::DimGray, [float]$margin, [float]115)

        $startIndex = $pageIndex * $linesPerPage
        $endIndex = [math]::Min($startIndex + $linesPerPage, $wrappedLines.Count)
        $y = $bodyTop

        for ($lineIndex = $startIndex; $lineIndex -lt $endIndex; $lineIndex++) {
          $graphics.DrawString([string]$wrappedLines[$lineIndex], $bodyFont, [System.Drawing.Brushes]::Black, [float]$margin, [float]$y)
          $y += $lineHeight
        }

        $bitmap.Save($memoryStream, [System.Drawing.Imaging.ImageFormat]::Jpeg)
        [void]$jpegPages.Add($memoryStream.ToArray())
      } finally {
        $memoryStream.Dispose()
        $graphics.Dispose()
        $bitmap.Dispose()
      }
    }
  } finally {
    $metaFont.Dispose()
    $titleFont.Dispose()
    $bodyFont.Dispose()
    $measureGraphics.Dispose()
    $measureBitmap.Dispose()
  }

  $pdfStream = New-Object System.IO.MemoryStream

  try {
    Write-AsciiBytes $pdfStream "%PDF-1.4`n"
    $objectCount = 2 + ($jpegPages.Count * 3)
    $offsets = New-Object long[] ($objectCount + 1)

    for ($objectId = 1; $objectId -le $objectCount; $objectId++) {
      $offsets[$objectId] = $pdfStream.Position
      Write-AsciiBytes $pdfStream ("{0} 0 obj`n" -f $objectId)

      if ($objectId -eq 1) {
        Write-AsciiBytes $pdfStream "<< /Type /Catalog /Pages 2 0 R >>`n"
      } elseif ($objectId -eq 2) {
        $kids = @()
        for ($pageIndex = 0; $pageIndex -lt $jpegPages.Count; $pageIndex++) {
          $kids += ("{0} 0 R" -f (3 + ($pageIndex * 3)))
        }
        Write-AsciiBytes $pdfStream ("<< /Type /Pages /Count {0} /Kids [ {1} ] >>`n" -f $jpegPages.Count, ($kids -join " "))
      } else {
        $relativeId = $objectId - 3
        $pageIndex = [math]::Floor($relativeId / 3)
        $objectType = $relativeId % 3
        $pageObjectId = 3 + ($pageIndex * 3)
        $imageObjectId = $pageObjectId + 1
        $contentObjectId = $pageObjectId + 2

        if ($objectType -eq 0) {
          Write-AsciiBytes $pdfStream ("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /XObject << /Im1 {0} 0 R >> >> /Contents {1} 0 R >>`n" -f $imageObjectId, $contentObjectId)
        } elseif ($objectType -eq 1) {
          $imageBytes = [byte[]]$jpegPages[$pageIndex]
          Write-AsciiBytes $pdfStream ("<< /Type /XObject /Subtype /Image /Width {0} /Height {1} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length {2} >>`nstream`n" -f $pageWidth, $pageHeight, $imageBytes.Length)
          $pdfStream.Write($imageBytes, 0, $imageBytes.Length)
          Write-AsciiBytes $pdfStream "`nendstream`n"
        } else {
          $content = "q 595 0 0 842 0 0 cm /Im1 Do Q"
          Write-AsciiBytes $pdfStream ("<< /Length {0} >>`nstream`n{1}`nendstream`n" -f $content.Length, $content)
        }
      }

      Write-AsciiBytes $pdfStream "endobj`n"
    }

    $xrefOffset = $pdfStream.Position
    Write-AsciiBytes $pdfStream ("xref`n0 {0}`n" -f ($objectCount + 1))
    Write-AsciiBytes $pdfStream "0000000000 65535 f `n"

    for ($objectId = 1; $objectId -le $objectCount; $objectId++) {
      Write-AsciiBytes $pdfStream ("{0:D10} 00000 n `n" -f $offsets[$objectId])
    }

    Write-AsciiBytes $pdfStream ("trailer`n<< /Size {0} /Root 1 0 R >>`nstartxref`n{1}`n%%EOF" -f ($objectCount + 1), $xrefOffset)
    [System.IO.File]::WriteAllBytes($OutputPath, $pdfStream.ToArray())
  } finally {
    $pdfStream.Dispose()
  }
}

function Update-CommentAuditPdf {
  param(
    [string]$RequestNo,
    [bool]$CommentLogChanged
  )

  if ([string]::IsNullOrWhiteSpace($RequestNo)) {
    return
  }

  $safeRequestNo = Get-SafeFileName $RequestNo
  $commentLogPath = Join-Path $logsPath ($safeRequestNo + "-Comments.log")
  $commentPdfPath = Join-Path $logsPath ($safeRequestNo + "-Comments.pdf")

  if (-not (Test-Path -LiteralPath $commentLogPath)) {
    return
  }

  if ($CommentLogChanged -or -not (Test-Path -LiteralPath $commentPdfPath)) {
    try {
      $commentLines = Get-Content -LiteralPath $commentLogPath -Encoding UTF8
      New-ImagePdf $commentLines ("GM Approval Comments - " + $RequestNo) $commentPdfPath
      Write-RequestLog $RequestNo ("Comments audit PDF updated: {0}" -f $commentPdfPath)
    } catch {
      Write-RequestLog $RequestNo ("Comments audit PDF could not be updated: {0}" -f $_.Exception.Message)
    }
  }
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

function Get-AuditFieldDefinitions {
  return @(
    @{ Name = "Title"; Label = "Title"; IsComment = $false },
    @{ Name = "RequestDate"; Label = "Request Date"; IsComment = $false },
    @{ Name = "Department"; Label = "Department"; IsComment = $false },
    @{ Name = "RequestDetails"; Label = "Request Details"; IsComment = $false },
    @{ Name = "Status"; Label = "Status"; IsComment = $false },
    @{ Name = "OfficeManagerComment"; Label = "Office Manager Comment"; IsComment = $true },
    @{ Name = "HODComment"; Label = "HOD Comment"; IsComment = $true },
    @{ Name = "GMComment"; Label = "GM Comment"; IsComment = $true },
    @{ Name = "SharedFolderPath"; Label = "Shared Folder Path"; IsComment = $false },
    @{ Name = "PDFFileUrl"; Label = "PDF File URL"; IsComment = $false },
    @{ Name = "SignedPDFUrl"; Label = "Signed PDF URL"; IsComment = $false },
    @{ Name = "GMApprovalDetected"; Label = "GM Approval Detected"; IsComment = $false },
    @{ Name = "GMApprovalConfirmed"; Label = "GM Approval Confirmed"; IsComment = $false }
  )
}

function Get-ItemSnapshot {
  param([Microsoft.SharePoint.SPListItem]$Item)

  return @{
    Id = [string]$Item.ID
    Title = Get-FieldString $Item "Title"
    RequestNo = Get-FieldString $Item "RequestNo"
    RequestDate = Get-FieldString $Item "RequestDate"
    Department = Get-FieldString $Item "Department"
    RequestDetails = Get-FieldString $Item "RequestDetails"
    Status = Get-FieldString $Item "Status"
    OfficeManagerComment = Get-FieldString $Item "OfficeManagerComment"
    HODComment = Get-FieldString $Item "HODComment"
    GMComment = Get-FieldString $Item "GMComment"
    SharedFolderPath = Get-FieldString $Item "SharedFolderPath"
    PDFFileUrl = Get-FieldString $Item "PDFFileUrl"
    SignedPDFUrl = Get-FieldString $Item "SignedPDFUrl"
    GMApprovalDetected = Get-FieldString $Item "GMApprovalDetected"
    GMApprovalConfirmed = Get-FieldString $Item "GMApprovalConfirmed"
    Created = ([datetime]$Item["Created"]).ToString("o")
    CreatedBy = Get-UserFieldString $Item "Author"
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

  $commentLogChanged = Add-CommentAuditBlock $RequestNo $Previous $Current
  Update-CommentAuditPdf $RequestNo $commentLogChanged

  $now = Get-Date
  $auditFields = Get-AuditFieldDefinitions
  $changedFields = @()

  if ($null -eq $Previous) {
    $lines = @(
      "================================================",
      ("Audit event : Request tracking started"),
      ("Audit time  : {0}" -f $now.ToString("yyyy-MM-dd HH:mm:ss")),
      ("Request ID  : {0}" -f $Current.Id),
      ("Request No  : {0}" -f (Format-AuditValue $Current.RequestNo)),
      ("Created     : {0}" -f (Format-AuditValue $Current.Created)),
      ("Created By  : {0}" -f (Format-AuditValue $Current.CreatedBy)),
      ("Modified    : {0}" -f (Format-AuditValue $Current.Modified)),
      ("Modified By : {0}" -f (Format-AuditValue $Current.ModifiedBy)),
      "------------------------------------------------"
    )

    foreach ($field in $auditFields) {
      $fieldName = $field.Name
      $label = $field.Label
      $lines += ("{0} : {1}" -f $label, (Format-AuditValue $Current[$fieldName]))
    }

    $lines += ""
    Write-RequestLogBlock $RequestNo $lines
    return
  }

  foreach ($field in $auditFields) {
    $fieldName = $field.Name
    $oldValue = [string]$Previous[$fieldName]
    $newValue = [string]$Current[$fieldName]

    if ($oldValue -ne $newValue) {
      $changedFields += $field
    }
  }

  if ($changedFields.Count -eq 0) {
    return
  }

  $lines = @(
    "================================================",
    ("Audit event : Request changed"),
    ("Audit time  : {0}" -f $now.ToString("yyyy-MM-dd HH:mm:ss")),
    ("Request ID  : {0}" -f $Current.Id),
    ("Request No  : {0}" -f (Format-AuditValue $Current.RequestNo)),
    ("Modified    : {0}" -f (Format-AuditValue $Current.Modified)),
    ("Modified By : {0}" -f (Format-AuditValue $Current.ModifiedBy)),
    "------------------------------------------------"
  )

  foreach ($field in $changedFields) {
    $fieldName = $field.Name
    $label = $field.Label
    $oldValue = Format-AuditValue ([string]$Previous[$fieldName])
    $newValue = Format-AuditValue ([string]$Current[$fieldName])

    if ($field.IsComment) {
      $lines += ("{0} version" -f $label)
      $lines += ("{0} : {1}" -f $label, $oldValue)
      $lines += ("{0} changed to : {1}" -f $label, $newValue)
    } else {
      $lines += ("{0} : {1}" -f $label, $oldValue)
      $lines += ("{0} changed to : {1}" -f $label, $newValue)
    }

    $lines += "."
  }

  $lines += ""
  Write-RequestLogBlock $RequestNo $lines
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
  $exactFile = $null
  $workflowFile = $null
  $requestFile = $null
  $legacyFile = $null
  $fallbackFile = $null
  $expectedFileName = $RequestNo + ".pdf"

  foreach ($libraryItem in $items) {
    if ($null -eq $libraryItem.File) {
      continue
    }

    $documentType = (Get-FieldString $libraryItem "DocumentType").Trim()

    if ($null -eq $fallbackFile -and
        $documentType -ine "HOD PDF" -and
        $documentType -ine "Secretary PDF") {
      $fallbackFile = $libraryItem.File
    }

    if ($null -eq $exactFile -and $libraryItem.File.Name -ieq $expectedFileName) {
      $exactFile = $libraryItem.File
    }

    if ($null -eq $workflowFile -and $documentType -ieq "Workflow PDF") {
      $workflowFile = $libraryItem.File
    }

    if ($null -eq $requestFile -and $documentType -ieq "Request PDF") {
      $requestFile = $libraryItem.File
    }

    if ($null -eq $legacyFile -and [string]::IsNullOrWhiteSpace($documentType)) {
      $legacyFile = $libraryItem.File
    }
  }

  if ($null -ne $exactFile) {
    return $exactFile
  }

  if ($null -ne $workflowFile) {
    return $workflowFile
  }

  if ($null -ne $requestFile) {
    return $requestFile
  }

  if ($null -ne $legacyFile) {
    return $legacyFile
  }

  return $fallbackFile
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

function Move-LocalFileReplacingDestination {
  param(
    [string]$SourcePath,
    [string]$DestinationPath
  )

  if (Test-Path -LiteralPath $DestinationPath) {
    Remove-Item -LiteralPath $DestinationPath -Force
  }

  Move-Item -LiteralPath $SourcePath -Destination $DestinationPath
}

function Get-FileContentHash {
  param(
    [string]$Path,
    [string]$RequestNo
  )

  try {
    $stream = [System.IO.File]::Open(
      $Path,
      [System.IO.FileMode]::Open,
      [System.IO.FileAccess]::Read,
      [System.IO.FileShare]::ReadWrite
    )
    $sha256 = [System.Security.Cryptography.SHA256]::Create()

    try {
      $hashBytes = $sha256.ComputeHash($stream)
      return ([System.BitConverter]::ToString($hashBytes)).Replace("-", "")
    } finally {
      $sha256.Dispose()
      $stream.Dispose()
    }
  } catch {
    Write-RequestLog $RequestNo ("Could not read PDF for signature detection: {0}" -f $_.Exception.Message)
    return ""
  }
}

function Get-SharePointFileContentHash {
  param(
    [Microsoft.SharePoint.SPFile]$File,
    [string]$RequestNo
  )

  if ($null -eq $File) {
    Write-RequestLog $RequestNo "Could not compare the Pending PDF because the SharePoint library file was not found."
    return ""
  }

  try {
    $fileBytes = $File.OpenBinary()
    $sha256 = [System.Security.Cryptography.SHA256]::Create()

    try {
      $hashBytes = $sha256.ComputeHash($fileBytes)
      return ([System.BitConverter]::ToString($hashBytes)).Replace("-", "")
    } finally {
      $sha256.Dispose()
    }
  } catch {
    Write-RequestLog $RequestNo ("Could not read the SharePoint PDF for signature detection: {0}" -f $_.Exception.Message)
    return ""
  }
}

function Update-SignedWorkflowAttachmentIfPresent {
  param(
    [Microsoft.SharePoint.SPListItem]$Item,
    [string]$RequestNo,
    [byte[]]$PdfBytes
  )

  $expectedAttachmentName = $RequestNo + ".pdf"

  try {
    $matchingAttachmentName = $null

    foreach ($attachmentName in $Item.Attachments) {
      if ([string]$attachmentName -ieq $expectedAttachmentName) {
        $matchingAttachmentName = [string]$attachmentName
        break
      }
    }

    if ([string]::IsNullOrWhiteSpace($matchingAttachmentName)) {
      Write-RequestLog $RequestNo ("Signed workflow PDF attachment was not present. Expected attachment: {0}" -f $expectedAttachmentName)
      return
    }

    $attachmentUrlPrefix = [string]$Item.Attachments.UrlPrefix
    $attachmentUrl = $attachmentUrlPrefix.TrimEnd("/") + "/" + $matchingAttachmentName
    $attachmentFile = $Item.Web.GetFile($attachmentUrl)

    if ($null -eq $attachmentFile -or -not $attachmentFile.Exists) {
      Write-RequestLog $RequestNo ("Signed workflow PDF attachment was listed but its file was missing: {0}" -f $matchingAttachmentName)
      return
    }

    $attachmentFile.SaveBinary($PdfBytes)
    Write-RequestLog $RequestNo ("Signed workflow PDF attachment updated: {0}" -f $matchingAttachmentName)
  } catch {
    Write-RequestLog $RequestNo ("Signed workflow PDF attachment could not be updated: {0}" -f $_.Exception.Message)
    throw
  }
}

function Set-RequestPdfUrls {
  param(
    [Microsoft.SharePoint.SPListItem]$Item,
    [string]$RequestNo,
    [string]$PdfUrl
  )

  try {
    $pdfUrlValue = New-Object Microsoft.SharePoint.SPFieldUrlValue
    $pdfUrlValue.Url = $PdfUrl
    $pdfUrlValue.Description = $RequestNo + " signed PDF"

    $signedUrlValue = New-Object Microsoft.SharePoint.SPFieldUrlValue
    $signedUrlValue.Url = $PdfUrl
    $signedUrlValue.Description = $RequestNo + " signed PDF"

    $Item["PDFFileUrl"] = $pdfUrlValue
    $Item["SignedPDFUrl"] = $signedUrlValue
    $Item.Update()
  } catch {
    Write-RequestLog $RequestNo ("The edited SharePoint PDF was published, but its URL fields could not be updated: {0}" -f $_.Exception.Message)
    throw
  }
}

function Confirm-GmSignedPdf {
  param(
    [Microsoft.SharePoint.SPListItem]$Item,
    [Microsoft.SharePoint.SPFile]$SharePointFile,
    [string]$RequestNo,
    [string]$PendingFilePath,
    [string]$SignedFilePath,
    [hashtable]$Snapshot
  )

  if ($null -eq $SharePointFile) {
    throw ("Cannot publish the signed PDF for {0}. The SharePoint library file was not found." -f $RequestNo)
  }

  $signedPdfUrl = $Item.Web.Site.MakeFullUrl($SharePointFile.ServerRelativeUrl)
  $signedBytes = [System.IO.File]::ReadAllBytes($PendingFilePath)
  $SharePointFile.SaveBinary($signedBytes)
  Update-SignedWorkflowAttachmentIfPresent $Item $RequestNo $signedBytes
  Set-RequestPdfUrls $Item $RequestNo $signedPdfUrl

  $Item["Status"] = "GM Signed Pending Office Manager Confirmation"
  $Item["GMApprovalDetected"] = $true
  $Item["SharedFolderPath"] = (($ShareUncRoot.TrimEnd("\")) + "\Signed\" + $RequestNo + ".pdf")
  $Item.Update()

  Move-LocalFileReplacingDestination $PendingFilePath $SignedFilePath
  Write-RequestLog $RequestNo ("GM signature file change confirmed. Edited PDF moved from Pending to Signed: {0}" -f $SignedFilePath)
  Write-RequestLog $RequestNo ("Edited PDF published to SharePoint: {0}" -f $signedPdfUrl)
  Write-RequestLog $RequestNo "SharePoint status changed to GM Signed Pending Office Manager Confirmation."

  $Snapshot["PendingSignatureConfirmed"] = "true"
  $Snapshot["PendingCandidateHash"] = ""
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

function Get-AvailableListTitles {
  param([Microsoft.SharePoint.SPWeb]$Web)

  $titles = @()

  try {
    foreach ($list in $Web.Lists) {
      if (-not $list.Hidden) {
        $titles += [string]$list.Title
      }
    }
  } catch {
    return ("Could not enumerate lists. Current Windows identity '{0}' received: {1}" -f [System.Security.Principal.WindowsIdentity]::GetCurrent().Name, $_.Exception.Message)
  }

  return ($titles | Sort-Object) -join ", "
}

function Get-RequiredList {
  param(
    [Microsoft.SharePoint.SPWeb]$Web,
    [string]$ListTitle,
    [string]$Purpose
  )

  $list = $null

  try {
    $list = $Web.Lists.TryGetList($ListTitle)
  } catch {
    $list = $null
  }

  if ($null -eq $list) {
    try {
      foreach ($candidate in $Web.Lists) {
        if ($candidate.Title -ieq $ListTitle) {
          $list = $candidate
          break
        }
      }
    } catch {
      throw ("Could not access {0} named '{1}' in site '{2}'. Current Windows identity '{3}' received access denied while reading lists. Grant this account SharePoint site/list permission and SPShellAdmin access to the content database. Original error: {4}" -f $Purpose, $ListTitle, $Web.Url, [System.Security.Principal.WindowsIdentity]::GetCurrent().Name, $_.Exception.Message)
    }
  }

  if ($null -eq $list) {
    $availableLists = Get-AvailableListTitles $Web
    throw ("Could not find {0} named '{1}' in site '{2}'. Available visible lists/libraries: {3}" -f $Purpose, $ListTitle, $Web.Url, $availableLists)
  }

  return $list
}

function Sync-RequestFiles {
  param(
    [Microsoft.SharePoint.SPListItem]$Item,
    [Microsoft.SharePoint.SPList]$Library,
    [hashtable]$Snapshot,
    [hashtable]$Previous
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

    if (-not (Test-Path -LiteralPath $pendingFilePath)) {
      return
    }

    $currentHash = Get-FileContentHash $pendingFilePath $requestNo

    if ([string]::IsNullOrWhiteSpace($currentHash)) {
      return
    }

    $sharePointHash = Get-SharePointFileContentHash $pdfFile $requestNo

    if ([string]::IsNullOrWhiteSpace($sharePointHash)) {
      return
    }

    $previousBaselineHash = ""

    if ($null -ne $Previous) {
      $previousBaselineHash = [string]$Previous["PendingBaselineHash"]
    }

    if ($currentHash -eq $sharePointHash -and
        ([string]::IsNullOrWhiteSpace($previousBaselineHash) -or $currentHash -eq $previousBaselineHash)) {
      $Snapshot["PendingBaselineHash"] = $sharePointHash
      return
    }

    $lastWriteUtc = [System.IO.File]::GetLastWriteTimeUtc($pendingFilePath)
    $stableAgeSeconds = ([datetime]::UtcNow - $lastWriteUtc).TotalSeconds

    if ($stableAgeSeconds -lt $SignedFileStableSeconds) {
      Write-RequestLog $requestNo ("Edited Pending PDF detected, but it is only {0:N1} seconds old. Waiting until it is stable for {1} seconds." -f $stableAgeSeconds, $SignedFileStableSeconds)
      return
    }

    Confirm-GmSignedPdf $Item $pdfFile $requestNo $pendingFilePath $signedFilePath $Snapshot
    return
  }

  if ($status -eq "GM Signed Pending Office Manager Confirmation" -or
      $status -eq "Approved by GM" -or
      $status -eq "Pending Secretary Information" -or
      $status -eq "Pending HOD Information") {
    if (Test-Path -LiteralPath $pendingFilePath) {
      if ($null -eq $pdfFile) {
        throw ("Cannot finish moving the signed PDF for {0}. The SharePoint library file was not found." -f $requestNo)
      }

      $signedBytes = [System.IO.File]::ReadAllBytes($pendingFilePath)
      $pdfFile.SaveBinary($signedBytes)
      Update-SignedWorkflowAttachmentIfPresent $Item $requestNo $signedBytes

      $signedPdfUrl = $Item.Web.Site.MakeFullUrl($pdfFile.ServerRelativeUrl)
      Set-RequestPdfUrls $Item $requestNo $signedPdfUrl

      $Item["SharedFolderPath"] = (($ShareUncRoot.TrimEnd("\")) + "\Signed\" + $requestNo + ".pdf")
      $Item.Update()

      Move-LocalFileReplacingDestination $pendingFilePath $signedFilePath
      Write-RequestLog $requestNo ("Moved remaining Pending PDF to Signed: {0}" -f $signedFilePath)
      Write-RequestLog $requestNo ("Edited PDF published to SharePoint: {0}" -f $signedPdfUrl)
    }

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
    Write-GeneralLog ("Sync started as {0}" -f [System.Security.Principal.WindowsIdentity]::GetCurrent().Name)

    $web = Get-SPWeb $SiteUrl
    $requests = Get-RequiredList $web $RequestsListName "requests list"
    $library = Get-RequiredList $web $DocumentsLibraryName "documents library"

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
      Sync-RequestFiles $item $library $snapshot $previous
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

if (-not $SkipExecution) {
  do {
    try {
      Invoke-GMApprovalFileSync
    } catch {
      if ($PollSeconds -le 0) {
        throw
      }
    }

    if ($PollSeconds -le 0) {
      break
    }

    Start-Sleep -Seconds $PollSeconds
  } while ($true)
}

<# 
Suggested continuous 10-second scheduled task configuration.
Run once as Administrator on the SharePoint server:

$credential = Get-Credential "TEST\Administrator"
$action = New-ScheduledTaskAction -Execute "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -Argument '-NoProfile -ExecutionPolicy Bypass -File "C:\GMApprovalDashboard\tools\GMApprovalFileSync.ps1" -PollSeconds 10'
$trigger = New-ScheduledTaskTrigger -AtStartup
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -RestartCount 999 -RestartInterval (New-TimeSpan -Minutes 1) -ExecutionTimeLimit ([TimeSpan]::Zero) -MultipleInstances IgnoreNew

Stop-ScheduledTask -TaskName "GM Approval File Sync" -ErrorAction SilentlyContinue
Set-ScheduledTask -TaskName "GM Approval File Sync" -Action $action -Trigger $trigger -Settings $settings -User $credential.UserName -Password $credential.GetNetworkCredential().Password
Start-ScheduledTask -TaskName "GM Approval File Sync"
#>
