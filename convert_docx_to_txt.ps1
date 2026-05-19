$ErrorActionPreference = "Stop"

$desktopDir = [Environment]::GetFolderPath("Desktop")

# Find the most recent .docx on Desktop (avoid Korean literals; PS 5.1 encoding issues)
$candidate = Get-ChildItem -LiteralPath $desktopDir -Filter "*.docx" |
  Sort-Object LastWriteTime -Descending |
  Select-Object -First 1

if ($null -eq $candidate) {
  throw "No matching .docx found in: $desktopDir"
}

$docxPath = $candidate.FullName
$outPath  = Join-Path -Path $PSScriptRoot -ChildPath "dev_plan_extracted.txt"

if (!(Test-Path -LiteralPath $docxPath)) {
  throw "DOCX not found: $docxPath"
}

$word = $null
$doc = $null

try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0

  # (FileName, ConfirmConversions, ReadOnly)
  $doc = $word.Documents.Open($docxPath, $false, $true)

  # 7 = wdFormatUnicodeText (UTF-16)
  $wdFormatUnicodeText = 7
  # Prefer SaveAs2; avoid [ref] which breaks with COM interop here
  $doc.SaveAs2($outPath, $wdFormatUnicodeText)

  # Also export plain text via .Content.Text as fallback (ensures something is written)
  if (!(Test-Path -LiteralPath $outPath)) {
    $txt = $doc.Content.Text
    [System.IO.File]::WriteAllText($outPath, $txt, [System.Text.Encoding]::UTF8)
  }
} finally {
  if ($doc -ne $null) { $doc.Close($false) | Out-Null }
  if ($word -ne $null) { $word.Quit() | Out-Null }
  [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc) | Out-Null
  [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word) | Out-Null
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}

Write-Output "OK: wrote $outPath"

