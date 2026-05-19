$ErrorActionPreference = "Stop"

$desktopDir = [Environment]::GetFolderPath("Desktop")
$outPath = Join-Path -Path $PSScriptRoot -ChildPath "dev_plan_extracted.txt"

# pick most recent docx on Desktop
$docx = Get-ChildItem -LiteralPath $desktopDir -Filter "*.docx" |
  Sort-Object LastWriteTime -Descending |
  Select-Object -First 1

if ($null -eq $docx) {
  throw "No .docx found in: $desktopDir"
}

# DOCX is a zip. Read main document XML.
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead($docx.FullName)
try {
  $entry = $zip.GetEntry("word/document.xml")
  if ($null -eq $entry) { throw "word/document.xml not found inside docx" }

  $stream = $entry.Open()
  try {
    $sr = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8, $true)
    $xmlText = $sr.ReadToEnd()
  } finally {
    $stream.Dispose()
  }
} finally {
  $zip.Dispose()
}

[xml]$xml = $xmlText
$ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main") | Out-Null

$paras = $xml.SelectNodes("//w:document/w:body/w:p", $ns)
$lines = New-Object System.Collections.Generic.List[string]

foreach ($p in $paras) {
  $texts = $p.SelectNodes(".//w:t", $ns)
  if ($texts -eq $null -or $texts.Count -eq 0) {
    $lines.Add("")
    continue
  }
  $sb = New-Object System.Text.StringBuilder
  foreach ($t in $texts) {
    [void]$sb.Append($t.InnerText)
  }
  $line = $sb.ToString().Trim()
  $lines.Add($line)
}

[System.IO.File]::WriteAllLines($outPath, $lines, [System.Text.Encoding]::UTF8)
Write-Output ("OK: extracted from {0} -> {1} ({2} lines)" -f $docx.Name, $outPath, $lines.Count)

