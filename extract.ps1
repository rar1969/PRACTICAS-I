Add-Type -AssemblyName System.IO.Compression.FileSystem
$docxPath = Join-Path $PSScriptRoot "TFI_Bitacora_Rol_Docente (1).docx"
$zip = [System.IO.Compression.ZipFile]::OpenRead($docxPath)
$entry = $zip.Entries | Where-Object { $_.FullName -eq 'word/document.xml' }
$stream = $entry.Open()
$reader = New-Object System.IO.StreamReader($stream)
$xmlContent = $reader.ReadToEnd()
$reader.Close()
$zip.Dispose()

$xml = [xml]$xmlContent
$ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
$ns.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")

$paragraphs = $xml.SelectNodes("//w:p", $ns)
$lines = @()
foreach ($p in $paragraphs) {
    $text = $p.InnerText
    if (-not [string]::IsNullOrWhiteSpace($text)) {
        $lines += $text
    }
}
$outputPath = Join-Path $PSScriptRoot "extracted_text.txt"
[System.IO.File]::WriteAllLines($outputPath, $lines, [System.Text.Encoding]::UTF8)
Write-Host "Extracted $($lines.Count) lines to $outputPath"
