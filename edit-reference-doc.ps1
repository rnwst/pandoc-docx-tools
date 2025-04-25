Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Set paths relative to current directory
$sourceDir = Join-Path $PWD "reference-doc"
$zipPath = Join-Path $PWD "reference-doc.zip"
$docxPath = Join-Path $PWD "reference-doc.docx"

# Cleanup any existing files
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
if (Test-Path $docxPath) { Remove-Item $docxPath -Force }

# Create the ZIP archive
$zipFileStream = [System.IO.File]::Create($zipPath)
$zip = New-Object System.IO.Compression.ZipArchive($zipFileStream, [System.IO.Compression.ZipArchiveMode]::Create)

# Get all files recursively
$files = Get-ChildItem -Path $sourceDir -Recurse -File

foreach ($file in $files) {
    # Get relative path in ZIP using Unix-style slashes
    $relativePath = $file.FullName.Substring($sourceDir.Length + 1).Replace("\", "/")

    # The contents of `word/media/*` need not be compressed, as the various
    # image formats already have built-in compression (this also speeds up the
    # script).
    $compressionLevel = if ($relativePath -like "word/media/*") {
        [System.IO.Compression.CompressionLevel]::NoCompression
    } else {
        [System.IO.Compression.CompressionLevel]::Optimal
    }

    # Create the entry with correct compression
    $entry = $zip.CreateEntry($relativePath, $compressionLevel)

    # Copy file contents
    $entryStream = $entry.Open()
    $fileStream = $file.OpenRead()
    $fileStream.CopyTo($entryStream)
    $fileStream.Dispose()
    $entryStream.Dispose()
}

$zip.Dispose()
$zipFileStream.Dispose()

# Rename the .zip to .docx
Rename-Item -Path $zipPath -NewName (Split-Path $docxPath -Leaf)

# Open with Microsoft Word
ii $docxPath
