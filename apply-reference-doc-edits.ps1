# Set paths
$docxPath = Join-Path $PWD "reference-doc.docx"
$tempZip = Join-Path ([System.IO.Path]::GetTempPath()) "reference-doc-temp.zip"
$tempExtractDir = Join-Path ([System.IO.Path]::GetTempPath()) "reference-doc-temp"
$targetDir = Join-Path $PWD "reference-doc"

# Ensure the .docx file exists
if (-Not (Test-Path $docxPath)) {
    Write-Error "The file reference-doc.docx does not exist."
    exit
}

# Cleanup any previous temp extract dir
if (Test-Path $tempExtractDir) {
    Remove-Item -LiteralPath $tempExtractDir -Recurse -Force
}

# Copy the .docx as a .zip so Expand-Archive will accept it
Copy-Item -LiteralPath $docxPath -Destination $tempZip -Force

# Extract to temp dir
Expand-Archive -LiteralPath $tempZip -DestinationPath $tempExtractDir -Force

# Build hashsets of relative file paths
function Get-RelativePaths($baseDir) {
    $baseDirFull = (Get-Item -LiteralPath $baseDir).FullName.TrimEnd('\')
    Get-ChildItem -LiteralPath $baseDir -Recurse -File | ForEach-Object {
        $fullPath = $_.FullName
        $relPath = $fullPath.Substring($baseDirFull.Length + 1)
        $relPath -replace '^[\\/]+', ''  # remove any leading slashes
    }
}

$sourceFiles = Get-RelativePaths $tempExtractDir
$targetFiles = Get-RelativePaths $targetDir

# Delete files in targetDir that are not in sourceFiles
$filesToDelete = $targetFiles | Where-Object { $sourceFiles -notcontains $_ }
foreach ($relPath in $filesToDelete) {
    $fullPath = Join-Path $targetDir $relPath
    Write-Host "Deleting $relPath since it is not present in the updated reference-doc.docx"
    Remove-Item -LiteralPath $fullPath -Force
}

# Copy new and updated files from source to target
foreach ($relPath in $sourceFiles) {
    $sourcePath = Join-Path $tempExtractDir $relPath
    $targetPath = Join-Path $targetDir $relPath

    # The files in `docProps` and `word\settings.xml` contain non-functional
    # changes on every save and should therefore not be updated.
    $protectedFiles = @(
        "docProps\app.xml",
        "docProps\core.xml",
        "word\settings.xml",
        "word\glossary\settings.xml"
    )

    if ($relPath -in $protectedFiles -and (Test-Path -LiteralPath $targetPath)) {
        Write-Host "Skipping $relPath as it likely contains only non-functional changes"
        continue
    }

    # Ensure target directory exists
    $targetSubdir = Split-Path $targetPath
    if (-Not (Test-Path $targetSubdir)) {
        New-Item -ItemType Directory -Path $targetSubdir -Force | Out-Null
    }

    # Copy file
    Write-Host "Copying $relPath"
    Copy-Item -LiteralPath $sourcePath -Destination $targetPath
}

# Cleanup
Remove-Item -LiteralPath $tempZip -Force
Remove-Item -LiteralPath $tempExtractDir -Recurse -Force

# Format XML files (and suppress output)
powershell -NoProfile -Command "& '.\format-openxml.ps1'" | Out-Null

Write-Host "Successfully extracted contents of reference-doc.docx into reference-doc/"
