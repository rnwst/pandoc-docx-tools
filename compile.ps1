# Get the first argument as the file name.
$mdFile = $args[0]

# Check if the argument was provided.
if (-not $mdFile) {
    Write-Host "Usage: .\Run-With-File.ps1 <filename>"
    exit 1
}

# Check if the markdown file exists.
if (-not (Test-Path $mdFile)) {
    Write-Host "Error: File '$mdFile' does not exist."
    exit 1
}

# Check if the file exists.
if (-not (Test-Path $mdFile)) {
    Write-Host "Error: File '$mdFile' does not exist."
    exit 1
}

# Check if the file extension is `.md`.
if ([System.IO.Path]::GetExtension($mdFile) -ne ".md") {
    Write-Host "Error: File '$mdFile' is not a Markdown file."
    exit 1
}

if (-not (Test-Path 'reference-doc.docx')) {
    Write-Host "Error: File 'reference-doc.docx' does not exist. Run '.\edit-reference-doc.ps1' and try again."
    exit 1
}

$docxFile = [System.IO.Path]::ChangeExtension($mdFile, ".docx")
$resourcePath = Split-Path $mdFile

pandoc `
    --reference-doc=reference-doc.docx `
    --template=template.openxml `
    --number-sections --toc --lot --lof `
    --citeproc `
    --metadata=link-citations:true `
    -t docx+native_numbering `
    --resource-path=$resourcePath `
    $mdFile -o $docxFile

# If pandoc succeeds, open file in Word.
if ($?) {
    Invoke-Item $docxFile
}
