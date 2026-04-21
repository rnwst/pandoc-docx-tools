#!/usr/bin/env bash

# Build a reference Word document from its constituent files

set -euo pipefail

# Set paths relative to current directory
source_dir="$PWD/reference-doc"
zip_path="$PWD/reference-doc.zip"
docx_path="$PWD/reference-doc.docx"

# Cleanup any existing files
rm -f "$zip_path" "$docx_path"

# Create the ZIP archive.
# First pass: add everything except `word/media/*` with default compression.
(
  cd "$source_dir"
  zip -q -r "$zip_path" . -x "word/media/*"
)

# Second pass: add `word/media/*` without compression.
if [[ -d "$source_dir/word/media" ]]; then
  (
    cd "$source_dir"
    zip -q -0 -r "$zip_path" "word/media"
  )
fi

# Rename the .zip to .docx
mv "$zip_path" "$docx_path"

echo "Successfully constructed reference-doc.docx from the contents of reference-doc/"

# Open with Microsoft Word
# ii $docxPath
