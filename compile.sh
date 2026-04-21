#!/usr/bin/env bash

# Compile a Markdown document into a Word document

set -euo pipefail

# Get the first argument as the file name.
MDFILE=$1

# Check if the argument was provided.
if [ -z "${MDFILE}" ]; then
    echo "Usage: .\compile.sh <filename>"
    exit 1
fi

# Check if the markdown file exists.
if ! [ -f "${MDFILE}" ]; then
    echo "Error: File '${MDFILE}' does not exist."
    exit 1
fi

FILENAME=$(basename -- "${MDFILE}")
EXTENSION="${FILENAME##*.}"
FILENAME="${FILENAME%.*}"

# Check if the file extension is `.md`.
if ! [ "${EXTENSION}" == "md" ]; then
    echo "Error: File '${MDFILE}' is not a Markdown file."
    exit 1
fi

if ! [ -f "reference-doc.docx" ]; then
    echo "Error: File 'reference-doc.docx' does not exist. Run '.\edit-reference-doc.ps1' and try again."
    exit 1
fi

DOCXFILE="${FILENAME}.docx"
RESOURCEPATH=$(dirname -- "${MDFILE}")

pandoc \
    --reference-doc=reference-doc.docx \
    --template=template.openxml \
    --number-sections --toc --lot --lof \
    --citeproc \
    --metadata=link-citations:true \
    -t docx+native_numbering \
    --resource-path=${RESOURCEPATH} \
    ${MDFILE} -o ${DOCXFILE}

echo "Successfully created ${DOCXFILE} from ${MDFILE}"
