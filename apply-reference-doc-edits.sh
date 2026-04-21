#!/usr/bin/env bash

# Extract a reference Word document into its constituent files

set -euo pipefail

docx_path="$PWD/reference-doc.docx"
target_dir="$PWD/reference-doc"
temp_dir="$(mktemp -d)"
temp_extract_dir="$temp_dir/reference-doc-temp"

cleanup() {
  rm -rf "$temp_dir"
}
trap cleanup EXIT

if [[ ! -f "$docx_path" ]]; then
  echo "The file reference-doc.docx does not exist." >&2
  exit 1
fi

mkdir -p "$temp_extract_dir"
unzip -q "$docx_path" -d "$temp_extract_dir"

mapfile -t source_files < <(cd "$temp_extract_dir" && find . -type f -printf '%P\n' | sort)
mapfile -t target_files < <(cd "$target_dir" && find . -type f -printf '%P\n' | sort)

declare -A source_set=()
for rel_path in "${source_files[@]}"; do
  source_set["$rel_path"]=1
done

for rel_path in "${target_files[@]}"; do
  if [[ -z "${source_set[$rel_path]+x}" ]]; then
    full_path="$target_dir/$rel_path"
    echo "Deleting $rel_path since it is not present in the updated reference-doc.docx"
    rm -f "$full_path"
  fi
done

protected_files=(
  "docProps/app.xml"
  "docProps/core.xml"
  "word/settings.xml"
  "word/glossary/settings.xml"
)

declare -A protected_set=()
for protected in "${protected_files[@]}"; do
  protected_set["$protected"]=1
done

for rel_path in "${source_files[@]}"; do
  source_path="$temp_extract_dir/$rel_path"
  target_path="$target_dir/$rel_path"

  if [[ -n "${protected_set[$rel_path]+x}" && -f "$target_path" ]]; then
    echo "Skipping $rel_path as it likely contains only non-functional changes"
    continue
  fi

  mkdir -p "$(dirname "$target_path")"
  echo "Copying $rel_path"
  cp "$source_path" "$target_path"
done

echo "Successfully extracted contents of reference-doc.docx into reference-doc/"
