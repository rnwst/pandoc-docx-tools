<#
.SYNOPSIS
Formats the contents of an unzipped DOCX to improve readability and enable effective version control.

.DESCRIPTION
This script recursively processes XML and .rels files in a given folder.
Each XML element is placed on a new line. The indent is set at two spaces.
XML namespace declarations on the root element are placed on separate lines.
Superfluous element attributes such as w14:paraId and w:rsidR are removed.

.PARAMETER Path
(Optional) The path to the folder containing the XML files to format. Defaults to '.\reference-doc' if not provided.

.EXAMPLE
.\format-openxml.ps1
Formats files in the default '.\reference-doc' folder.

.EXAMPLE
.\format-openxml.ps1 .\template
Formats files in the specified folder.

.NOTES
Author: R. N. West
PowerShell version: 5.1+

.LINK
https://www.brandwares.com/downloads/Open-XML-Explained.pdf is a good introduction to OpenXML.
See http://officeopenxml.com/ for an OpenXML element reference.
#>


[CmdletBinding()]
param (
    [Parameter(Position = 0)]
    [string]$Path = ".\reference-doc"
)


if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
    Write-Warning "The directory '$Path' does not exist. Exiting script."
    return
}


# These attributes appear to be superfluous, change on every save, and should therefore be removed to make version-control easier.
function Remove-Attributes {
    param ($node)
    if ($node.Attributes) {
        $attrsToRemove = @('w14:paraId', 'w14:textId', 'wp14:editId', 'wp14:anchorId', 'w:storeItemID')
        
        # Collect attribute names that start with 'w:rsid'
        foreach ($attr in $node.Attributes) {
            if ($attr.Name -like 'w:rsid*') {
                $attrsToRemove += $attr.Name
            }
        }

        # Remove the matching attributes
        foreach ($attrName in $attrsToRemove) {
            $node.Attributes.RemoveNamedItem($attrName) | Out-Null
        }
    }

    foreach ($child in $node.ChildNodes) {
        Remove-Attributes -node $child
    }
}


# These elements appear to be superfluous, change on every save, and should therefore be removed to make version-control easier.
function Remove-Elements {
    param ($xml)
    $eltsToRemove = @("w:rsid", "w:id")
    $namespaceManager = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $namespaceManager.AddNamespace("w", $xml.DocumentElement.NamespaceURI)
    foreach ($eltName in $eltsToRemove) {
        $elts = $xml.SelectNodes("//$eltName", $namespaceManager)
        foreach ($elt in $elts) {
            $elt.ParentNode.RemoveChild($elt) | Out-Null
        }
    }
}


Get-ChildItem -Force -LiteralPath $Path  -File -Recurse | Where-Object { $_.Name -match '\.xml$' -or $_.Name -match '\.rels$' } | ForEach-Object {
    Write-Host "Formatting $($_.FullName)"

    # Read XML.
    $xml = [xml](Get-Content -LiteralPath $_.FullName -Encoding utf8)

    # Remove superfluous attributes.
    Remove-Attributes -node $xml.DocumentElement

    # Remove superfluous elements.
    Remove-Elements -xml $xml

    # Set encoding to avoid BoM at beginning (git sees this as a change otherwise).
    $utf8NoBomEncoding = New-Object System.Text.UTF8Encoding($false)
    $xmlWriterSettings = New-Object System.Xml.XmlWriterSettings
    $xmlWriterSettings.Encoding = $utf8NoBomEncoding

    # Indent nodes 2 spaces.
    $xmlWriterSettings.Indent = $true

    # Write XML to string.
    $stringWriter = New-Object System.IO.StringWriter
    $xmlWriter = [System.Xml.XmlWriter]::Create($stringWriter, $xmlWriterSettings)
    $xml.WriteContentTo($xmlWriter)
    $xmlWriter.Flush()
    $xmlWriter.Close()
    $xmlString = $stringWriter.ToString()

    # Put XML namespace declarations on the root element on new lines.
    # First, put the first declaration on a new line.
    $xmlString = $xmlString -replace '(?m)^(<[\w:]+) (xmlns(:\w+)?="[^"]*")', "`$1`r`n  `$2"
    # Now the remaining ones. This needs to be done recursively due to overlapping patterns.
    do {
        $prev = $xmlString
        $xmlString = $xmlString -replace '(?m)^  (xmlns(?::\w+)?="[^"]*") (xmlns(:\w+)?="[^"]*")', "  `$1`r`n  `$2"
    } while ($xmlString -ne $prev)
    # Put the `mc:Ignorable` attributes on a new line as well.
    $xmlString = $xmlString -replace '(?m)^  (xmlns(?::\w+)?="[^"]*") (mc:Ignorable="[^"]*")', "  `$1`r`n  `$2"

    # Write formatted XML back to file.
    # Why not Set-Content -LiteralPath $_.FullName -Encoding utf8 $xmlString?
    # Because it adds a BOM to the file.
    $null = New-Item -Force $_.FullName -Value ($xmlString + "`r`n")    
}
