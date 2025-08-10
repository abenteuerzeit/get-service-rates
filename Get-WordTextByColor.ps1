#Requires -Modules Microsoft.PowerShell.Management, Microsoft.PowerShell.Utility

<#
.SYNOPSIS
    Extracts text from a Word document based on its color.

.DESCRIPTION
    This script opens a Microsoft Word document, iterates through its content (including headers, footers, footnotes, and endnotes),
    and groups text fragments by their font color. It then returns an object containing the text, word count, and fragment count for each color found.

.PARAMETER FilePath
    The path to the Word document (.doc or .docx) to be processed. This parameter is mandatory.

.PARAMETER IncludePositions
    A switch parameter that, if present, includes the start and end character positions for each text fragment in the output.

.EXAMPLE
    .\Get-WordTextByColor.ps1 -FilePath "C:\path\to\document.docx"
    Processes the specified document and outputs the text grouped by color.

.EXAMPLE
    .\Get-WordTextByColor.ps1 -FilePath ".\mydoc.docx" -IncludePositions -Verbose
    Processes the document, includes character positions in the output, and shows verbose status messages during execution.

.EXAMPLE
    .\Get-WordTextByColor.ps1 -FilePath "report.docx" -Debug
    Processes the document and shows detailed debug messages for troubleshooting.

.OUTPUTS
    An array of PSCustomObjects. Each object represents a color and contains the fragments, word count, and fragment count for that color.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$FilePath,

    [switch]$IncludePositions,

    [switch]$Help
)

if ($Help -or [string]::IsNullOrWhiteSpace($FilePath)) {
    Write-Host "Usage:" -ForegroundColor Green
    Write-Host "  .\Get-WordTextByColor.ps1 <FilePath> [-IncludePositions] [-Verbose] [-Debug]" -ForegroundColor White
    Write-Host ""
    Write-Host "Parameters:" -ForegroundColor Green
    Write-Host "  FilePath          - Path to Word document (.docx/.doc)" -ForegroundColor White
    Write-Host "  -IncludePositions - Include character positions in output" -ForegroundColor White
    Write-Host "  -Verbose          - Show verbose status messages" -ForegroundColor White
    Write-Host "  -Debug            - Show detailed debug information for troubleshooting" -ForegroundColor White
    Write-Host "  -Help             - Show this help message" -ForegroundColor White
    exit
}

function Get-WordTextByColor {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        [switch]$IncludePositions
    )

    if (-not (Test-Path $FilePath)) {
        throw "File does not exist: $FilePath"
    }

    if (-not ($FilePath -match '\.docx?$')) {
        throw "File must be a Word document (.doc or .docx): $FilePath"
    }

    $AbsoluteFilePath = (Resolve-Path -Path $FilePath).Path
    $colorGroups = @{}
    $wordApp = $null
    $document = $null

    try {
        Write-Verbose "Creating Word Application COM object..."
        $wordApp = New-Object -ComObject Word.Application
        $wordApp.Visible = $false
        $wordApp.DisplayAlerts = 0

        Write-Verbose "Opening document: $AbsoluteFilePath"
        $document = $wordApp.Documents.Open($AbsoluteFilePath, $false, $true)

        Write-Verbose "Processing document content..."
        $processedFragments = 0

        Write-Verbose "Processing main document body..."
        $processedFragments = ProcessRangeByColor -range $document.Content -colorGroups ([ref]$colorGroups) -processedFragments $processedFragments -includePositions $IncludePositions -source "Main"

        if ($document.Footnotes.Count -gt 0) {
            Write-Verbose "Processing $($document.Footnotes.Count) footnotes..."
            for ($f = 1; $f -le $document.Footnotes.Count; $f++) {
                $footnote = $document.Footnotes.Item($f)
                Write-Debug "Processing footnote $f"
                $processedFragments = ProcessRangeByColor -range $footnote.Range -colorGroups ([ref]$colorGroups) -processedFragments $processedFragments -includePositions $IncludePositions -source "Footnote $f"
            }
        }

        if ($document.Endnotes.Count -gt 0) {
            Write-Verbose "Processing $($document.Endnotes.Count) endnotes..."
            for ($e = 1; $e -le $document.Endnotes.Count; $e++) {
                $endnote = $document.Endnotes.Item($e)
                Write-Debug "Processing endnote $e"
                $processedFragments = ProcessRangeByColor -range $endnote.Range -colorGroups ([ref]$colorGroups) -processedFragments $processedFragments -includePositions $IncludePositions -source "Endnote $e"
            }
        }

        Write-Verbose "Processing headers and footers across $($document.Sections.Count) sections..."
        for ($s = 1; $s -le $document.Sections.Count; $s++) {
            $section = $document.Sections.Item($s)
            
            foreach ($headerType in 1..3) {
                try {
                    $header = $section.Headers.Item($headerType)
                    if ($header.Exists -and $header.Range.Text.Trim().Length -gt 0) {
                        Write-Debug "Processing section $s header type $headerType"
                        $processedFragments = ProcessRangeByColor -range $header.Range -colorGroups ([ref]$colorGroups) -processedFragments $processedFragments -includePositions $IncludePositions -source "Header S$s T$headerType"
                    }
                } catch { Write-Debug "No header of type $headerType in section $s." }
            }
            
            foreach ($footerType in 1..3) {
                try {
                    $footer = $section.Footers.Item($footerType)
                    if ($footer.Exists -and $footer.Range.Text.Trim().Length -gt 0) {
                        Write-Debug "Processing section $s footer type $footerType"
                        $processedFragments = ProcessRangeByColor -range $footer.Range -colorGroups ([ref]$colorGroups) -processedFragments $processedFragments -includePositions $IncludePositions -source "Footer S$s T$footerType"
                    }
                } catch { Write-Debug "No footer of type $footerType in section $s." }
            }
        }

        Write-Verbose "Processing complete. Found $($colorGroups.Keys.Count) different colors, $processedFragments total fragments."

        $results = @{}
        foreach ($color in $colorGroups.Keys) {
            $fragments = $colorGroups[$color]
            $wordCount = 0
            
            foreach ($fragment in $fragments) {
                $text = if ($fragment -is [hashtable]) { $fragment.Text } else { $fragment }
                $words = $text -split '\s+' | Where-Object { $_ -ne '' }
                $wordCount += $words.Count
            }
            
            $results[$color] = @{
                Fragments = $fragments
                WordCount = $wordCount
                FragmentCount = $fragments.Count
            }
        }
        
        Write-Debug "--- Color Breakdown ---"
        foreach ($color in $results.Keys | Sort-Object) {
            $fragmentCount = $results[$color].FragmentCount
            $wordCount = $results[$color].WordCount
            Write-Debug "  $color`: $fragmentCount fragments, $wordCount words"
        }
        
        return $results
        
    } catch {
        Write-Error "Failed to process Word document: $_"
        throw
    } finally {
        if ($document) {
            Write-Verbose "Closing document..."
            $document.Close($false) # Close without saving changes
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($document) | Out-Null
        }
        
        if ($wordApp) {
            Write-Verbose "Closing Word application..."
            $wordApp.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wordApp) | Out-Null
        }
        
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function ProcessRangeByColor {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $range,
        [Parameter(Mandatory = $true)]
        [ref]$colorGroups,
        [Parameter(Mandatory = $true)]
        [int]$processedFragments,
        [Parameter(Mandatory = $true)]
        [bool]$includePositions,
        [Parameter(Mandatory = $true)]
        [string]$source
    )
    
    $text = $range.Text
    if ([string]::IsNullOrWhiteSpace($text)) {
        return $processedFragments
    }
    
    $startPos = $range.Start
    $endPos = $range.End
    $currentPos = $startPos
    
    Write-Debug "Processing '$source' range: $startPos to $endPos ($($text.Length) chars)"
    
    while ($currentPos -lt $endPos) {
        try {
            $charRange = $range.Document.Range($currentPos, [Math]::Min($currentPos + 1, $endPos))
            $char = $charRange.Text
            
            if ([string]::IsNullOrEmpty($char) -or $char -match '[\r\n\v\f]') {
                $currentPos++
                continue
            }
            
            $fontColor = $charRange.Font.Color
            $colorKey = Get-ColorName -ColorValue $fontColor
            
            $fragmentStart = $currentPos
            $fragmentEnd = $currentPos
            $chunkSize = 50
            
            while ($fragmentEnd -lt $endPos - 1) {
                $testEnd = [Math]::Min($fragmentEnd + $chunkSize, $endPos - 1)
                $testRange = $range.Document.Range($fragmentEnd + 1, $testEnd + 1)
                
                if ($testRange.Font.Color -eq $fontColor) {
                    $fragmentEnd = $testEnd
                    $chunkSize = [Math]::Min($chunkSize * 2, 500)
                } else {
                    $chunkSize = 1
                    while ($fragmentEnd -lt $testEnd) {
                        $nextRange = $range.Document.Range($fragmentEnd + 1, $fragmentEnd + 2)
                        if ($nextRange.Font.Color -eq $fontColor) {
                            $fragmentEnd++
                        } else {
                            break # Found the color change.
                        }
                    }
                    break # Exit the outer loop since we found the end.
                }
            }
            
            $fragmentRange = $range.Document.Range($fragmentStart, $fragmentEnd + 1)
            $textFragment = $fragmentRange.Text
            
            $textFragment = ($textFragment -replace '[\r\n\v\f]' , " ").Trim()
            
            if (![string]::IsNullOrWhiteSpace($textFragment)) {
                if (-not $colorGroups.Value.ContainsKey($colorKey)) {
                    $colorGroups.Value[$colorKey] = @()
                }
                
                $fragment = if ($includePositions) {
                    @{
                        Text = $textFragment
                        StartPosition = $fragmentStart
                        EndPosition = $fragmentEnd
                        Length = $textFragment.Length
                        Source = $source
                    }
                } else {
                    $textFragment
                }
                
                $colorGroups.Value[$colorKey] += $fragment
                $processedFragments++
                
                $preview = $textFragment.Substring(0, [Math]::Min(50, $textFragment.Length))
                Write-Debug "  Added fragment: '$preview...' ($colorKey) from '$source'"
            }
            
            $currentPos = $fragmentEnd + 1
            
        } catch {
            Write-Debug "Error processing character at position $currentPos in '$source`: $_. Skipping."
            $currentPos++ 
        }
    }
    
    return $processedFragments
}

function Get-ColorName {
    param([long]$ColorValue)
    
    # Map common WdColor index values to names.
    switch ($ColorValue) {
        -16777216 { return "Automatic" }
        0 { return "Black" }
        255 { return "Red" }
        65280 { return "Green" }
        16711680 { return "Blue" }
        65535 { return "Yellow" }
        16711935 { return "Magenta" }
        16776960 { return "Cyan" }
        8388608 { return "DarkBlue" }
        32768 { return "DarkGreen" }
        128 { return "DarkRed" }
        8421376 { return "DarkCyan" }
        32896 { return "DarkYellow" }
        8388736 { return "DarkMagenta" }
        12632256 { return "LightGray" }
        8421504 { return "DarkGray" }
        16777215 { return "White" }
        default {
            if ($ColorValue -lt 0) {
                return "ThemeColor($ColorValue)"
            }
            $red = $ColorValue -band 0xFF
            $green = ($ColorValue -shr 8) -band 0xFF
            $blue = ($ColorValue -shr 16) -band 0xFF
            return "RGB({0:X2}{1:X2}{2:X2})" -f $red, $green, $blue
        }
    }
}

try {
    $results = Get-WordTextByColor -FilePath $FilePath -IncludePositions:$IncludePositions
    
    if ($results.Keys.Count -eq 0) {
        Write-Warning "No colored text fragments were found. Try running with -Debug to see processing details."
        exit 1
    }
    
    $outputObjects = @()
    
    foreach ($color in $results.Keys | Sort-Object) {
        $colorData = $results[$color]
        Write-Host "`nColor: $color" -ForegroundColor Magenta
        Write-Host "Word Count: $($colorData.WordCount)" -ForegroundColor Yellow
        Write-Host "Fragments: $($colorData.FragmentCount)" -ForegroundColor White
        
        $colorGroup = [PSCustomObject]@{
            Color = $color
            WordCount = $colorData.WordCount
            FragmentCount = $colorData.FragmentCount
            Fragments = $colorData.Fragments
        }
        $outputObjects += $colorGroup
        
        $examples = $colorData.Fragments | Select-Object -First 3
        foreach ($fragment in $examples) {
            $previewText = if ($fragment -is [hashtable]) { $fragment.Text } else { $fragment }
            $preview = $previewText.Substring(0, [Math]::Min(50, $previewText.Length))
            
            if ($fragment -is [hashtable]) {
                Write-Host "  - `"$preview...`" (Pos: $($fragment.StartPosition)-$($fragment.EndPosition), $($fragment.Source))" -ForegroundColor Gray
            } else {
                Write-Host "  - `"$preview...`"" -ForegroundColor Gray
            }
        }
        
        if ($colorData.FragmentCount -gt 3) {
            Write-Host "  ... and $($colorData.FragmentCount - 3) more fragments" -ForegroundColor DarkGray
        }
    }
    
    return $outputObjects
    
} catch {
    Write-Error "Script execution failed: $_"
    exit 1
}
