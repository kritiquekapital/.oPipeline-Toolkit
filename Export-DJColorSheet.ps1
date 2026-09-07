<#
.SYNOPSIS
    Fresh export of DJ tag-sheet metadata (current Genre etc.) PLUS the decoded
    Serato track Color, straight from the file tags -- no stale CSV involved.

.DESCRIPTION
    This is a read-only companion to DJ-TagSheet-v4.ps1. It re-scans -Root from
    scratch (so Genre/etc. reflect whatever you've most recently tagged in
    Serato) and additionally decodes each file's Serato track color.

    SERATO COLOR FORMAT
        Serato stores per-track color in a "Serato Markers_" ID3 GEOB frame,
        using a custom 4-byte encoding ("serato32") -- NOT plain RGB bytes.
        On top of that, the value Serato *stores* is not the value Serato
        *displays* in the library column; there's a second transform.
        Source: https://github.com/Holzhaus/serato-tags (the same reverse-
        engineering docs Mixxx's own Serato import is built on).

        This script:
          1. reads the raw "Serato Markers_" GEOB frame bytes via TagLib-Sharp
          2. skips the version + cue/loop entries to find the trailing 4-byte
             track color field
          3. decodes serato32 -> 24-bit "stored" RGB
          4. applies Serato's stored->displayed transform -> the color you'd
             actually see in Serato's library view
          5. maps that displayed hex to the closest match in Serato's known
             20-swatch track-color palette, and reports the name AND the hex,
             plus how close the match was (should be 0 for real Serato colors)

        IMPORTANT: the palette *names* (Magenta, Pink, Red-Orange, ...) are
        not something Serato stores anywhere -- I matched them to your
        documented anchor palette by hue order. Sanity-check the "Validation"
        section this script prints against known anchor tracks before
        trusting ColorName on the full library.

.EXAMPLE
    .\Export-DJColorSheet.ps1 -Root "E:\Music"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Root,

    [string]$OutputDir,   # where the CSV goes. Defaults to $Root if not given (old behavior).

    [string[]]$Extensions = @('.mp3'),   # Serato Markers_ is an ID3v2/MP3(& AIFF) feature; see note below for other formats

    [string]$TagLibPath
)

$ErrorActionPreference = 'Stop'

function Ensure-TagLibSharp {
    param([string]$ExplicitPath)
    if ([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'taglib-sharp' -or $_.GetName().Name -eq 'TagLibSharp' }) {
        return
    }
    if ($ExplicitPath) {
        if (-not (Test-Path -LiteralPath $ExplicitPath)) { throw "TagLibPath does not exist: $ExplicitPath" }
        Add-Type -Path (Resolve-Path -LiteralPath $ExplicitPath)
        return
    }
    $libDir = Join-Path $PSScriptRoot 'lib'
    $dllPath = Join-Path $libDir 'TagLibSharp.dll'
    if (-not (Test-Path -LiteralPath $dllPath)) {
        Write-Host 'TagLibSharp.dll is not here yet. Downloading it once...' -ForegroundColor Yellow
        New-Item -ItemType Directory -Path $libDir -Force | Out-Null
        $packagePath = Join-Path $libDir 'TagLibSharp.nupkg'
        $extractPath = Join-Path $libDir '_taglib_extract'
        try {
            Invoke-WebRequest -Uri 'https://www.nuget.org/api/v2/package/TagLibSharp' -OutFile $packagePath -UseBasicParsing
            if (Test-Path -LiteralPath $extractPath) { Remove-Item -LiteralPath $extractPath -Recurse -Force }
            # FIX: Expand-Archive refuses anything not literally named ".zip" --
            # it checks the file EXTENSION, not the actual content, and a
            # .nupkg is a real zip file under a different name. Extract via
            # .NET's ZipFile class instead, which doesn't care about the name.
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            [System.IO.Compression.ZipFile]::ExtractToDirectory($packagePath, $extractPath)
            $candidates = Get-ChildItem -LiteralPath $extractPath -Recurse -Filter 'TagLibSharp.dll'
            $candidate = if ($PSVersionTable.PSVersion.Major -lt 6) {
                $candidates | Sort-Object @{ Expression = { if ($_.FullName -match '\\net4') {0} elseif ($_.FullName -match '\\netstandard') {1} else {2} } } | Select-Object -First 1
            } else {
                $candidates | Sort-Object @{ Expression = { if ($_.FullName -match '\\netstandard') {0} elseif ($_.FullName -match '\\net4') {1} else {2} } } | Select-Object -First 1
            }
            if (-not $candidate) { throw 'TagLibSharp.dll was not found inside the NuGet package.' }
            Copy-Item -LiteralPath $candidate.FullName -Destination $dllPath -Force
        } finally {
            Remove-Item -LiteralPath $packagePath -Force -ErrorAction SilentlyContinue
            Remove-Item -LiteralPath $extractPath -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    Add-Type -Path $dllPath
}

# ---------------------------------------------------------------------------
# Serato "Serato Markers_" -> track color decoding
# ---------------------------------------------------------------------------

# Serato's 20-swatch track-color palette: stored(picker) hex -> displayed(library) hex -> name.
# Names are OUR mapping (hue-order match to the documented anchor palette), not
# something Serato itself labels -- validate against known anchor tracks.
$Script:ColorPalette = @(
    [PSCustomObject]@{ Stored = 0xFF99FF; Displayed = 0x993399; Name = 'Magenta' }
    [PSCustomObject]@{ Stored = 0xFF99DD; Displayed = 0x993377; Name = 'Pink' }
    [PSCustomObject]@{ Stored = 0xFF99BB; Displayed = 0x993355; Name = 'Pink-Red' }
    [PSCustomObject]@{ Stored = 0xFF9999; Displayed = 0x993333; Name = 'Red' }
    [PSCustomObject]@{ Stored = 0xFFBB99; Displayed = 0x995533; Name = 'Red-Orange' }
    [PSCustomObject]@{ Stored = 0xFFDD99; Displayed = 0x997733; Name = 'Orange' }
    [PSCustomObject]@{ Stored = 0xFFFF99; Displayed = 0x999933; Name = 'Gold' }
    [PSCustomObject]@{ Stored = 0xDDFF99; Displayed = 0x779933; Name = 'Olive' }
    [PSCustomObject]@{ Stored = 0xBBFF99; Displayed = 0x559933; Name = 'Green' }
    [PSCustomObject]@{ Stored = 0x99FF99; Displayed = 0x339933; Name = 'Lime' }
    [PSCustomObject]@{ Stored = 0x99FFBB; Displayed = 0x339955; Name = 'Lime-Teal' }
    [PSCustomObject]@{ Stored = 0x99FFDD; Displayed = 0x339977; Name = 'Teal' }
    [PSCustomObject]@{ Stored = 0x99FFFF; Displayed = 0x339999; Name = 'Cyan' }
    [PSCustomObject]@{ Stored = 0x99DDFF; Displayed = 0x337799; Name = 'Sky Blue' }
    [PSCustomObject]@{ Stored = 0x99BBFF; Displayed = 0x335599; Name = 'Blue' }
    [PSCustomObject]@{ Stored = 0x9999FF; Displayed = 0x333399; Name = 'Blue-Dark' }
    [PSCustomObject]@{ Stored = 0xBB99FF; Displayed = 0x553399; Name = 'Purple' }
    [PSCustomObject]@{ Stored = 0xDD99FF; Displayed = 0x773399; Name = 'Violet' }
    [PSCustomObject]@{ Stored = 0xBBBBBB; Displayed = 0x555555; Name = 'Gray' }
    [PSCustomObject]@{ Stored = 0x999999; Displayed = 0x090909; Name = 'Black' }
)

# Files known (from color_genres.xlsx anchors) to carry a specific color, for
# self-validation. Matched by filename substring.
$Script:AnchorTracks = @(
    [PSCustomObject]@{ Match = 'knock kncok'; Expected = 'Magenta' }
    [PSCustomObject]@{ Match = 'Torn, out of time'; Expected = 'Lime' }
    [PSCustomObject]@{ Match = 'On the gas, shoot to kill'; Expected = 'Teal' }
    [PSCustomObject]@{ Match = 'Moment vierra cloud'; Expected = 'Cyan' }
    [PSCustomObject]@{ Match = 'Inside Out'; Expected = 'Sky Blue' }
    [PSCustomObject]@{ Match = 'fall back to nothing'; Expected = 'Blue-Dark' }
)

function Decode-Serato32 {
    param([byte]$B1, [byte]$B2, [byte]$B3, [byte]$B4)
    $val = ([uint32]($B1 -band 0x7F) -shl 21) -bor `
           ([uint32]($B2 -band 0x7F) -shl 14) -bor `
           ([uint32]($B3 -band 0x7F) -shl 7)  -bor `
           ([uint32]($B4 -band 0x7F))
    return $val -band 0xFFFFFF
}

function Get-DisplayedTrackColorCode {
    param([uint32]$Stored)
    if ($Stored -eq 0xFFFFFF) { return $null }  # "no color set" sentinel
    if ($Stored -eq 0x999999) { return 0x090909 }
    if ($Stored -eq 0x000000) { return 0x333333 }
    if ($Stored -lt 0x666666) { return $Stored + 0x99999A }
    return $Stored - 0x666666
}

function Get-ColorName {
    param([uint32]$DisplayedCode)
    $match = $Script:ColorPalette | Where-Object { $_.Displayed -eq $DisplayedCode } | Select-Object -First 1
    if ($match) { return [PSCustomObject]@{ Name = $match.Name; Distance = 0 } }

    # Not an exact palette hit (unusual for real Serato data) -- report nearest
    # by RGB distance so it's visible rather than silently blank.
    $r1 = ($DisplayedCode -shr 16) -band 0xFF; $g1 = ($DisplayedCode -shr 8) -band 0xFF; $b1 = $DisplayedCode -band 0xFF
    $best = $null; $bestDist = [double]::MaxValue
    foreach ($p in $Script:ColorPalette) {
        $r2 = ($p.Displayed -shr 16) -band 0xFF; $g2 = ($p.Displayed -shr 8) -band 0xFF; $b2 = $p.Displayed -band 0xFF
        $dist = [math]::Sqrt((($r1 - $r2) * ($r1 - $r2)) + (($g1 - $g2) * ($g1 - $g2)) + (($b1 - $b2) * ($b1 - $b2)))
        if ($dist -lt $bestDist) { $bestDist = $dist; $best = $p.Name }
    }
    return [PSCustomObject]@{ Name = "$best (approx)"; Distance = [math]::Round($bestDist, 1) }
}

function Get-SeratoTrackColor {
    <#
    Returns $null if the file has no "Serato Markers_" tag at all (never
    color-tagged in Serato), or a PSCustomObject with StoredHex, DisplayedHex,
    ColorName if it does.

    FIX: this used to call [TagLib.File]::Create($Path) itself, meaning the
    file got opened TWICE per track -- once here, once again by the caller
    for the ordinary tag fields (Genre/BPM/etc). Doubled disk I/O and
    TagLib overhead across the whole library for no reason. Now it takes
    the already-open Id3v2 tag object instead of reopening the file.
    #>
    param($Id3)

    try {
        if (-not $Id3) { return $null }

        $frame = [TagLib.Id3v2.GeneralEncapsulatedObjectFrame]::Get($Id3, 'Serato Markers_', $false)
        if (-not $frame) {
            # Legacy tag isn't present. If only Markers2 exists, flag that distinctly
            # rather than silently reporting "not tagged" -- it means color data
            # likely IS there, just needs a different (not-yet-built) decoder.
            $markers2 = [TagLib.Id3v2.GeneralEncapsulatedObjectFrame]::Get($Id3, 'Serato Markers2', $false)
            if ($markers2) {
                return [PSCustomObject]@{ StoredHex = ''; DisplayedHex = ''; ColorName = '(Markers2 only -- not decoded)'; MatchDistance = '' }
            }
            return $null
        }

        [byte[]]$data = $null
        $dataLen = $frame.Data.Count
        $data = New-Object byte[] $dataLen
        for ($i = 0; $i -lt $dataLen; $i++) { $data[$i] = $frame.Data[$i] }
        if ($data.Length -lt 6) { return $null }

        # bytes 0-1: version (expect 0x01 0x01) -- not validated strictly, just skipped
        $pos = 2

        # bytes 2-5: big-endian u32 entry count
        $count = ([uint32]$data[$pos] -shl 24) -bor ([uint32]$data[$pos + 1] -shl 16) -bor `
                 ([uint32]$data[$pos + 2] -shl 8) -bor [uint32]$data[$pos + 3]
        $pos += 4

        # each marker entry is exactly 22 bytes (5 start-pos + 5 end-pos + 6 unknown + 4 color + 1 type + 1 locked)
        $pos += ($count * 22)

        if ($data.Length -lt ($pos + 4)) { return $null }

        $stored = Decode-Serato32 -B1 $data[$pos] -B2 $data[$pos + 1] -B3 $data[$pos + 2] -B4 $data[$pos + 3]
        $displayed = Get-DisplayedTrackColorCode -Stored $stored

        if ($null -eq $displayed) {
            return [PSCustomObject]@{ StoredHex = ('{0:X6}' -f $stored); DisplayedHex = ''; ColorName = '(no color set)'; MatchDistance = 0 }
        }

        $nameInfo = Get-ColorName -DisplayedCode $displayed
        return [PSCustomObject]@{
            StoredHex     = ('{0:X6}' -f $stored)
            DisplayedHex  = ('{0:X6}' -f $displayed)
            ColorName     = $nameInfo.Name
            MatchDistance = $nameInfo.Distance
        }
    } catch {
        return $null
    }
}

function Get-StringValue {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [System.Array]) { return (($Value | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join '; ') }
    return [string]$Value
}

# ---------------------------------------------------------------------------
# Main scan
# ---------------------------------------------------------------------------
if (-not (Test-Path -LiteralPath $Root -PathType Container)) { throw "Root folder does not exist: $Root" }
Ensure-TagLibSharp -ExplicitPath $TagLibPath

$Root = (Resolve-Path -LiteralPath $Root).Path
$extLower = $Extensions | ForEach-Object { $_.ToLowerInvariant() }

Write-Host "Scanning $Root ..." -ForegroundColor Cyan
$files = @(Get-ChildItem -LiteralPath $Root -File -Recurse -ErrorAction SilentlyContinue |
    Where-Object { $extLower -contains $_.Extension.ToLowerInvariant() })

if ($files.Count -eq 0) {
    Write-Warning "No matching audio files found under $Root"
    exit 0
}

# --- One-time frame diagnostic on the first couple of files, so we find out
#     immediately whether "Serato Markers_" is really the frame name in use,
#     instead of silently getting zero matches across the whole library again. ---
Write-Host "`n--- Frame diagnostic (first 2 files) ---" -ForegroundColor Yellow
foreach ($diagFile in ($files | Select-Object -First 2)) {
    try {
        $diagTagFile = [TagLib.File]::Create($diagFile.FullName)
        $diagId3 = $diagTagFile.GetTag([TagLib.TagTypes]::Id3v2, $false)
        Write-Host "`n$($diagFile.Name)" -ForegroundColor Cyan
        if (-not $diagId3) {
            Write-Host "  (no ID3v2 tag at all)"
        } else {
            $allFrameIds = $diagId3.GetFrames() | ForEach-Object { $_.FrameId.ToString() }
            Write-Host "  All frame IDs: $($allFrameIds -join ', ')"

            foreach ($wantedDesc in @('Serato Markers_', 'Serato Markers2')) {
                $gf = [TagLib.Id3v2.GeneralEncapsulatedObjectFrame]::Get($diagId3, $wantedDesc, $false)
                if ($gf) {
                    Write-Host "  FOUND via Get(): Description='$wantedDesc'  DataLength=$($gf.Data.Count)" -ForegroundColor Green
                } else {
                    Write-Host "  NOT found via Get(): Description='$wantedDesc'" -ForegroundColor DarkYellow
                }
            }
        }
        $diagTagFile.Dispose()
    } catch {
        Write-Host "  (error reading file: $($_.Exception.Message))"
    }
}
Write-Host "--- end diagnostic ---`n" -ForegroundColor Yellow

Write-Host "Reading fresh Genre + Color data from $($files.Count) file(s)..." -ForegroundColor Cyan
$rows = New-Object System.Collections.Generic.List[object]
$index = 0
$anchorHits = New-Object System.Collections.Generic.List[object]

foreach ($file in $files) {
    $index++
    if (($index % 25) -eq 0 -or $index -eq $files.Count) {
        Write-Progress -Activity 'Exporting fresh color/genre sheet' -Status "$index of $($files.Count)" -PercentComplete (($index / $files.Count) * 100)
    }

    $tag = $null
    $tagFile = $null
    try {
        $tagFile = [TagLib.File]::Create($file.FullName)
        $tag = $tagFile.Tag
    } catch {
        continue
    }

    # FIX: reuse this same open file's Id3v2 tag for color decoding instead
    # of reopening the file a second time (see Get-SeratoTrackColor).
    $id3ForColor = $tagFile.GetTag([TagLib.TagTypes]::Id3v2, $false)
    $color = Get-SeratoTrackColor -Id3 $id3ForColor
    $tagFile.Dispose()

    $row = [PSCustomObject]@{
        FullPath      = $file.FullName
        FileName      = $file.Name
        Artist        = Get-StringValue $tag.Performers
        Title         = $tag.Title
        Genre         = Get-StringValue $tag.Genres
        BPM           = $tag.BeatsPerMinute
        InitialKey    = $tag.InitialKey
        Grouping      = $tag.Grouping
        Comment       = $tag.Comment
        ColorStoredHex    = if ($color) { $color.StoredHex } else { '' }
        ColorDisplayedHex = if ($color) { $color.DisplayedHex } else { '' }
        ColorName         = if ($color) { $color.ColorName } else { '(not tagged)' }
        ColorMatchDistance = if ($color) { $color.MatchDistance } else { '' }
    }
    $rows.Add($row)

    foreach ($anchor in $Script:AnchorTracks) {
        if ($file.Name -like "*$($anchor.Match)*") {
            $anchorHits.Add([PSCustomObject]@{
                File = $file.Name; Expected = $anchor.Expected
                Decoded = if ($color) { $color.ColorName } else { '(not tagged)' }
            })
        }
    }
}
Write-Progress -Activity 'Exporting fresh color/genre sheet' -Completed

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
if (-not $OutputDir) { $OutputDir = $Root }
if (-not (Test-Path -LiteralPath $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }
$outPath = Join-Path $OutputDir "DJ_ColorSheet_$timestamp.csv"
$rows | Export-Csv -LiteralPath $outPath -NoTypeInformation -Encoding UTF8

Write-Host "`nDone: $($rows.Count) track(s) exported." -ForegroundColor Green
Write-Host $outPath -ForegroundColor Green

if ($anchorHits.Count -gt 0) {
    Write-Host "`n--- Anchor validation (check these match your documented colors) ---" -ForegroundColor Yellow
    $anchorHits | Format-Table -AutoSize
} else {
    Write-Host "`nNo anchor tracks found under this Root -- couldn't self-validate the color mapping this run." -ForegroundColor Yellow
}

$taggedCount = ($rows | Where-Object { $_.ColorName -ne '(not tagged)' }).Count
Write-Host "`n$taggedCount of $($rows.Count) files have a Serato track color set." -ForegroundColor Cyan
