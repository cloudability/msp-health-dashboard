# Prunes timestamped dashboard backups.
# Keeps the newest $PerDay files for each of the newest $MaxDays distinct days; deletes the rest.
# Only touches files named MSP-CCB-Performance-Dashboard_yyyy-MM-dd_HH-mm-ss.html.
param(
    [Parameter(Mandatory = $true)][string]$Dir,
    [int]$MaxDays = 10,
    [int]$PerDay  = 2
)

if (-not (Test-Path -LiteralPath $Dir)) { exit 0 }

$pattern = '^MSP-CCB-Performance-Dashboard_(\d{4}-\d{2}-\d{2})_\d{2}-\d{2}-\d{2}\.html$'

$backups = Get-ChildItem -LiteralPath $Dir -File | ForEach-Object {
    if ($_.Name -match $pattern) {
        [pscustomobject]@{ File = $_; Day = $Matches[1]; Name = $_.Name }
    }
}
if (-not $backups) { exit 0 }

$keepDays = $backups | Select-Object -ExpandProperty Day -Unique | Sort-Object -Descending | Select-Object -First $MaxDays

$keep = @()
foreach ($day in $keepDays) {
    $keep += $backups | Where-Object { $_.Day -eq $day } | Sort-Object Name -Descending | Select-Object -First $PerDay
}
$keepNames = $keep | Select-Object -ExpandProperty Name

$removed = 0
foreach ($b in $backups) {
    if ($keepNames -notcontains $b.Name) {
        Remove-Item -LiteralPath $b.File.FullName -Force
        $removed++
    }
}
Write-Host "  Backup cleanup: kept $($keepNames.Count), removed $removed."
