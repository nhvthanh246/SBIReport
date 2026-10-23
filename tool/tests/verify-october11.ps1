param(
  [string]$ReportFile = 'AltiuslinkVN-20261011.xlsx',
  [string]$ExpectedFile = 'october11-excel-expected.json',
  [string]$Label = 'October 11 (historical rules)'
)
$ErrorActionPreference = 'Stop'
$excel = $null
$book = $null
function Retry-Excel([scriptblock]$Action) {
  for ($attempt=0; $attempt -lt 60; $attempt++) {
    try { return (& $Action) } catch {
      $cause = $_.Exception
      $busy = $false
      while ($null -ne $cause) {
        if ($cause.HResult -eq -2147418111) { $busy = $true; break }
        $cause = $cause.InnerException
      }
      if (-not $busy -or $attempt -eq 59) { throw }
      Start-Sleep -Milliseconds 500
    }
  }
}
function Same-Value($Actual, $Expected) {
  if ($Actual -is [ValueType] -and $Expected -is [ValueType]) {
    return [Math]::Abs([double]$Actual-[double]$Expected) -lt 0.00000001
  }
  $a = ([string]$Actual).Replace("`r`r`n","`n").Replace("`r`n","`n").Replace("`r","`n").Trim()
  $b = ([string]$Expected).Replace("`r`r`n","`n").Replace("`r`n","`n").Replace("`r","`n").Trim()
  return $a -ceq $b
}
try {
  $expected = Get-Content -Raw -Encoding UTF8 "$PSScriptRoot\artifacts\$ExpectedFile" | ConvertFrom-Json
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.AskToUpdateLinks = $false
  $excel.AutomationSecurity = 3
  $book = Retry-Excel { $excel.Workbooks.Open("$PSScriptRoot\artifacts\$ReportFile",0,$true) }
  Write-Output "Refreshing $Label pivots and calculating formulas..."
  $caches = Retry-Excel { $book.PivotCaches() }
  foreach ($cache in $caches) { Retry-Excel { if ($cache.SourceType -eq 1) { $cache.Refresh() } } }
  Retry-Excel { $excel.CalculateFullRebuild() }
  $checked = 0
  $mismatches = @()
  foreach ($range in $expected.ranges) {
    $sheet = $book.Worksheets.Item($range.sheet)
    $values = $sheet.Range($range.anchor).Resize($range.rows.Count,$range.rows[0].Count).Value2
    for ($i=0; $i -lt $range.rows.Count; $i++) {
      for ($j=0; $j -lt $range.rows[$i].Count; $j++) {
        if (-not (Same-Value $values[($i+1),($j+1)] $range.rows[$i][$j])) {
          $mismatches += "$($range.sheet) offset $i,$j from $($range.anchor)"
        }
        $checked++
      }
    }
  }
  foreach ($cell in $expected.cells) {
    $value = $book.Worksheets.Item($cell.sheet).Range($cell.cell).Value2
    if (-not (Same-Value $value $cell.value)) { $mismatches += "$($cell.sheet)!$($cell.cell): $value vs $($cell.value)" }
    $checked++
  }
  $s5 = $book.Worksheets.Item('Sheet5')
  $labels = $s5.Range('A5:A70').Value2
  $helpers = $s5.Range('O5:V70').Value2
  $blank = 0
  for ($i=1; $i -le 66; $i++) {
    if ([string]$labels[$i,1] -ne '') { continue }
    for ($j=1; $j -le 8; $j++) {
      if ([string]$helpers[$i,$j] -ne '') { throw 'Nonempty spare helper' }
      $blank++
    }
  }
  if ($mismatches.Count) { $mismatches | Select-Object -First 15 | Write-Output; throw "$($mismatches.Count) mismatches after Excel recalculation" }
  Write-Output "PASS: $Label Excel refresh - $checked cells match expected values; $blank spare helper cells stay blank."
} finally {
  if ($book) { Retry-Excel { $book.Close($false) }; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
  if ($excel) { Retry-Excel { $excel.Quit() }; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
