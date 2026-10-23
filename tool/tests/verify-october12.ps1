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
try {
  $excel = New-Object -ComObject Excel.Application
  $excel.Visible = $false
  $excel.DisplayAlerts = $false
  $excel.AskToUpdateLinks = $false
  $excel.AutomationSecurity = 3
  foreach ($file in @('AltiuslinkVN-20261012.xlsx','chain12-reference-fixture.xlsx')) {
    Write-Output "Opening $file in hidden Excel..."
    $book = Retry-Excel { $excel.Workbooks.Open("$PSScriptRoot\artifacts\$file",0,$true) }
    $caches = Retry-Excel { $book.PivotCaches() }
    foreach ($cache in $caches) { Retry-Excel { if ($cache.SourceType -eq 1) { $cache.Refresh() } } }
    Retry-Excel { $excel.CalculateFullRebuild() }
    $daily = $book.Worksheets.Item('Daily KPI Result')
    $monthly = $book.Worksheets.Item('202610 KPI Result')
    $byEmail = @{}
    for ($r=6; $r -le 60; $r++) {
      $email = ([string]$monthly.Cells.Item($r,4).Value2).Trim().ToLowerInvariant()
      if ($email.Contains('@')) { $byEmail[$email] = $monthly.Cells.Item($r,16).Value2 }
    }
    $checked = 0
    for ($r=15; $r -le 60; $r++) {
      $email = ([string]$daily.Cells.Item($r,4).Value2).Trim().ToLowerInvariant()
      if (-not $byEmail.ContainsKey($email)) { continue }
      if ([Math]::Abs([double]$daily.Cells.Item($r,23).Value2-[double]$byEmail[$email]) -gt 0.00000001) {
        $details = @{}
        foreach ($col in @('E','F','H','I','N','O','W')) { $details[$col] = $daily.Range("$col$r").Value2 }
        $details | ConvertTo-Json | Write-Output
        throw "Daily/monthly KPI mismatch at row $r ($email): Excel=$($daily.Cells.Item($r,23).Value2), tool=$($byEmail[$email])"
      }
      $checked++
    }
    if ($checked -ne 39) { throw "Wrong KPI count: $checked" }
    if ($daily.Range('E14').Formula -ne '=SUM(E15:E42)') { throw 'Changed team SUM formula' }
    if ([Math]::Abs([double]$daily.Range('W14').Value2-[double]$monthly.Range('P7').Value2) -gt 0.00000001) {
      throw 'Team KPI differs after Excel recalculation'
    }
    $other = $book.Worksheets.Item('3. Other Tasks & Working time')
    if ($file -eq 'AltiuslinkVN-20261012.xlsx') {
      $hours = $other.Range('G94').Value2
      if ($hours -is [string] -or [Math]::Abs([double]$hours-6.12) -gt 0.00000001) { throw 'Hours are still text' }
      if ([Math]::Abs([double]$other.Range('F94').Value2-8.002352941176472) -gt 0.00000001) { throw 'Wrong total hours' }
      $names = $book.Worksheets.Item('2. SF Case info').Range('F7:F389').Value2
      $quang = 0
      for ($r=1; $r -le 383; $r++) { if ([string]$names[$r,1] -eq 'LUU NGOC QUANG') { $quang++ } }
      if ($quang -ne 4) { throw 'Quang identities changed' }
    } else {
      if ([Math]::Abs([double]$monthly.Range('P7').Value2-1.2773099582093352) -gt 0.00000001) { throw 'Reference-input team KPI differs' }
    }
    $s5 = $book.Worksheets.Item('Sheet5')
    $labels = $s5.Range('A5:A70').Value2
    $helpers = $s5.Range('O5:V70').Value2
    for ($i=1; $i -le 66; $i++) {
      if ([string]$labels[$i,1] -ne '') { continue }
      for ($j=1; $j -le 8; $j++) { if ([string]$helpers[$i,$j] -ne '') { throw 'Nonempty spare helper' } }
    }
    Write-Output "PASS: $file - 39 individual KPI plus team agree after Excel refresh; hours, identities and blank helpers verified."
    Retry-Excel { $book.Close($false) }
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    $book = $null
  }
} finally {
  if ($book) { Retry-Excel { $book.Close($false) }; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
  if ($excel) { Retry-Excel { $excel.Quit() }; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
