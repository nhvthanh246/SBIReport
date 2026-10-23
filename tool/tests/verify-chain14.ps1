$ErrorActionPreference = 'Stop'
$excel = $null
$book = $null
function Invoke-ExcelRetry([scriptblock]$Action) {
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
  foreach ($file in @('AltiuslinkVN-20260914.xlsx', 'chain14-reference-fixture.xlsx')) {
    Write-Output "Opening $file in hidden Excel..."
    $book = Invoke-ExcelRetry { $excel.Workbooks.Open("$PSScriptRoot\artifacts\$file", 0, $false) }
    $caches = Invoke-ExcelRetry { $book.PivotCaches() }
    foreach ($cache in $caches) { Invoke-ExcelRetry { if ($cache.SourceType -eq 1) { $cache.Refresh() } } }
    Invoke-ExcelRetry { $excel.CalculateFullRebuild() }
    $mail = $book.Worksheets.Item('mail SF')
    $sf = $book.Worksheets.Item('2. SF Case info')
    if ($mail.Range('C43').Value2 -ne 'NGUYEN THI KHANH LY (3M)') { throw 'Missing mail SF correction' }
    $corrected = 0
    for ($r=7; $r -le 518; $r++) {
      if ($sf.Cells.Item($r,7).Value2 -eq 'ly.ntk1@altius-link.com.vn') { $corrected++ }
    }
    if ($corrected -ne 18) { throw "Corrected email count changed after refresh: $corrected" }
    $monthly = $book.Worksheets.Item('202609 KPI Result')
    $daily = $book.Worksheets.Item('Daily KPI Result')
    $monthByEmail = @{}
    for ($r=6; $r -le 60; $r++) {
      $email = ([string]$monthly.Cells.Item($r,4).Value2).Trim().ToLowerInvariant()
      if ($email.Contains('@')) { $monthByEmail[$email] = $monthly.Cells.Item($r,18).Value2 }
    }
    $checked = 0
    for ($r=15; $r -le 60; $r++) {
      $email = ([string]$daily.Cells.Item($r,4).Value2).Trim().ToLowerInvariant()
      if (-not $monthByEmail.ContainsKey($email)) { continue }
      $a = [double]$daily.Cells.Item($r,23).Value2
      $b = [double]$monthByEmail[$email]
      if ([Math]::Abs($a-$b) -gt 0.000001) { throw "Daily/monthly mismatch at daily row ${r}: $a vs $b" }
      $checked++
    }
    if ($checked -ne 38) { throw "Expected 38 individual KPI checks, got $checked" }
    if ([Math]::Abs([double]$daily.Range('W14').Value2-[double]$monthly.Range('R7').Value2) -gt 0.000001) { throw 'Team KPI mismatch' }
    if ($file -eq 'chain14-reference-fixture.xlsx') {
      if ([Math]::Abs([double]$daily.Range('W36').Value2-2.204761904761905) -gt 0.000001) { throw 'Setsumekai KPI mismatch' }
      if ($daily.Range('E36').Value2 -ne 201) { throw 'Setsumekai weight mismatch' }
    }
    $s5 = $book.Worksheets.Item('Sheet5')
    $labels = $s5.Range('A5:A70').Value2
    $helpers = $s5.Range('O5:V70').Value2
    for ($i=1; $i -le 66; $i++) {
      if ([string]$labels[$i,1] -ne '') { continue }
      for ($j=1; $j -le 8; $j++) {
        if ([string]$helpers[$i,$j] -ne '') { throw "Nonempty helper below pivot, row $($i+4)" }
      }
    }
    Write-Output "PASS: $file - 18 emails, 39 recalculated KPI results, blank spare helpers."
    Invoke-ExcelRetry { $book.Close($false) }
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)
    $book = $null
  }
} finally {
  if ($book) { Invoke-ExcelRetry { $book.Close($false) }; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
  if ($excel) { Invoke-ExcelRetry { $excel.Quit() }; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
