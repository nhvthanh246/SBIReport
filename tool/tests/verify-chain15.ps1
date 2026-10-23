$ErrorActionPreference = 'Stop'
$excel = $null
$book = $null
function Retry-Excel([scriptblock]$Action) {
  for ($attempt=0; $attempt -lt 30; $attempt++) {
    try { return (& $Action) } catch {
      $cause = $_.Exception
      $busy = $false
      while ($null -ne $cause) {
        if ($cause.HResult -eq -2147418111) { $busy = $true; break }
        $cause = $cause.InnerException
      }
      if (-not $busy -or $attempt -eq 29) { throw }
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
  $book = Retry-Excel { $excel.Workbooks.Open("$PSScriptRoot\artifacts\AltiuslinkVN-20260915.xlsx",0,$true) }
  $caches = Retry-Excel { $book.PivotCaches() }
  foreach ($cache in $caches) { Retry-Excel { if ($cache.SourceType -eq 1) { $cache.Refresh() } } }
  Retry-Excel { $excel.CalculateFullRebuild() }
  if ($book.Worksheets.Item('1. CSKH').Range('F11').Value2 -ne '2026-09-15 21:30:00+07:00') { throw 'Incorrect Start Time' }
  $daily = $book.Worksheets.Item('Daily KPI Result')
  $monthly = $book.Worksheets.Item('202609 KPI Result')
  $byEmail = @{}
  for ($r=6; $r -le 60; $r++) {
    $email = [string]$monthly.Cells.Item($r,4).Value2
    if ($email.Contains('@')) { $byEmail[$email.Trim().ToLowerInvariant()] = $monthly.Cells.Item($r,19).Value2 }
  }
  $checked = 0
  for ($r=15; $r -le 60; $r++) {
    $email = ([string]$daily.Cells.Item($r,4).Value2).Trim().ToLowerInvariant()
    if (-not $byEmail.ContainsKey($email)) { continue }
    if ([Math]::Abs([double]$daily.Cells.Item($r,23).Value2-[double]$byEmail[$email]) -gt 0.000001) { throw "KPI mismatch at daily row $r" }
    $checked++
  }
  if ($checked -ne 38) { throw "Wrong individual KPI count: $checked" }
  if ([Math]::Abs([double]$daily.Range('W14').Value2-[double]$monthly.Range('S7').Value2) -gt 0.000001) { throw 'Team KPI mismatch' }
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
  Write-Output "PASS: day 15 Excel refresh - First response Start Time, 39 daily/monthly KPI results, $blank blank helper cells."
} finally {
  if ($book) { Retry-Excel { $book.Close($false) }; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
  if ($excel) { Retry-Excel { $excel.Quit() }; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
