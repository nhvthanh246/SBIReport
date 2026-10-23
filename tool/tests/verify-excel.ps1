param([string]$Path = "$PSScriptRoot\artifacts\AltiuslinkVN-20261004.xlsx")
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
  Write-Output 'Opening generated workbook in Excel...'
  $book = Invoke-ExcelRetry { $excel.Workbooks.Open($Path, 0, $false) }
  Write-Output 'Refreshing pivots...'
  $caches = Invoke-ExcelRetry { $book.PivotCaches() }
  foreach ($cache in $caches) { Invoke-ExcelRetry { if ($cache.SourceType -eq 1) { $cache.Refresh() } } }
  Invoke-ExcelRetry { $excel.CalculateFullRebuild() }
  Write-Output 'Checking calculated values...'
  $hours = $book.Worksheets.Item('3. Other Tasks & Working time')
  $s5 = $book.Worksheets.Item('Sheet5')
  $daily = $book.Worksheets.Item('Daily KPI Result')
  $summary = $book.Worksheets.Item('Summarize_team VietNam')
  $monthly = $book.Worksheets.Item('202610 KPI Result')
  $labels = $s5.Range('A5:A70').Value2
  $helpers = $s5.Range('O5:V70').Value2
  $blankHelperCells = 0
  $helperTotalFound = $false
  for ($i=1; $i -le 66; $i++) {
    if ([string]$labels[$i,1] -eq '') {
      for ($j=1; $j -le 8; $j++) {
        if ([string]$helpers[$i,$j] -ne '') { throw "Nonempty helper under blank pivot row: row $($i+4), column $($j+14)" }
        $blankHelperCells++
      }
    } elseif ([string]$labels[$i,1] -eq 'Grand Total') {
      $helperTotalFound = $true
      if ($helpers[$i,1] -ne 1782 -or $helpers[$i,3] -ne 633) { throw 'Helper Grand Total changed' }
    }
  }
  if (-not $helperTotalFound -or $blankHelperCells -eq 0) { throw 'Missing Grand Total or spare rows in test fixture' }
  $monthlyByEmail = @{}
  for ($r=6; $r -le 60; $r++) {
    $email = [string]$monthly.Cells.Item($r,4).Value2
    if ($email.Contains('@')) { $monthlyByEmail[$email.Trim().ToLowerInvariant()] = $r }
  }
  $kpiMismatches = @()
  $kpiChecked = 0
  for ($r=15; $r -le 60; $r++) {
    $email = ([string]$daily.Cells.Item($r,4).Value2).Trim().ToLowerInvariant()
    if (-not $monthlyByEmail.ContainsKey($email)) { continue }
    $kpiChecked++
    $dayValue = [double]$daily.Cells.Item($r,23).Value2
    $monthValue = [double]$monthly.Cells.Item($monthlyByEmail[$email],8).Value2
    if ([Math]::Abs($dayValue-$monthValue) -gt 0.000001) { $kpiMismatches += "row ${r}: daily=$dayValue monthly=$monthValue" }
  }
  $jpRow = 0
  $monthlyJpRow = 0
  for ($r=5; $r -le 70; $r++) { if ($s5.Cells.Item($r,1).Value2 -eq 'bao.nq@altius-link.com.vn') { $jpRow = $r } }
  for ($r=6; $r -le 60; $r++) { if ($monthly.Cells.Item($r,4).Value2 -eq 'bao.nq@altius-link.com.vn') { $monthlyJpRow = $r } }
  $result = [ordered]@{
    hours = @($hours.Range('F6').Value2, $hours.Range('F14').Value2, $hours.Range('F30').Value2)
    otherHours = @($hours.Range('M13').Value2, $hours.Range('M21').Value2, $hours.Range('M33').Value2)
    jpHelper = @($s5.Cells.Item($jpRow,15).Value2, $s5.Cells.Item($jpRow,17).Value2, $s5.Cells.Item($jpRow,20).Value2, $s5.Cells.Item($jpRow,21).Value2)
    dailyJp = @($daily.Range('K54').Value2, $daily.Range('W54').Value2)
    monthlyJp = $monthly.Cells.Item($monthlyJpRow,8).Value2
    cardStatus = $summary.Range('I36').Value2
    topicTotal = $summary.Range('I32').Value2
    sheets = $book.Worksheets.Count
    pivots = $book.PivotCaches().Count
    kpiChecked = $kpiChecked
    kpiMismatches = $kpiMismatches
    blankHelperCells = $blankHelperCells
  }
  $book.SaveCopyAs("$PSScriptRoot\artifacts\AltiuslinkVN-20261004-recalculated.xlsx")
  $result | ConvertTo-Json -Depth 5
  if ($result.hours[0] -ne 4 -or $result.hours[1] -ne 8 -or [Math]::Abs($result.hours[2]-8) -gt 0.000001) { throw 'Hours differ from screenshots' }
  if ($result.jpHelper[0] -ne 20 -or $result.jpHelper[1] -ne 21 -or $result.jpHelper[2] -ne 1 -or $result.jpHelper[3] -ne 24) { throw 'JP channels differ from screenshots' }
  if ($result.dailyJp[0] -ne 24 -or [Math]::Round($result.dailyJp[1]*100) -ne 113) { throw 'Daily JP KPI differs from screenshot' }
  if ([Math]::Abs($result.monthlyJp-$result.dailyJp[1]) -gt 0.000001) { throw 'Monthly JP KPI differs from daily KPI' }
  if ($result.cardStatus -ne 621) { throw 'Card status differs from screenshot' }
  if ($kpiMismatches.Count) { throw 'Monthly/daily KPI mismatch' }
  Write-Output 'PASS: Excel refresh and calculation match the October 4 screenshots.'
} finally {
  if ($book) { Invoke-ExcelRetry { $book.Close($false) }; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
  if ($excel) { Invoke-ExcelRetry { $excel.Quit() }; [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
}
