<#
  Apply-ManualMatches.ps1
  ---------------------------------------------------------------------------
  Applies approved manual match tickets to a Summary workbook the recon has
  already produced, IN PLACE:

      - the matched rows are deleted from "Internal outstanding" / "External
        outstanding"
      - the same rows are appended to "Matched Pairs" as Manual legs
      - the outcome of every ticket is appended to "Manual Match Log"

  It never rebuilds the workbook. The recon owns that; this only edits what the
  recon published, so the template, the Summary formulas and the formatting are
  untouched. The corrected balances and Check_Balance update by themselves
  because the totals in row 5 of each outstanding sheet are open-ended
  SUM(C8:C1048576) formulas over the data.

  The journal is NOT modified. It is the durable rule: the next full recon run
  re-reads it and re-applies the same tickets while rebuilding from scratch.
  Running this twice on the same workbook is harmless - the second time the
  items are no longer outstanding, so every ticket is reported as already
  matched and nothing changes.

  Usage
      .\Apply-ManualMatches.ps1                      # newest Summary under the recon root
      .\Apply-ManualMatches.ps1 -Summary <path>
      .\Apply-ManualMatches.ps1 -WhatIf              # report only, save nothing
#>
[CmdletBinding()]
param(
  [string]$Summary,
  [string]$Journal,
  [string]$ReconRoot = "C:\Users\kelvinch\Documents\FINANCE AUTOMATED RECON\Nostro",
  [string]$LibPath,
  [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------- libraries --
if (-not $LibPath) {
  $LibPath = Join-Path (Split-Path -Parent $PSCommandPath) 'lib'
}
function Load-Lib([string]$name, [string]$nugetId, [string]$ver, [string]$tfm) {
  $local = Join-Path $LibPath "$name.dll"
  if (Test-Path $local) { Add-Type -Path $local; return }
  $cache = Join-Path $env:USERPROFILE ".nuget\packages\$nugetId\$ver\lib\$tfm\$name.dll"
  if (Test-Path $cache) { Add-Type -Path $cache; return }
  throw "$name.dll not found. Put it in '$LibPath' (copy from the UiPath project's package cache)."
}
Load-Lib 'DocumentFormat.OpenXml' 'documentformat.openxml' '2.7.2' 'net46'
Load-Lib 'ClosedXML' 'closedxml' '0.95.4' 'net46'

# ------------------------------------------------------------------- inputs --
if (-not $Summary) {
  if (-not (Test-Path $ReconRoot)) { throw "Recon root not found: $ReconRoot" }
  $Summary = Get-ChildItem $ReconRoot -Recurse -Filter '*Summary.xlsx' -ErrorAction SilentlyContinue |
             Where-Object { $_.Name -notlike '~$*' } |
             Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName
  if (-not $Summary) { throw "No Summary workbook found under $ReconRoot" }
}
if (-not (Test-Path $Summary)) { throw "Summary not found: $Summary" }
Write-Host "Summary : $Summary"

$account = $null
if ((Split-Path -Leaf $Summary) -match '(\d{6,})') { $account = $Matches[1] }
if (-not $Journal) {
  if (-not $account) { throw "Could not read the account number from the file name; pass -Journal explicitly." }
  $Journal = Join-Path $ReconRoot "Manual Matches\$account.csv"
}
if (-not (Test-Path $Journal)) { Write-Host "No journal at $Journal - nothing to apply."; return }
Write-Host "Journal : $Journal"

# ------------------------------------------------------------- journal read --
$legs = Import-Csv -Path $Journal
$approved = @($legs | Where-Object { $_.Status -and $_.Status.Trim().ToUpper() -eq 'APPROVED' })
Write-Host ("Journal has {0} leg(s), {1} approved." -f @($legs).Count, $approved.Count)
if ($approved.Count -eq 0) { Write-Host "Nothing approved - stopping."; return }

# ------------------------------------------------------------ workbook read --
$wb = New-Object ClosedXML.Excel.XLWorkbook($Summary)
function Sheet([string]$n){ if($wb.Worksheets.Contains($n)){ $wb.Worksheet($n) } else { $null } }
$wsInt = Sheet 'Internal outstanding'
$wsExt = Sheet 'External outstanding'
if (-not $wsInt -or -not $wsExt) { throw "Workbook has no outstanding sheets - is this a recon Summary?" }

# headers on row 7, data from row 8
function ReadSheet($ws){
  $cols=@{}
  $last=$ws.LastColumnUsed(); if(-not $last){ return @{Cols=$cols;Rows=@()} }
  $nc=$last.ColumnNumber()
  for($c=1;$c -le $nc;$c++){
    $h=$ws.Cell(7,$c).GetString().Trim()
    if($h -and -not $cols.ContainsKey($h)){ $cols[$h]=$c }
  }
  $rows=New-Object System.Collections.ArrayList
  $lastRow=$ws.LastRowUsed(); if(-not $lastRow){ return @{Cols=$cols;Rows=$rows} }
  $nr=$lastRow.RowNumber()
  for($r=8;$r -le $nr;$r++){
    $id = if($cols.ContainsKey('Item Id')){ $ws.Cell($r,$cols['Item Id']).GetString().Trim() } else { '' }
    if(-not $id){ continue }
    $dr=0.0; $cr=0.0
    [void][double]::TryParse($ws.Cell($r,$cols['DR']).GetString().Replace(',',''), [ref]$dr)
    [void][double]::TryParse($ws.Cell($r,$cols['CR']).GetString().Replace(',',''), [ref]$cr)
    [void]$rows.Add([pscustomobject]@{ Row=$r; Id=$id; DR=[Math]::Round($dr,2); CR=[Math]::Round($cr,2) })
  }
  @{ Cols=$cols; Rows=$rows }
}
$int = ReadSheet $wsInt
$ext = ReadSheet $wsExt
if (-not $int.Cols.ContainsKey('Item Id')) {
  throw "'Internal outstanding' has no Item Id column. This workbook predates manual matching - re-run the recon once to stamp ids."
}
$ixInt=@{}; foreach($r in $int.Rows){ if(-not $ixInt.ContainsKey($r.Id)){ $ixInt[$r.Id]=$r } }
$ixExt=@{}; foreach($r in $ext.Rows){ if(-not $ixExt.ContainsKey($r.Id)){ $ixExt[$r.Id]=$r } }
Write-Host ("Outstanding: {0} ledger, {1} statement." -f $int.Rows.Count, $ext.Rows.Count)

# ------------------------------------------------------------- apply tickets --
$groups=[ordered]@{}
foreach($l in $approved){
  $ref=("" + $l.'Match Ref').Trim(); if(-not $ref){ continue }
  if(-not $groups.Contains($ref)){ $groups[$ref]=New-Object System.Collections.ArrayList }
  [void]$groups[$ref].Add($l)
}

$claimedInt=@{}; $claimedExt=@{}
$applied=New-Object System.Collections.ArrayList
$log=New-Object System.Collections.ArrayList

foreach($ref in $groups.Keys){
  $li=New-Object System.Collections.ArrayList
  $si=New-Object System.Collections.ArrayList
  $bad=''
  foreach($leg in $groups[$ref]){
    $side=("" + $leg.Side).Trim().ToUpper()
    $id  =("" + $leg.'Item Id').Trim()
    if(-not $id){ $bad='leg with no Item Id'; break }
    if($side -eq 'CBS' -or $side -eq 'LEDGER'){
      if(-not $ixInt.ContainsKey($id)){ $bad="ledger item not outstanding: $id"; break }
      if($claimedInt.ContainsKey($id)){ $bad="ledger item already matched: $id"; break }
      [void]$li.Add($ixInt[$id])
    } elseif($side -eq 'STATEMENT' -or $side -eq 'SWIFT'){
      if(-not $ixExt.ContainsKey($id)){ $bad="statement item not outstanding: $id"; break }
      if($claimedExt.ContainsKey($id)){ $bad="statement item already matched: $id"; break }
      [void]$si.Add($ixExt[$id])
    } else { $bad="unrecognised Side: $side"; break }
  }

  $lNet=0.0; foreach($x in $li){ $lNet += $x.DR - $x.CR }
  $sNet=0.0; foreach($x in $si){ $sNet += $x.CR - $x.DR }
  $lNet=[Math]::Round($lNet,2); $sNet=[Math]::Round($sNet,2)

  if(-not $bad -and ($li.Count -eq 0 -or $si.Count -eq 0)){ $bad='a ticket needs at least one leg on each side' }
  if(-not $bad -and $lNet -ne $sNet){ $bad=("does not balance: ledger net {0:F2} vs statement net {1:F2}" -f $lNet,$sNet) }

  if($bad){
    [void]$log.Add([pscustomobject]@{Ref=$ref;Outcome='REJECTED';Detail=$bad;L=$li.Count;S=$si.Count;LNet=$lNet;SNet=$sNet})
    Write-Host ("  {0,-20} REJECTED  {1}" -f $ref,$bad) -ForegroundColor Yellow
    continue
  }
  foreach($x in $li){ $claimedInt[$x.Id]=$true }
  foreach($x in $si){ $claimedExt[$x.Id]=$true }
  [void]$applied.Add([pscustomobject]@{Ref=$ref;L=$li;S=$si;Net=$lNet;Reason=("" + $groups[$ref][0].Reason);By=("" + $groups[$ref][0].'Approved By')})
  [void]$log.Add([pscustomobject]@{Ref=$ref;Outcome='APPLIED';Detail=("{0} ledger : {1} statement" -f $li.Count,$si.Count);L=$li.Count;S=$si.Count;LNet=$lNet;SNet=$sNet})
  Write-Host ("  {0,-20} APPLIED   {1} ledger : {2} statement, net {3:N2}" -f $ref,$li.Count,$si.Count,$lNet) -ForegroundColor Green
}

if($applied.Count -eq 0){ Write-Host "No ticket could be applied - workbook untouched."; $wb.Dispose(); return }

# the invariant, checked again here: equal value must leave both sides
$remL=0.0; $remS=0.0
foreach($a in $applied){ foreach($x in $a.L){ $remL += $x.DR - $x.CR }; foreach($x in $a.S){ $remS += $x.CR - $x.DR } }
if([Math]::Round($remL-$remS,2) -ne 0){
  $wb.Dispose()
  throw ("Refusing to write: manual matches would move Check_Balance by {0:N2}" -f ($remL-$remS))
}
Write-Host ("Net removed from each side: ledger {0:N2}, statement {1:N2}" -f $remL,$remS)

# ------------------------------------------------------- append Matched Pairs --
$wsMP = Sheet 'Matched Pairs'
if(-not $wsMP){
  $wsMP=$wb.Worksheets.Add('Matched Pairs')
  $hdr=@('Match Id','Matched On','Source','TRN Date','Value Date','TRN Reference','Source Reference','UETR No','DR','CR','Narrative','Item Id')
  for($c=0;$c -lt $hdr.Count;$c++){ $wsMP.Cell(1,$c+1).Value=$hdr[$c] }
  $wsMP.Range($wsMP.Cell(1,1),$wsMP.Cell(1,$hdr.Count)).Style.Font.Bold=$true
}
$mpNext = if($wsMP.LastRowUsed()){ $wsMP.LastRowUsed().RowNumber()+1 } else { 2 }

# References like 2623113064853000 are numeric-looking but are NOT numbers. Force
# the text columns to text before writing or ClosedXML stores them as doubles and
# they come back as 2.623113064853E+15.
foreach($c in @(1,2,3,4,5,6,7,8,11,12)){ $wsMP.Column($c).Style.NumberFormat.Format = '@' }

function LegValues($ws,$cols,$row,$extraName){
  $g={param($n) if($cols.ContainsKey($n)){ $ws.Cell($row,$cols[$n]).GetString() } else { '' }}
  [pscustomobject]@{
    Date=(&$g 'TRN Date'); Value=(&$g 'Value Date'); R1=(&$g 'TRN Reference'); R2=(&$g 'Source Reference')
    U=(&$g 'UETR No'); Narr=(&$g 'Narrative'); Id=(&$g 'Item Id')
  }
}
foreach($a in $applied){
  foreach($x in $a.L){
    $v=LegValues $wsInt $int.Cols $x.Row
    [void]$wsMP.Cell($mpNext,1).SetValue([string]$a.Ref); [void]$wsMP.Cell($mpNext,2).SetValue([string]("Manual " + $a.Ref))
    [void]$wsMP.Cell($mpNext,3).SetValue([string]"CBS")
    [void]$wsMP.Cell($mpNext,4).SetValue([string]$v.Date); [void]$wsMP.Cell($mpNext,5).SetValue([string]$v.Value)
    [void]$wsMP.Cell($mpNext,6).SetValue([string]$v.R1); [void]$wsMP.Cell($mpNext,7).SetValue([string]$v.R2)
    [void]$wsMP.Cell($mpNext,8).SetValue([string]$v.U)
    $wsMP.Cell($mpNext,9).Value=$x.DR;                $wsMP.Cell($mpNext,10).Value=$x.CR
    [void]$wsMP.Cell($mpNext,11).SetValue([string]$v.Narr); [void]$wsMP.Cell($mpNext,12).SetValue([string]$v.Id)
    $mpNext++
  }
  foreach($x in $a.S){
    $v=LegValues $wsExt $ext.Cols $x.Row
    [void]$wsMP.Cell($mpNext,1).SetValue([string]$a.Ref); [void]$wsMP.Cell($mpNext,2).SetValue([string]("Manual " + $a.Ref))
    [void]$wsMP.Cell($mpNext,3).SetValue([string]"STATEMENT")
    [void]$wsMP.Cell($mpNext,4).SetValue([string]$v.Date); [void]$wsMP.Cell($mpNext,5).SetValue([string]$v.Value)
    [void]$wsMP.Cell($mpNext,6).SetValue([string]$v.R1); [void]$wsMP.Cell($mpNext,7).SetValue([string]$v.R2)
    [void]$wsMP.Cell($mpNext,8).SetValue([string]$v.U)
    $wsMP.Cell($mpNext,9).Value=$x.DR;                $wsMP.Cell($mpNext,10).Value=$x.CR
    [void]$wsMP.Cell($mpNext,11).SetValue([string]$v.Narr); [void]$wsMP.Cell($mpNext,12).SetValue([string]$v.Id)
    $mpNext++
  }
}

# --------------------------------------------------- delete outstanding rows --
# bottom-up, so the row numbers captured above stay valid as the sheet shifts up
$delInt = @($applied | ForEach-Object { $_.L } | ForEach-Object { $_.Row } | Sort-Object -Descending -Unique)
$delExt = @($applied | ForEach-Object { $_.S } | ForEach-Object { $_.Row } | Sort-Object -Descending -Unique)
foreach($r in $delInt){ $wsInt.Row($r).Delete() }
foreach($r in $delExt){ $wsExt.Row($r).Delete() }
Write-Host ("Removed {0} ledger and {1} statement row(s) from the outstanding sheets." -f $delInt.Count,$delExt.Count)

# -------------------------------------------------------- Manual Match Log ---
$wsLog = Sheet 'Manual Match Log'
if(-not $wsLog){
  $wsLog=$wb.Worksheets.Add('Manual Match Log')
  $wsLog.Cell(1,1).Value="Manual match journal: $Journal"
  $wsLog.Cell(1,1).Style.Font.Bold=$true
  $h=@('Match Ref','Outcome','Detail','Ledger Legs','Statement Legs','Ledger Net','Statement Net','Applied By','Applied On','Source')
  for($c=0;$c -lt $h.Count;$c++){ $wsLog.Cell(3,$c+1).Value=$h[$c] }
  $wsLog.Range($wsLog.Cell(3,1),$wsLog.Cell(3,$h.Count)).Style.Font.Bold=$true
}
$logNext = if($wsLog.LastRowUsed()){ $wsLog.LastRowUsed().RowNumber()+1 } else { 4 }
$stamp=(Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
foreach($e in $log){
  $who=($applied | Where-Object { $_.Ref -eq $e.Ref } | Select-Object -First 1 -ExpandProperty By)
  $wsLog.Cell($logNext,1).Value=$e.Ref;   $wsLog.Cell($logNext,2).Value=$e.Outcome
  $wsLog.Cell($logNext,3).Value=$e.Detail;$wsLog.Cell($logNext,4).Value=$e.L
  $wsLog.Cell($logNext,5).Value=$e.S;     $wsLog.Cell($logNext,6).Value=$e.LNet
  $wsLog.Cell($logNext,7).Value=$e.SNet;  $wsLog.Cell($logNext,8).Value=("" + $who)
  $wsLog.Cell($logNext,9).Value=$stamp;   $wsLog.Cell($logNext,10).Value='portal'
  $logNext++
}

# ------------------------------------------------------------------- save ----
if($WhatIf){
  Write-Host "-WhatIf: nothing written." -ForegroundColor Cyan
  $wb.Dispose(); return
}
$backupName = 'none'
try {
  $dir  = Split-Path -Parent $Summary
  $stem = [System.IO.Path]::GetFileNameWithoutExtension($Summary)
  $bakDir = Join-Path $dir 'backup'
  if (-not (Test-Path $bakDir)) { New-Item -ItemType Directory -Force $bakDir | Out-Null }
  $backup = Join-Path $bakDir ($stem + '_' + (Get-Date).ToString('yyyyMMdd_HHmmss') + '.xlsx')
  Copy-Item -LiteralPath $Summary -Destination $backup -Force
  $backupName = Split-Path -Leaf $backup
} catch {
  # a backup is a courtesy, not a precondition - never block the apply on it
  Write-Warning "Could not write a backup ($($_.Exception.Message)). Continuing."
}
$wb.Save()
$wb.Dispose()
Write-Host ""
Write-Host ("Applied {0} ticket(s). Backup: {1}" -f $applied.Count, $backupName) -ForegroundColor Green
Write-Host "The Summary totals and Check_Balance recalculate from the SUM formulas when the workbook is opened."
