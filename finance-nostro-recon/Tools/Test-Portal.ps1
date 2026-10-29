<#
  Test-Portal.ps1
  ---------------------------------------------------------------------------
  Runs the portal on THIS machine so you can exercise the whole flow before
  anything is deployed to robot.nmbtz.com.

  It serves Deploy\portal\portal.html and implements the same operations as
  api.aspx - whoami, runs, data, file, tickets - against a real recon folder.
  The missing-pages operations - pages, upload, unupload - and the sign-in
  sessions and audit log are not mirrored: api.aspx's own server code is
  compiled here and called, so they cannot drift. Create <ReconRoot>\Missing
  Pages and <ReconRoot>\Audit first, as IT does on the real share.
  No IIS, no admin: HttpListener on localhost is enough.

  What this DOES prove
    - the portal finds runs and accounts on a server rather than your PC
    - downloads of Summary / Workings come off the server
    - a ticket POSTs and lands in Manual Matches\<account>.csv in the right
      format, stamped with the Windows user the server sees
    - Apply-ManualMatches.ps1 then consumes that journal

  What it does NOT prove
    - api.aspx itself (that is C#; compile-check it with csc.exe)
    - the IIS web.config: handlers, Windows Authentication, app-pool identity
      and folder permissions all still need testing on the server

  Usage
      .\Test-Portal.ps1
      .\Test-Portal.ps1 -ReconRoot "D:\some\fixture\Nostro" -Port 8731
  Stop with Ctrl+C, or browse to /shutdown.
#>
[CmdletBinding()]
param(
  [string]$ReconRoot = "C:\Users\kelvinch\Documents\FINANCE AUTOMATED RECON\Nostro",
  [int]$Port = 8731,
  [string]$PortalFile
)

$ErrorActionPreference = 'Stop'
if (-not $PortalFile) {
  $PortalFile = Join-Path (Split-Path -Parent (Split-Path -Parent $PSCommandPath)) 'Deploy\portal\portal.html'
}
if (-not (Test-Path $PortalFile)) { throw "portal.html not found at $PortalFile" }
if (-not (Test-Path $ReconRoot))  { throw "ReconRoot not found: $ReconRoot" }
$ReconRoot = (Resolve-Path $ReconRoot).Path.TrimEnd('\')

$SKIP = @('backup','manual matches','missing pages','audit')

# api.aspx's server script, compiled as a Page subclass with public entry points
# for the missing-pages helpers (which are private statics inside it).
$apiFile = Join-Path (Split-Path -Parent $PortalFile) 'api.aspx'
$apiScript = [regex]::Match([IO.File]::ReadAllText($apiFile), '(?s)<script runat="server">(.*?)</script>').Groups[1].Value
if (-not $apiScript) { throw "no server script found in $apiFile" }
$fw = [Runtime.InteropServices.RuntimeEnvironment]::GetRuntimeDirectory()
Add-Type -ReferencedAssemblies (Join-Path $fw 'System.Web.dll'),(Join-Path $fw 'System.Configuration.dll'),'System.Xml','System.Core' -TypeDefinition @"
using System; using System.IO; using System.Text; using System.Text.RegularExpressions;
using System.Collections.Generic; using System.Linq; using System.Globalization; using System.Web;
public class PortalApiTest : System.Web.UI.Page
{
$apiScript
    static Exception Plain(Exception e) { return (e is PageError || e is AuditError) ? new InvalidOperationException(e.Message) : e; }
    public static string StartSession(string user) { return SessionStart(user); }
    public static bool Live(string sid, string user) { return SessionLive(sid, user); }
    public static bool EndSession(string sid) { return SessionEnd(sid); }
    public static int Idle() { return IdleLimit(); }
    public static string[][] Expire() { return SessionsExpire(IdleLimit()).Select(s => new[] { s.User, s.Started.ToString("yyyy-MM-dd HH:mm:ss"), s.LastSeen.ToString("HH:mm:ss") }).ToArray(); }
    public static bool IsPageEvent(string a) { return Array.IndexOf(PageEvents, a) >= 0; }
    public static void Ready(string root) { try { AuditReady(root); } catch (Exception e) { throw Plain(e); } }
    public static void Log(string root, string user, string client, string action, string detail, string result) { try { AuditAppend(root, user, client, action, Detail(detail), result); } catch (Exception e) { throw Plain(e); } }
    public static string ListPages(string root) { try { return PagesJson(root); } catch (Exception e) { throw Plain(e); } }
    public static string Upload(string root, string std, string name, byte[] body, string user) { try { return SavePage(root, std, name, body, user); } catch (Exception e) { throw Plain(e); } }
    public static void Unupload(string root, string std, string name, string user) { try { RemovePage(root, std, name, user); } catch (Exception e) { throw Plain(e); } }
}
"@
function Inside([string]$full){ $full.ToLower().StartsWith(($ReconRoot + '\').ToLower()) }
function JsonStr([string]$s){
  if($null -eq $s){ return '' }
  ($s -replace '\\','\\' -replace '"','\"' -replace "[`r`n`t]",' ')
}
function Reply($res,[string]$json){
  $res.ContentType='application/json'
  $b=[Text.Encoding]::UTF8.GetBytes($json); $res.OutputStream.Write($b,0,$b.Length)
}
function Pick([string]$dir,[string]$pattern){
  Get-ChildItem -LiteralPath $dir -Filter $pattern -File -ErrorAction SilentlyContinue |
    Where-Object { -not $_.Name.StartsWith('~$') } |
    Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
}
function AcctOf([string]$name){ if($name -match '(\d{6,})'){ $Matches[1] } else { '' } }
function Ms([datetime]$utc){ [long]($utc - [datetime]'1970-01-01').TotalMilliseconds }
function Newest($files){ @($files | Where-Object { $_ -and -not $_.Name.StartsWith('~$') } | Sort-Object LastWriteTimeUtc -Descending)[0] }

# Same rules as api.aspx: every account of a bank is written into one
# <date>\<bank> folder, each with its own recon-data-<account>.json; the old
# shared recon-data.json is read for the account named inside it.
function DataAccount($f){
  if($f.Name -match '^recon-data-([A-Za-z0-9_\-]+)\.json$'){ return $Matches[1] }
  if($f.Name -ne 'recon-data.json'){ return '' }
  $head = ''
  try{ $rd = New-Object IO.StreamReader($f.FullName); $buf = New-Object char[] 2048; $n = $rd.Read($buf,0,2048); $rd.Close(); $head = New-Object string($buf,0,$n) }catch{}
  if($head -match '"account"\s*:\s*"([^"]*)"'){ $Matches[1] } else { '' }
}
function AccountsIn($bankDir){
  $files = @(Get-ChildItem -LiteralPath $bankDir.FullName -File -ErrorAction SilentlyContinue)
  $sums  = @($files | Where-Object { $_.Name -like '*Summary.xlsx' })
  $wrks  = @($files | Where-Object { $_.Name -like '*Workings.xlsx' })
  $datas = @($files | Where-Object { $_.Name -like 'recon-data*.json' })
  $dataBy = @{}
  foreach($f in ($datas | Where-Object { $_.Name -notmatch '^recon-data-.+\.json$' })){ $a = DataAccount $f; if($a){ $dataBy[$a] = $f } }
  foreach($f in ($datas | Where-Object { $_.Name -match '^recon-data-.+\.json$' })){ $a = DataAccount $f; if($a){ $dataBy[$a] = $f } }
  $accts = New-Object System.Collections.ArrayList
  foreach($f in @($sums + $wrks)){ $a = AcctOf $f.Name; if($a -and -not $accts.Contains($a)){ [void]$accts.Add($a) } }
  foreach($a in $dataBy.Keys){ if(-not $accts.Contains($a)){ [void]$accts.Add($a) } }
  foreach($a in $accts){
    $sum = Newest ($sums | Where-Object { (AcctOf $_.Name) -eq $a })
    $wrk = Newest ($wrks | Where-Object { (AcctOf $_.Name) -eq $a })
    $data = $dataBy[$a]
    $named = @($sum,$wrk | Where-Object { $_ })[0]
    $ccy = ''; if($named -and $named.Name -cmatch '_([A-Z]{3})_\d{6,}'){ $ccy = $Matches[1] }
    $when = (@($data,$sum,$wrk | Where-Object { $_ } | ForEach-Object { Ms $_.LastWriteTimeUtc }) | Measure-Object -Maximum).Maximum
    [pscustomobject]@{ account=$a; ccy=$ccy; bank=$bankDir.Name; dir=$bankDir.FullName; data=$data; summary=$sum; workings=$wrk; when=[long]$when }
  }
}
function GetRuns {
  $out = New-Object System.Collections.ArrayList
  foreach($day in Get-ChildItem -LiteralPath $ReconRoot -Directory -ErrorAction SilentlyContinue){
    if($SKIP -contains $day.Name.ToLower()){ continue }
    # order by the RUN DATE, not file time - applying manual matches rewrites a Summary
    $d=[datetime]::MinValue
    $dated = [datetime]::TryParseExact($day.Name,[string[]]@('d-MMM-yyyy','dd-MMM-yyyy'),[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::None,[ref]$d)
    foreach($bank in Get-ChildItem -LiteralPath $day.FullName -Directory -ErrorAction SilentlyContinue){
      if($SKIP -contains $bank.Name.ToLower()){ continue }
      foreach($r in AccountsIn $bank){
        $r | Add-Member id "$($day.Name)/$($bank.Name)"
        $r | Add-Member date $day.Name
        $r | Add-Member ord $(if($dated){ Ms $d }else{ $r.when })
        [void]$out.Add($r)
      }
    }
  }
  $out | Sort-Object ord,when -Descending
}
function RunFor([string]$rel,[string]$acct){
  $dir = Get-Item -LiteralPath (ResolveRun $rel)
  $all = @(AccountsIn $dir)
  if(-not $acct -and $all.Count -eq 1){ return $all[0] }
  if($acct -notmatch '^[A-Za-z0-9_\-]{1,64}$'){ throw 'bad account' }
  $hit = @($all | Where-Object { $_.account -eq $acct })[0]
  if(-not $hit){ throw "no account $acct in this run" }
  $hit
}
function ResolveRun([string]$rel){
  if(-not $rel -or $rel.Contains('..')){ throw 'bad run' }
  $full=[IO.Path]::GetFullPath((Join-Path $ReconRoot ($rel -replace '/','\')))
  if(-not (Inside $full) -or -not (Test-Path -LiteralPath $full)){ throw 'bad run' }
  $full
}

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$Port/")
$listener.Start()

$me = [Security.Principal.WindowsIdentity]::GetCurrent().Name
Write-Host ""
Write-Host "  Nostro portal (local test)      http://localhost:$Port/" -ForegroundColor Green
Write-Host "  Recon root : $ReconRoot"
Write-Host "  Signed in  : $me"
Write-Host "  Runs found : $((GetRuns | Measure-Object).Count)"
Write-Host "  Stop with Ctrl+C or browse to /shutdown"
Write-Host ""

try{
  while($listener.IsListening){
    $ctx = $listener.GetContext()
    $req = $ctx.Request; $res = $ctx.Response
    $res.Headers.Add('Cache-Control','no-store')
    $op = $req.QueryString['op']
    $act = ''; $client = ''
    try{
      if($req.Url.AbsolutePath -eq '/shutdown'){
        $res.ContentType='text/plain'; $b=[Text.Encoding]::UTF8.GetBytes('stopping')
        $res.OutputStream.Write($b,0,$b.Length); $res.Close(); break
      }
      elseif($req.Url.AbsolutePath -notlike '*api.aspx'){
        $html=[IO.File]::ReadAllBytes($PortalFile)
        $res.ContentType='text/html; charset=utf-8'
        $res.OutputStream.Write($html,0,$html.Length)
      }
      else{
      # ---- api.aspx: sessions and audit as api.aspx does them ----
      $sid = if($req.Cookies['nostro_sid']){ $req.Cookies['nostro_sid'].Value } else { '' }
      $client = $req.RemoteEndPoint.Address.ToString()
      $post = $req.HttpMethod -eq 'POST'
      $act = ''; $detail = ''
      foreach($s in [PortalApiTest]::Expire()){
        try{ [PortalApiTest]::Log($ReconRoot,$s[0],'','SESSION_EXPIRED',"signed in $($s[1]), no activity since $($s[2])",'ended by the server') }catch{}
      }
      $live = $true
      if(@('whoami','login','logout') -notcontains $op){
        $live = [PortalApiTest]::Live($sid,$me)
        if($live){ [PortalApiTest]::Ready($ReconRoot) }
        else{
          try{ [PortalApiTest]::Log($ReconRoot,$me,$client,'DENIED',"op=$op",'no live portal session') }catch{}
          $res.StatusCode=440; Reply $res '{"error":"Your portal session has ended - sign in again.","signedOut":true}'
        }
      }
      if(-not $live){ }
      elseif($op -eq 'whoami'){
        Reply $res ('{"user":"'+(JsonStr $me)+'","session":'+$(if([PortalApiTest]::Live($sid,$me)){'true'}else{'false'})+',"idleMinutes":'+[PortalApiTest]::Idle()+'}')
      }
      elseif($op -eq 'login'){
        if(-not $post){ throw 'login must be POSTed' }
        [PortalApiTest]::Ready($ReconRoot)
        [void][PortalApiTest]::EndSession($sid)
        $new=[PortalApiTest]::StartSession($me)
        [PortalApiTest]::Log($ReconRoot,$me,$client,'SIGN_IN',"idle limit $([PortalApiTest]::Idle()) min",'OK')
        $res.AppendHeader('Set-Cookie',"nostro_sid=$new; Path=/; HttpOnly; SameSite=Strict")
        Write-Host "  + $me signed in" -ForegroundColor Green
        Reply $res ('{"user":"'+(JsonStr $me)+'","idleMinutes":'+[PortalApiTest]::Idle()+'}')
      }
      elseif($op -eq 'logout'){
        if(-not $post){ throw 'logout must be POSTed' }
        $idleOut = $req.QueryString['reason'] -eq 'idle'
        if([PortalApiTest]::EndSession($sid)){
          [PortalApiTest]::Log($ReconRoot,$me,$client,'SIGN_OUT',$(if($idleOut){"no activity for $([PortalApiTest]::Idle()) min"}else{'signed out'}),$(if($idleOut){'signed out for inactivity'}else{'OK'}))
          Write-Host "  - $me signed out$(if($idleOut){' (idle)'})" -ForegroundColor Green
        }
        $res.AppendHeader('Set-Cookie','nostro_sid=; Path=/; HttpOnly; SameSite=Strict; Max-Age=0')
        Reply $res '{"signedOut":true}'
      }
      elseif($op -eq 'ping'){ Reply $res '{"ok":true}' }
      elseif($op -eq 'event'){
        if(-not $post){ throw 'event must be POSTed' }
        $ev=($req.QueryString['action']+'').ToUpper()
        if(-not [PortalApiTest]::IsPageEvent($ev)){ throw 'unknown event' }
        $act=$ev; $detail=(New-Object IO.StreamReader($req.InputStream,[Text.Encoding]::UTF8)).ReadToEnd()
        Reply $res '{"ok":true}'
      }
      elseif($op -eq 'runs'){
        $act='LIST_RUNS'
        $res.ContentType='application/json'
        $items = GetRuns | ForEach-Object {
          '{"id":"'+(JsonStr $_.id)+'","date":"'+(JsonStr $_.date)+'","bank":"'+(JsonStr $_.bank)+
          '","account":"'+(JsonStr $_.account)+'","ccy":"'+(JsonStr $_.ccy)+'","when":'+$_.when+
          ',"data":'+$(if($_.data){'true'}else{'false'})+
          ',"summary":'+$(if($_.summary){'true'}else{'false'})+
          ',"workings":'+$(if($_.workings){'true'}else{'false'})+',"ord":'+$_.ord+'}'
        }
        $detail="$(@($items).Count) account run(s)"
        $b=[Text.Encoding]::UTF8.GetBytes('['+($items -join ',')+']')
        $res.OutputStream.Write($b,0,$b.Length)
      }
      elseif($op -eq 'data'){
        $act='OPEN_RUN'; $detail="$($req.QueryString['run']) account $($req.QueryString['account'])"
        $f=(RunFor $req.QueryString['run'] $req.QueryString['account']).data
        if(-not $f){ throw 'this account has no portal data in this run' }
        $res.ContentType='application/json'
        $b=[IO.File]::ReadAllBytes($f.FullName); $res.OutputStream.Write($b,0,$b.Length)
      }
      elseif($op -eq 'file'){
        $kind=($req.QueryString['kind']+'').ToLower()
        $act='DOWNLOAD'; $detail="$kind of $($req.QueryString['run']) account $($req.QueryString['account'])"
        $r=RunFor $req.QueryString['run'] $req.QueryString['account']
        $f = if($kind -eq 'summary'){ $r.summary } elseif($kind -eq 'workings'){ $r.workings } else { $null }
        if(-not $f){ throw 'no such file for this run' }
        $detail="$kind $($f.Name)"
        $res.ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        $res.AddHeader('Content-Disposition','attachment; filename="'+($f.Name -replace '"','')+'"')
        $b=[IO.File]::ReadAllBytes($f.FullName); $res.OutputStream.Write($b,0,$b.Length)
      }
      elseif($op -eq 'tickets'){
        if($req.HttpMethod -ne 'POST'){ throw 'tickets must be POSTed' }
        $acct=$req.QueryString['account']
        $act='SAVE_TICKETS'; $detail="account $acct"
        if($acct -notmatch '^[A-Za-z0-9_\-]{1,64}$'){ throw 'bad account' }
        $body=(New-Object IO.StreamReader($req.InputStream,[Text.Encoding]::UTF8)).ReadToEnd()
        $stamp=(Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        $lines=@()
        foreach($raw in ($body -split "`n")){
          $line=$raw.Trim(); if(-not $line){ continue }
          $p=$line.Split('|')
          if($p.Count -lt 3){ throw 'each leg must be MatchRef|Side|ItemId|Reason' }
          $mref=$p[0].Trim(); $side=$p[1].Trim().ToUpper(); $item=$p[2].Trim()
          $reason=$(if($p.Count -gt 3){ ($p[3..($p.Count-1)] -join '|').Trim() }else{ '' })
          if($mref -notmatch '^[A-Za-z0-9_\-]{1,40}$'){ throw "bad Match Ref: $mref" }
          if($side -ne 'CBS' -and $side -ne 'STATEMENT'){ throw "bad Side: $side" }
          if($item -notmatch '^[A-Za-z0-9_\-]{1,64}$'){ throw "bad Item Id: $item" }
          $esc={param($s) $s=($s -replace "[`r`n]",' '); if($s -match '[,"]'){ '"'+($s -replace '"','""')+'"' }else{ $s }}
          # the recorded user is the one the SERVER sees, never what the client sent
          $lines += (@((&$esc $mref),$side,(&$esc $item),'APPROVED',(&$esc $me),$stamp,(&$esc $me),$stamp,(&$esc $reason)) -join ',')
        }
        if(-not $lines){ throw 'no legs supplied' }
        $detail="account ${acct}: $($lines.Count) leg(s), tickets "+((@($lines | ForEach-Object { $_.Split(',')[0] }) | Select-Object -Unique) -join ' ')
        $dir=Join-Path $ReconRoot 'Manual Matches'
        New-Item -ItemType Directory -Force $dir | Out-Null
        $path=Join-Path $dir "$acct.csv"
        if(-not (Test-Path -LiteralPath $path)){
          Set-Content -LiteralPath $path -Encoding UTF8 -Value 'Match Ref,Side,Item Id,Status,Proposed By,Proposed On,Approved By,Approved On,Reason'
        }
        Add-Content -LiteralPath $path -Encoding UTF8 -Value $lines
        Write-Host ("  + {0} leg(s) -> {1}" -f $lines.Count,$path) -ForegroundColor Green
        $res.ContentType='application/json'
        $b=[Text.Encoding]::UTF8.GetBytes('{"saved":'+$lines.Count+',"journal":"'+(JsonStr $path)+'","user":"'+(JsonStr $me)+'"}')
        $res.OutputStream.Write($b,0,$b.Length)
      }
      elseif($op -eq 'pages'){
        $act='LIST_PAGES'
        $res.ContentType='application/json'
        $b=[Text.Encoding]::UTF8.GetBytes([PortalApiTest]::ListPages($ReconRoot))
        $res.OutputStream.Write($b,0,$b.Length)
      }
      elseif($op -eq 'upload' -or $op -eq 'unupload'){
        $act=$(if($op -eq 'upload'){'UPLOAD_PAGE'}else{'TAKE_BACK_PAGE'}); $detail="$($req.QueryString['std']) $($req.QueryString['name'])"
        if($req.HttpMethod -ne 'POST'){ throw "$op must be POSTed" }
        $res.ContentType='application/json'
        if($op -eq 'unupload'){
          [PortalApiTest]::Unupload($ReconRoot,$req.QueryString['std'],$req.QueryString['name'],$me)
          $out='{"removed":"'+(JsonStr $req.QueryString['name'])+'"}'
        } else {
          $ms=New-Object IO.MemoryStream; $req.InputStream.CopyTo($ms)
          $saved=[PortalApiTest]::Upload($ReconRoot,$req.QueryString['std'],$req.QueryString['name'],$ms.ToArray(),$me)
          Write-Host ("  + page {0} -> Missing Pages\{1}" -f $saved,$req.QueryString['std']) -ForegroundColor Green
          $detail="$($req.QueryString['std']) $($req.QueryString['name']) saved as $saved ($($ms.Length) bytes)"
          $out='{"saved":"'+(JsonStr $saved)+'","user":"'+(JsonStr $me)+'"}'
        }
        $b=[Text.Encoding]::UTF8.GetBytes($out); $res.OutputStream.Write($b,0,$b.Length)
      }
      else{ throw 'unknown op' }
      if($act){ $a=$act; $act=''; [PortalApiTest]::Log($ReconRoot,$me,$client,$a,$detail,'OK') }
      }
    }
    catch{
      # a .NET method's exception reaches PowerShell wrapped; report the real one
      $msg = if($_.Exception.InnerException){ $_.Exception.InnerException.Message } else { $_.Exception.Message }
      if($act){ try{ [PortalApiTest]::Log($ReconRoot,$me,$client,$act,$detail,"FAILED: $msg") }catch{} }
      $res.StatusCode=400; $res.ContentType='application/json'
      $b=[Text.Encoding]::UTF8.GetBytes('{"error":"'+(JsonStr $msg)+'"}')
      try{ $res.OutputStream.Write($b,0,$b.Length) }catch{}
      Write-Host ("  ! " + $msg) -ForegroundColor Yellow
    }
    try{ $res.Close() }catch{}
  }
}
finally{
  if($listener.IsListening){ $listener.Stop() }
  $listener.Close()
  Write-Host "stopped."
}
