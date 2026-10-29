<#
  Install-Portal.ps1
  ---------------------------------------------------------------------------
  One-time setup of the Nostro manual matching portal as a child APPLICATION
  of the Orchestrator site, so it answers at  https://robot.nmbtz.com/portal

  Run by Main.xaml when PORTAL_SETUP = True (the recon does not run), or by
  hand from an elevated PowerShell on the Orchestrator server. Safe to re-run:
  every step checks what is there and only adds or corrects it.

  What it changes - and nothing else:
    - copies portal.html, api.aspx and web.config to -PhysicalPath, with
      ReconRoot in web.config set to -ReconRoot (an existing web.config there
      is kept as web.config.bak-<timestamp>)
    - creates or corrects the app pool -AppPoolName (.NET v4.0, Integrated)
    - creates or corrects the application -AppPath under the Orchestrator site
    - for -AppPath ONLY: unlocks the handlers section (the portal's web.config
      clears Orchestrator's handlers), turns Anonymous authentication off and
      Windows authentication on
    - grants the app pool identity Read on -ReconRoot and Modify on
      <ReconRoot>\Manual Matches (created if missing)

  It does not touch the Orchestrator site's own settings, bindings,
  certificate or app pool, and it does not install Windows features: if IIS
  Windows Authentication or ASP.NET 4.x is missing it stops and says so.

  Exit codes: 0 set up and answering; 2 set up but the HTTPS check failed;
              1 nothing or only part was set up (see the log).

  Usage
      .\Install-Portal.ps1 -ReconRoot "D:\FINANCE AUTOMATED RECON\Nostro"
      .\Install-Portal.ps1 -ReconRoot ... -SiteName "UiPath Orchestrator"
      .\Install-Portal.ps1 -ReconRoot ... -CheckOnly     # report, change nothing
#>
[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$ReconRoot,
  [string]$HostName     = "robot.nmbtz.com",
  [string]$SiteName,
  [string]$AppPath      = "/portal",
  [string]$PhysicalPath = "C:\inetpub\NostroPortal",
  [string]$AppPoolName  = "NostroPortal",
  [string]$SourceDir    = $PSScriptRoot,
  [switch]$CheckOnly
)

$ErrorActionPreference = 'Stop'
$appcmd = Join-Path $env:windir 'system32\inetsrv\appcmd.exe'

# Windows PowerShell 5.1 can leave $PSScriptRoot empty in a param() default when
# the script is started with -File, so fall back to the script's own location.
if (-not $SourceDir) { $SourceDir = $PSScriptRoot }
if (-not $SourceDir -and $MyInvocation.MyCommand.Path) { $SourceDir = Split-Path -Parent $MyInvocation.MyCommand.Path }

function Step([string]$m) { Write-Output ("[{0}] {1}" -f (Get-Date -Format 'HH:mm:ss'), $m) }

function Invoke-AppCmd {
  param([Parameter(ValueFromRemainingArguments = $true)][string[]]$CmdArgs)
  $out = & $appcmd @CmdArgs 2>&1 | Out-String
  if ($LASTEXITCODE -ne 0) { throw ("appcmd " + ($CmdArgs -join ' ') + " failed (" + $LASTEXITCODE + "): " + $out.Trim()) }
  return $out
}

# Picks the site that serves https://<HostName>: the -SiteName given, else the
# one whose https binding names the host, else the only site listening on 443.
function Select-Site([xml]$sites, [string]$wanted, [string]$hostName) {
  $all = @($sites.appcmd.SITE)
  if ($wanted) {
    $s = $all | Where-Object { $_.'SITE.NAME' -eq $wanted }
    if (-not $s) { throw "No IIS site named '$wanted'. Sites: " + (($all | ForEach-Object { $_.'SITE.NAME' + ' [' + $_.bindings + ']' }) -join '; ') }
    return $s.'SITE.NAME'
  }
  $byHost = @($all | Where-Object { ($_.bindings -split ',') -match ('^https/[^,]*:443:' + [regex]::Escape($hostName) + '$') })
  if ($byHost.Count -eq 1) { return $byHost[0].'SITE.NAME' }
  $on443 = @($all | Where-Object { ($_.bindings -split ',') -match '^https/[^,]*:443:' })
  if ($byHost.Count -eq 0 -and $on443.Count -eq 1) { return $on443[0].'SITE.NAME' }
  throw "Could not tell which IIS site serves https://$hostName - pass -SiteName. Sites: " + (($all | ForEach-Object { $_.'SITE.NAME' + ' [' + $_.bindings + ']' }) -join '; ')
}

# Writes ReconRoot into the portal's web.config, adding the key if it is absent.
function Set-ReconRoot([string]$configPath, [string]$root) {
  [xml]$cfg = Get-Content -Raw -LiteralPath $configPath
  $app = $cfg.configuration.SelectSingleNode('appSettings')
  if (-not $app) { $app = $cfg.configuration.AppendChild($cfg.CreateElement('appSettings')) }
  $add = $app.SelectSingleNode("add[@key='ReconRoot']")
  if (-not $add) {
    $add = $app.AppendChild($cfg.CreateElement('add'))
    $add.SetAttribute('key', 'ReconRoot')
  }
  $add.SetAttribute('value', $root)
  $cfg.Save($configPath)
}

if ($MyInvocation.InvocationName -eq '.') { return }   # dot-sourced for testing: functions only

$failed = $false
$siteOk = $false
try {
  Step "Nostro portal setup on $env:COMPUTERNAME as $env:USERDOMAIN\$env:USERNAME$(if ($CheckOnly) { ' (CHECK ONLY - nothing will be changed)' })"

  # ------------------------------------------------------------ prerequisites --
  # Is this the machine behind the portal URL? Informational only: behind a load
  # balancer the addresses differ even on the right server.
  $hostIps = @()
  try { $hostIps = @([Net.Dns]::GetHostAddresses($HostName) | ForEach-Object { $_.IPAddressToString }) } catch { }
  $myIps = @([Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces() | ForEach-Object { $_.GetIPProperties().UnicastAddresses } | ForEach-Object { $_.Address.IPAddressToString })
  $match = @($hostIps | Where-Object { $myIps -contains $_ }).Count -gt 0
  Step ("$HostName resolves to " + $(if ($hostIps.Count) { $hostIps -join ', ' } else { '(nothing)' }) + "; this server is " + ($myIps -join ', ') + $(if ($match) { ' - match' } else { ' - NO match: if Orchestrator is not hosted here, run this on the Orchestrator server instead' }))

  # IIS first: if it is not here, nothing else matters.
  if (-not (Test-Path $appcmd)) { throw "IIS is not installed on $env:COMPUTERNAME ($appcmd not found). PORTAL_SETUP has to run on the server that hosts Orchestrator ($HostName)." }

  $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
  if (-not $isAdmin -and -not $CheckOnly) {
    # A local admin under UAC still gets a filtered token in an unattended job:
    # Administrators is then present but "used for deny only".
    $admLine = @((whoami /groups) | Where-Object { $_ -match 'S-1-5-32-544' })
    $me = "$env:USERDOMAIN\$env:USERNAME"
    if ($admLine.Count -gt 0 -and $admLine[0] -match 'deny only') {
      throw "$me IS a local administrator here, but this process is not elevated - UAC gives unattended jobs a filtered token, and a job cannot raise it itself. Run this script once by hand from PowerShell opened with 'Run as administrator' on $env:COMPUTERNAME (see the usage at the top of the script)."
    }
    throw "$me is not a local administrator on $env:COMPUTERNAME, which changing IIS requires. Either add it to the local Administrators group for the setup, or run this script once by hand as an administrator from an elevated PowerShell."
  }

  $global = Invoke-AppCmd list config '-section:system.webServer/globalModules'
  $missing = @()
  if ($global -notmatch 'WindowsAuthenticationModule') { $missing += 'Web-Windows-Auth (IIS Windows Authentication)' }
  if ($global -notmatch 'ManagedEngineV4\.0_64bit')   { $missing += 'Web-Asp-Net45 (ASP.NET 4.x, needed by api.aspx)' }
  if ($missing.Count -gt 0) { throw ("Missing IIS features: " + ($missing -join ', ') + ". Install them (Install-WindowsFeature Web-Windows-Auth, Web-Asp-Net45) and re-run.") }
  Step "IIS, Windows Authentication and ASP.NET 4.x are present"

  if (-not $SourceDir) { throw "Could not tell which folder holds the portal files - pass -SourceDir with the folder containing portal.html, api.aspx and web.config." }
  foreach ($f in 'portal.html', 'api.aspx', 'web.config') {
    if (-not (Test-Path -LiteralPath (Join-Path $SourceDir $f))) { throw "Portal file $f not found in $SourceDir" }
  }
  Step "Portal files from $SourceDir"

  $ReconRoot = $ReconRoot.Trim().TrimEnd('\')
  if (-not (Test-Path -LiteralPath $ReconRoot)) { throw "ReconRoot '$ReconRoot' does not exist on this server. Point it at the folder the recon writes its dated run folders to (Main's REPORT_ROOT)." }
  if ($ReconRoot.StartsWith('\\')) {
    Step "WARNING - ReconRoot is a network path. The app pool reaches it as the server's computer account ($env:USERDOMAIN\$env:COMPUTERNAME`$), which needs Read on the share and Modify on 'Manual Matches' and 'Missing Pages' there - this script can only set rights on local folders."
  }

  [xml]$sites = Invoke-AppCmd list site /xml
  $site = Select-Site $sites $SiteName $HostName
  # A stopped site here means the live Orchestrator runs elsewhere - adding the
  # portal to it would publish nothing.
  $siteState = (@($sites.appcmd.SITE) | Where-Object { $_.'SITE.NAME' -eq $site }).state
  if ($siteState -and $siteState -ne 'Started') {
    throw "IIS site '$site' on $env:COMPUTERNAME is $siteState, so this machine is not serving https://$HostName. Run this on the server that is (the one behind $(if ($hostIps.Count) { $hostIps -join ', ' } else { $HostName }))."
  }
  $siteOk = $true
  $appId = $site + $AppPath
  Step "Orchestrator site: '$site'  ->  application '$appId'  ->  $PhysicalPath"

  if ($CheckOnly) {
    $pools = [xml](Invoke-AppCmd list apppool /xml)
    $apps  = [xml](Invoke-AppCmd list app /xml)
    Step ("App pool '$AppPoolName': " + $(if (@($pools.appcmd.APPPOOL) | Where-Object { $_.'APPPOOL.NAME' -eq $AppPoolName }) { 'exists' } else { 'would be created' }))
    Step ("Application '$appId': " + $(if (@($apps.appcmd.APP) | Where-Object { $_.'APP.NAME' -eq $appId }) { 'exists' } else { 'would be created' }))
    Step "Check only - stopping before any change."
    exit 0
  }

  # ------------------------------------------------------------------- files --
  New-Item -ItemType Directory -Force -Path $PhysicalPath | Out-Null
  $existing = Join-Path $PhysicalPath 'web.config'
  if (Test-Path -LiteralPath $existing) {
    $bak = $existing + '.bak-' + (Get-Date -Format 'yyyyMMdd-HHmmss')
    Copy-Item -LiteralPath $existing -Destination $bak
    Step "Kept the previous web.config as $bak"
  }
  foreach ($f in 'portal.html', 'api.aspx', 'web.config') {
    Copy-Item -LiteralPath (Join-Path $SourceDir $f) -Destination (Join-Path $PhysicalPath $f) -Force
  }
  Set-ReconRoot $existing $ReconRoot
  Step "Copied portal files; web.config ReconRoot = $ReconRoot"

  # ---------------------------------------------------------------- app pool --
  $pools = [xml](Invoke-AppCmd list apppool /xml)
  if (@($pools.appcmd.APPPOOL) | Where-Object { $_.'APPPOOL.NAME' -eq $AppPoolName }) {
    Invoke-AppCmd set apppool $AppPoolName '/managedRuntimeVersion:v4.0' '/managedPipelineMode:Integrated' | Out-Null
    Step "App pool '$AppPoolName' already existed - set to .NET v4.0, Integrated"
  } else {
    Invoke-AppCmd add apppool "/name:$AppPoolName" '/managedRuntimeVersion:v4.0' '/managedPipelineMode:Integrated' | Out-Null
    Step "Created app pool '$AppPoolName' (.NET v4.0, Integrated)"
  }

  # ------------------------------------------------------------- application --
  $apps = [xml](Invoke-AppCmd list app /xml)
  if (@($apps.appcmd.APP) | Where-Object { $_.'APP.NAME' -eq $appId }) {
    Invoke-AppCmd set app $appId "/applicationPool:$AppPoolName" | Out-Null
    Invoke-AppCmd set vdir "$appId/" "/physicalPath:$PhysicalPath" | Out-Null
    Step "Application '$appId' already existed - pointed it at $PhysicalPath and '$AppPoolName'"
  } else {
    Invoke-AppCmd add app "/site.name:$site" "/path:$AppPath" "/physicalPath:$PhysicalPath" "/applicationPool:$AppPoolName" | Out-Null
    Step "Created application '$appId'"
  }

  # ------------------------------------------------- per-application settings --
  # Each of these is written as a <location path="$appId"> in applicationHost,
  # so the Orchestrator site itself is unchanged.
  Invoke-AppCmd unlock config $appId '-section:system.webServer/handlers' | Out-Null
  Invoke-AppCmd set config $appId '-section:system.webServer/security/authentication/anonymousAuthentication' '/enabled:False' '/commit:apphost' | Out-Null
  Invoke-AppCmd set config $appId '-section:system.webServer/security/authentication/windowsAuthentication' '/enabled:True' '/commit:apphost' | Out-Null
  Step "For $appId only: handlers unlocked, Anonymous auth OFF, Windows auth ON"

  # ------------------------------------------------------------- permissions --
  $identity = "IIS AppPool\$AppPoolName"
  if (-not $ReconRoot.StartsWith('\\')) {
    $mm = Join-Path $ReconRoot 'Manual Matches'
    New-Item -ItemType Directory -Force -Path $mm | Out-Null
    & icacls $ReconRoot /grant "${identity}:(OI)(CI)RX" | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "icacls could not grant Read on $ReconRoot" }
    & icacls $mm /grant "${identity}:(OI)(CI)M" | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "icacls could not grant Modify on $mm" }
    & icacls $PhysicalPath /grant "${identity}:(OI)(CI)RX" | Out-Null
    Step "Granted $identity Read on $ReconRoot and Modify on $mm"
  }

  Invoke-AppCmd recycle apppool "/apppool.name:$AppPoolName" | Out-Null
} catch {
  $failed = $true
  Step ("FAILED: " + $_.Exception.Message)
}

if ($failed) {
  Step "Setup did not complete. Nothing after the failed step was done; fix the cause and run PORTAL_SETUP again."
  exit 1
}

# ------------------------------------------------------------------ verify --
try {
  [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
  $base = "https://$HostName$AppPath"
  $who = Invoke-WebRequest -Uri "$base/api.aspx?op=whoami" -UseDefaultCredentials -UseBasicParsing -TimeoutSec 60
  $page = Invoke-WebRequest -Uri "$base/" -UseDefaultCredentials -UseBasicParsing -TimeoutSec 60
  Step ("Checked $base/ -> HTTP " + $page.StatusCode + "; api.aspx?op=whoami -> HTTP " + $who.StatusCode + " " + $who.Content)
  # whoami proves sign-in only. Signing in proves the audit log can be written,
  # and listing the runs that the server can read ReconRoot. All three are
  # recorded in the audit log against the account running this script.
  try {
    $null = Invoke-WebRequest -Uri "$base/api.aspx?op=login" -Method Post -SessionVariable portal -UseDefaultCredentials -UseBasicParsing -TimeoutSec 60
    $runs = Invoke-WebRequest -Uri "$base/api.aspx?op=runs" -WebSession $portal -UseDefaultCredentials -UseBasicParsing -TimeoutSec 120
    $null = Invoke-WebRequest -Uri "$base/api.aspx?op=logout" -Method Post -WebSession $portal -UseDefaultCredentials -UseBasicParsing -TimeoutSec 60
    Step ("api.aspx sign-in, ?op=runs -> HTTP " + $runs.StatusCode + ", " + $runs.Content.Length + " bytes, sign-out - all recorded in the audit log")
  } catch {
    $body = ""
    try { $body = (New-Object IO.StreamReader($_.Exception.Response.GetResponseStream())).ReadToEnd() } catch { }
    Step ("Setup finished, but signing in to the portal or listing the runs failed: " + $_.Exception.Message + " " + $body)
    Step "Users will see that message in the portal. Fix it (usually the Audit folder, or rights on ReconRoot) and refresh."
    exit 2
  }
  Step "DONE - users open $base/ in their browser and are signed in as themselves."
  exit 0
} catch {
  Step ("Setup finished, but the HTTPS check of https://$HostName$AppPath failed: " + $_.Exception.Message)
  Step "Open https://$HostName$AppPath/ from a workstation to confirm; a 401 there means Windows auth is not reaching the browser, a 500.19 means a locked IIS section."
  exit 2
}
