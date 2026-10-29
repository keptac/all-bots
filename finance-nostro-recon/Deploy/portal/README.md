# Nostro manual matching portal

The portal is a small web page where users review outstanding items and raise manual match tickets. It is served at **https://robot.nmbtz.com/portal/** as its own IIS application inside the Orchestrator site.

Users only need a browser. Windows authentication tells the portal who they are; they then press **Sign in** to start a portal session, which ends when they sign out or after 15 minutes without activity. Every ticket records the Windows user that IIS authenticated, and every action is written to an [audit log](#sign-in-and-audit-log).

## How it is deployed

| | Machine | What it does |
|---|---|---|
| Robot | `NMB-DC-RBT-V002` (10.200.225.20) | Runs the recon as `NMBTZ\ROBOT` and writes the reports to `D:\FINANCE AUTOMATED RECON\Nostro`, which is shared as `\\NMB-DC-RBT-V002\NostroRecon`. |
| Orchestrator | `NMB-DC-RBT-V001` (behind `robot.nmbtz.com`, 10.200.200.107) | Hosts IIS. The `/portal` application reads the reports through the share and writes tickets to `Manual Matches`. |

```
Browser ── https://robot.nmbtz.com/portal ──▶ V001: IIS /portal (app pool NostroPortal)
                                                   │ reads runs, writes Manual Matches\<account>.csv
                                                   ▼
                                      \\NMB-DC-RBT-V002\NostroRecon
                                                   ▲
V002: recon (REPORT_ROOT) ─────────────────────────┘  D:\FINANCE AUTOMATED RECON\Nostro
```

On V001 the portal lives in `C:\inetpub\NostroPortal`, and its `web.config` sets `ReconRoot` = `\\NMB-DC-RBT-V002\NostroRecon`.

> V002 also has an old Orchestrator IIS site, bound to `https/10.200.225.20:443:robot.nmbtz.com`. That site is **Stopped** on purpose and is not what serves `robot.nmbtz.com`. Do not start it, and do not install the portal into it.

## Files

| File | Purpose |
|---|---|
| `portal.html` | The whole user interface: page, styles and script in one file. |
| `api.aspx` | Server side: who the user is, the list of runs (one entry per account), an account's portal data, Summary and Workings downloads, appending tickets to the journal, missing statement page uploads, sign-in sessions and the audit log. |
| `web.config` | IIS settings for `/portal` only: handlers, Windows sign-in, security headers, and `ReconRoot`. |
| `Install-Portal.ps1` | Sets up or updates the IIS application. It is safe to re-run. |

### What the recon writes for the portal

Every account of a bank goes into the same run folder, `<date>\<BANK>\`:

- `<BANK>_<CCY>_<ACCOUNT> HH_mm Summary.xlsx`
- `<BANK>_<CCY>_<ACCOUNT> HH_mm Workings.xlsx`
- `recon-data-<ACCOUNT>.json`, the outstanding items the portal shows

Runs from before this layout wrote a single `recon-data.json` per bank folder. Each account overwrote it, so only the last account of each bank was visible. The portal still reads those old files, for the one account named inside each.

## Setup, as it was done

Run the steps on V002 and V001 in an **elevated** PowerShell (Run as administrator). In PowerShell, keep the single quotes in the commands below: in double quotes, `$` starts a variable, so a `$` such as the one in `V001$` breaks the command.

### On V002 (robot machine)

1. Move the reports out of the robot user's profile:
   ```
   robocopy "C:\Users\ROBOT\Documents\FINANCE AUTOMATED RECON\Nostro" "D:\FINANCE AUTOMATED RECON\Nostro" /E /COPY:DAT
   ```
2. Set the folder permissions:
   - The robot needs its own rights, because its jobs don't run elevated.
   - V001's computer account is what the portal uses to read the reports.
   - The admin account is the one used to run the setup on V001.
   ```
   icacls 'D:\FINANCE AUTOMATED RECON\Nostro' /grant 'NMBTZ\ROBOT:(OI)(CI)M' 'NMBTZ\NMB-DC-RBT-V001$:(OI)(CI)RX' 'NMBTZ\emmanuel.charles:(OI)(CI)RX'
   icacls 'D:\FINANCE AUTOMATED RECON\Nostro\Manual Matches' /grant 'NMBTZ\NMB-DC-RBT-V001$:(OI)(CI)M'
   New-Item -ItemType Directory -Force 'D:\FINANCE AUTOMATED RECON\Nostro\Missing Pages'
   icacls 'D:\FINANCE AUTOMATED RECON\Nostro\Missing Pages' /grant 'NMBTZ\NMB-DC-RBT-V001$:(OI)(CI)M'
   New-Item -ItemType Directory -Force 'D:\FINANCE AUTOMATED RECON\Nostro\Audit'
   icacls 'D:\FINANCE AUTOMATED RECON\Nostro\Audit' /grant 'NMBTZ\NMB-DC-RBT-V001$:(OI)(CI)M'
   icacls 'D:\FINANCE AUTOMATED RECON\Nostro\Audit' /deny 'NMBTZ\ROBOT:(OI)(CI)(W,D)'
   ```
   Lines 3 and 4 are for [missing statement pages](#missing-statement-pages), and the last three for the [audit log](#sign-in-and-audit-log). If the portal is already set up, run just the lines you haven't run yet. **The portal refuses to work until the `Audit` folder exists**, because it does nothing it cannot record.
3. Share the folder:
   ```
   New-SmbShare -Name NostroRecon -Path 'D:\FINANCE AUTOMATED RECON\Nostro' -ChangeAccess 'NMBTZ\NMB-DC-RBT-V001$' -ReadAccess 'NMBTZ\emmanuel.charles' -FullAccess 'BUILTIN\Administrators'
   ```
4. In Orchestrator, set the process arguments `REPORT_ROOT` = `D:\FINANCE AUTOMATED RECON\Nostro` and `PORTAL_SETUP` = `False`.

### On V001 (Orchestrator server)

5. Create `C:\PortalSetup`, copy the four portal files into it (copy and paste over RDP works), and clear the downloaded-file flag:
   ```
   Get-ChildItem C:\PortalSetup | Unblock-File
   ```
6. Check that the share is reachable:
   ```
   net use \\NMB-DC-RBT-V002\NostroRecon
   ```
7. Do a dry run, which changes nothing:
   ```
   powershell -NoProfile -ExecutionPolicy Bypass -File "C:\PortalSetup\Install-Portal.ps1" -ReconRoot "\\NMB-DC-RBT-V002\NostroRecon" -SourceDir "C:\PortalSetup" -CheckOnly
   ```
   Confirm that the IIS features are present, that the site it picked is the live (Started) Orchestrator site, and that there is no `FAILED` line.
8. Run the setup itself:
   ```
   powershell -NoProfile -ExecutionPolicy Bypass -File "C:\PortalSetup\Install-Portal.ps1" -ReconRoot "\\NMB-DC-RBT-V002\NostroRecon" -SourceDir "C:\PortalSetup"
   ```
   A warning that ReconRoot is a network path is expected. Step 2 already gave the rights it asks for.
9. Test from a workstation: open https://robot.nmbtz.com/portal/.

### What the setup changes on V001

- It copies the three portal files to `C:\inetpub\NostroPortal`. An existing `web.config` there is kept as `web.config.bak-<timestamp>`.
- It creates the `NostroPortal` app pool (.NET v4.0, Integrated) and the `/portal` application under the Orchestrator site.
- For `/portal` only, it unlocks the handlers section, turns Anonymous authentication off and turns Windows authentication on.

It does not touch Orchestrator's own site, bindings, certificate or app pool, and it does not install Windows features.

### Why not `PORTAL_SETUP = True`?

Main.xaml has a `PORTAL_SETUP` argument that runs this script from Orchestrator. It cannot work in this deployment, for two reasons:

- The robot (V002) is not the IIS server (V001).
- `NMBTZ\ROBOT` is a local admin, but unattended jobs get a UAC-filtered token and cannot change IIS.

The job is still useful for one thing. Its log and email print the exact script path and commands to run by hand.

## Updating the UI

The portal is plain files, so an update is a file copy. Nothing needs restarting and the recon is not affected.

1. Edit `Deploy/portal/portal.html` (and `api.aspx` if needed) in this repository.
2. Test it locally:
   ```
   .\Tools\Test-Portal.ps1 -ReconRoot "<a copy of a recon folder>"
   ```
   This serves the page on `http://localhost:8731` against real run folders, with no IIS and no admin rights. It mirrors `api.aspx` but is not `api.aspx`, and it doesn't test the IIS settings.
3. Publish the package, then trigger the Nostro process **once** with **`PORTAL_UPDATE` = `True`**. The recon does not run. See [Automated updates](#automated-updates).
4. Hard-refresh the page (Ctrl+F5). The browser may keep the old `portal.html` cached. The data itself is never cached, because `api.aspx` sends no-cache headers.

**If the change needs new data:** the page reads `recon-data-<account>.json`, which the recon writes from `Matching.xaml` ("Write recon report"). A field the page doesn't receive yet has to be added there too. It only appears after the next recon run.

### Automated updates

`PORTAL_UPDATE = True` runs `Portal Update.xaml` instead of the recon:

1. It copies `portal.html` and `api.aspx` from the package to `PORTAL_SHARE` (default `\\NMB-DC-RBT-V001\NostroPortal`). A file that is already identical is skipped.
2. Before overwriting a file, it keeps the live one as `<name>.<yyyyMMdd-HHmmss>.bak` beside it, and keeps the newest 10 of each. IIS doesn't serve `.bak` files.
3. It never copies `web.config`, which holds `ReconRoot` and the sign-in settings.
4. It fetches `https://robot.nmbtz.com/portal/portal.html` and compares it byte for byte with the package copy, so "UPDATED" means users really get the new page.
5. It emails the result with one of three subjects:
   - **UPDATED:** copied and confirmed live.
   - **COPIED, LIVE CHECK FAILED:** the files were copied, but check in a browser.
   - **UPDATE FAILED:** the job is marked Faulted.

Set `PORTAL_UPDATE` back to `False` afterwards. If `PORTAL_SETUP` is also `True`, setup runs instead.

**One-time setup on V001**, in an elevated PowerShell, so the robot can write to the portal folder and nowhere else:
```
icacls 'C:\inetpub\NostroPortal' /grant 'NMBTZ\ROBOT:(OI)(CI)M'
icacls 'C:\inetpub\NostroPortal\web.config' /deny 'NMBTZ\ROBOT:(W,D)'
New-SmbShare -Name NostroPortal -Path 'C:\inetpub\NostroPortal' -ChangeAccess 'NMBTZ\ROBOT' -FullAccess 'BUILTIN\Administrators'
```
The second line keeps the robot from ever changing or deleting `web.config`. Check from V002, as the robot or any account with the same rights:
```
Test-Path '\\NMB-DC-RBT-V001\NostroPortal\web.config'
```

**Rolling back:** on V001, copy the `.bak` you want over the live file, for example `portal.html.20261026-151426.bak` over `portal.html`.

**Changes to `web.config` or IIS** are not automated. Copy the four files to `C:\PortalSetup` on V001 and re-run setup step 8. It backs up the live `web.config`, re-applies `ReconRoot`, and is safe to repeat. **Never copy the repository's `web.config` over the live one:** its `ReconRoot` points at a development folder.

## Missing statement pages

Finance users can't reach the servers, so when a statement page never reaches the robot they upload it on the portal: **Missing pages** in the header.

1. They choose the kind of statement (ISO 20022 camt.053 `.xml`, legacy MT940/950 `.txt`, or legacy detailed `.txt`) and upload one or more files.
2. `api.aspx` checks each file and saves it to `Missing Pages\<SWIFTMX|SWIFT|SWIFTDETAIL>` on the recon share. It refuses:
   - the wrong file type, including a `.zip` (unzip it first)
   - a camt file that holds no statement
   - a file already waiting, or one already reconciled
   - any file over 4 MB
3. The next unattended run stages every waiting page with that day's statements, whatever the file's date. After the reports are written, it moves the pages to `Missing Pages\Processed\<run time>`. The run email lists each page, who uploaded it, and what happened to it.
4. Until that run, the portal lists the page as waiting and the uploader can **take it back**. If the run fails, the pages stay waiting and the next run tries again.

Every upload and take-back is recorded in `Missing Pages\upload-log.csv` with the Windows user IIS authenticated.

**Permissions:** V001's computer account needs Modify on `Missing Pages` (setup step 2). The portal can't create the folder itself, because it only has Read on the rest of the share.

## Sign-in and audit log

**Signing in.** Windows authentication already tells the portal who is at the browser, so there is no password to type. The first page shows the Windows account and a **Sign in** button, which starts a portal session. The session ends when the user:

- presses **Sign out**, or
- does nothing with the mouse or keyboard in any portal tab for 15 minutes. A warning with a countdown appears one minute before.

The page then reloads to the sign-in screen, so nothing from the session stays on screen. The server refuses every other request without a live session for the same Windows user, so a session can't be carried to another account. To change the limit, add `<add key="IdleMinutes" value="20" />` to `appSettings` in the live `web.config` on V001 (5 to 240).

Sessions are held in memory on V001. Recycling the `NostroPortal` app pool, or deploying a new `api.aspx`, signs everyone out. They only need to press Sign in again; tickets not yet saved stay in their browser.

**The audit log** is `Audit\portal-audit-<yyyy-MM>.csv` on the recon share, one file per month, with times in V001's local time:

| Column | What it holds |
|---|---|
| `When` | Date and time of the action. |
| `User` | The Windows user IIS authenticated, never a name the page sent. |
| `Client` | The user's IP address as V001 sees it. |
| `Action` | What was done (below). |
| `Detail` | Which run, account, file or tickets. |
| `Result` | `OK`, or `FAILED <code>: <reason>`. |

| Action | Recorded when |
|---|---|
| `SIGN_IN`, `SIGN_OUT` | The user signs in, signs out, or is signed out for inactivity. |
| `SESSION_EXPIRED` | The server ends a session nobody signed out of, such as a closed browser. |
| `DENIED` | A request arrives without a live session. |
| `LIST_RUNS`, `OPEN_RUN`, `DOWNLOAD` | The run list loads, a run is opened, a Summary or Workings is downloaded. |
| `SAVE_TICKETS` | Tickets are saved to the journal, with the account, leg count and Match Refs. |
| `LIST_PAGES`, `UPLOAD_PAGE`, `TAKE_BACK_PAGE` | Missing statement pages are listed, uploaded or taken back. |
| `SELECT_ACCOUNT`, `VIEW_MATCHING`, `VIEW_HISTORY`, `VIEW_PAGES` | The user changes account or screen. |
| `SUGGEST`, `ADD_TICKET`, `REMOVE_TICKET`, `DOWNLOAD_JOURNAL` | The user asks for suggestions, builds or removes a ticket before saving, or downloads the journal instead of saving it. |

Lines are only ever appended. Before every action the server checks that it can write the log; **if it can't, the action is refused** and the user sees why. Two things cause that:

- **The `Audit` folder doesn't exist.** Run the `Audit` lines in setup step 2.
- **Someone has the month's file open in Excel**, which locks it. Open a copy instead.

`ROBOT` is denied write access to `Audit` in step 2, so the robot account can't change the log.

## When something goes wrong

| Symptom | Cause | Fix |
|---|---|---|
| `/portal` redirects to `/identity/` | The request reached Orchestrator, not the portal app: `/portal` is missing on the server that answers, or was installed into the stopped site on V002. | Check `appcmd list app` on V001 for `<site>/portal`. |
| `curl` cannot connect on `127.0.0.1:443` | The site is bound to one IP (`https/<ip>:443:robot.nmbtz.com`), not all addresses. | Test against the server's own IP instead. |
| `No mapping between account names and security IDs` (error 1332) | A name in `icacls` or `New-SmbShare` doesn't exist in the domain. | Use `whoami` and `hostname` for the real names. A computer account ends in `$`. |
| `Variable reference is not valid` | A `$` inside double quotes in PowerShell. | Use single quotes. |
| `-File ... does not exist`, or `UnauthorizedAccessException` reading the setup folder | The copy to V001 failed, or the folder's permissions exclude you. | Create `C:\PortalSetup` yourself and paste the files in. |
| `Cannot bind argument to parameter 'Path' because it is an empty string` | Older script: `$PSScriptRoot` is empty under `-File`. | Pass `-SourceDir`. Fixed in the current script. |
| `FAILED: ... is not elevated` | PowerShell is not running as administrator. | Reopen it with Run as administrator. |
| The account dropdown shows only one account per bank | A run from before per-account data files: each account of a bank overwrote the same `recon-data.json`. | Runs from the current recon have every account. Re-run a date with `STATEMENT_DATE` to rebuild it. |
| `PORTAL_UPDATE` email: "Cannot reach \\…\NostroPortal" | The one-time share on V001 is missing, or the robot has no rights on it. | Run the three one-time commands under [Automated updates](#automated-updates). |
| `PORTAL_UPDATE` email: "LIVE CHECK FAILED" | The files were copied, but the page fetched from `robot.nmbtz.com` differs, or the fetch failed. | Open the portal and press Ctrl+F5. If it is still old, compare `C:\inetpub\NostroPortal\portal.html` with the repository copy. |
| Portal loads but shows no runs | V001 can't read the share. | Re-check steps 2 and 3, then `net use` from V001. |
| `Saving` a ticket fails | V001's computer account lacks Modify on `Manual Matches`. | Re-run the second `icacls` line in step 2. |
| Missing pages: "The Missing Pages folder has not been set up" | The folder doesn't exist on the share. | Run the `New-Item` and `icacls` lines for `Missing Pages` in step 2. |
| "The Audit folder has not been set up on the recon share yet" | The portal refuses to act without its audit log. | Run the `Audit` lines in step 2. |
| "The audit log … is locked by another program" | The month's audit file is open in Excel. | Close it. Open a copy to read it. |
| Everyone is back at the sign-in screen at once | The app pool recycled, or a new `api.aspx` was deployed. Sessions are held in memory. | Nothing to fix. Users press Sign in again. |
| "The portal server could not load the reconciliation runs" | `api.aspx?op=runs` failed. The message under it is the server's own error. | Act on that message. "was refused access" names the account and path: grant it Read on the share (steps 2 and 3), or Modify for `Manual Matches` / `Missing Pages`. |
| "The portal server (…) was refused access: …" | The named account can't read or write the named path. | Steps 2 and 3. `Manual Matches` and `Missing Pages` need Modify, everything else Read. |
| A bare "500 - Internal server error" page | An older `api.aspx` whose error message IIS hid from other machines. | Deploy the current `api.aspx`, or open the same URL on V001 itself to see the message. |
| 401 in the browser | Windows sign-in isn't reaching the browser. | Add `robot.nmbtz.com` to the Local intranet zone. |
| 500.19 | An IIS section is locked for `/portal`. | Send the error detail. The script unlocks only `handlers`. |
