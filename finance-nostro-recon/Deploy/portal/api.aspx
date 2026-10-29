<%@ Page Language="C#" EnableViewState="false" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Text" %>
<%@ Import Namespace="System.Text.RegularExpressions" %>
<%@ Import Namespace="System.Collections.Generic" %>
<%@ Import Namespace="System.Linq" %>
<%@ Import Namespace="System.Globalization" %>
<script runat="server">
//
//  Server side of the Nostro manual matching portal.
//
//  The people using this are on their own workstations; the reconciliation
//  output lives on this server. So the browser cannot reach the files directly -
//  everything it needs comes through here, over the same authenticated HTTPS
//  origin that serves the page.
//
//  Operations (all require an authenticated Windows identity):
//     ?op=whoami                        who IIS says you are, and whether you have a portal session
//     POST ?op=login                    start a portal session (the sign-in page's button)
//     POST ?op=logout&reason=<manual|idle>   end it
//  Everything below also needs a live portal session for that same Windows user:
//     POST ?op=ping                     the page is in use - keeps the session from going idle
//     POST ?op=event&action=<ACTION>    record something done in the page alone (body = detail)
//     ?op=runs                                    index of runs: one entry per account per run
//     ?op=data&run=<rel>&account=<acct>           that account's recon data
//     ?op=file&run=<rel>&account=<acct>&kind=summary   download its Summary / Workings
//     POST ?op=tickets&account=<acct>             append approved legs to that account's journal
//     ?op=pages                                   missing statement pages waiting for the robot, and recent batches
//     POST ?op=upload&std=<SWIFTMX|SWIFT|SWIFTDETAIL>&name=<file>   body = one page file
//     POST ?op=unupload&std=<std>&name=<file>     take back a page the robot has not used yet
//
//  Finance cannot reach the servers, so a statement page that never arrived is
//  downloaded by hand and uploaded here. It lands in <ReconRoot>\Missing Pages\<std>,
//  which the next unattended run stages alongside the normal files and then moves
//  to Missing Pages\Processed. Who uploaded what is kept in Missing Pages\upload-log.csv.
//
//  A run folder is <date>\<bank>, and every account of that bank is written into
//  it: "<BANK>_<CCY>_<ACCOUNT> HH_mm Summary.xlsx", the matching Workings, and
//  "recon-data-<ACCOUNT>.json". Runs before the per-account data file wrote one
//  "recon-data.json" per folder, which each account overwrote; it is still read,
//  for the account named inside it.
//
//  Every action is appended to <ReconRoot>\Audit\portal-audit-<yyyy-MM>.csv with
//  the Windows user IIS authenticated. The log is checked to be writable BEFORE an
//  action runs: if it cannot be written, nothing is done.
//
//  Every path is resolved and then checked to be inside ReconRoot, and only the
//  three known file shapes can be read. The journal is APPEND ONLY and the user
//  recorded on a ticket is the one IIS authenticated - never a name the browser
//  supplied, which would be worthless for audit.
//

static readonly Regex SafeId  = new Regex(@"^[A-Za-z0-9_\-]{1,64}$", RegexOptions.Compiled);
static readonly Regex SafeRef = new Regex(@"^[A-Za-z0-9_\-]{1,40}$", RegexOptions.Compiled);
static readonly Regex AcctRx  = new Regex(@"(\d{6,})", RegexOptions.Compiled);
static readonly Regex CcyRx   = new Regex(@"_([A-Z]{3})_\d{6,}", RegexOptions.Compiled);
static readonly Regex DataRx  = new Regex(@"^recon-data-([A-Za-z0-9_\-]+)\.json$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
static readonly Regex JsonAcctRx = new Regex("\"account\"\\s*:\\s*\"([^\"]*)\"", RegexOptions.Compiled);

string Root
{
    get
    {
        string r = System.Configuration.ConfigurationManager.AppSettings["ReconRoot"];
        if (string.IsNullOrEmpty(r)) throw new Exception("ReconRoot is not configured in web.config");
        return Path.GetFullPath(r).TrimEnd(Path.DirectorySeparatorChar);
    }
}

bool Inside(string full)
{
    string root = Root + Path.DirectorySeparatorChar;
    return full.StartsWith(root, StringComparison.OrdinalIgnoreCase);
}

string ResolveRun(string rel)
{
    if (string.IsNullOrEmpty(rel)) throw new Exception("run is required");
    if (rel.Contains("..")) throw new Exception("bad run");
    string full = Path.GetFullPath(Path.Combine(Root, rel.Replace('/', Path.DirectorySeparatorChar)));
    if (!Inside(full) || !Directory.Exists(full)) throw new Exception("bad run");
    return full;
}

static string J(string s)
{
    if (s == null) return "";
    var b = new StringBuilder();
    foreach (char c in s)
    {
        if (c == '"' || c == '\\') b.Append('\\').Append(c);
        else if (c == '\n' || c == '\r' || c == '\t') b.Append(' ');
        else if (c < 32) { }
        else b.Append(c);
    }
    return b.ToString();
}

static string Csv(string s)
{
    s = (s ?? "").Replace("\r", " ").Replace("\n", " ");
    return (s.IndexOf(',') >= 0 || s.IndexOf('"') >= 0) ? "\"" + s.Replace("\"", "\"\"") + "\"" : s;
}


class AccountRun
{
    public string Id, Date, Bank, Acct, Ccy;
    public long When, Ord;
    public FileInfo Data, Sum, Wrk;
}

static string AcctOf(string name)
{
    var m = AcctRx.Match(name ?? "");
    return m.Success ? m.Groups[1].Value : "";
}

static long Ms(DateTime utc) { return (long)(utc - new DateTime(1970, 1, 1)).TotalMilliseconds; }

// The account a recon-data file belongs to: in the name for the per-account
// file, inside the JSON for the old single recon-data.json.
static string DataAccount(FileInfo f)
{
    var m = DataRx.Match(f.Name);
    if (m.Success) return m.Groups[1].Value;
    if (!f.Name.Equals("recon-data.json", StringComparison.OrdinalIgnoreCase)) return "";
    try
    {
        using (var rd = new StreamReader(f.FullName, Encoding.UTF8))
        {
            var buf = new char[2048];
            int n = rd.Read(buf, 0, buf.Length);
            var mm = JsonAcctRx.Match(new string(buf, 0, n));
            return mm.Success ? mm.Groups[1].Value : "";
        }
    }
    catch { return ""; }
}

static FileInfo Newest(IEnumerable<FileInfo> files)
{
    return files.Where(f => !f.Name.StartsWith("~$")).OrderByDescending(f => f.LastWriteTimeUtc).FirstOrDefault();
}

// Every account the recon wrote into one <date>\<bank> folder.
static List<AccountRun> AccountsIn(DirectoryInfo bank)
{
    var files = bank.GetFiles();
    var sums  = files.Where(f => f.Name.EndsWith("Summary.xlsx", StringComparison.OrdinalIgnoreCase)).ToList();
    var wrks  = files.Where(f => f.Name.EndsWith("Workings.xlsx", StringComparison.OrdinalIgnoreCase)).ToList();
    var datas = files.Where(f => f.Name.StartsWith("recon-data", StringComparison.OrdinalIgnoreCase)
                              && f.Name.EndsWith(".json", StringComparison.OrdinalIgnoreCase)).ToList();

    // the per-account file wins over the old shared one for the same account
    var dataBy = new Dictionary<string, FileInfo>();
    foreach (var f in datas.Where(f => !DataRx.IsMatch(f.Name)))
    { string a = DataAccount(f); if (a.Length > 0) dataBy[a] = f; }
    foreach (var f in datas.Where(f => DataRx.IsMatch(f.Name)))
    { string a = DataAccount(f); if (a.Length > 0) dataBy[a] = f; }

    var accts = new List<string>();
    foreach (var f in sums.Concat(wrks))
    { string a = AcctOf(f.Name); if (a.Length > 0 && !accts.Contains(a)) accts.Add(a); }
    foreach (var a in dataBy.Keys) if (!accts.Contains(a)) accts.Add(a);

    var outp = new List<AccountRun>();
    foreach (var a in accts)
    {
        string acct = a;
        var r = new AccountRun { Acct = acct, Bank = bank.Name };
        r.Sum  = Newest(sums.Where(f => AcctOf(f.Name) == acct));
        r.Wrk  = Newest(wrks.Where(f => AcctOf(f.Name) == acct));
        r.Data = dataBy.ContainsKey(acct) ? dataBy[acct] : null;
        var named = r.Sum ?? r.Wrk;
        var cm = named != null ? CcyRx.Match(named.Name) : Match.Empty;
        r.Ccy = cm.Success ? cm.Groups[1].Value : "";
        r.When = new[] { r.Data, r.Sum, r.Wrk }.Where(f => f != null).Max(f => Ms(f.LastWriteTimeUtc));
        outp.Add(r);
    }
    return outp;
}

// root\<date>\<bank>\ , skipping the folders the recon writes for its own use
List<AccountRun> Runs()
{
    var outp = new List<AccountRun>();
    var root = new DirectoryInfo(Root);
    if (!root.Exists) return outp;

    foreach (var day in root.GetDirectories().Where(x => !IsSkipped(x.Name)))
    {
        // Order by the RUN DATE in the folder name, not by file time: applying
        // manual matches rewrites a Summary, which would otherwise make an old
        // run look like the newest one.
        DateTime runDay;
        bool dated = DateTime.TryParseExact(day.Name, new[]{"d-MMM-yyyy","dd-MMM-yyyy"},
                         CultureInfo.InvariantCulture, DateTimeStyles.None, out runDay);

        foreach (var bank in day.GetDirectories().Where(x => !IsSkipped(x.Name)))
            foreach (var r in AccountsIn(bank))
            {
                r.Id = day.Name + "/" + bank.Name;
                r.Date = day.Name;
                r.Ord = dated ? Ms(runDay) : r.When;
                outp.Add(r);
            }
    }
    return outp.OrderByDescending(r => r.Ord).ThenByDescending(r => r.When).ToList();
}

// One account of one run, from the run folder and the account in the request.
AccountRun RunFor(string run, string acct)
{
    var dir = new DirectoryInfo(ResolveRun(run));
    var all = AccountsIn(dir);
    acct = acct ?? "";
    if (acct.Length == 0 && all.Count == 1) return all[0];
    if (!SafeId.IsMatch(acct)) Fail(400, "bad account");
    var hit = all.FirstOrDefault(x => x.Acct == acct);
    if (hit == null) Fail(404, "no account " + acct + " in this run");
    return hit;
}

static bool IsSkipped(string name)
{
    return name.Equals("backup", StringComparison.OrdinalIgnoreCase)
        || name.Equals("Manual Matches", StringComparison.OrdinalIgnoreCase)
        || name.Equals("Missing Pages", StringComparison.OrdinalIgnoreCase)
        || name.Equals("Audit", StringComparison.OrdinalIgnoreCase);
}

// ---- missing pages (begin) --------------------------------------------------
// Static and free of Request/Response so it can be compile-tested outside IIS.

class PageError : Exception { public PageError(string m) : base(m) { } }

// The folder each standard is read from, and the only extension its reader takes.
static readonly string[] PageStds = { "SWIFTMX", "SWIFT", "SWIFTDETAIL" };
static string PageExt(string std) { return std == "SWIFTMX" ? ".xml" : ".txt"; }
static string PageLabel(string std)
{
    return std == "SWIFTMX" ? "ISO 20022 (camt.053)" : std == "SWIFT" ? "Legacy statement (MT940/950)" : "Legacy detailed transactions";
}

static string DropRoot(string root)
{
    string d = Path.Combine(root, "Missing Pages");
    // Created once by IT with Modify for this server (see README); the portal
    // only has Read on the rest of the share, so it cannot make it itself.
    if (!Directory.Exists(d)) throw new PageError("The Missing Pages folder has not been set up on the recon share yet - ask IT to run the Missing Pages step in Deploy\\portal\\README.md.");
    return d;
}

static string CheckStd(string std)
{
    std = (std ?? "").Trim().ToUpperInvariant();
    if (Array.IndexOf(PageStds, std) < 0) throw new PageError("Choose which kind of statement the file is.");
    return std;
}

// A file name the robot and Windows will both take, kept recognisable.
// The last path segment, without Path.GetFileName, which throws on characters
// such as < > that a browser may still send.
static string LeafName(string name)
{
    name = (name ?? "").Replace('/', '\\');
    return name.Substring(name.LastIndexOf('\\') + 1);
}

static string CleanName(string name, string std)
{
    name = LeafName(name);
    int dot = name.LastIndexOf('.');
    string ext = dot >= 0 ? name.Substring(dot) : "";
    if (!ext.Equals(PageExt(std), StringComparison.OrdinalIgnoreCase))
        throw new PageError(name + ": " + PageLabel(std) + " pages must be " + PageExt(std) + " files" +
            (ext.Equals(".zip", StringComparison.OrdinalIgnoreCase) ? " - unzip it and upload the files inside." : "."));
    string stem = Regex.Replace(name.Substring(0, dot), @"[^A-Za-z0-9 ._()\-]", "_").Trim(' ', '.');
    if (stem.Length == 0) stem = "page";
    if (stem.Length > 120) stem = stem.Substring(0, 120);
    return stem + PageExt(std);
}

static string Sha(byte[] b)
{
    using (var h = System.Security.Cryptography.SHA256.Create()) return BitConverter.ToString(h.ComputeHash(b)).Replace("-", "");
}
static string ShaFile(string p)
{
    using (var h = System.Security.Cryptography.SHA256.Create())
    using (var s = new FileStream(p, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        return BitConverter.ToString(h.ComputeHash(s)).Replace("-", "");
}

// Catches the wrong file before it costs a night's run: a camt page must be a
// bank-to-customer statement, a legacy page plain text.
static void CheckContent(string std, string name, byte[] b)
{
    if (b.Length == 0) throw new PageError(name + " is empty.");
    if (std == "SWIFTMX")
    {
        bool stmt = false;
        try
        {
            var set = new System.Xml.XmlReaderSettings { DtdProcessing = System.Xml.DtdProcessing.Prohibit, XmlResolver = null };
            using (var rd = System.Xml.XmlReader.Create(new MemoryStream(b), set))
                while (rd.Read())
                    if (rd.NodeType == System.Xml.XmlNodeType.Element && rd.LocalName == "Stmt") stmt = true;
        }
        catch (System.Xml.XmlException ex) { throw new PageError(name + " is not a readable XML file (" + ex.Message + ")."); }
        if (!stmt) throw new PageError(name + " is XML but holds no camt.053 statement (no Stmt element) - check it is the statement page.");
    }
    else if (Array.IndexOf(b, (byte)0) >= 0)
        throw new PageError(name + " is not a text file.");
}

// Every page already waiting or already processed, by content - the same page
// sent twice would be reconciled twice.
static Dictionary<string, string> KnownPages(string drop, long size)
{
    var known = new Dictionary<string, string>();
    foreach (var f in Directory.GetFiles(drop, "*", SearchOption.AllDirectories))
    {
        var fi = new FileInfo(f);
        if (fi.Length != size || fi.Name.StartsWith(".")) continue;
        try { string h = ShaFile(f); if (!known.ContainsKey(h)) known[h] = f.Substring(drop.Length + 1); } catch { }
    }
    return known;
}

static void LogPage(string drop, string user, string action, string std, string file, long bytes, string sha)
{
    string log = Path.Combine(drop, "upload-log.csv");
    string line = string.Join(",", new[] {
        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture), Csv(user), action, std, Csv(file),
        bytes.ToString(CultureInfo.InvariantCulture), sha }) + "\r\n";
    for (int i = 0; ; i++)
    {
        try
        {
            if (!File.Exists(log)) File.WriteAllText(log, "When,User,Action,Standard,File,Bytes,SHA256\r\n", Encoding.UTF8);
            File.AppendAllText(log, line, Encoding.UTF8);
            return;
        }
        catch (IOException) { if (i >= 4) throw; System.Threading.Thread.Sleep(200); }
    }
}

// Saves one uploaded page into Missing Pages\<std>. Returns the name it was saved as.
static string SavePage(string root, string std, string name, byte[] body, string user)
{
    std = CheckStd(std);
    string drop = DropRoot(root);
    string clean = CleanName(name, std);
    CheckContent(std, clean, body);

    string sha = Sha(body);
    var known = KnownPages(drop, body.Length);
    if (known.ContainsKey(sha))
    {
        string where = known[sha];
        throw new PageError(clean + " was already uploaded" +
            (where.StartsWith("Processed\\", StringComparison.OrdinalIgnoreCase) ? " and reconciled (" + where + ")" : " and is waiting for the next run") + ".");
    }

    string dir = Path.Combine(drop, std);
    Directory.CreateDirectory(dir);
    string target = Path.Combine(dir, clean);
    for (int n = 2; File.Exists(target); n++)
        target = Path.Combine(dir, Path.GetFileNameWithoutExtension(clean) + " (" + n + ")" + PageExt(std));

    // Written under a dot name the robot ignores, then renamed, so a run that
    // starts mid-upload never stages half a file.
    string part = Path.Combine(dir, ".upload-" + Guid.NewGuid().ToString("N") + ".partial");
    File.WriteAllBytes(part, body);
    File.Move(part, target);
    LogPage(drop, user, "UPLOAD", std, Path.GetFileName(target), body.Length, sha);
    return Path.GetFileName(target);
}

// Takes back a page that is still waiting; processed pages stay as the record.
static void RemovePage(string root, string std, string name, string user)
{
    std = CheckStd(std);
    string drop = DropRoot(root);
    string file = LeafName(name);
    if (file.Length == 0 || file.StartsWith(".") || file.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
        throw new PageError("bad file name");
    string full = Path.GetFullPath(Path.Combine(drop, std, file));
    if (!string.Equals(Path.GetDirectoryName(full), Path.Combine(drop, std), StringComparison.OrdinalIgnoreCase))
        throw new PageError("bad file name");
    if (!File.Exists(full)) throw new PageError(file + " is no longer waiting - the robot may already have used it.");
    string sha = ShaFile(full);
    long len = new FileInfo(full).Length;
    File.Delete(full);
    LogPage(drop, user, "REMOVE", std, file, len, sha);
}

// Who uploaded each waiting file, from the log (last UPLOAD wins).
static Dictionary<string, string> Uploaders(string drop)
{
    var by = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    string log = Path.Combine(drop, "upload-log.csv");
    if (!File.Exists(log)) return by;
    using (var fs = new FileStream(log, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
    using (var rd = new StreamReader(fs, Encoding.UTF8))
    {
        string line; bool first = true;
        while ((line = rd.ReadLine()) != null)
        {
            if (first) { first = false; continue; }
            var c = SplitCsv(line);
            if (c.Count >= 5 && c[2] == "UPLOAD") by[c[3] + "\\" + c[4]] = c[1] + " at " + c[0];
        }
    }
    return by;
}

static List<string> SplitCsv(string line)
{
    var cells = new List<string>(); var cur = new StringBuilder(); bool q = false;
    for (int i = 0; i < line.Length; i++)
    {
        char ch = line[i];
        if (q) { if (ch == '"' && i + 1 < line.Length && line[i + 1] == '"') { cur.Append('"'); i++; } else if (ch == '"') q = false; else cur.Append(ch); }
        else if (ch == '"') q = true;
        else if (ch == ',') { cells.Add(cur.ToString()); cur.Clear(); }
        else cur.Append(ch);
    }
    cells.Add(cur.ToString());
    return cells;
}

// JSON for the page: what is waiting, and the newest processed batches.
static string PagesJson(string root)
{
    string drop = DropRoot(root);
    var by = Uploaders(drop);
    var sb = new StringBuilder("{\"waiting\":[");
    bool first = true;
    foreach (var std in PageStds)
    {
        string dir = Path.Combine(drop, std);
        if (!Directory.Exists(dir)) continue;
        foreach (var fi in new DirectoryInfo(dir).GetFiles("*", SearchOption.AllDirectories)
                              .Where(f => !f.Name.StartsWith(".")).OrderBy(f => f.Name))
        {
            string rel = fi.FullName.Substring(dir.Length + 1);
            string who; by.TryGetValue(std + "\\" + rel, out who);
            if (!first) sb.Append(',');
            first = false;
            sb.Append("{\"std\":\"").Append(std).Append("\",\"label\":\"").Append(J(PageLabel(std)))
              .Append("\",\"name\":\"").Append(J(rel)).Append("\",\"size\":").Append(fi.Length)
              .Append(",\"when\":").Append(Ms(fi.LastWriteTimeUtc)).Append(",\"by\":\"").Append(J(who ?? "")).Append("\"}");
        }
    }
    sb.Append("],\"processed\":[");
    first = true;
    string proc = Path.Combine(drop, "Processed");
    if (Directory.Exists(proc))
        foreach (var batch in new DirectoryInfo(proc).GetDirectories().OrderByDescending(d => d.Name).Take(10))
        {
            var files = batch.GetFiles("*", SearchOption.AllDirectories).Select(f => f.FullName.Substring(batch.FullName.Length + 1)).OrderBy(f => f);
            if (!first) sb.Append(',');
            first = false;
            sb.Append("{\"batch\":\"").Append(J(batch.Name)).Append("\",\"files\":[")
              .Append(string.Join(",", files.Select(f => "\"" + J(f) + "\"").ToArray())).Append("]}");
        }
    return sb.Append("]}").ToString();
}
// ---- missing pages (end) ----------------------------------------------------

// ---- sessions and audit (begin) ---------------------------------------------
// Static and free of Request/Response too, for the same reason.
//
// Windows sign-in says WHO is at the browser. A portal session says they chose
// to start work, and it ends when they sign out or go idle. Every other
// operation needs a live session belonging to the Windows user of that same
// request, so a session cookie carried to another account is worthless.
// Sessions are held in memory: an app-pool recycle or a new api.aspx signs
// everyone out, which costs them one click.

class AuditError : Exception { public AuditError(string m) : base(m) { } }

class PortalSession { public string User; public DateTime Started, LastSeen; }
static readonly Dictionary<string, PortalSession> Sessions = new Dictionary<string, PortalSession>(StringComparer.Ordinal);
const string SessionCookie = "nostro_sid";

// Minutes without activity before a session ends: IdleMinutes in web.config, else 15.
static int IdleLimit()
{
    int m;
    return int.TryParse(System.Configuration.ConfigurationManager.AppSettings["IdleMinutes"], out m) && m >= 5 && m <= 240 ? m : 15;
}

static string SessionStart(string user)
{
    var b = new byte[32];
    using (var rng = new System.Security.Cryptography.RNGCryptoServiceProvider()) rng.GetBytes(b);
    string sid = BitConverter.ToString(b).Replace("-", "");
    lock (Sessions) Sessions[sid] = new PortalSession { User = user, Started = DateTime.Now, LastSeen = DateTime.Now };
    return sid;
}

// True, and the session counts as used now, when sid is a live session of user.
static bool SessionLive(string sid, string user)
{
    lock (Sessions)
    {
        PortalSession s;
        if (string.IsNullOrEmpty(sid) || !Sessions.TryGetValue(sid, out s) || !s.User.Equals(user, StringComparison.OrdinalIgnoreCase)) return false;
        s.LastSeen = DateTime.Now;
        return true;
    }
}

static bool SessionEnd(string sid)
{
    lock (Sessions) return !string.IsNullOrEmpty(sid) && Sessions.Remove(sid);
}

// Removes and returns the sessions idle past the limit. The server allows a
// minute more than the page, which signs itself out first and records why;
// what is left here is mostly browsers closed without signing out.
static List<PortalSession> SessionsExpire(int idleMinutes)
{
    var cut = DateTime.Now.AddMinutes(-(idleMinutes + 1));
    lock (Sessions)
    {
        var gone = Sessions.Where(kv => kv.Value.LastSeen < cut).ToList();
        foreach (var kv in gone) Sessions.Remove(kv.Key);
        return gone.Select(kv => kv.Value).ToList();
    }
}

static readonly object AuditLock = new object();
const string AuditHead = "When,User,Client,Action,Detail,Result\r\n";

// One file per month, in V001's local time. Lines are only ever appended.
static string AuditPath(string root)
{
    string d = Path.Combine(root, "Audit");
    // Created once by IT with Modify for this server, like Missing Pages.
    if (!Directory.Exists(d)) throw new AuditError("The Audit folder has not been set up on the recon share yet - ask IT to run the Audit step in Deploy\\portal\\README.md. The portal does nothing it cannot record.");
    return Path.Combine(d, "portal-audit-" + DateTime.Now.ToString("yyyy-MM", CultureInfo.InvariantCulture) + ".csv");
}

static AuditError AuditLocked(string path)
{
    return new AuditError("The audit log " + path + " is locked by another program - is it open in Excel? Close it (open a copy instead) and try again.");
}

// Proves the log can be written BEFORE an action runs, so nothing is done
// that could not be recorded.
static void AuditReady(string root)
{
    string path = AuditPath(root);
    lock (AuditLock)
    {
        try
        {
            if (!File.Exists(path)) File.WriteAllText(path, AuditHead, new UTF8Encoding(true));
            else using (new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.Read)) { }
        }
        catch (IOException) { throw AuditLocked(path); }
    }
}

static void AuditAppend(string root, string user, string client, string action, string detail, string result)
{
    string path = AuditPath(root);
    string line = string.Join(",", new[] {
        DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture), Csv(user), Csv(client),
        action, Csv(detail), Csv(result) }) + "\r\n";
    lock (AuditLock)
        for (int i = 0; ; i++)
        {
            try
            {
                bool fresh = !File.Exists(path);
                using (var s = new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.Read))
                using (var w = new StreamWriter(s, new UTF8Encoding(fresh)))
                {
                    if (fresh) w.Write(AuditHead);
                    w.Write(line);
                }
                return;
            }
            catch (IOException) { if (i >= 4) throw AuditLocked(path); System.Threading.Thread.Sleep(200); }
        }
}

// What the page may record on its own: things done without calling the server.
static readonly string[] PageEvents = {
    "SELECT_ACCOUNT", "VIEW_MATCHING", "VIEW_HISTORY", "VIEW_PAGES", "SUGGEST",
    "ADD_TICKET", "REMOVE_TICKET", "DOWNLOAD_JOURNAL"
};

static string Detail(string s)
{
    s = Regex.Replace(s ?? "", @"[\x00-\x1F]", " ").Trim();
    return s.Length > 500 ? s.Substring(0, 500) + "..." : s;
}
// ---- sessions and audit (end) -----------------------------------------------

// The action this request is recording, set before it runs. Fail() records
// it as FAILED; a request that completes records it as OK.
string auditUser = "", act = "", actDetail = "";

string Sid()
{
    var c = Request.Cookies[SessionCookie];
    return c == null ? "" : (c.Value ?? "");
}

// HttpOnly, and SameSite=Strict so no other site can make the browser act here.
void SetSid(string sid)
{
    string path = string.IsNullOrEmpty(Request.ApplicationPath) ? "/" : Request.ApplicationPath;
    Response.AppendHeader("Set-Cookie", SessionCookie + "=" + sid + "; Path=" + path + "; HttpOnly; SameSite=Strict" +
        (Request.IsSecureConnection ? "; Secure" : "") + (sid.Length == 0 ? "; Max-Age=0" : ""));
}

void Audited(string result)
{
    if (act.Length == 0) return;
    string a = act; act = "";
    AuditAppend(Root, auditUser, Request.UserHostAddress ?? "", a, actDetail, result);
}

// 440 (IIS's "login time-out") rather than 401: Windows authentication turns
// every 401 into a fresh sign-in challenge.
void FailSignedOut()
{
    Response.Clear();
    Response.StatusCode = 440;
    Response.TrySkipIisCustomErrors = true;
    Response.ContentType = "application/json";
    Response.Write("{\"error\":\"Your portal session has ended - sign in again.\",\"signedOut\":true}");
    Response.End();
}

void Fail(int code, string msg)
{
    if (act.Length > 0)
    {
        try { Audited("FAILED " + code + ": " + msg); } catch { }
        act = "";
    }
    Response.Clear();
    Response.StatusCode = code;
    // IIS replaces the body of an error response for any client but the server
    // itself, leaving the user a bare "500 - Internal server error". The message
    // here is the only clue to what is wrong, so keep it.
    Response.TrySkipIisCustomErrors = true;
    Response.ContentType = "application/json";
    Response.Write("{\"error\":\"" + J(msg) + "\"}");
    Response.End();
}
</script>
<%
    Response.Cache.SetCacheability(HttpCacheability.NoCache);
    Response.Cache.SetNoStore();

    string user = (User != null && User.Identity != null && User.Identity.IsAuthenticated)
                  ? (User.Identity.Name ?? "") : "";
    if (user.Length == 0) { Fail(401, "not authenticated - enable Windows Authentication on /portal"); return; }

    string op = (Request["op"] ?? "").ToLowerInvariant();
    string client = Request.UserHostAddress ?? "";
    bool post = string.Equals(Request.HttpMethod, "POST", StringComparison.OrdinalIgnoreCase);
    auditUser = user;

    try
    {
        // Sessions that ran out are recorded against their own user. Best
        // effort: a missing Audit folder is reported by the request's own check.
        int idle = IdleLimit();
        foreach (var s in SessionsExpire(idle))
            try { AuditAppend(Root, s.User, "", "SESSION_EXPIRED", "signed in " + s.Started.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) +
                              ", no activity since " + s.LastSeen.ToString("HH:mm:ss", CultureInfo.InvariantCulture), "ended by the server"); } catch { }

        if (op != "whoami" && op != "login" && op != "logout")
        {
            if (!SessionLive(Sid(), user))
            {
                try { AuditAppend(Root, user, client, "DENIED", "op=" + Detail(op), "no live portal session"); } catch { }
                FailSignedOut();
            }
            AuditReady(Root);
        }

        if (op == "whoami")
        {
            Response.ContentType = "application/json";
            Response.Write("{\"user\":\"" + J(user) + "\",\"session\":" + (SessionLive(Sid(), user) ? "true" : "false") +
                           ",\"idleMinutes\":" + idle + "}");
        }
        else if (op == "login")
        {
            if (!post) Fail(405, "login must be POSTed");
            AuditReady(Root);
            SessionEnd(Sid());
            act = "SIGN_IN"; actDetail = "idle limit " + idle + " min";
            string sid = SessionStart(user);
            try { Audited("OK"); } catch { SessionEnd(sid); throw; }   // no session that was not recorded
            SetSid(sid);
            Response.ContentType = "application/json";
            Response.Write("{\"user\":\"" + J(user) + "\",\"idleMinutes\":" + idle + "}");
        }
        else if (op == "logout")
        {
            if (!post) Fail(405, "logout must be POSTed");
            bool idleOut = (Request["reason"] ?? "") == "idle";
            bool had = SessionEnd(Sid());
            SetSid("");
            if (had)
            {
                act = "SIGN_OUT"; actDetail = idleOut ? "no activity for " + idle + " min" : "signed out";
                Audited(idleOut ? "signed out for inactivity" : "OK");
            }
            Response.ContentType = "application/json";
            Response.Write("{\"signedOut\":true}");
        }
        else if (op == "ping")
        {
            // SessionLive above has already counted this as activity
            Response.ContentType = "application/json";
            Response.Write("{\"ok\":true}");
        }
        else if (op == "event")
        {
            if (!post) Fail(405, "event must be POSTed");
            string ev = (Request["action"] ?? "").ToUpperInvariant();
            if (Array.IndexOf(PageEvents, ev) < 0) Fail(400, "unknown event");
            string body;
            using (var rd = new StreamReader(Request.InputStream, Encoding.UTF8)) body = rd.ReadToEnd();
            act = ev; actDetail = Detail(body);
            Audited("OK");
            Response.ContentType = "application/json";
            Response.Write("{\"ok\":true}");
        }
        else if (op == "runs")
        {
            act = "LIST_RUNS"; actDetail = "";
            Response.ContentType = "application/json";
            var sb = new StringBuilder("[");
            bool first = true;
            var all = Runs();
            actDetail = all.Count + " account run(s)";
            foreach (var r in all)
            {
                if (!first) sb.Append(',');
                first = false;
                sb.Append("{\"id\":\"").Append(J(r.Id))
                  .Append("\",\"date\":\"").Append(J(r.Date))
                  .Append("\",\"bank\":\"").Append(J(r.Bank))
                  .Append("\",\"account\":\"").Append(J(r.Acct))
                  .Append("\",\"ccy\":\"").Append(J(r.Ccy))
                  .Append("\",\"when\":").Append(r.When.ToString(CultureInfo.InvariantCulture))
                  .Append(",\"data\":").Append(r.Data != null ? "true" : "false")
                  .Append(",\"summary\":").Append(r.Sum != null ? "true" : "false")
                  .Append(",\"workings\":").Append(r.Wrk != null ? "true" : "false")
                  .Append(",\"ord\":").Append(r.Ord.ToString(CultureInfo.InvariantCulture))
                  .Append("}");
            }
            Response.Write(sb.Append(']').ToString());
        }
        else if (op == "data")
        {
            act = "OPEN_RUN"; actDetail = Detail(Request["run"] + " account " + Request["account"]);
            var r = RunFor(Request["run"], Request["account"]);
            if (r.Data == null) Fail(404, "this account has no portal data in this run");
            Response.ContentType = "application/json";
            Response.TransmitFile(r.Data.FullName);
        }
        else if (op == "file")
        {
            string kind = (Request["kind"] ?? "").ToLowerInvariant();
            act = "DOWNLOAD"; actDetail = Detail(kind + " of " + Request["run"] + " account " + Request["account"]);
            var r = RunFor(Request["run"], Request["account"]);
            FileInfo f = kind == "summary" ? r.Sum
                       : kind == "workings" ? r.Wrk
                       : null;
            if (f == null) Fail(404, "no such file for this run");
            actDetail = Detail(kind + " " + f.Name);
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("Content-Disposition", "attachment; filename=\"" + f.Name.Replace("\"", "") + "\"");
            Response.TransmitFile(f.FullName);
        }
        else if (op == "tickets")
        {
            if (!string.Equals(Request.HttpMethod, "POST", StringComparison.OrdinalIgnoreCase))
                Fail(405, "tickets must be POSTed");

            string acct = Request["account"] ?? "";
            act = "SAVE_TICKETS"; actDetail = Detail("account " + acct);
            if (!SafeId.IsMatch(acct)) Fail(400, "bad account");

            string body;
            using (var rd = new StreamReader(Request.InputStream, Encoding.UTF8)) body = rd.ReadToEnd();

            string stamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            var lines = new List<string>();
            foreach (var raw in (body ?? "").Split('\n'))
            {
                string line = raw.Trim();
                if (line.Length == 0) continue;
                var p = line.Split('|');
                if (p.Length < 3) Fail(400, "each leg must be MatchRef|Side|ItemId|Reason");

                string mref = p[0].Trim(), side = p[1].Trim().ToUpperInvariant(), item = p[2].Trim();
                string reason = p.Length > 3 ? string.Join("|", p.Skip(3).ToArray()).Trim() : "";

                if (!SafeRef.IsMatch(mref)) Fail(400, "bad Match Ref: " + mref);
                if (side != "CBS" && side != "STATEMENT") Fail(400, "bad Side: " + side);
                if (!SafeId.IsMatch(item)) Fail(400, "bad Item Id: " + item);
                if (reason.Length > 300) reason = reason.Substring(0, 300);

                // the recorded user is the authenticated one, never anything the client sent
                lines.Add(string.Join(",", new[]{
                    Csv(mref), side, Csv(item), "APPROVED",
                    Csv(user), stamp, Csv(user), stamp, Csv(reason)
                }));
            }
            if (lines.Count == 0) Fail(400, "no legs supplied");
            actDetail = Detail("account " + acct + ": " + lines.Count + " leg(s), tickets " +
                string.Join(" ", lines.Select(l => l.Split(',')[0]).Distinct().ToArray()));

            string dir = Path.Combine(Root, "Manual Matches");
            Directory.CreateDirectory(dir);
            string path = Path.Combine(dir, acct + ".csv");
            if (!Inside(Path.GetFullPath(path))) Fail(400, "bad account");

            if (!File.Exists(path))
                File.WriteAllText(path,
                    "Match Ref,Side,Item Id,Status,Proposed By,Proposed On,Approved By,Approved On,Reason\r\n",
                    Encoding.UTF8);

            File.AppendAllText(path, string.Join("\r\n", lines.ToArray()) + "\r\n", Encoding.UTF8);

            Response.ContentType = "application/json";
            Response.Write("{\"saved\":" + lines.Count + ",\"journal\":\"" + J(path) + "\",\"user\":\"" + J(user) + "\"}");
        }
        else if (op == "pages")
        {
            act = "LIST_PAGES"; actDetail = "";
            Response.ContentType = "application/json";
            Response.Write(PagesJson(Root));
        }
        else if (op == "upload" || op == "unupload")
        {
            act = op == "upload" ? "UPLOAD_PAGE" : "TAKE_BACK_PAGE";
            actDetail = Detail(Request["std"] + " " + Request["name"]);
            if (!post) Fail(405, op + " must be POSTed");
            Response.ContentType = "application/json";
            if (op == "unupload")
            {
                RemovePage(Root, Request["std"], Request["name"], user);
                Response.Write("{\"removed\":\"" + J(Request["name"]) + "\"}");
            }
            else
            {
                // one file per request, so ASP.NET's default 4 MB request limit is per page
                byte[] body;
                using (var ms = new MemoryStream()) { Request.InputStream.CopyTo(ms); body = ms.ToArray(); }
                string saved = SavePage(Root, Request["std"], Request["name"], body, user);
                actDetail = Detail(Request["std"] + " " + Request["name"] + " saved as " + saved + " (" + body.Length + " bytes)");
                Response.Write("{\"saved\":\"" + J(saved) + "\",\"user\":\"" + J(user) + "\"}");
            }
        }
        else
        {
            Fail(400, "unknown op");
        }
        Audited("OK");
    }
    catch (System.Threading.ThreadAbortException) { }   // Response.End
    catch (PageError ex) { Fail(400, ex.Message); }
    catch (AuditError ex) { Fail(503, ex.Message); }
    catch (UnauthorizedAccessException ex)
    {
        // Any op can hit this - reading the run folders as much as writing a
        // journal, a page or the audit log - so say which path and which account.
        string who = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
        Fail(500, "The portal server (" + who + ") was refused access: " + ex.Message
                + " Ask IT to grant it Read on the recon share, and Modify on Manual Matches, Missing Pages and Audit (Deploy\\portal\\README.md).");
    }
    catch (Exception ex) { Fail(500, ex.Message); }
%>
