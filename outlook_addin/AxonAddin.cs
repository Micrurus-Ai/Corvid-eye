// Axon Outlook add-in — adds two ribbon buttons ("File with Axon", "Download with Axon").
// Implemented as a managed COM add-in against the REAL Office interop interfaces
// (Extensibility.IDTExtensibility2, Microsoft.Office.Core.IRibbonExtensibility). Those types are
// EMBEDDED at build time (csc /link), so the compiled DLL is self-contained and needs no PIAs on
// the user's machine. Outlook objects are used late-bound (dynamic), so there's no hard dependency
// on the Outlook object model either.
//
// Phase 1: the buttons appear and report the selected email's subject (proves load + email access).

using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;

namespace Axon.OutlookAddin
{
    // Converts a managed Image into the COM IPictureDisp the ribbon's getImage callback needs.
    internal class RibbonImage : AxHost
    {
        private RibbonImage() : base("00000000-0000-0000-0000-000000000000") { }
        public static stdole.IPictureDisp Get(System.Drawing.Image img)
        {
            return (stdole.IPictureDisp)AxHost.GetIPictureDispFromPicture(img);
        }
    }

    [ComVisible(true)]
    [Guid("7B2C9E14-6A3D-4F58-9C21-3E5A1B7D4F60")]
    [ProgId("Axon.OutlookAddin")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]   // exposes the ribbon callbacks via IDispatch
    public partial class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private object _app;   // Outlook.Application (late-bound)

        // --- IDTExtensibility2 ---
        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom) { _app = Application; }
        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom) { _app = null; }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { StartReminderService(); }
        public void OnBeginShutdown(ref Array custom) { try { if (_reminderTimer != null) _reminderTimer.Stop(); } catch { } }

        // --- IRibbonExtensibility ---
        public string GetCustomUI(string RibbonID)
        {
            // Axon is in the right-click menu AND, since right-click is an extra click on the actions
            // people use all day, as a real button group on the mail Home tab and on an open message.
            // A ribbon button is also what makes the Quick Access Toolbar possible: right-click the
            // button -> "Add to Quick Access Toolbar" and Outlook gives it Alt+1..9 for free.
            if (RibbonID == "Microsoft.Outlook.Explorer")
                return CtxUI(MailRibbon("TabMail"),
                             CtxMenu("ContextMenuMailItem") + CtxMenu("ContextMenuReadOnlyMailText"));
            if (RibbonID == "Microsoft.Outlook.Mail.Read")
                return CtxUI(MailRibbon("TabReadMessage"), CtxMenu("ContextMenuReadOnlyMailText"));
            if (RibbonID == "Microsoft.Outlook.Mail.Compose")
                return ComposeRibbon();   // compose body right-click is Word's, so use a ribbon button
            return null;
        }

        // customUI's schema fixes the child order: <ribbon> before <contextMenus>. Getting it wrong makes
        // Office reject the WHOLE document, which would silently take the right-click menu with it.
        private string CtxUI(string ribbon, string menus)
        {
            return "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>" +
                   (string.IsNullOrEmpty(ribbon) ? "" : ribbon) +
                   "<contextMenus>" + menus + "</contextMenus></customUI>";
        }

        // An "Axon" group on the given built-in tab. Every button is the same size — one small icon with
        // its label beside it — so the group reads as one set rather than two. Office fills columns three
        // at a time, and the separator splits it deliberately: the two everyday actions in the first
        // column, the three that call the model in the second. Settings sits in the dialog box launcher,
        // the corner arrow that is where Office puts a group's options.
        // No insertAfterMso — pinning the group beside a built-in one means a single idMso that is absent
        // in some Outlook build drops the whole customUI document, taking the right-click menu with it.
        // Button ids must be unique per tab, hence the suffix.
        private string MailRibbon(string tabMso)
        {
            string s = tabMso;
            return "<ribbon><tabs><tab idMso='" + s + "'>" +
                   "<group id='axonGroup_" + s + "' label='Axon'>" +
                   Btn("axonMove_r_" + s, "Move", "XM", "GetMoveImage", "OnFile", false,
                       "Move with Axon", "File this email in an Outlook folder. Axon suggests the folders you already file this sender in.") +
                   Btn("axonDownload_r_" + s, "Download", "XD", "GetDownloadImage", "OnDownload", false,
                       "Download with Axon", "Save this email to the Sales archive on disk, in the order folder it belongs to.") +
                   "<separator id='axonSep_r_" + s + "'/>" +
                   Btn("axonSummarize_r_" + s, "Summarize", "XS", "GetSummarizeImage", "OnSummarize", false,
                       "Summarize with Axon", "Read the whole thread and write a short summary of it.") +
                   Btn("axonReply_r_" + s, "Reply", "XR", "GetReplyImage", "OnReply", false,
                       "Reply with Axon", "Draft a reply in your own tone, ready for you to check and send.") +
                   Btn("axonSchedule_r_" + s, "Schedule", "XC", "GetScheduleImage", "OnSchedule", false,
                       "Schedule with Axon", "Turn this email into a meeting or an appointment.") +
                   "<dialogBoxLauncher>" +
                   "<button id='axonSettings_r_" + s + "' keytip='XG' screentip='Axon settings' " +
                   "supertip='Archive folders, what Download saves, and the country codes.' onAction='OnSettings'/>" +
                   "</dialogBoxLauncher>" +
                   "</group></tab></tabs></ribbon>";
        }

        // One ribbon button. Every button carries a screentip and supertip: the hover card with a bold
        // title and a sentence under it is most of what makes a group read as finished rather than homemade.
        private static string Btn(string id, string label, string keytip, string image, string action,
                                  bool large, string tip, string superTip)
        {
            return "<button id='" + id + "' label='" + label + "' keytip='" + keytip + "'" +
                   (large ? " size='large'" : "") +
                   " screentip='" + Xml(tip) + "' supertip='" + Xml(superTip) + "'" +
                   " getImage='" + image + "' onAction='" + action + "'/>";
        }

        // Attribute values are single-quoted above, so anything quotable has to be escaped or the whole
        // customUI document fails to parse and Office silently shows none of it.
        private static string Xml(string s)
        {
            return (s ?? "").Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
                            .Replace("'", "&apos;").Replace("\"", "&quot;");
        }

        // A small "Send Later" button in an Axon group on the compose Message tab (reliable — the
        // compose body's right-click menu belongs to the Word editor and can't be extended).
        private string ComposeRibbon()
        {
            return "<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>" +
                   "<ribbon><tabs><tab idMso='TabNewMailMessage'>" +
                   "<group id='axonComposeGroup' label='Axon'>" +
                   "<button id='axonWriteBtn' label='Write with Axon' size='large' " +
                   "getImage='GetWriteImage' onAction='OnWriteEmail'/>" +
                   "<button id='axonSendLaterBtn' label='Send Later' size='large' " +
                   "getImage='GetSendLaterImage' onAction='OnSendLater'/>" +
                   "</group></tab></tabs></ribbon></customUI>";
        }

        // Axon's right-click items for a given Office context-menu id (button ids must be unique).
        // Summarize/Reply have no icon (avoids any invalid-image risk); Move/Download keep theirs.
        private string CtxMenu(string idMso)
        {
            return "<contextMenu idMso='" + idMso + "'>" +
                   "<menuSeparator id='axonSep_" + idMso + "'/>" +
                   "<button id='axonSummarize_" + idMso + "' label='Summarize with Disassist' getImage='GetSummarizeImage' onAction='OnSummarize'/>" +
                   "<button id='axonReply_" + idMso + "' label='Reply with Disassist' getImage='GetReplyImage' onAction='OnReply'/>" +
                   "<button id='axonSchedule_" + idMso + "' label='Schedule with Disassist' getImage='GetScheduleImage' onAction='OnSchedule'/>" +
                   "<button id='axonFollowUp_" + idMso + "' label='Follow up with Disassist' getImage='GetFollowUpImage' onAction='OnFollowUp'/>" +
                   "<button id='axonAttach_" + idMso + "' label='Forward as attachment' getImage='GetAttachImage' onAction='OnAttachEmail'/>" +
                   "<button id='axonMove_" + idMso + "' label='Move with Disassist' getImage='GetMoveImage' onAction='OnFile'/>" +
                   "<button id='axonDownload_" + idMso + "' label='Download with Disassist' getImage='GetDownloadImage' onAction='OnDownload'/>" +
                   "<menuSeparator id='axonSep2_" + idMso + "'/>" +
                   "<button id='axonSettings_" + idMso + "' label='Disassist Settings' getImage='GetSettingsImage' onAction='OnSettings'/>" +
                   "</contextMenu>";
        }


        // --- custom ribbon image (the Axon-branded Move icon, distinct from Outlook's built-ins) ---
        private System.Drawing.Image _moveIcon, _downloadIcon, _summarizeIcon, _replyIcon, _scheduleIcon,
                                     _followUpIcon, _sendLaterIcon, _writeIcon, _attachIcon, _settingsIcon;

        private System.Drawing.Image LoadIcon(string file, ref System.Drawing.Image cache)
        {
            if (cache == null)
            {
                string dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                byte[] bytes = File.ReadAllBytes(Path.Combine(dir, file));   // ReadAllBytes -> no file lock
                cache = System.Drawing.Image.FromStream(new System.IO.MemoryStream(bytes));
            }
            return cache;
        }

        public stdole.IPictureDisp GetMoveImage(object control)
        {
            try { return RibbonImage.Get(LoadIcon("axon-move.png", ref _moveIcon)); } catch { return null; }
        }

        public stdole.IPictureDisp GetDownloadImage(object control)
        {
            try { return RibbonImage.Get(LoadIcon("axon-download.png", ref _downloadIcon)); } catch { return null; }
        }

        public stdole.IPictureDisp GetSummarizeImage(object control)
        {
            try { return RibbonImage.Get(LoadIcon("axon-summarize.png", ref _summarizeIcon)); } catch { return null; }
        }

        public stdole.IPictureDisp GetReplyImage(object control)
        {
            try { return RibbonImage.Get(LoadIcon("axon-reply.png", ref _replyIcon)); } catch { return null; }
        }

        public stdole.IPictureDisp GetScheduleImage(object control)
        {
            try { return RibbonImage.Get(LoadIcon("axon-schedule.png", ref _scheduleIcon)); } catch { return null; }
        }

        public stdole.IPictureDisp GetFollowUpImage(object control)
        {
            try { return RibbonImage.Get(LoadIcon("axon-followup.png", ref _followUpIcon)); } catch { return null; }
        }

        public stdole.IPictureDisp GetSendLaterImage(object control)
        {
            try { return RibbonImage.Get(LoadIcon("axon-sendlater.png", ref _sendLaterIcon)); } catch { return null; }
        }

        public stdole.IPictureDisp GetWriteImage(object control)
        {
            try { return RibbonImage.Get(LoadIcon("axon-write.png", ref _writeIcon)); } catch { return null; }
        }

        public stdole.IPictureDisp GetAttachImage(object control)
        {
            try { return RibbonImage.Get(LoadIcon("axon-attach.png", ref _attachIcon)); } catch { return null; }
        }

        public stdole.IPictureDisp GetSettingsImage(object control)
        {
            try { return RibbonImage.Get(LoadIcon("axon-settings.png", ref _settingsIcon)); } catch { return null; }
        }

        // --- ribbon button callbacks (Office invokes these by name via IDispatch) ---
        public void OnFile(object control)
        {
            try
            {
                object m = GetSelectedMail();
                if (m == null) { Ui.Notify("Select an email first.", "Axon intelligence"); return; }
                dynamic mail = m;
                // Walking every store to list the folders costs 1-5 seconds and used to happen HERE, before
                // the dialog could appear, on every single click. The list is now read from disk instead and
                // refreshed behind the open dialog, so the wait is paid once rather than every time.
                string[] folders = LoadFolderCache();
                if (folders.Length == 0)
                {
                    folders = EnumerateInboxFolders();   // first ever use: nothing cached yet
                    SaveFolderCache();
                }
                if (folders.Length == 0)
                {
                    Ui.Notify("No folders found to move emails into. Create a folder in Outlook first, then try again.", "Axon intelligence");
                    return;
                }
                string subject = ""; try { subject = (string)mail.Subject; } catch { }
                string senderName = ""; try { senderName = (string)mail.SenderName; } catch { }
                string sender = senderName;
                try { string em = (string)mail.SenderEmailAddress; if (!string.IsNullOrEmpty(em)) sender = (senderName + " <" + em + ">").Trim(); } catch { }
                string body = ""; try { body = (string)mail.Body; } catch { }
                // Show the picker IMMEDIATELY; fetch AI suggestions on a background thread and fill them in.
                var dlg = new FolderPicker(subject, folders);
                var worker = new System.Threading.Thread(() =>
                {
                    Filing fil = SuggestFiling(subject, sender, senderName, body, folders);
                    dlg.SetSuggestions(fil.matches, fil.newFolder);
                });
                worker.IsBackground = true;
                worker.Start();
                // Separately, re-walk the stores behind the dialog and refresh both the cache and the list
                // the user is looking at. A folder added in Outlook shows up as soon as this lands, so the
                // cache never needs a TTL and nobody has to remember to press anything. It runs apart from
                // the suggestions thread so a slow walk cannot hold the suggestions back.
                var refresh = new System.Threading.Thread(() =>
                {
                    try
                    {
                        var fresh = EnumerateInboxFolders();
                        SaveFolderCache();
                        if (fresh.Length > 0) dlg.SetFolders(fresh);
                    }
                    catch { }
                });
                refresh.IsBackground = true;
                refresh.Start();
                try
                {
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        if (!string.IsNullOrEmpty(dlg.CreateFolder))
                        {
                            dynamic dest = CreateInboxSubfolder(dlg.CreateFolder.Trim());
                            if (dest != null) mail.Move(dest);
                        }
                        else if (!string.IsNullOrEmpty(dlg.Chosen))
                        {
                            MoveTo(mail, dlg.Chosen);
                        }
                    }
                }
                finally { dlg.Dispose(); }
            }
            catch (Exception ex) { Ui.Notify("Axon error: " + ex.Message, "Axon intelligence"); }
        }

        // Maps each folder's display path (e.g. "Clients / Acme / Invoices") to its EntryID, so we can
        // move into deeply-nested folders unambiguously (folder names may even contain "/").
        private System.Collections.Generic.Dictionary<string, string> _folderMap =
            new System.Collections.Generic.Dictionary<string, string>();

        // ALL folders under the Inbox, recursively, as display paths (in-process COM, fast).
        // Every folder the user could file mail into. This used to list ONLY the subfolders of the default
        // Inbox, so a mailbox that keeps its filing folders BESIDE the Inbox at the store root — Outlook's
        // other perfectly normal layout — was told "You have no Inbox subfolders yet" while being full of
        // folders. Order: the Inbox subtree first (display names unchanged, so filing history learned from
        // SenderFolders still matches), then the rest of that mailbox, then any other open store (a shared
        // mailbox, an online archive, a .pst) prefixed with its name.
        private const int MaxFolders = 500;   // a shared mailbox can hold thousands; keep the picker usable

        // The subset of _folderMap that lives in the user's OWN mailbox. Ranking a folder means querying it
        // (how many mails from this sender does it hold, what is in it lately), and that is a round trip
        // per folder. In the main mailbox those hit the local cached copy; in a shared mailbox or an online
        // archive they go to the server and cost hundreds of ms EACH, which is what makes the suggestions
        // take "quite some time". Those folders stay in the picker — they are simply not probed.
        private System.Collections.Generic.HashSet<string> _primaryFolders =
            new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private string[] EnumerateInboxFolders()
        {
            // Built into fresh dictionaries and only published at the end, so the refresh running behind an
            // open dialog can never be seen half-written by the suggestion thread reading these.
            _folderMap = new System.Collections.Generic.Dictionary<string, string>();
            _folderPaths = new System.Collections.Generic.Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            _primaryFolders = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var list = new System.Collections.Generic.List<string>();
            try
            {
                dynamic ns = ((dynamic)_app).GetNamespace("MAPI");

                // Outlook's own folders, identified by id rather than by name so the check still holds on a
                // Dutch or French Outlook (Postvak IN, Verwijderde items, ...). Deleted/Outbox/Sent/Drafts/
                // Conflicts/Sync/Local+Server failures/Junk/RSS/To-Do are dead ends — neither offered nor
                // descended into. The Inbox is a special case: it is not itself a filing target, but its
                // subfolders are the usual ones, so we descend into it WITHOUT adding a level to the
                // display path (which is what keeps names like "hr" identical to before).
                // Each store is asked for its OWN defaults; ns.GetDefaultFolder only knows the main
                // mailbox, which is why a second store used to offer its Deleted Items and Sync Issues.
                // Matched on FolderPath, NOT EntryID: Outlook hands out short-term entry ids, so the id
                // for a folder fetched via GetDefaultFolder does not necessarily equal the id for the same
                // folder reached by walking Folders a moment later, and the skip silently missed. The path
                // still comes from GetDefaultFolder rather than a hardcoded name, so it stays correct on a
                // non-English Outlook.
                var skipTree = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var skipSelf = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
                Action<dynamic> addDefaults = store =>
                {
                    foreach (int id in new[] { 3, 4, 5, 16, 19, 20, 21, 22, 23, 25, 28 })
                        try { skipTree.Add((string)store.GetDefaultFolder(id).FolderPath); } catch { }
                    try { skipSelf.Add((string)store.GetDefaultFolder(6).FolderPath); } catch { }
                };

                string mainRoot = null;
                try { mainRoot = (string)ns.GetDefaultFolder(6).Parent.EntryID; } catch { }
                foreach (int id in new[] { 3, 4, 5, 16, 19, 20, 21, 22, 23, 25, 28 })
                    try { skipTree.Add((string)ns.GetDefaultFolder(id).FolderPath); } catch { }
                try { skipSelf.Add((string)ns.GetDefaultFolder(6).FolderPath); } catch { }

                var roots = new System.Collections.Generic.List<System.Collections.Generic.KeyValuePair<string, object>>();
                try
                {
                    foreach (dynamic st in ns.Stores)
                    {
                        dynamic root; string rid;
                        try { root = st.GetRootFolder(); rid = (string)root.EntryID; } catch { continue; }
                        addDefaults(st);
                        string nm = ""; try { nm = (string)st.DisplayName; } catch { }
                        bool isMain = mainRoot != null && string.Equals(rid, mainRoot, StringComparison.OrdinalIgnoreCase);
                        // The main mailbox contributes unprefixed names and goes first; other stores are
                        // prefixed so "Inbox / Archive" and a shared mailbox's Archive stay tellable apart.
                        var entry = new System.Collections.Generic.KeyValuePair<string, object>(isMain ? "" : (nm ?? ""), root);
                        if (isMain) roots.Insert(0, entry); else roots.Add(entry);
                    }
                }
                catch { }
                if (roots.Count == 0)   // no store enumerable — fall back to just the Inbox subtree
                    try { roots.Add(new System.Collections.Generic.KeyValuePair<string, object>("", ns.GetDefaultFolder(6).Parent)); } catch { }

                foreach (var r in roots)
                {
                    int before = list.Count;
                    CollectFolders(r.Value, r.Key, list, 0, skipTree, skipSelf);
                    // r.Key == "" is the main mailbox (it is the only root added unprefixed).
                    if (r.Key.Length == 0)
                        for (int i = before; i < list.Count; i++) _primaryFolders.Add(list[i]);
                }
            }
            catch { }
            return list.ToArray();
        }

        // A folder mail can actually be moved into: it holds mail items, and it isn't one of Outlook's
        // hidden bookkeeping folders (Working Set, Quick Step Settings, Yammer Root, Conversation Action
        // Settings...) which are otherwise indistinguishable from ordinary folders. Both tests read MAPI
        // properties rather than names, so they don't depend on Outlook's display language.
        private static bool IsMailFolder(dynamic f)
        {
            try { if ((int)f.DefaultItemType != 0) return false; } catch { return false; }   // 0 = olMailItem
            try { if ((bool)f.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x10F4000B")) return false; }
            catch { }   // PR_ATTR_HIDDEN missing = not hidden
            try
            {
                // PR_CONTAINER_CLASS: mail folders are "IPF.Note". Anything else set (IPF.Files,
                // IPF.Configuration, IPF.Note.OutlookHomepage for RSS) is not a filing target. Unset is
                // allowed through — some genuine folders carry no class.
                string k = (string)f.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3613001F");
                if (!string.IsNullOrEmpty(k) && !k.Equals("IPF.Note", StringComparison.OrdinalIgnoreCase)) return false;
            }
            catch { }
            return true;
        }

        private void CollectFolders(dynamic parent, string prefix, System.Collections.Generic.List<string> list,
            int depth, System.Collections.Generic.HashSet<string> skipTree,
            System.Collections.Generic.HashSet<string> skipSelf)
        {
            if (depth > 8 || list.Count >= MaxFolders) return;
            dynamic subs;
            try { subs = parent.Folders; } catch { return; }
            foreach (dynamic f in subs)
            {
                if (list.Count >= MaxFolders) return;
                string fpath = null; try { fpath = (string)f.FolderPath; } catch { }
                if (fpath != null && skipTree != null && skipTree.Contains(fpath)) continue;   // dead end
                if (fpath != null && skipSelf != null && skipSelf.Contains(fpath))
                {   // the Inbox: not a target itself, but its subfolders are — and at THIS prefix level
                    CollectFolders(f, prefix, list, depth, skipTree, skipSelf);
                    continue;
                }
                string eid = null; try { eid = (string)f.EntryID; } catch { }
                string name; try { name = (string)f.Name; } catch { continue; }
                if (!IsMailFolder(f)) continue;
                string disp = prefix == "" ? name : prefix + " / " + name;
                // Two folders can now share a display name (Inbox\Archive and the root Archive). Keep both:
                // the map is keyed by what the picker shows, so the second needs a distinct label.
                if (_folderMap.ContainsKey(disp))
                {
                    string bas = disp;
                    for (int n = 2; _folderMap.ContainsKey(disp); n++) disp = bas + " (" + n + ")";
                }
                if (eid != null) _folderMap[disp] = eid;
                if (fpath != null) _folderPaths[disp] = fpath;
                list.Add(disp);
                CollectFolders(f, disp, list, depth + 1, skipTree, skipSelf);   // nested subfolders
            }
        }

        // Where the learned folder list lives between Outlook sessions.
        private static string FolderCachePath()
        {
            return Path.Combine(AxonDataDir(), "folders.json");
        }

        // Display path -> the folder's FolderPath ("\\mailbox\Inbox\hr"). Kept ALONGSIDE the EntryID map
        // because entry ids are not dependable across enumerations, let alone across sessions: the id for a
        // folder fetched one way did not match the id for the same folder reached another way, which is what
        // silently broke the skip list earlier. FolderPath is derived from names and stays put, so it is the
        // key that survives being written to disk; the EntryID is only ever the fast path.
        private System.Collections.Generic.Dictionary<string, string> _folderPaths =
            new System.Collections.Generic.Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        private string[] LoadFolderCache()
        {
            try
            {
                string p = FolderCachePath();
                if (!File.Exists(p)) return new string[0];
                var js = new System.Web.Script.Serialization.JavaScriptSerializer();
                var d = js.DeserializeObject(File.ReadAllText(p)) as System.Collections.Generic.Dictionary<string, object>;
                var arr = d != null && d.ContainsKey("folders") ? d["folders"] as object[] : null;
                if (arr == null) return new string[0];
                var map = new System.Collections.Generic.Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                var paths = new System.Collections.Generic.Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                var prim = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var list = new System.Collections.Generic.List<string>();
                foreach (var o in arr)
                {
                    var e = o as System.Collections.Generic.Dictionary<string, object>;
                    if (e == null || !e.ContainsKey("n") || e["n"] == null) continue;
                    string n = e["n"].ToString();
                    if (map.ContainsKey(n)) continue;
                    map[n] = e.ContainsKey("id") && e["id"] != null ? e["id"].ToString() : "";
                    paths[n] = e.ContainsKey("p") && e["p"] != null ? e["p"].ToString() : "";
                    if (e.ContainsKey("m") && e["m"] != null && e["m"].ToString() == "1") prim.Add(n);
                    list.Add(n);
                }
                if (list.Count == 0) return new string[0];
                _folderMap = map; _folderPaths = paths; _primaryFolders = prim;
                return list.ToArray();
            }
            catch { }
            return new string[0];
        }

        private void SaveFolderCache()
        {
            try
            {
                var map = _folderMap; var paths = _folderPaths; var prim = _primaryFolders;
                if (map == null || map.Count == 0) return;
                var sb = new System.Text.StringBuilder("{\"folders\":[");
                bool first = true;
                foreach (var kv in map)
                {
                    string path; if (paths == null || !paths.TryGetValue(kv.Key, out path)) path = "";
                    if (!first) sb.Append(",");
                    first = false;
                    sb.Append("{\"n\":").Append(JsStr(kv.Key))
                      .Append(",\"p\":").Append(JsStr(path))
                      .Append(",\"id\":").Append(JsStr(kv.Value))
                      .Append(",\"m\":\"").Append(prim != null && prim.Contains(kv.Key) ? "1" : "0").Append("\"}");
                }
                sb.Append("]}");
                File.WriteAllText(FolderCachePath(), sb.ToString(), new System.Text.UTF8Encoding(false));
            }
            catch { }
        }

        private static string JsStr(string s)
        {
            var sb = new System.Text.StringBuilder("\"");
            foreach (char c in s ?? "")
            {
                if (c == '"' || c == '\\') sb.Append('\\').Append(c);
                else if (c < ' ') sb.Append("\\u").Append(((int)c).ToString("x4"));
                else sb.Append(c);
            }
            return sb.Append('"').ToString();
        }

        // Re-walk the stores now and write the result to disk. Behind the Settings button, for anyone who
        // just made a folder and wants it immediately instead of after the next Move.
        public string RelearnFolders()
        {
            var folders = EnumerateInboxFolders();
            SaveFolderCache();
            int stores = 1;
            try
            {
                var seen = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var kv in _folderPaths)
                {
                    string p = kv.Value ?? "";
                    int i = p.StartsWith("\\\\") ? p.IndexOf('\\', 2) : -1;
                    if (i > 0) seen.Add(p.Substring(2, i - 2));
                }
                if (seen.Count > 0) stores = seen.Count;
            }
            catch { }
            return folders.Length == 0
                ? "No folders found."
                : "Learned " + folders.Length + " folder" + (folders.Length == 1 ? "" : "s")
                  + " across " + stores + " mailbox" + (stores == 1 ? "" : "es") + ".";
        }

        // The folder behind a display path: its EntryID first, and if that no longer resolves — a cached id
        // from an earlier session, a folder since moved — walk the FolderPath instead.
        private dynamic ResolveFolder(dynamic ns, string display)
        {
            string eid;
            if (_folderMap != null && _folderMap.TryGetValue(display, out eid) && !string.IsNullOrEmpty(eid))
            { try { dynamic f = ns.GetFolderFromID(eid); if (f != null) return f; } catch { } }
            string path;
            if (_folderPaths != null && _folderPaths.TryGetValue(display, out path) && !string.IsNullOrEmpty(path))
            { try { return FolderByPath(ns, path); } catch { } }
            return null;
        }

        // "\\Mailbox name\Inbox\hr" -> the folder, by walking names from that store's root.
        // A folder whose own name contains a slash comes back from Outlook with it escaped — the real
        // folder "Energy Efficiency / Training Materials" appears in FolderPath as
        // "Energy Efficiency %2F Training Materials" — so each segment has to be decoded before it is
        // compared with Folder.Name, or that folder is never found. Only %2F is decoded: running the
        // segment through a general URL-unescape would mangle any name with a literal % in it.
        private static dynamic FolderByPath(dynamic ns, string path)
        {
            string p = (path ?? "").TrimStart('\\');
            var segs = p.Split('\\');
            for (int i = 0; i < segs.Length; i++)
                segs[i] = System.Text.RegularExpressions.Regex.Replace(segs[i], "%2F", "/",
                              System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (segs.Length < 1) return null;
            dynamic root = null;
            foreach (dynamic st in ns.Stores)
            {
                string nm = ""; try { nm = (string)st.DisplayName; } catch { }
                if (string.Equals(nm, segs[0], StringComparison.OrdinalIgnoreCase))
                { try { root = st.GetRootFolder(); } catch { } break; }
            }
            if (root == null) return null;
            dynamic cur = root;
            for (int i = 1; i < segs.Length; i++)
            {
                dynamic next = null;
                foreach (dynamic f in cur.Folders)
                {
                    string n = ""; try { n = (string)f.Name; } catch { }
                    if (string.Equals(n, segs[i], StringComparison.OrdinalIgnoreCase)) { next = f; break; }
                }
                if (next == null) return null;
                cur = next;
            }
            return cur;
        }

        // Move the email into the chosen folder (looked up from its display path).
        private void MoveTo(dynamic mail, string display)
        {
            try
            {
                dynamic ns = ((dynamic)_app).GetNamespace("MAPI");
                dynamic dest = ResolveFolder(ns, display);
                if (dest != null) { mail.Move(dest); return; }
                Ui.Notify("Folder not found: " + display + "\n\nIt may have been renamed or removed. "
                          + "Settings > Learn my folders will refresh the list.", "Axon intelligence");
            }
            catch (Exception ex) { Ui.Notify("Couldn't move: " + ex.Message, "Axon intelligence"); }
        }

        private class Filing { public string[] matches = new string[0]; public string newFolder = ""; }

        // Suggest folders via an OpenAI-COMPATIBLE chat API. `api_base` can point at OpenAI's cloud
        // or any local server that speaks the same protocol (Ollama, vLLM, LM Studio, LocalAI, ...),
        // so the same add-in works for cloud or fully on-site deployments — only the config differs.
        // If the API is unreachable the picker still lists every folder; only AI ranking is skipped.
        private Filing SuggestFiling(string subject, string sender, string senderName, string body, string[] folders)
        {
            var result = new Filing();
            try { TrySuggestViaApi(subject, sender, body, folders, result); } catch { }

            // Strongest signal: folders where this sender's emails ALREADY live (where the user has filed
            // them before). Put those on top of the AI's topic guesses; it self-improves as you file.
            try
            {
                var hist = SenderFolders(senderName);
                if (hist.Count > 0)
                {
                    var merged = new System.Collections.Generic.List<string>(hist);
                    if (result.matches != null)
                        foreach (var m in result.matches)
                            if (!merged.Exists(x => string.Equals(x, m, StringComparison.OrdinalIgnoreCase))) merged.Add(m);
                    if (merged.Count > 5) merged = merged.GetRange(0, 5);
                    result.matches = merged.ToArray();
                }
            }
            catch { }

            // Never leave the 'create new folder' box empty when nothing matched — propose a name.
            if ((result.matches == null || result.matches.Length == 0) && string.IsNullOrWhiteSpace(result.newFolder))
                result.newFolder = FallbackFolderName(subject, sender);
            return result;
        }

        // Rank the user's folders by how many emails from this sender they already contain (the user's
        // own filing history). Returns up to 3 folder display-paths, most-used first.
        private System.Collections.Generic.List<string> SenderFolders(string senderName)
        {
            var ranked = new System.Collections.Generic.List<string>();
            try
            {
                var map = _folderMap; var prim = _primaryFolders;   // snapshot: a refresh may swap these
                if (string.IsNullOrWhiteSpace(senderName) || map == null || map.Count == 0) return ranked;
                dynamic ns = ((dynamic)_app).GetNamespace("MAPI");
                string filter = "[SenderName] = '" + senderName.Replace("'", "''") + "'";
                var counts = new System.Collections.Generic.Dictionary<string, int>();
                int scanned = 0;
                foreach (var kv in new System.Collections.Generic.List<System.Collections.Generic.KeyValuePair<string, string>>(map))
                {
                    if (scanned >= 60) break;   // bound the scan
                    if (prim != null && !prim.Contains(kv.Key)) continue;   // don't query other stores (slow)
                    scanned++;
                    try
                    {
                        dynamic f = ResolveFolder(ns, kv.Key);
                        if (f == null) continue;
                        int c = 0; try { c = (int)f.Items.Restrict(filter).Count; } catch { }
                        if (c > 0) counts[kv.Key] = c;
                    }
                    catch { }
                }
                var keys = new System.Collections.Generic.List<string>(counts.Keys);
                keys.Sort((a, b) => counts[b].CompareTo(counts[a]));
                for (int i = 0; i < keys.Count && i < 3; i++) ranked.Add(keys[i]);
            }
            catch { }
            return ranked;
        }

        // Derive a sensible new-folder name from the subject (or sender) when the model doesn't give one.
        private static string FallbackFolderName(string subject, string sender)
        {
            string s = (subject ?? "").Trim();
            s = System.Text.RegularExpressions.Regex.Replace(s, @"^((RE|FW|FWD|AW|VS|TR)\s*:\s*)+", "",
                                                             System.Text.RegularExpressions.RegexOptions.IgnoreCase).Trim();
            foreach (var sep in new[] { " | ", " - ", " – ", " — ", ": " })
            { int i = s.IndexOf(sep); if (i > 2) { s = s.Substring(0, i).Trim(); break; } }
            var words = s.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (words.Length > 4) s = string.Join(" ", words, 0, 4);
            if (string.IsNullOrWhiteSpace(s)) s = (sender ?? "").Trim();
            foreach (var c in Path.GetInvalidFileNameChars()) s = s.Replace(c.ToString(), "");
            s = s.Trim();
            if (s.Length > 40) s = s.Substring(0, 40).Trim();
            return string.IsNullOrWhiteSpace(s) ? "New folder" : s;
        }

        // Describe each folder for the model as its path + a few recent subjects (what actually lives
        // in it), so it can match on folder PURPOSE, not just the name. Bounded so it stays responsive.
        private string FolderSamples(string[] folders)
        {
            var sb = new System.Text.StringBuilder();
            dynamic ns = null; try { ns = ((dynamic)_app).GetNamespace("MAPI"); } catch { }
            int sampled = 0;
            foreach (var path in folders)
            {
                sb.Append("- ").Append(path).Append("\n");
                var map = _folderMap; var prim = _primaryFolders;   // snapshot: a refresh may swap these
                if (ns == null || sampled >= 40 || map == null || !map.ContainsKey(path)) continue;
                if (prim != null && !prim.Contains(path)) continue;   // other stores: listed, not probed
                sampled++;
                try
                {
                    dynamic fld = ResolveFolder(ns, path);
                    if (fld == null) continue;
                    dynamic items = fld.Items;
                    try { items.Sort("[ReceivedTime]", true); } catch { }
                    var subs = new System.Collections.Generic.List<string>();
                    dynamic m = null; try { m = items.GetFirst(); } catch { }
                    int k = 0;
                    while (m != null && subs.Count < 3 && k < 12)
                    {
                        k++;
                        try
                        {
                            if ((int)m.Class == 43)
                            {
                                string s = ""; try { s = (string)m.Subject; } catch { }
                                if (!string.IsNullOrWhiteSpace(s)) { s = s.Trim(); if (s.Length > 60) s = s.Substring(0, 60); subs.Add(s); }
                            }
                        }
                        catch { }
                        try { m = items.GetNext(); } catch { m = null; }
                    }
                    if (subs.Count > 0) sb.Append("    recent: ").Append(string.Join("; ", subs.ToArray())).Append("\n");
                }
                catch { }
            }
            return sb.ToString();
        }

        private bool TrySuggestViaApi(string subject, string sender, string body, string[] folders, Filing result)
        {
            var js = new System.Web.Script.Serialization.JavaScriptSerializer();
            string b = body ?? "";
            if (b.Length > 4000) b = b.Substring(0, 4000);
            string prompt =
                "You are filing an email into one of the user's Outlook folders.\n\n" +
                "Email\n  Subject: " + subject + "\n  From: " + sender + "\n  Body:\n" + b + "\n\n" +
                "The user's folders (each '- ' line is a folder path; 'recent:' shows a few subjects already " +
                "filed there, so you can tell what belongs in it):\n" + FolderSamples(folders) + "\n" +
                "READ THE BODY to understand what the email is really about — the subject alone is not enough. " +
                "Then pick the folder whose PURPOSE fits, judging from BOTH its path AND the kind of emails already " +
                "in it (the 'recent:' samples), plus who the email is from (the sender's company/domain is a strong " +
                "hint). List the best-fitting folders first (up to 5), ONLY genuinely good fits (do not pad). If none " +
                "clearly fit, leave matches empty and you MUST propose a short new_folder name based on the topic or " +
                "sender — never leave both empty.\n" +
                "matches must be folder paths copied EXACTLY from the '- ' lines (do NOT include the 'recent:' text).\n" +
                "Reply with ONLY JSON: {\"matches\": [up to 5 exact folder paths, best first], " +
                "\"new_folder\": \"a short (1-3 word) new folder name; REQUIRED whenever matches is empty\"}";
            string text = ModelComplete(prompt, 0);
            if (string.IsNullOrEmpty(text)) return false;
            var mt = System.Text.RegularExpressions.Regex.Match(text, "\\{[\\s\\S]*\\}");
            if (!mt.Success) return false;
            ParseSuggestion(mt.Value, folders, result, js);
            return true;
        }

        // Map the model's JSON onto the real folder names (case-insensitive); if it "invents" a new
        // folder that already exists, treat it as a match instead.
        private void ParseSuggestion(string json, string[] folders, Filing result, System.Web.Script.Serialization.JavaScriptSerializer js)
        {
            var d = (System.Collections.Generic.Dictionary<string, object>)js.DeserializeObject(json);
            var lower = new System.Collections.Generic.Dictionary<string, string>();
            foreach (var f in folders) if (f != null) lower[f.ToLowerInvariant()] = f;
            var matches = new System.Collections.Generic.List<string>();
            if (d.ContainsKey("matches") && d["matches"] is object[])
                foreach (var x in (object[])d["matches"])
                {
                    if (x == null) continue;
                    string k = x.ToString().Trim().ToLowerInvariant();
                    if (lower.ContainsKey(k) && !matches.Contains(lower[k])) matches.Add(lower[k]);
                }
            string nf = (d.ContainsKey("new_folder") && d["new_folder"] != null) ? d["new_folder"].ToString().Trim() : "";
            if (nf.Length > 0)
            {
                string nfl = nf.ToLowerInvariant();
                if (lower.ContainsKey(nfl)) { if (!matches.Contains(lower[nfl])) matches.Add(lower[nfl]); }
                else result.newFolder = nf;
            }
            if (matches.Count > 5) matches = matches.GetRange(0, 5);
            result.matches = matches.ToArray();
        }

        // Create (or reuse) a top-level Inbox subfolder by name; returns the folder or null.
        private dynamic CreateInboxSubfolder(string name)
        {
            try
            {
                dynamic ns = ((dynamic)_app).GetNamespace("MAPI");
                dynamic inbox = ns.GetDefaultFolder(6);
                foreach (dynamic f in inbox.Folders)
                    if (string.Equals((string)f.Name, name, StringComparison.OrdinalIgnoreCase)) return f;
                dynamic made = inbox.Folders.Add(name);
                // The cached list predates this folder — re-learn in the background so it is offered next
                // time rather than only after the refresh that runs behind the following Move.
                var t = new System.Threading.Thread(() => { try { EnumerateInboxFolders(); SaveFolderCache(); } catch { } });
                t.IsBackground = true; t.Start();
                return made;
            }
            catch (Exception ex) { Ui.Notify("Couldn't create folder: " + ex.Message, "Axon intelligence"); return null; }
        }

        // --- COM (de)registration: add/remove the Outlook add-in registry entry ---
        private const string AddinKey = @"Software\Microsoft\Office\Outlook\AddIns\Axon.OutlookAddin";

        [ComRegisterFunction]
        public static void RegisterFunction(Type t)
        {
            using (var k = Registry.CurrentUser.CreateSubKey(AddinKey))
            {
                k.SetValue("FriendlyName", "Axon intelligence");
                k.SetValue("Description", "File and download emails with Axon intelligence");
                k.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
            }
        }

        [ComUnregisterFunction]
        public static void UnregisterFunction(Type t)
        {
            try { Registry.CurrentUser.DeleteSubKey(AddinKey, false); } catch { }
        }
    }

    // Shared look-and-feel for the Axon dialogs (black & white, flat, Segoe UI).
    internal static class Ui
    {
        public static readonly System.Drawing.Color Ink = System.Drawing.Color.FromArgb(51, 51, 58);     // softened text (dark grey, not pure black)
        public static readonly System.Drawing.Color AccentBg = System.Drawing.Color.FromArgb(74, 74, 84);   // button fill — dark grey
        public static readonly System.Drawing.Color AccentHover = System.Drawing.Color.FromArgb(92, 92, 104);
        public static readonly System.Drawing.Color Muted = System.Drawing.Color.FromArgb(130, 130, 140);
        public static readonly System.Drawing.Color SelBg = System.Drawing.Color.FromArgb(236, 238, 248);
        public static readonly System.Drawing.Color Line = System.Drawing.Color.FromArgb(220, 220, 226);

        public static Label Title(string t, int x, int y, int w)
        { return new Label { Text = t, Left = x, Top = y, Width = w, Height = 26, ForeColor = Ink, Font = new System.Drawing.Font("Segoe UI", 13F, System.Drawing.FontStyle.Bold) }; }
        public static Label Sub(string t, int x, int y, int w)
        { return new Label { Text = t, Left = x, Top = y, Width = w, Height = 20, ForeColor = Muted, AutoEllipsis = true }; }
        public static Label Caption(string t, int x, int y, int w)
        { return new Label { Text = t, Left = x, Top = y, Width = w, Height = 15, ForeColor = Muted, Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Bold) }; }
        public static Label Hint(string t, int x, int y, int w, int h)
        { return new Label { Text = t, Left = x, Top = y, Width = w, Height = h, ForeColor = Muted }; }

        public static Button Accent(string text, int x, int y, int w, int h)
        {
            var b = new Button { Text = text, Left = x, Top = y, Width = w, Height = h, FlatStyle = FlatStyle.Flat, BackColor = AccentBg, ForeColor = System.Drawing.Color.White, Font = new System.Drawing.Font("Segoe UI", 9.5F, System.Drawing.FontStyle.Bold), Cursor = Cursors.Hand };
            b.FlatAppearance.BorderSize = 0;
            b.FlatAppearance.MouseOverBackColor = AccentHover;
            return b;
        }
        public static Button Subtle(string text, int x, int y, int w, int h)
        {
            var b = new Button { Text = text, Left = x, Top = y, Width = w, Height = h, FlatStyle = FlatStyle.Flat, BackColor = System.Drawing.Color.White, ForeColor = System.Drawing.Color.FromArgb(40, 40, 40), Font = new System.Drawing.Font("Segoe UI", 9.5F), Cursor = Cursors.Hand };
            b.FlatAppearance.BorderColor = Line;
            b.FlatAppearance.BorderSize = 1;
            return b;
        }
        // Full-width left-aligned accent button used for suggestion rows.
        public static Button RowBtn(string text, int x, int y, int w, int h)
        {
            var b = Accent(text, x, y, w, h);
            b.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            b.Padding = new Padding(14, 0, 0, 0);
            return b;
        }
        public static string Leaf(string p)
        {
            if (string.IsNullOrEmpty(p)) return p;
            int i = p.LastIndexOf(" / ", StringComparison.Ordinal);
            if (i >= 0) return p.Substring(i + 3);
            string t = p.TrimEnd('\\', '/');
            int j = Math.Max(t.LastIndexOf('\\'), t.LastIndexOf('/'));
            return j >= 0 ? t.Substring(j + 1) : t;
        }
        // Owner-draw a folder list row as: leaf name (bold) + full path (gray).
        public static void DrawFolderItem(ListBox list, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            string path = list.Items[e.Index].ToString();
            bool sel = (e.State & DrawItemState.Selected) != 0;
            using (var bg = new System.Drawing.SolidBrush(sel ? SelBg : System.Drawing.Color.White))
                e.Graphics.FillRectangle(bg, e.Bounds);
            using (var nf = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Bold))
            using (var pf = new System.Drawing.Font("Segoe UI", 8F))
            using (var ink = new System.Drawing.SolidBrush(Ink))
            using (var mut = new System.Drawing.SolidBrush(Muted))
            {
                e.Graphics.DrawString(Leaf(path), nf, ink, e.Bounds.Left + 12, e.Bounds.Top + 6);
                e.Graphics.DrawString(path, pf, mut, e.Bounds.Left + 12, e.Bounds.Top + 26);
            }
        }

        // Branded replacement for MessageBox — matches the Axon dialog style.
        public static void Notify(string message, string title = "Axon intelligence")
        {
            using (var f = new Form())
            {
                f.Text = title;
                f.FormBorderStyle = FormBorderStyle.FixedDialog;
                f.StartPosition = FormStartPosition.CenterScreen;
                f.MaximizeBox = false; f.MinimizeBox = false; f.ShowInTaskbar = false;
                f.BackColor = System.Drawing.Color.White;
                f.Font = new System.Drawing.Font("Segoe UI", 9.5F);
                f.ClientSize = new System.Drawing.Size(400, 170);
                int W = f.ClientSize.Width, H = f.ClientSize.Height;
                f.Controls.Add(Title("Axon intelligence", 22, 18, W - 44));
                f.Controls.Add(new Label { Text = message ?? "", Left = 22, Top = 52, Width = W - 44, Height = H - 52 - 54, ForeColor = Ink });
                var ok = Accent("OK", W - 22 - 96, H - 46, 96, 32);
                ok.DialogResult = DialogResult.OK;
                f.Controls.Add(ok);
                f.AcceptButton = ok; f.CancelButton = ok;
                f.ShowDialog();
            }
        }
    }
}
