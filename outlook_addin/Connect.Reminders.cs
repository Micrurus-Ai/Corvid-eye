// Axon Outlook add-in — In-Outlook Follow-up / Send-Later reminder service.  (partial of Connect; split out of AxonAddin.cs.)
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
    public partial class Connect
    {
        // --- Reminder service --------------------------------------------------------------------
        // Fires Follow-up / Send-Later reminders from INSIDE Outlook, so they work even without the
        // floating dot. Reads the same JSON the add-in writes and uses the SAME 'seen' files as the
        // dot (keys match), and it stands down if the dot is running — so the two never double-fire.
        private System.Windows.Forms.Timer _reminderTimer;

        private void StartReminderService()
        {
            try
            {
                _reminderTimer = new System.Windows.Forms.Timer();
                _reminderTimer.Interval = 60000;                 // check every minute
                _reminderTimer.Tick += (o, e) => CheckRemindersSafe();
                _reminderTimer.Start();
                var first = new System.Windows.Forms.Timer();    // and once, ~10s after Outlook opens
                first.Interval = 10000;
                first.Tick += (o, e) => { first.Stop(); first.Dispose(); CheckRemindersSafe(); };
                first.Start();
            }
            catch { }
        }

        private void CheckRemindersSafe() { try { CheckReminders(); } catch { } }

        private static bool DotIsRunning()
        {
            try { return System.Diagnostics.Process.GetProcessesByName("AxonIntelligence").Length > 0; }
            catch { return false; }
        }

        private static string AxonDir()
        {
            return AxonDataDir();
        }

        private static System.Collections.Generic.List<System.Collections.Generic.Dictionary<string, object>> LoadArray(string path)
        {
            var outList = new System.Collections.Generic.List<System.Collections.Generic.Dictionary<string, object>>();
            try
            {
                if (!File.Exists(path)) return outList;
                var js = new System.Web.Script.Serialization.JavaScriptSerializer();
                var arr = js.DeserializeObject(File.ReadAllText(path)) as object[];
                if (arr != null)
                    foreach (var o in arr)
                    {
                        var d = o as System.Collections.Generic.Dictionary<string, object>;
                        if (d != null) outList.Add(d);
                    }
            }
            catch { }
            return outList;
        }

        private static System.Collections.Generic.HashSet<string> LoadSeen(string path)
        {
            var set = new System.Collections.Generic.HashSet<string>();
            try
            {
                if (!File.Exists(path)) return set;
                var js = new System.Web.Script.Serialization.JavaScriptSerializer();
                var arr = js.DeserializeObject(File.ReadAllText(path)) as object[];
                if (arr != null) foreach (var o in arr) if (o != null) set.Add(o.ToString());
            }
            catch { }
            return set;
        }

        private static void SaveSeen(string path, System.Collections.Generic.HashSet<string> set)
        {
            try
            {
                var js = new System.Web.Script.Serialization.JavaScriptSerializer();
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                File.WriteAllText(path, js.Serialize(new System.Collections.Generic.List<string>(set)));
            }
            catch { }
        }

        private static string F(System.Collections.Generic.Dictionary<string, object> d, string k)
        {
            object v; return (d != null && d.TryGetValue(k, out v) && v != null) ? v.ToString() : "";
        }

        private void CheckReminders()
        {
            if (DotIsRunning()) return;   // the dot owns reminders when it's running; avoid duplicates
            DateTime now = DateTime.Now;
            string dir = AxonDir();

            // Follow-ups (seen key: id|due, matching the dot).
            var fItems = LoadArray(Path.Combine(dir, "followups.json"));
            if (fItems.Count > 0)
            {
                string seenPath = Path.Combine(dir, "followups_seen.json");
                var seen = LoadSeen(seenPath);
                bool changed = false;
                foreach (var it in fItems)
                {
                    string due = F(it, "due");
                    string key = F(it, "id") + "|" + due;
                    if (seen.Contains(key)) continue;
                    DateTime d;
                    if (!DateTime.TryParse(due, out d) || d > now) continue;
                    DateTime created; DateTime.TryParse(F(it, "created"), out created);
                    if (HasReply(F(it, "id"), created)) { seen.Add(key); changed = true; continue; }
                    string who = F(it, "who"), subj = F(it, "subject");
                    ShowReminder("Follow up" + (string.IsNullOrEmpty(who) ? "" : " with " + who),
                                 string.IsNullOrEmpty(subj) ? "(no subject)" : subj);
                    seen.Add(key); changed = true;
                }
                if (changed) SaveSeen(seenPath, seen);
            }

            // Scheduled sends (seen keys: to|when|remind and to|when|sent, matching the dot).
            var sItems = LoadArray(Path.Combine(dir, "scheduled.json"));
            if (sItems.Count > 0)
            {
                string seenPath = Path.Combine(dir, "scheduled_seen.json");
                var seen = LoadSeen(seenPath);
                bool changed = false;
                foreach (var it in sItems)
                {
                    string bas = F(it, "to") + "|" + F(it, "when");
                    string subj = F(it, "subject");
                    string remindAt = F(it, "remind_at");
                    if (!string.IsNullOrEmpty(remindAt))
                    {
                        string rk = bas + "|remind"; DateTime rt;
                        if (!seen.Contains(rk) && DateTime.TryParse(remindAt, out rt) && rt <= now)
                        {
                            ShowReminder("Axon will send this scheduled email at " + F(it, "when") + ".",
                                string.IsNullOrEmpty(subj) ? "Scheduled email" : subj);
                            seen.Add(rk); changed = true;
                        }
                    }
                    string sk = bas + "|sent"; DateTime w;
                    if (!seen.Contains(sk) && DateTime.TryParse(F(it, "when"), out w) && w <= now)
                    {
                        ShowReminder("This scheduled email has now been sent from your Outbox.",
                            string.IsNullOrEmpty(subj) ? "Scheduled email" : subj);
                        seen.Add(sk); changed = true;
                    }
                }
                if (changed) SaveSeen(seenPath, seen);
            }
        }

        // Best-effort: did a message in this conversation arrive after the follow-up was set?
        private bool HasReply(string entryId, DateTime created)
        {
            if (string.IsNullOrEmpty(entryId)) return false;
            try
            {
                dynamic app = _app; if (app == null) return false;
                dynamic ns = app.GetNamespace("MAPI");
                dynamic item = ns.GetItemFromID(entryId);
                if (item == null) return false;
                dynamic conv = null; try { conv = item.GetConversation(); } catch { }
                if (conv == null) return false;
                dynamic table = conv.GetTable();
                try { table.Columns.Add("ReceivedTime"); } catch { }
                while (!(bool)table.EndOfTable)
                {
                    dynamic row = table.GetNextRow();
                    try
                    {
                        object rto = row["ReceivedTime"];
                        if (rto != null && Convert.ToDateTime(rto) > created.AddMinutes(1)) return true;
                    }
                    catch { }
                }
            }
            catch { }
            return false;
        }

        // A reliable reminder styled like Outlook's own "1 Reminder" window: light, item in bold,
        // Dismiss + Snooze. Non-modal + topmost so it's dependable (Outlook's native one is flaky).
        private void ShowReminder(string kind, string item)
        {
            try
            {
                var f = new Form
                {
                    Text = "1 Reminder",
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    MaximizeBox = false, MinimizeBox = false,
                    ShowInTaskbar = false, TopMost = true,
                    BackColor = System.Drawing.Color.White,
                    Font = new System.Drawing.Font("Segoe UI", 9F),
                    ClientSize = new System.Drawing.Size(410, 152),
                    StartPosition = FormStartPosition.CenterScreen
                };
                int W = f.ClientSize.Width;
                f.Controls.Add(new Label
                {
                    Text = "🔔", Left = 16, Top = 18, Width = 34, Height = 34,   // 🔔
                    Font = new System.Drawing.Font("Segoe UI Emoji", 15F),
                    ForeColor = System.Drawing.Color.FromArgb(198, 118, 28)
                });
                f.Controls.Add(new Label
                {
                    Text = string.IsNullOrEmpty(item) ? "(no subject)" : item,
                    Left = 58, Top = 18, Width = W - 74, Height = 24, AutoEllipsis = true,
                    Font = new System.Drawing.Font("Segoe UI", 10.5F, System.Drawing.FontStyle.Bold),
                    ForeColor = System.Drawing.Color.FromArgb(28, 28, 28)
                });
                f.Controls.Add(new Label
                {
                    Text = kind, Left = 58, Top = 46, Width = W - 74, Height = 34, AutoEllipsis = true,
                    ForeColor = System.Drawing.Color.FromArgb(96, 96, 96)
                });
                f.Controls.Add(new Label { Left = 0, Top = 100, Width = W, Height = 1, BackColor = System.Drawing.Color.FromArgb(224, 224, 224) });

                var dismiss = new Button { Text = "Dismiss", Left = W - 16 - 96, Top = 112, Width = 96, Height = 30, FlatStyle = FlatStyle.System };
                dismiss.Click += (o, e) => { try { f.Close(); } catch { } };
                var snooze = new Button { Text = "Snooze 5 min", Left = W - 16 - 96 - 8 - 116, Top = 112, Width = 116, Height = 30, FlatStyle = FlatStyle.System };
                snooze.Click += (o, e) =>
                {
                    try { f.Close(); } catch { }
                    var re = new System.Windows.Forms.Timer { Interval = 5 * 60 * 1000 };
                    re.Tick += (o2, e2) => { re.Stop(); re.Dispose(); ShowReminder(kind, item); };
                    re.Start();
                };
                f.Controls.Add(dismiss); f.Controls.Add(snooze);
                f.AcceptButton = dismiss;
                f.FormClosed += (o, e) => { try { f.Dispose(); } catch { } };
                f.Show();
            }
            catch { }
        }
    }
}
