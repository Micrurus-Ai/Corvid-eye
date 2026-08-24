// Axon Outlook add-in — Ribbon actions: Summarize, Reply, Schedule, Attach, Write, Send Later, Follow up.  (partial of Connect; split out of AxonAddin.cs.)
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
        public void OnSummarize(object control)
        {
            try
            {
                object m = GetSelectedMail();
                if (m == null) { Ui.Notify("Select an email first.", "Axon intelligence"); return; }
                dynamic mail = m;
                string subj = ""; try { subj = (string)mail.Subject; } catch { }
                string body = ""; try { body = (string)mail.Body; } catch { }
                body = TrimBody(body, SummaryBodyMax);
                string bodyCopy = body;
                // A reply thread is a big body that also needs long-context reasoning to build the
                // timeline. The backup provider handles those markedly faster, so send threads there.
                bool thread = IsThread(subj, bodyCopy);
                bool wantsReply;
                using (var dlg = new SummaryDialog(subj,
                    lang => ModelComplete(BuildSummaryPrompt(subj, bodyCopy, lang), 0.3, thread)))
                {
                    dlg.ShowDialog();
                    wantsReply = dlg.WantsReply;
                }
                if (wantsReply) OnReply(control);   // jump straight into the reply flow on the same email
            }
            catch (Exception ex) { Ui.Notify("Axon error: " + ex.Message, "Axon intelligence"); }
        }

        // True when the email carries earlier messages: either the subject is a reply/forward, or the
        // body contains a quoted message header. Multilingual, since the office writes NL / FR / EN.
        private static bool IsThread(string subject, string body)
        {
            string s = (subject ?? "").TrimStart();
            string[] prefixes = { "RE:", "RE :", "FW:", "FWD:", "TR:", "ANTW:", "AW:", "VS:" };
            foreach (var p in prefixes)
                if (s.StartsWith(p, StringComparison.OrdinalIgnoreCase)) return true;

            string b = body ?? "";
            string[] marks = { "\nFrom:", "\nVan:", "\nDe :", "\nDe:", "\nSent:", "\nVerzonden:", "\nEnvoyé",
                               "-----Original", "-----Oorspronkelijk", "-----Message d'origine" };
            foreach (var m in marks)
                if (b.IndexOf(m, StringComparison.OrdinalIgnoreCase) >= 0) return true;
            return false;
        }

        // No practical limit: 150k characters is roughly 39k tokens, and the longest thread anyone
        // actually sends is a small fraction of that. A hard ceiling still has to exist, because past it
        // the model call fails outright (context window / rate limit) and the user gets NO answer at all
        // — worse than trimming. Measured against both providers: 150k is comfortable on each; 400k
        // rate-limits the primary.
        private const int SummaryBodyMax = 150000;
        // A reply needs the thread's history, but not at summary depth — 50k chars (~13k tokens) covers
        // any real back-and-forth and keeps drafting snappy.
        private const int ReplyBodyMax = 50000;

        // If a thread exceeds the ceiling, keep the NEWEST messages (top of the body) and the OLDEST
        // (bottom) and drop only the middle, so both ends of the conversation survive.
        private static string TrimBody(string body, int max)
        {
            if (string.IsNullOrEmpty(body) || body.Length <= max) return body;
            int half = max / 2;
            return body.Substring(0, half)
                 + "\n\n[... middle of this very long thread omitted ...]\n\n"
                 + body.Substring(body.Length - half);
        }

        private string BuildSummaryPrompt(string subject, string body, string lang)
        {
            string langLine = string.IsNullOrEmpty(lang)
                ? "Write the summary in the same language as the email."
                : "Write the summary in " + lang + ".";
            return "Summarize this email THOROUGHLY in PLAIN TEXT only — no markdown, no asterisks, no '#'. " +
                "Capture all the important information: who is involved, what they are asking or saying, and every " +
                "specific detail that matters (names, dates, amounts, quantities, order/reference numbers, deadlines, " +
                "links, attachments, and any questions raised). Do NOT over-compress — it is better to keep a detail " +
                "than to drop it. Use this layout, with a blank line between sections:\n\n" +
                "Gist: <2-3 sentences giving the full picture>\n\n" +
                "Conversation:\n" +
                "- On 25 June, Jan replied: <what he said>\n" +
                "- On 26 June, Disan sent: <what she said>\n" +
                "(ONLY include this section when the email is a thread with MORE THAN ONE message. Walk the " +
                "replies in time order, OLDEST FIRST. One line per message: start with 'On <day> <month>' " +
                "(add the time only if two messages share a day), then WHO it was and whether they sent, replied " +
                "or forwarded, then what they actually said. Never merge two messages into one line, and never " +
                "invent a date — use only the dates in the email. If it is a single message with no replies, " +
                "OMIT this whole section, heading included.)\n\n" +
                "Key points:\n- <point>\n- <point>\n- <point>\n(the substance that is not already obvious from the " +
                "conversation above — decisions, numbers, names, links, attachments, open questions. Do not repeat " +
                "the thread line by line.)\n\n" +
                "Action: <what the reader should do, including any deadline or detail, or 'None'>\n\n" +
                langLine + "\n\nSubject: " + subject + "\n\n" + body;
        }

        public void OnReply(object control)
        {
            try
            {
                object m = GetSelectedMail();
                if (m == null) { Ui.Notify("Select an email first.", "Axon intelligence"); return; }
                dynamic mail = m;
                string subj = ""; try { subj = (string)mail.Subject; } catch { }
                string sender = ""; try { sender = (string)mail.SenderName; } catch { }
                string body = ""; try { body = (string)mail.Body; } catch { }
                // Read the whole back-and-forth, not just the newest message, so the draft knows what was
                // already said. Threads go to the backup provider first — it is much faster on big bodies.
                body = TrimBody(body, ReplyBodyMax);
                string me = ""; try { me = (string)((dynamic)_app).Session.CurrentUser.Name; } catch { }
                string bodyCopy = body;
                bool thread = IsThread(subj, bodyCopy);
                // Ask the user HOW they want to reply, then Axon drafts to that instruction.
                string draft;
                using (var prompt = new ReplyPrompt(subj,
                    (instr, lang) => ModelComplete(BuildReplyPrompt(subj, sender, bodyCopy, instr, me, lang, thread), 0.4, thread)))
                {
                    if (prompt.ShowDialog() != DialogResult.OK) return;
                    draft = prompt.Draft;
                }
                if (string.IsNullOrEmpty(draft)) { Ui.Notify("Couldn't draft a reply (model unavailable).", "Axon intelligence"); return; }
                dynamic reply = mail.Reply();
                // Draft at the top, a horizontal line, then Outlook's quoted original (keeps formatting).
                try
                {
                    string html = (string)reply.HTMLBody;
                    // No <hr> here — Outlook's reply already has its own divider above the quoted original.
                    string draftHtml = "<div style='font-family:Calibri,sans-serif;font-size:11pt;'>" +
                        ToHtml(draft) + "<br></div>";
                    int bi = html.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
                    int gt = bi >= 0 ? html.IndexOf('>', bi) : -1;
                    reply.HTMLBody = gt >= 0 ? html.Substring(0, gt + 1) + draftHtml + html.Substring(gt + 1)
                                             : draftHtml + html;
                }
                catch { try { reply.Body = draft; } catch { } }
                reply.Display();   // open the draft for the user to review + send (never auto-sent)
            }
            catch (Exception ex) { Ui.Notify("Axon error: " + ex.Message, "Axon intelligence"); }
        }

        // Plain text -> minimal HTML (escape + line breaks) for inserting into the reply.
        private static string ToHtml(string text)
        {
            string e = (text ?? "").Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
            return e.Replace("\r\n", "\n").Replace("\n", "<br>");
        }

        private string BuildReplyPrompt(string subject, string sender, string body, string instruction, string me, string lang, bool thread)
        {
            // The user is in charge. When they say what to reply, that instruction decides the content and
            // the email is only background — the draft must not wander off answering points they never
            // raised. This matters most on a thread, where there is a lot of old material to be tempted by.
            string how = string.IsNullOrWhiteSpace(instruction)
                ? "Write a concise, professional, courteous reply that addresses the points raised."
                : "The USER'S INSTRUCTION below decides what this reply says. Follow it exactly and cover " +
                  "everything it asks for, and nothing it does not. Treat the email itself as background " +
                  "context only: do NOT answer other questions from it, and do NOT raise points the user " +
                  "did not ask you to raise.\nUSER'S INSTRUCTION: " + instruction;
            // On a thread, reply to the LATEST message; earlier ones are history, not open questions.
            string threadLine = thread
                ? " This email is a thread: the NEWEST message is at the top and older replies follow below. " +
                  "Reply to the newest message. Use the earlier messages only as background — do not re-answer " +
                  "points that were already settled earlier in the thread. Never quote the thread back."
                : "";
            // Detect the language from the TEXT, never from a name: 'Hélène' or 'Jan' is not evidence that
            // an English email is French or Dutch, and getting this wrong sends a reply in the wrong language.
            string langLine = string.IsNullOrEmpty(lang)
                ? "Write the reply in the SAME language as the email body. Determine that language from the " +
                  "WORDS of the most recent message only. Ignore the language of people's names, signatures, " +
                  "companies and addresses — a French- or Dutch-looking name does NOT mean the email is in " +
                  "French or Dutch. If the message is written in English, reply in English."
                : "Write the ENTIRE reply (greeting, message, and sign-off) in " + lang + ", regardless of the email's language.";
            string tone = MyToneGuide();
            string toneLine = string.IsNullOrEmpty(tone) ? "" : " IMPORTANT: write in the USER'S OWN VOICE, following this style profile exactly — same greeting, sign-off, vocabulary, recurring phrases, punctuation and level of formality, so it reads as if the user wrote it themselves:\n" + tone + "\n";
            return how + threadLine + " Begin with an appropriate greeting addressed to the sender by first " +
                   "name, and end with a short closing line such as 'Kind regards,' (matching the user's usual " +
                   "sign-off phrase). " + langLine + toneLine + NoSignatureLine +
                   " Use a natural tone. Output ONLY the reply text itself (greeting, message, closing line) — " +
                   "no subject line and no quoted original.\n\n" +
                   "Sender: " + sender + "\nSubject: " + subject + "\n\n" + body;
        }

        // The user's learned writing style (saved by the dot's "learn my tone" to %APPDATA%\AxonOutlook\tone.txt).
        private static string MyToneGuide()
        {
            try
            {
                string p = Path.Combine(AxonDataDir(), "tone.txt");
                if (File.Exists(p)) { string s = File.ReadAllText(p).Trim(); if (s.Length > 0) return s; }
            }
            catch { }
            return "";
        }

        public void OnSchedule(object control)
        {
            try
            {
                object m = GetSelectedMail();
                if (m == null) { Ui.Notify("Select an email first.", "Axon intelligence"); return; }
                dynamic mail = m;
                string subj = ""; try { subj = (string)mail.Subject; } catch { }
                string sender = ""; try { sender = (string)mail.SenderName; } catch { }
                string body = ""; try { body = (string)mail.Body; } catch { }
                if (body.Length > 6000) body = body.Substring(0, 6000);
                // Ask the user WHEN (any timeframe — "next Tue 2pm", "week of 15 Sept", "in 3 months",
                // or blank to suggest soon) and how long. Axon then drafts the invite to match.
                string when; int dur;
                using (var sp = new SchedulePrompt(subj))
                {
                    if (sp.ShowDialog() != DialogResult.OK) return;
                    when = sp.When; dur = sp.DurationMinutes;
                }
                string today = DateTime.Now.ToString("yyyy-MM-dd (dddd)");
                string busy = MyBusyBlocks((dynamic)_app, 14);   // next 2 weeks, for near-term conflict checks
                string text = ModelComplete(BuildSchedulePrompt(subj, sender, body, today, busy, when), 0.2);
                if (string.IsNullOrEmpty(text)) { Ui.Notify("Couldn't propose a meeting (model unavailable).", "Axon intelligence"); return; }
                var mt = System.Text.RegularExpressions.Regex.Match(text, "\\{[\\s\\S]*\\}");
                if (!mt.Success) { Ui.Notify("Couldn't read the proposed meeting details.", "Axon intelligence"); return; }
                var js = new System.Web.Script.Serialization.JavaScriptSerializer();
                var d = (System.Collections.Generic.Dictionary<string, object>)js.DeserializeObject(mt.Value);

                string mSubject = Field(d, "subject");
                if (string.IsNullOrWhiteSpace(mSubject)) mSubject = "Meeting: " + subj;
                DateTime start;
                if (!DateTime.TryParse(Field(d, "start"), System.Globalization.CultureInfo.InvariantCulture,
                        System.Globalization.DateTimeStyles.None, out start))
                    start = DateTime.Now.Date.AddDays(1).AddHours(10);
                string loc = Field(d, "location");
                string agenda = Field(d, "agenda");

                dynamic app = _app;
                dynamic appt = app.CreateItem(1);   // olAppointmentItem
                appt.MeetingStatus = 1;             // olMeeting -> a sendable invite
                appt.Subject = mSubject;
                appt.Start = start;
                appt.Duration = dur;
                if (!string.IsNullOrEmpty(loc)) appt.Location = loc;
                if (!string.IsNullOrEmpty(agenda)) appt.Body = agenda;
                // Invite everyone on the email: sender + To as Required, Cc as Optional (minus yourself).
                string meSmtp = ""; try { meSmtp = SmtpOf(app.Session.CurrentUser.AddressEntry); } catch { }
                var seen = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (!string.IsNullOrEmpty(meSmtp)) seen.Add(meSmtp);
                Action<string, int> addAttendee = (addr, type) =>
                {
                    if (string.IsNullOrEmpty(addr) || !seen.Add(addr)) return;
                    try { dynamic rr = appt.Recipients.Add(addr); rr.Type = type; } catch { }
                };
                try { addAttendee(SmtpOf(mail.Sender), 1); } catch { }            // sender -> Required
                try
                {
                    foreach (dynamic r in mail.Recipients)
                    {
                        string a = ""; try { a = SmtpOf(r.AddressEntry); } catch { }
                        if (string.IsNullOrEmpty(a)) { try { a = (string)r.Address; } catch { } }
                        int t = 1; try { t = (int)r.Type; } catch { }            // mail: 1=To, 2=Cc, 3=Bcc
                        addAttendee(a, t == 2 ? 2 : 1);                           // Cc -> Optional, else Required
                    }
                }
                catch { }
                try { appt.Recipients.ResolveAll(); } catch { }
                // Single clean window: the pre-filled invite (already at a free time). Outlook's own
                // Scheduling Assistant inside it shows everyone's free/busy if the user wants to adjust.
                appt.Display();   // review + send (never auto-sent)
            }
            catch (Exception ex) { Ui.Notify("Axon error: " + ex.Message, "Axon intelligence"); }
        }

        private static string Field(System.Collections.Generic.Dictionary<string, object> d, string k)
        {
            return (d != null && d.ContainsKey(k) && d[k] != null) ? d[k].ToString() : "";
        }

        // Resolve a sender/recipient AddressEntry to its SMTP address (handles Exchange senders).
        private static string SmtpOf(dynamic ae)
        {
            try
            {
                if (ae == null) return "";
                try { dynamic eu = ae.GetExchangeUser(); if (eu != null) { string s = (string)eu.PrimarySmtpAddress; if (!string.IsNullOrEmpty(s)) return s; } } catch { }
                try { dynamic pa = ae.PropertyAccessor; string s = (string)pa.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001F"); if (!string.IsNullOrEmpty(s)) return s; } catch { }
                try { string s = (string)ae.Address; if (!string.IsNullOrEmpty(s)) return s; } catch { }
            }
            catch { }
            return "";
        }

        private string BuildSchedulePrompt(string subject, string sender, string body, string today, string busy, string when)
        {
            string target = string.IsNullOrWhiteSpace(when)
                ? "Pick the soonest suitable time during business hours (Mon–Fri, 09:00–17:00)."
                : "The user wants the meeting around: \"" + when + "\". Resolve that relative to today into a " +
                  "concrete date and time during business hours (Mon–Fri, 09:00–17:00).";
            string avail = string.IsNullOrEmpty(busy)
                ? ""
                : " If the chosen time falls within the next 2 weeks, do NOT overlap these busy blocks:\n" + busy;
            return "Read this email and propose a meeting with the sender. Today is " + today + ". " + target +
                avail + "\nIf the email itself proposes a specific time and it fits, prefer that.\n\n" +
                "Reply with ONLY JSON:\n{\"subject\": \"a clear meeting title\", " +
                "\"start\": \"YYYY-MM-DDTHH:MM\" (24h), \"location\": \"a place, or 'Microsoft Teams', or empty\", " +
                "\"agenda\": \"a short 1-3 line agenda\"}.\n\n" +
                "From: " + sender + "\nSubject: " + subject + "\n\n" + body;
        }

        // The user's busy blocks over the next `days` days (recurrences expanded), for conflict-free
        // scheduling. Best-effort: returns "" if the calendar can't be read.
        private static string MyBusyBlocks(dynamic app, int days)
        {
            try
            {
                dynamic items = app.Session.GetDefaultFolder(9).Items;   // 9 = olFolderCalendar
                items.IncludeRecurrences = true;
                items.Sort("[Start]");
                DateTime from = DateTime.Now;
                DateTime to = from.AddDays(days);
                var ci = System.Globalization.CultureInfo.CurrentCulture;
                string filter = "[Start] < '" + to.ToString("g", ci) + "' AND [End] > '" + from.ToString("g", ci) + "'";
                var sb = new System.Text.StringBuilder();
                int n = 0;
                foreach (dynamic a in items.Restrict(filter))
                {
                    try
                    {
                        DateTime s = (DateTime)a.Start;
                        DateTime e = (DateTime)a.End;
                        bool allday = false; try { allday = (bool)a.AllDayEvent; } catch { }
                        string subj = ""; try { subj = (string)a.Subject; } catch { }
                        sb.AppendLine("- " + s.ToString("ddd yyyy-MM-dd HH:mm") + "-" + e.ToString("HH:mm") +
                            (allday ? " (all day)" : "") + (string.IsNullOrEmpty(subj) ? "" : " : " + subj));
                    }
                    catch { }
                    if (++n >= 60) break;
                }
                return sb.ToString().Trim();
            }
            catch { return ""; }
        }

        public void OnAttachEmail(object control)
        {
            try
            {
                dynamic app = _app;
                var items = new System.Collections.Generic.List<dynamic>();
                // Prefer the current selection (may be several); fall back to a single open email.
                try
                {
                    dynamic exp = app.ActiveExplorer();
                    if (exp != null && exp.Selection != null)
                    {
                        foreach (dynamic it in exp.Selection)
                        {
                            try { if ((int)it.Class == 43) items.Add(it); } catch { }
                        }
                    }
                }
                catch { }
                if (items.Count == 0)
                {
                    object m = GetSelectedMail();
                    if (m != null) items.Add((dynamic)m);
                }
                if (items.Count == 0) { Ui.Notify("Select an email first.", "Axon intelligence"); return; }

                dynamic nm = app.CreateItem(0);   // olMailItem
                string firstSubj = "";
                foreach (dynamic it in items)
                {
                    try { nm.Attachments.Add(it); } catch { }   // attach the message as a .msg
                    if (firstSubj == "") { try { firstSubj = (string)it.Subject; } catch { } }
                }
                nm.Subject = items.Count > 1 ? ("FW: " + items.Count + " emails") : ("FW: " + firstSubj);
                nm.Display();   // open the new email for the user to add a recipient/note and send
            }
            catch (Exception ex) { Ui.Notify("Axon error: " + ex.Message, "Axon intelligence"); }
        }

        public void OnWriteEmail(object control)
        {
            try
            {
                dynamic app = _app;
                dynamic insp = app.ActiveInspector();
                if (insp == null) { Ui.Notify("Open a new email first, then click Write with Axon.", "Axon intelligence"); return; }
                dynamic item = insp.CurrentItem;
                int cls = 0; try { cls = (int)item.Class; } catch { }
                if (item == null || cls != 43) { Ui.Notify("Write with Axon works on an email you're composing.", "Axon intelligence"); return; }
                string me = ""; try { me = (string)app.Session.CurrentUser.Name; } catch { }

                string draft = null;
                using (var prompt = new WritePrompt(
                    (instr, lang) => ModelComplete(BuildWritePrompt(instr, me, lang), 0.5)))
                {
                    if (prompt.ShowDialog() != DialogResult.OK) return;
                    draft = prompt.Draft;
                }
                if (string.IsNullOrEmpty(draft)) { Ui.Notify("Couldn't write the email (model unavailable).", "Axon intelligence"); return; }

                // Parse the leading To/Cc/Bcc/Subject headers — allowing blank lines BETWEEN them —
                // then the body is everything after the last header line.
                string to = null, cc = null, bcc = null, subject = null;
                var lines = draft.Replace("\r\n", "\n").Split('\n');
                int lastHeader = -1;
                int scanTo = Math.Min(lines.Length, 12);
                for (int i = 0; i < scanTo; i++)
                {
                    string ln = lines[i].Trim();
                    if (ln.Length == 0) continue;   // skip blank lines between headers
                    if (to == null && ln.StartsWith("To:", StringComparison.OrdinalIgnoreCase)) { to = ln.Substring(3).Trim(); lastHeader = i; continue; }
                    if (cc == null && ln.StartsWith("Cc:", StringComparison.OrdinalIgnoreCase)) { cc = ln.Substring(3).Trim(); lastHeader = i; continue; }
                    if (bcc == null && ln.StartsWith("Bcc:", StringComparison.OrdinalIgnoreCase)) { bcc = ln.Substring(4).Trim(); lastHeader = i; continue; }
                    if (subject == null && ln.StartsWith("Subject:", StringComparison.OrdinalIgnoreCase)) { subject = ln.Substring(8).Trim(); lastHeader = i; continue; }
                    break;   // a real content line -> headers are done
                }
                int bstart = lastHeader + 1;
                while (bstart < lines.Length && lines[bstart].Trim().Length == 0) bstart++;   // skip blank(s) after headers
                var bodySb = new System.Text.StringBuilder();
                for (int j = bstart; j < lines.Length; j++) bodySb.AppendLine(lines[j]);
                string body = bodySb.ToString().Trim();
                if (string.IsNullOrEmpty(body)) body = draft;   // fallback if no headers were found

                // Add To / Cc / Bcc as proper Recipients and resolve each against the address book
                // (GAL + Contacts). Adding + Resolve() per recipient is far more reliable than setting
                // the .To string and calling ResolveAll.
                to = StripNone(to); cc = StripNone(cc); bcc = StripNone(bcc);
                bool anyRecip = false;
                anyRecip |= AddRecipients(item, to, 1);    // olTo
                anyRecip |= AddRecipients(item, cc, 2);    // olCC
                anyRecip |= AddRecipients(item, bcc, 3);   // olBCC
                if (anyRecip) { try { item.Recipients.ResolveAll(); } catch { } try { item.Save(); } catch { } }
                if (!string.IsNullOrEmpty(subject))
                {
                    try { string cur = (string)item.Subject; if (string.IsNullOrWhiteSpace(cur)) item.Subject = subject; } catch { }
                }
                // Write the body into the open compose window (HTML) for the user to review + send.
                try
                {
                    string html = (string)item.HTMLBody;
                    string draftHtml = "<div style='font-family:Calibri,sans-serif;font-size:11pt;'>" + ToHtml(body) + "</div>";
                    int bi = html.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
                    int gt = bi >= 0 ? html.IndexOf('>', bi) : -1;
                    item.HTMLBody = gt >= 0 ? html.Substring(0, gt + 1) + draftHtml + html.Substring(gt + 1) : draftHtml + html;
                }
                catch { try { item.Body = body; } catch { } }
            }
            catch (Exception ex) { Ui.Notify("Axon error: " + ex.Message, "Axon intelligence"); }
        }

        // Add each name/address (';'-separated) to the compose item as a recipient of the given Outlook
        // type (1=To, 2=Cc, 3=Bcc) and resolve it against the address book so a bare NAME becomes a real
        // address. Returns true if any were added.
        private bool AddRecipients(dynamic item, string names, int type)
        {
            if (string.IsNullOrWhiteSpace(names)) return false;
            bool added = false;
            foreach (var raw in names.Split(';'))
            {
                string nm = raw.Trim();
                if (nm.Length == 0) continue;
                try
                {
                    dynamic r = item.Recipients.Add(nm);
                    r.Type = type;
                    try { r.Resolve(); } catch { }
                    bool ok = false; try { ok = (bool)r.Resolved; } catch { }
                    // A bare name Outlook couldn't resolve (e.g. a first name) -> look it up in Contacts.
                    if (!ok && nm.IndexOf('@') < 0)
                    {
                        string email = LookupContactEmail(nm);
                        if (!string.IsNullOrEmpty(email))
                        {
                            try { r.Delete(); } catch { }
                            try { dynamic r2 = item.Recipients.Add(email); r2.Type = type; try { r2.Resolve(); } catch { } } catch { }
                        }
                    }
                    added = true;
                }
                catch { }
            }
            return added;
        }

        // Resolve a bare name to an email. Prefers an address the user has ACTUALLY corresponded with
        // (Sent/Inbox history), then their Contacts, then the company directory (GAL).
        private string LookupContactEmail(string name)
        {
            try { string h = LookupFromHistory(name); if (!string.IsNullOrEmpty(h)) return h; } catch { }
            try
            {
                dynamic app = _app; if (app == null) return "";
                dynamic items = app.GetNamespace("MAPI").GetDefaultFolder(10).Items;   // olFolderContacts
                string q = name.Replace("'", "''");
                foreach (var filter in new[] { "[FirstName] = '" + q + "'", "[FullName] = '" + q + "'", "[CompanyName] = '" + q + "'" })
                {
                    try { string e = ContactEmail(items.Restrict(filter).GetFirst()); if (!string.IsNullOrEmpty(e)) return e; } catch { }
                }
                // fallback: scan Contacts for a first-name or full-name match
                dynamic it = null; try { it = items.GetFirst(); } catch { }
                int n = 0;
                while (it != null && n < 3000)
                {
                    try
                    {
                        string first = ""; try { first = (string)it.FirstName; } catch { }
                        string full = ""; try { full = (string)it.FullName; } catch { }
                        if (string.Equals(first, name, StringComparison.OrdinalIgnoreCase) ||
                            (!string.IsNullOrEmpty(full) && full.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0))
                        { string e = ContactEmail(it); if (!string.IsNullOrEmpty(e)) return e; }
                    }
                    catch { }
                    try { it = items.GetNext(); } catch { it = null; }
                    n++;
                }

                // Also search the company directory (GAL) — this is where colleagues usually live.
                string g = LookupGalEmail(app, name);
                if (!string.IsNullOrEmpty(g)) return g;
            }
            catch { }
            return "";
        }

        // Search the Global Address List for a person whose name matches, and return their SMTP address.
        private string LookupGalEmail(dynamic app, string name)
        {
            try
            {
                dynamic entries = app.GetNamespace("MAPI").GetGlobalAddressList().AddressEntries;
                int count = 0; try { count = (int)entries.Count; } catch { }
                string exact = "", partial = "";
                for (int i = 1; i <= count && i <= 5000; i++)
                {
                    dynamic ae = null; try { ae = entries[i]; } catch { }
                    if (ae == null) continue;
                    string enm = ""; try { enm = (string)ae.Name; } catch { }
                    if (string.IsNullOrEmpty(enm)) continue;
                    if (string.Equals(enm, name, StringComparison.OrdinalIgnoreCase))
                    { string e = ExchangeSmtp(ae); if (!string.IsNullOrEmpty(e)) return e; }
                    if (partial == "" && (enm.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0
                         || enm.Split(' ')[0].Equals(name, StringComparison.OrdinalIgnoreCase)))
                        partial = ExchangeSmtp(ae);
                }
                return partial;
            }
            catch { return ""; }
        }

        private static string ExchangeSmtp(dynamic ae)
        {
            try { dynamic ex = ae.GetExchangeUser(); if (ex != null) { string s = (string)ex.PrimarySmtpAddress; if (!string.IsNullOrEmpty(s) && s.IndexOf('@') > 0) return s; } } catch { }
            try { string a = (string)ae.Address; if (!string.IsNullOrEmpty(a) && a.IndexOf('@') > 0) return a; } catch { }
            return "";
        }

        private static string ContactEmail(dynamic c)
        {
            if (c == null) return "";
            try { string e = (string)c.Email1Address; if (!string.IsNullOrEmpty(e) && e.IndexOf('@') > 0) return e; } catch { }
            try { string e = (string)c.Email2Address; if (!string.IsNullOrEmpty(e) && e.IndexOf('@') > 0) return e; } catch { }
            return "";
        }

        // Find the address the user has actually corresponded with for this name: scan recent Sent
        // (people they email) and Inbox (people who email them), tally SMTP addresses whose display
        // name matches, and return the most-used one. This beats a blind directory guess.
        private string LookupFromHistory(string name)
        {
            try
            {
                dynamic app = _app; if (app == null) return "";
                dynamic ns = app.GetNamespace("MAPI");
                var tally = new System.Collections.Generic.Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

                // Sent: recipients I send to (weighted higher — it's who I choose to write).
                try
                {
                    dynamic items = ns.GetDefaultFolder(5).Items;   // olFolderSentMail
                    try { items.Sort("[SentOn]", true); } catch { }
                    dynamic m = null; try { m = items.GetFirst(); } catch { }
                    int n = 0;
                    while (m != null && n < 150)
                    {
                        try
                        {
                            if ((int)m.Class == 43)
                            {
                                dynamic recips = m.Recipients; int rc = 0; try { rc = (int)recips.Count; } catch { }
                                for (int i = 1; i <= rc; i++)
                                {
                                    dynamic r = null; try { r = recips[i]; } catch { }
                                    if (r == null) continue;
                                    string rn = ""; try { rn = (string)r.Name; } catch { }
                                    string ra = ""; try { ra = (string)r.Address; } catch { }   // cheap: SMTP or EX DN
                                    bool match = (!string.IsNullOrEmpty(rn) && rn.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0)
                                              || (!string.IsNullOrEmpty(ra) && ra.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0);
                                    if (match)
                                    {
                                        string e = SmtpFromRecipient(r);
                                        if (!string.IsNullOrEmpty(e)) tally[e] = (tally.ContainsKey(e) ? tally[e] : 0) + 2;
                                    }
                                }
                            }
                        }
                        catch { }
                        try { m = items.GetNext(); } catch { m = null; }
                        n++;
                    }
                }
                catch { }

                // Inbox: senders who write to me.
                try
                {
                    dynamic items = ns.GetDefaultFolder(6).Items;   // olFolderInbox
                    try { items.Sort("[ReceivedTime]", true); } catch { }
                    dynamic m = null; try { m = items.GetFirst(); } catch { }
                    int n = 0;
                    while (m != null && n < 150)
                    {
                        try
                        {
                            if ((int)m.Class == 43)
                            {
                                string sn = ""; try { sn = (string)m.SenderName; } catch { }
                                string sa = ""; try { sa = (string)m.SenderEmailAddress; } catch { }
                                bool match = (!string.IsNullOrEmpty(sn) && sn.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0)
                                          || (!string.IsNullOrEmpty(sa) && sa.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0);
                                if (match)
                                {
                                    string e = SmtpFromSender(m);
                                    if (!string.IsNullOrEmpty(e)) tally[e] = (tally.ContainsKey(e) ? tally[e] : 0) + 1;
                                }
                            }
                        }
                        catch { }
                        try { m = items.GetNext(); } catch { m = null; }
                        n++;
                    }
                }
                catch { }

                string best = ""; int bestN = 0;
                foreach (var kv in tally) if (kv.Value > bestN) { best = kv.Key; bestN = kv.Value; }
                return best;
            }
            catch { return ""; }
        }

        private static string SmtpFromRecipient(dynamic r)
        {
            try { string a = (string)r.Address; if (!string.IsNullOrEmpty(a) && a.IndexOf('@') > 0) return a; } catch { }
            try { return ExchangeSmtp(r.AddressEntry); } catch { }
            return "";
        }

        private static string SmtpFromSender(dynamic mail)
        {
            try
            {
                string t = ""; try { t = (string)mail.SenderEmailType; } catch { }
                if (string.Equals(t, "EX", StringComparison.OrdinalIgnoreCase))
                { try { string e = ExchangeSmtp(mail.Sender); if (!string.IsNullOrEmpty(e)) return e; } catch { } }
                string s = ""; try { s = (string)mail.SenderEmailAddress; } catch { }
                if (!string.IsNullOrEmpty(s) && s.IndexOf('@') > 0) return s;
            }
            catch { }
            return "";
        }

        // Drop placeholder values the model sometimes emits for empty recipient fields.
        private static string StripNone(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            string t = s.Trim();
            string low = t.ToLowerInvariant().Trim('(', ')', '[', ']', '<', '>', '.', ' ');
            if (low == "" || low == "none" || low == "blank" || low == "n/a" || low == "na" ||
                low.StartsWith("leave") || low.StartsWith("no ")) return "";
            return t;
        }

        private string BuildWritePrompt(string instruction, string me, string lang)
        {
            string langLine = string.IsNullOrEmpty(lang)
                ? "Write the email in the same language the description is written in."
                : "Write the ENTIRE email in " + lang + ".";
            string tone = MyToneGuide();
            string toneLine = string.IsNullOrEmpty(tone) ? "" : " IMPORTANT: write in the USER'S OWN VOICE, following this style profile exactly — same greeting, sign-off, vocabulary, recurring phrases, punctuation and level of formality, so it reads as if the user wrote it themselves:\n" + tone + "\n";
            return "Write a complete, professional email based on this description from the user:\n" + instruction +
                   "\n\nBegin with an appropriate greeting and end with a short closing line such as 'Kind regards,' " +
                   "(matching the user's usual sign-off phrase). " + langLine + toneLine + NoSignatureLine +
                   " At the very top, output these header lines, each on its own line:\n" +
                   "To: <main recipient(s)>\n" +
                   "Cc: <anyone the user asked to copy — include ONLY if the description mentions people to Cc, else omit this line>\n" +
                   "Bcc: <anyone the user asked to blind-copy — include ONLY if mentioned, else omit this line>\n" +
                   "Subject: <a short, clear subject>\n" +
                   "Put the main recipient in To and anyone the user says to 'cc'/'copy' in Cc ONLY (never in To); " +
                   "'bcc'/'blind copy' goes in Bcc. " +
                   "For recipients you may write just the person's NAME (e.g. 'Jan Peeters') — do not invent an email " +
                   "address; Outlook will look it up in the address book. Separate multiple recipients with semicolons. " +
                   "Then a blank line, then the email body (greeting by first name, message, closing line). Output only that — no notes.";
        }

        // Outlook appends the user's OWN signature (name, title, company, phone, logo) to every new mail
        // and reply, so the draft must not add one too — otherwise the name/signature appears twice. End
        // at the closing salutation and stop.
        private const string NoSignatureLine =
            " Do NOT add the sender's name after the closing line, and do NOT add any job title, company " +
            "name, phone number, address, website or signature block — Outlook automatically appends the " +
            "user's own signature. Stop right after the closing salutation.";

        public void OnSendLater(object control)
        {
            try
            {
                dynamic app = _app;
                dynamic insp = app.ActiveInspector();
                if (insp == null) { Ui.Notify("Open or start an email first, then use Send Later.", "Axon intelligence"); return; }
                dynamic item = insp.CurrentItem;
                int cls = 0; try { cls = (int)item.Class; } catch { }
                if (item == null || cls != 43) { Ui.Notify("Send Later works on an email you're composing.", "Axon intelligence"); return; }
                DateTime when; bool remind;
                using (var p = new WhenPrompt("Send Later", "When should Axon send this email?", true))
                {
                    if (p.ShowDialog() != DialogResult.OK) return;
                    when = p.When; remind = p.Remind;
                }
                if (when <= DateTime.Now.AddMinutes(1)) { Ui.Notify("Pick a time in the future.", "Axon intelligence"); return; }
                string subj = ""; try { subj = (string)item.Subject; } catch { }
                string to = ""; try { to = (string)item.To; } catch { }
                item.DeferredDeliveryTime = when;
                item.Send();   // moves to the Outbox and is sent automatically at that time
                WriteScheduled(subj, to, when, remind ? 10 : 0);   // dot notifies when it sends (+ heads-up)
                Ui.Notify("Scheduled — Axon will send this at " + when.ToString("ddd dd MMM, HH:mm") +
                          ". It waits in your Outbox until then (keep Outlook connected).", "Axon intelligence");
            }
            catch (Exception ex) { Ui.Notify("Axon error: " + ex.Message, "Axon intelligence"); }
        }

        public void OnFollowUp(object control)
        {
            try
            {
                object m = GetSelectedMail();
                if (m == null) { Ui.Notify("Select an email first.", "Axon intelligence"); return; }
                dynamic mail = m;
                DateTime when;
                using (var p = new WhenPrompt("Follow up", "When should Axon remind you to follow up?"))
                {
                    if (p.ShowDialog() != DialogResult.OK) return;
                    when = p.When;
                }
                try { mail.FlagRequest = "Follow up"; } catch { }
                try { mail.TaskStartDate = DateTime.Now; } catch { }
                try { mail.TaskDueDate = when; } catch { }
                try { mail.ReminderTime = when; mail.ReminderSet = true; } catch { }
                try { mail.FlagStatus = 2; } catch { }   // olFlagMarked (visual marker in Outlook)
                try { mail.Save(); } catch { }
                WriteFollowUp(mail, when);   // the dot pops a reliable reminder at 'when' (and skips it if replied)
                Ui.Notify("Follow-up set for " + when.ToString("ddd dd MMM, HH:mm") +
                          ". Axon will remind you then — unless they've already replied.", "Axon intelligence");
            }
            catch (Exception ex) { Ui.Notify("Axon error: " + ex.Message, "Axon intelligence"); }
        }

        // Record a follow-up so the always-running dot can pop a reliable reminder at the due time
        // (and skip it if a reply arrived). Shared file: %APPDATA%\AxonOutlook\followups.json.
        private void WriteFollowUp(dynamic mail, DateTime due)
        {
            try
            {
                string id = ""; try { id = (string)mail.EntryID; } catch { }
                string subject = ""; try { subject = (string)mail.Subject; } catch { }
                string who = ""; try { who = (string)mail.SenderName; } catch { }
                if (string.IsNullOrEmpty(who)) { try { who = (string)mail.To; } catch { } }
                string topic = ""; try { topic = (string)mail.ConversationTopic; } catch { }

                string dir = AxonDataDir();
                Directory.CreateDirectory(dir);
                string path = Path.Combine(dir, "followups.json");
                var js = new System.Web.Script.Serialization.JavaScriptSerializer();
                var items = new System.Collections.Generic.List<object>();
                try { if (File.Exists(path)) { var arr = js.DeserializeObject(File.ReadAllText(path)) as object[]; if (arr != null) items.AddRange(arr); } }
                catch { }
                items.Add(new System.Collections.Generic.Dictionary<string, object> {
                    { "id", id }, { "subject", subject }, { "who", who }, { "topic", topic },
                    { "due", due.ToString("yyyy-MM-ddTHH:mm:ss") },
                    { "created", DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss") },
                    { "notified", false },
                });
                File.WriteAllText(path, js.Serialize(items));
            }
            catch { }
        }

        // Record a scheduled ('Send Later') send so the always-running dot can notify the user when
        // its time arrives. Shared file: %APPDATA%\AxonOutlook\scheduled.json (append-only).
        private void WriteScheduled(string subject, string to, DateTime when, int remindBefore)
        {
            try
            {
                string dir = AxonDataDir();
                Directory.CreateDirectory(dir);
                string path = Path.Combine(dir, "scheduled.json");
                var js = new System.Web.Script.Serialization.JavaScriptSerializer();
                var items = new System.Collections.Generic.List<object>();
                try { if (File.Exists(path)) { var arr = js.DeserializeObject(File.ReadAllText(path)) as object[]; if (arr != null) items.AddRange(arr); } }
                catch { }
                var entry = new System.Collections.Generic.Dictionary<string, object> {
                    { "subject", subject ?? "" }, { "to", to ?? "" },
                    { "when", when.ToString("yyyy-MM-ddTHH:mm:ss") },
                    { "created", DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss") },
                };
                if (remindBefore > 0)
                    entry["remind_at"] = when.AddMinutes(-remindBefore).ToString("yyyy-MM-ddTHH:mm:ss");
                items.Add(entry);
                File.WriteAllText(path, js.Serialize(items));
            }
            catch { }
        }
    }
}
