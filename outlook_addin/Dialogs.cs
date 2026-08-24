// Axon Outlook add-in — WinForms dialog classes (moved out of AxonAddin.cs to keep it lean).
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
    // Move dialog — suggested folders, create-new-folder, and a searchable list. Opens instantly;
    // suggestions fill in async.
    // Settings dialog for the standalone add-in (no dot): email-archive rules + writing tone.
    internal class SettingsForm : Form
    {
        private readonly Connect _owner;
        private TextBox _tone;
        private ComboBox _mode;
        private Button _learn;
        private Button _relearn;
        private Label _folderHint;

        public SettingsForm(Connect owner)
        {
            _owner = owner;
            Text = "Axon intelligence — Settings";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false; MinimizeBox = false; ShowInTaskbar = false;
            Font = new System.Drawing.Font("Segoe UI", 9.5F);
            BuildPremiumUi();
            LoadAll();
        }

        private static readonly System.Drawing.Color SettingsBg = System.Drawing.Color.White;
        private static readonly System.Drawing.Color SettingsSurface = System.Drawing.Color.White;
        private static readonly System.Drawing.Color SettingsBorder = Ui.Line;
        private static readonly System.Drawing.Color SettingsText = Ui.Ink;
        private static readonly System.Drawing.Color SettingsMuted = Ui.Muted;

        private void BuildPremiumUi()
        {
            ClientSize = new System.Drawing.Size(560, 590);   // +120 for the Folders card below Writing tone
            BackColor = SettingsBg;
            int W = ClientSize.Width, x = 24, lw = W - 48;

            Controls.Add(SettingsTitle("Axon settings", x, 20, lw));
            Controls.Add(SettingsHint("Settings for the Outlook add-in.", x, 54, lw, 22));

            // Email archive: the folders and country codes are built in, so the only choice left here is
            // what gets saved when you download an email.
            var archive = SettingsCard(x, 92, lw, 116);
            Controls.Add(archive);
            archive.Controls.Add(SettingsSectionTitle("Saving emails", 18, 16, lw - 36));
            archive.Controls.Add(SettingsHint("When you download an email, save the message, its attachments, or both.", 18, 42, lw - 36, 28));
            archive.Controls.Add(SettingsFieldLabel("Save", 18, 78, 46));
            _mode = new ComboBox { Left = 66, Top = 74, Width = 260, Height = 30, DropDownStyle = ComboBoxStyle.DropDownList, FlatStyle = FlatStyle.Flat, BackColor = System.Drawing.Color.White, ForeColor = SettingsText };
            _mode.Items.AddRange(new object[] { "Both (email + attachments)", "Email (.msg)", "Attachments only" });
            archive.Controls.Add(_mode);

            var writing = SettingsCard(x, 224, lw, 170);
            Controls.Add(writing);
            writing.Controls.Add(SettingsSectionTitle("Writing tone", 18, 16, lw - 36));
            writing.Controls.Add(SettingsHint("How Axon writes replies and drafts. Type it, or learn it from your Sent mail.", 18, 42, lw - 36, 28));
            _tone = SettingsArea(18, 76, lw - 36, 62);
            writing.Controls.Add(_tone);
            _learn = SettingsGhostButton("Learn from my Sent emails", 18, 140, 210, 28);
            _learn.Click += (o, e) => LearnTone();
            writing.Controls.Add(_learn);

            // Move reads its folder list from disk and refreshes it behind the dialog, so this is only for
            // "I just made a folder and want it in the list now" rather than something anyone must remember.
            var foldersCard = SettingsCard(x, 404, lw, 104);
            Controls.Add(foldersCard);
            foldersCard.Controls.Add(SettingsSectionTitle("Folders", 18, 16, lw - 36));
            _folderHint = SettingsHint("Move remembers your folders so the list opens instantly. It refreshes itself; use this if you just made one.", 18, 42, lw - 36, 30);
            foldersCard.Controls.Add(_folderHint);
            _relearn = SettingsGhostButton("Learn my folders", 18, 70, 210, 28);
            _relearn.Click += (o, e) => RelearnFolders();
            foldersCard.Controls.Add(_relearn);

            Controls.Add(SettingsHint("Changes apply after saving.", x, ClientSize.Height - 58, 220, 20));
            var save = SettingsPrimaryButton("Save", W - 24 - 96, ClientSize.Height - 54, 96, 34);
            save.Click += (o, e) => { SaveAll(); DialogResult = DialogResult.OK; Close(); };
            var cancel = SettingsGhostButton("Cancel", save.Left - 10 - 96, ClientSize.Height - 54, 96, 34);
            cancel.Click += (o, e) => { DialogResult = DialogResult.Cancel; Close(); };
            Controls.Add(cancel);
            Controls.Add(save);
            // No AcceptButton: on a settings screen Enter should never save+close while you're still
            // typing (especially in the multiline list fields). Save is an explicit button click.
            // Escape still cancels, which is the expected, safe shortcut.
            CancelButton = cancel;
        }

        private static Panel SettingsCard(int x, int y, int w, int h)
        { return new Panel { Left = x, Top = y, Width = w, Height = h, BackColor = SettingsSurface, BorderStyle = BorderStyle.FixedSingle }; }
        private static Label SettingsTitle(string t, int x, int y, int w)
        { return new Label { Text = t, Left = x, Top = y, Width = w, Height = 34, ForeColor = SettingsText, Font = new System.Drawing.Font("Segoe UI", 16F, System.Drawing.FontStyle.Bold) }; }
        private static Label SettingsSectionTitle(string t, int x, int y, int w)
        { return new Label { Text = t, Left = x, Top = y, Width = w, Height = 24, ForeColor = SettingsText, Font = new System.Drawing.Font("Segoe UI", 11F, System.Drawing.FontStyle.Bold) }; }
        private static Label SettingsFieldLabel(string t, int x, int y, int w)
        { return new Label { Text = t, Left = x, Top = y, Width = w, Height = 22, ForeColor = SettingsText, Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold) }; }
        private static Label SettingsHint(string t, int x, int y, int w, int h)
        { return new Label { Text = t, Left = x, Top = y, Width = w, Height = h, ForeColor = SettingsMuted, AutoEllipsis = true, AutoSize = false }; }
        private static TextBox SettingsField(int x, int y, int w)
        { return new TextBox { Left = x, Top = y, Width = w, Height = 28, BorderStyle = BorderStyle.FixedSingle, BackColor = System.Drawing.Color.White, ForeColor = SettingsText }; }
        private static TextBox SettingsArea(int x, int y, int w, int h)
        // AcceptsReturn: these are one-entry-per-line lists, so Enter must add a new line here, not
        // trigger the form's default button and close the dialog.
        { return new TextBox { Left = x, Top = y, Width = w, Height = h, Multiline = true, AcceptsReturn = true, ScrollBars = ScrollBars.Vertical, BorderStyle = BorderStyle.FixedSingle, BackColor = System.Drawing.Color.White, ForeColor = SettingsText }; }
        private static Button SettingsPrimaryButton(string text, int x, int y, int w, int h)
        {
            var b = new Button { Text = text, Left = x, Top = y, Width = w, Height = h, FlatStyle = FlatStyle.Flat, BackColor = System.Drawing.Color.FromArgb(17, 17, 17), ForeColor = System.Drawing.Color.White, Font = new System.Drawing.Font("Segoe UI", 9.5F, System.Drawing.FontStyle.Bold), Cursor = Cursors.Hand };
            b.FlatAppearance.BorderSize = 0;
            b.FlatAppearance.MouseOverBackColor = System.Drawing.Color.FromArgb(48, 45, 40);
            return b;
        }
        private static Button SettingsGhostButton(string text, int x, int y, int w, int h)
        {
            var b = new Button { Text = text, Left = x, Top = y, Width = w, Height = h, FlatStyle = FlatStyle.Flat, BackColor = SettingsSurface, ForeColor = SettingsText, Font = new System.Drawing.Font("Segoe UI", 9.5F), Cursor = Cursors.Hand };
            b.FlatAppearance.BorderColor = SettingsBorder;
            b.FlatAppearance.BorderSize = 1;
            return b;
        }
        private Button SettingsBrowseBtn(TextBox target, int x, int y)
        {
            var b = SettingsGhostButton("Browse...", x, y, 94, 28);
            b.Click += (o, e) => { using (var fb = new FolderBrowserDialog()) { if (fb.ShowDialog() == DialogResult.OK) target.Text = fb.SelectedPath; } };
            return b;
        }
        private Button SettingsAddFolderBtn(TextBox target, int x, int y)
        {
            var b = SettingsGhostButton("Add...", x, y, 74, 28);
            b.Click += (o, e) =>
            {
                using (var fb = new FolderBrowserDialog())
                {
                    if (fb.ShowDialog() == DialogResult.OK)
                    {
                        string cur = target.Text ?? "";
                        if (cur.Length > 0 && !cur.EndsWith("\n")) cur += "\r\n";
                        target.Text = cur + "Label = " + fb.SelectedPath;
                    }
                }
            };
            return b;
        }

        private static string Dir() { return Connect.AxonDataDir(); }
        private static string Get(System.Collections.Generic.Dictionary<string, object> d, string k) { object v; return (d != null && d.TryGetValue(k, out v) && v != null) ? v.ToString() : ""; }
        private static TextBox Tb(int x, int y, int w) { return new TextBox { Left = x, Top = y, Width = w, BorderStyle = BorderStyle.FixedSingle }; }
        private static Label Lbl(string t, int x, int y) { return new Label { Text = t, Left = x, Top = y, Width = 510, Height = 16, ForeColor = System.Drawing.Color.FromArgb(80, 80, 80) }; }
        private static Label Cap(string t, int x, int y) { return new Label { Text = t, Left = x, Top = y, Width = 300, Height = 16, ForeColor = System.Drawing.Color.FromArgb(120, 120, 120), Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Bold) }; }
        private Button BrowseBtn(TextBox target, int x, int y)
        {
            var b = new Button { Text = "Browse…", Left = x, Top = y - 1, Width = 84, Height = 24, FlatStyle = FlatStyle.System };
            b.Click += (o, e) => { using (var fb = new FolderBrowserDialog()) { if (fb.ShowDialog() == DialogResult.OK) target.Text = fb.SelectedPath; } };
            return b;
        }

        private void LoadAll()
        {
            try
            {
                string p = Path.Combine(Dir(), "archive.json");
                if (File.Exists(p))
                {
                    var js = new System.Web.Script.Serialization.JavaScriptSerializer();
                    var d = js.DeserializeObject(File.ReadAllText(p)) as System.Collections.Generic.Dictionary<string, object>;
                    if (d != null)
                    {
                        string mode = Get(d, "save_mode").ToLowerInvariant();
                        _mode.SelectedIndex = mode == "email" ? 1 : (mode == "attachments" ? 2 : 0);
                    }
                }
            }
            catch { }
            if (_mode.SelectedIndex < 0) _mode.SelectedIndex = 0;
            try { string t = Path.Combine(Dir(), "tone.txt"); if (File.Exists(t)) _tone.Text = File.ReadAllText(t).Trim(); } catch { }
        }

        private void SaveAll()
        {
            try
            {
                Directory.CreateDirectory(Dir());
                string p = Path.Combine(Dir(), "archive.json");
                var js = new System.Web.Script.Serialization.JavaScriptSerializer();
                // Archive folders + country codes are built in now and no longer edited here, so keep any
                // existing archive.json values (in case someone set an override) and only update save_mode.
                var d = new System.Collections.Generic.Dictionary<string, object>();
                try { if (File.Exists(p)) d = js.DeserializeObject(File.ReadAllText(p)) as System.Collections.Generic.Dictionary<string, object> ?? d; } catch { }
                d["save_mode"] = new[] { "both", "email", "attachments" }[_mode.SelectedIndex < 0 ? 0 : _mode.SelectedIndex];
                File.WriteAllText(p, js.Serialize(d));
                File.WriteAllText(Path.Combine(Dir(), "tone.txt"), (_tone.Text ?? "").Trim());
            }
            catch (Exception ex) { Ui.Notify("Couldn't save settings: " + ex.Message, "Axon intelligence"); }
        }

        // Walking the stores takes seconds on a mailbox with shared boxes attached, so it runs off the UI
        // thread and the button says what it is doing rather than freezing the settings window.
        private void RelearnFolders()
        {
            _relearn.Enabled = false;
            string was = _relearn.Text;
            _relearn.Text = "Reading your folders…";
            var t = new System.Threading.Thread(() =>
            {
                string msg;
                try { msg = _owner.RelearnFolders(); }
                catch (Exception ex) { msg = "Couldn't read the folders: " + ex.Message; }
                try
                {
                    BeginInvoke(new Action(() =>
                    {
                        _folderHint.Text = msg;
                        _relearn.Text = was;
                        _relearn.Enabled = true;
                    }));
                }
                catch { }
            });
            t.IsBackground = true;
            t.SetApartmentState(System.Threading.ApartmentState.STA);
            t.Start();
        }

        private void LearnTone()
        {
            _learn.Enabled = false; _learn.Text = "Reading your Sent mail…";
            var t = new System.Threading.Thread(() =>
            {
                string tone = "";
                try { tone = _owner.LearnToneFromSent(); } catch { }
                try
                {
                    BeginInvoke(new Action(() =>
                    {
                        if (!string.IsNullOrEmpty(tone)) _tone.Text = tone;
                        else Ui.Notify("Couldn't read your Sent items (is Outlook set up?).", "Axon intelligence");
                        _learn.Enabled = true; _learn.Text = "Learn from my Sent emails";
                    }));
                }
                catch { }
            });
            t.IsBackground = true; t.Start();
        }
    }

    // Resizable summary dialog: pick the language (re-summarizes on change), and a Reply button to
    // jump straight into composing a reply to the same email.
    internal class SummaryDialog : Form
    {
        private readonly RichTextBox _box;
        private readonly ComboBox _lang;
        private readonly Func<string, string> _summarize;   // language -> summary text
        public bool WantsReply { get; private set; }

        public SummaryDialog(string subject, Func<string, string> summarize)
        {
            _summarize = summarize;
            Text = "Axon intelligence — Summary";
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new System.Drawing.Size(540, 460);
            MinimumSize = new System.Drawing.Size(440, 320);
            MaximizeBox = true; MinimizeBox = false; ShowIcon = false;
            BackColor = System.Drawing.Color.White;

            var langLbl = new Label { Left = 16, Top = 16, Width = 80, Height = 24, Text = "Summary in:",
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft, Anchor = AnchorStyles.Top | AnchorStyles.Left };
            _lang = new ComboBox { Left = 100, Top = 13, Width = 200, Height = 24,
                DropDownStyle = ComboBoxStyle.DropDownList, Anchor = AnchorStyles.Top | AnchorStyles.Left };
            _lang.Items.AddRange(new object[] { "Auto (match the email)", "English", "Nederlands (Dutch)", "Français (French)" });
            _lang.SelectedIndex = 0;
            _lang.SelectedIndexChanged += (o, e) => Regenerate();
            _box = new RichTextBox { Left = 16, Top = 48, Width = ClientSize.Width - 32, Height = ClientSize.Height - 100,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                ReadOnly = true, ScrollBars = RichTextBoxScrollBars.Vertical, BorderStyle = BorderStyle.FixedSingle,
                Text = "Summarizing…", Font = new System.Drawing.Font("Segoe UI", 9.75f), BackColor = System.Drawing.Color.White };
            var reply = new Button { Text = "Reply", Width = 90, Height = 28,
                Left = ClientSize.Width - 200, Top = ClientSize.Height - 40, Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            var close = new Button { Text = "Close", Width = 90, Height = 28, DialogResult = DialogResult.Cancel,
                Left = ClientSize.Width - 104, Top = ClientSize.Height - 40, Anchor = AnchorStyles.Bottom | AnchorStyles.Right };
            reply.Click += (o, e) => { WantsReply = true; DialogResult = DialogResult.OK; Close(); };
            Controls.AddRange(new Control[] { langLbl, _lang, _box, reply, close });
            CancelButton = close;
            Shown += (o, e) => Regenerate();
        }

        private static string LangCode(string display)
        {
            if (display.StartsWith("Eng")) return "English";
            if (display.StartsWith("Ned")) return "Dutch";
            if (display.StartsWith("Fr")) return "French";
            return "";   // Auto -> match the email
        }

        private void Regenerate()
        {
            string lang = LangCode(_lang.SelectedItem.ToString());
            SetText("Summarizing…");
            var th = new System.Threading.Thread(() =>
            {
                string s = _summarize(lang);
                SetText(string.IsNullOrEmpty(s) ? "Couldn't summarize (model unavailable)." : s);
            });
            th.IsBackground = true;
            th.Start();
        }

        private void SetText(string s)
        {
            if (IsDisposed) return;
            if (InvokeRequired) { try { BeginInvoke(new Action(() => SetText(s))); } catch { } return; }
            string text = SpaceBullets((s ?? "").Replace("**", "").Replace("\r\n", "\n")).Replace("\n", "\r\n");
            _box.Text = text;
            BoldSummaryHeading("Gist:");
            BoldSummaryHeading("Conversation:");   // only present when the email is a thread
            BoldSummaryHeading("Key points:");
            BoldSummaryHeading("Action:");
            _box.Select(0, 0);
        }

        // Give the points room to breathe: a blank line BETWEEN consecutive '- ' bullets only, so the
        // gaps that already precede a section heading don't get doubled.
        private static string SpaceBullets(string s)
        {
            var lines = (s ?? "").Split('\n');
            var sb = new System.Text.StringBuilder();
            for (int i = 0; i < lines.Length; i++)
            {
                sb.Append(lines[i]).Append('\n');
                bool bullet = lines[i].TrimStart().StartsWith("- ");
                bool nextBullet = i + 1 < lines.Length && lines[i + 1].TrimStart().StartsWith("- ");
                if (bullet && nextBullet) sb.Append('\n');
            }
            return sb.ToString().TrimEnd('\n');
        }

        private void BoldSummaryHeading(string heading)
        {
            int start = _box.Text.IndexOf(heading, StringComparison.OrdinalIgnoreCase);
            if (start < 0) return;
            _box.Select(start, heading.Length);
            _box.SelectionFont = new System.Drawing.Font(_box.Font, System.Drawing.FontStyle.Bold);
        }
    }

    // Asks the user HOW to reply + in which language, then drafts it (off the UI thread).
    internal class ReplyPrompt : Form
    {
        private readonly TextBox _box;
        private readonly ComboBox _lang;
        private readonly Func<string, string, string> _drafter;   // (instruction, language) -> draft
        private readonly Button _draftBtn;
        private readonly Label _status;
        public string Draft { get; private set; }

        public ReplyPrompt(string subject, Func<string, string, string> drafter)
        {
            _drafter = drafter;
            Text = "Axon intelligence — Reply";
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new System.Drawing.Size(500, 318);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false; ShowIcon = false;
            BackColor = System.Drawing.Color.White;

            var t1 = new Label { Left = 20, Top = 18, Width = 460, Height = 20, Text = "Reply to: " + subject,
                Font = new System.Drawing.Font("Segoe UI", 10f, System.Drawing.FontStyle.Bold), AutoEllipsis = true };
            var t2 = new Label { Left = 20, Top = 48, Width = 460, Height = 20,
                Text = "Tell Axon how to reply (or leave blank for a professional reply):",
                ForeColor = System.Drawing.Color.FromArgb(110, 110, 120) };
            _box = new TextBox { Left = 20, Top = 74, Width = 460, Height = 110, Multiline = true,
                ScrollBars = ScrollBars.Vertical, BorderStyle = BorderStyle.FixedSingle,
                Font = new System.Drawing.Font("Segoe UI", 9.75f) };
            var langLbl = new Label { Left = 20, Top = 200, Width = 60, Height = 22, Text = "Reply in:",
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft };
            _lang = new ComboBox { Left = 82, Top = 197, Width = 200, Height = 24, DropDownStyle = ComboBoxStyle.DropDownList };
            _lang.Items.AddRange(new object[] { "Auto (match the email)", "English", "Nederlands (Dutch)", "Français (French)" });
            _lang.SelectedIndex = 0;
            _status = new Label { Left = 20, Top = 240, Width = 260, Height = 20, Text = "",
                ForeColor = System.Drawing.Color.FromArgb(110, 110, 120) };
            _draftBtn = new Button { Text = "Draft reply", Left = 300, Top = 272, Width = 110, Height = 28 };
            var cancel = new Button { Text = "Cancel", Left = 415, Top = 272, Width = 65, Height = 28, DialogResult = DialogResult.Cancel };
            _draftBtn.Click += (o, e) => DoDraft();
            Controls.AddRange(new Control[] { t1, t2, _box, langLbl, _lang, _status, _draftBtn, cancel });
            AcceptButton = _draftBtn;
            CancelButton = cancel;
        }

        private static string LangCode(string display)
        {
            if (display.StartsWith("Eng")) return "English";
            if (display.StartsWith("Ned")) return "Dutch";
            if (display.StartsWith("Fr")) return "French";
            return "";   // Auto -> match the email
        }

        private void DoDraft()
        {
            _draftBtn.Enabled = false;
            _status.Text = "Drafting…";
            string instr = _box.Text;
            string lang = LangCode(_lang.SelectedItem.ToString());
            var th = new System.Threading.Thread(() =>
            {
                string d = _drafter(instr, lang);
                try { BeginInvoke(new Action(() => { Draft = d; DialogResult = DialogResult.OK; Close(); })); } catch { }
            });
            th.IsBackground = true;
            th.Start();
        }
    }

    // Asks the user WHAT the email should be about (+ language), then Axon writes it (off the UI thread).
    internal class WritePrompt : Form
    {
        private readonly TextBox _box;
        private readonly ComboBox _lang;
        private readonly Func<string, string, string> _writer;   // (description, language) -> email text
        private readonly Button _writeBtn;
        private readonly Label _status;
        public string Draft { get; private set; }

        public WritePrompt(Func<string, string, string> writer)
        {
            _writer = writer;
            Text = "Axon intelligence — Write email";
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new System.Drawing.Size(500, 322);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false; ShowIcon = false;
            BackColor = System.Drawing.Color.White;

            var t1 = new Label { Left = 20, Top = 16, Width = 460, Height = 20, Text = "Write an email with Axon",
                Font = new System.Drawing.Font("Segoe UI", 10f, System.Drawing.FontStyle.Bold) };
            var t2 = new Label { Left = 20, Top = 44, Width = 460, Height = 36,
                Text = "What should this email be about? Who it's to, the purpose, and any key points:",
                ForeColor = System.Drawing.Color.FromArgb(110, 110, 120) };
            _box = new TextBox { Left = 20, Top = 84, Width = 460, Height = 120, Multiline = true,
                ScrollBars = ScrollBars.Vertical, BorderStyle = BorderStyle.FixedSingle,
                Font = new System.Drawing.Font("Segoe UI", 9.75f) };
            var langLbl = new Label { Left = 20, Top = 218, Width = 60, Height = 22, Text = "Write in:",
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft };
            _lang = new ComboBox { Left = 82, Top = 215, Width = 210, Height = 24, DropDownStyle = ComboBoxStyle.DropDownList };
            _lang.Items.AddRange(new object[] { "Auto (match my description)", "English", "Nederlands (Dutch)", "Français (French)" });
            _lang.SelectedIndex = 0;
            _status = new Label { Left = 20, Top = 258, Width = 260, Height = 20, Text = "",
                ForeColor = System.Drawing.Color.FromArgb(110, 110, 120) };
            _writeBtn = new Button { Text = "Write email", Left = 300, Top = 286, Width = 110, Height = 28 };
            var cancel = new Button { Text = "Cancel", Left = 415, Top = 286, Width = 65, Height = 28, DialogResult = DialogResult.Cancel };
            _writeBtn.Click += (o, e) => DoWrite();
            Controls.AddRange(new Control[] { t1, t2, _box, langLbl, _lang, _status, _writeBtn, cancel });
            AcceptButton = _writeBtn;
            CancelButton = cancel;
        }

        private static string LangCode(string display)
        {
            if (display.StartsWith("Eng")) return "English";
            if (display.StartsWith("Ned")) return "Dutch";
            if (display.StartsWith("Fr")) return "French";
            return "";   // Auto
        }

        private void DoWrite()
        {
            if (_box.Text.Trim().Length == 0) { _status.Text = "Describe the email first."; return; }
            _writeBtn.Enabled = false;
            _status.Text = "Writing…";
            string instr = _box.Text;
            string lang = LangCode(_lang.SelectedItem.ToString());
            var th = new System.Threading.Thread(() =>
            {
                string d = _writer(instr, lang);
                try { BeginInvoke(new Action(() => { Draft = d; DialogResult = DialogResult.OK; Close(); })); } catch { }
            });
            th.IsBackground = true;
            th.Start();
        }
    }

    // A friendly "when?" picker: one-click preset buttons (the common case) plus a calendar +
    // time list for a custom time. Used by Send Later and Follow up.
    internal class WhenPrompt : Form
    {
        private static readonly string[] Presets =
            { "In 1 hour", "This evening (6 PM)", "Tomorrow morning (8 AM)", "In 3 days", "Next Monday (9 AM)" };
        private readonly DateTimePicker _date;
        private readonly ComboBox _time;
        private readonly CheckBox _remind;
        public DateTime When { get; private set; }
        public bool Remind { get; private set; }

        public WhenPrompt(string title, string prompt, bool showRemind = false)
        {
            Text = "Axon intelligence — " + title;
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false; ShowIcon = false;
            BackColor = System.Drawing.Color.White;
            int W = 380;

            var lbl = new Label { Left = 20, Top = 16, Width = W - 40, Height = 22, Text = prompt,
                Font = new System.Drawing.Font("Segoe UI", 10f, System.Drawing.FontStyle.Bold) };
            Controls.Add(lbl);

            int y = 50;
            for (int i = 0; i < Presets.Length; i++)
            {
                int idx = i;
                var b = new Button { Left = 20, Top = y, Width = W - 40, Height = 32, Text = "   " + Presets[i],
                    TextAlign = System.Drawing.ContentAlignment.MiddleLeft, FlatStyle = FlatStyle.System,
                    Font = new System.Drawing.Font("Segoe UI", 9.75f), Cursor = Cursors.Hand };
                b.Click += (o, e) => { When = Resolve(idx); Remind = _remind != null && _remind.Checked; DialogResult = DialogResult.OK; Close(); };
                Controls.Add(b);
                y += 36;
            }

            y += 8;
            var orLbl = new Label { Left = 20, Top = y + 3, Width = 56, Height = 22, Text = "Or pick:",
                ForeColor = System.Drawing.Color.FromArgb(110, 110, 120),
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft };
            DateTime soon = DateTime.Now.AddHours(1);
            _date = new DateTimePicker { Left = 80, Top = y, Width = 158, Format = DateTimePickerFormat.Short };
            _date.Value = soon.Date;   // default to TODAY — same-day sends are the common case
            _time = new ComboBox { Left = 246, Top = y, Width = 114, DropDownStyle = ComboBoxStyle.DropDown };
            for (int h = 7; h <= 19; h++) { _time.Items.Add(h.ToString("00") + ":00"); _time.Items.Add(h.ToString("00") + ":30"); }
            _time.Text = soon.ToString("HH:mm");   // ~1 hour from now, so it's already in the future
            Controls.AddRange(new Control[] { orLbl, _date, _time });
            y += 44;

            if (showRemind)
            {
                _remind = new CheckBox { Left = 20, Top = y, Width = W - 40, Height = 22,
                    Text = "Remind me ~10 min before it sends",
                    Font = new System.Drawing.Font("Segoe UI", 9.25f) };
                Controls.Add(_remind);
                y += 30;
            }

            var set = new Button { Text = "Set custom", Left = W - 20 - 95 - 85, Top = y, Width = 95, Height = 28 };
            var cancel = new Button { Text = "Cancel", Left = W - 20 - 80, Top = y, Width = 80, Height = 28, DialogResult = DialogResult.Cancel };
            set.Click += (o, e) =>
            {
                int hh = 9, mm = 0;
                var parts = (_time.Text ?? "09:00").Split(':');
                int.TryParse(parts.Length > 0 ? parts[0] : "9", out hh);
                if (parts.Length > 1) int.TryParse(parts[1], out mm);
                When = _date.Value.Date.AddHours(hh).AddMinutes(mm);
                Remind = _remind != null && _remind.Checked;
                DialogResult = DialogResult.OK; Close();
            };
            Controls.AddRange(new Control[] { set, cancel });
            ClientSize = new System.Drawing.Size(W, y + 44);
            CancelButton = cancel;
        }

        private static DateTime Resolve(int i)
        {
            DateTime n = DateTime.Now;
            switch (i)
            {
                case 0: return n.AddHours(1);
                case 1: { DateTime e = n.Date.AddHours(18); return e > n ? e : n.AddHours(1); }
                case 2: return n.Date.AddDays(1).AddHours(8);
                case 3: return n.Date.AddDays(3).AddHours(9);
                case 4: { int add = (8 - (int)n.DayOfWeek) % 7; if (add == 0) add = 7; return n.Date.AddDays(add).AddHours(9); }
                default: return n.AddHours(1);
            }
        }
    }

    // Asks the user WHEN to schedule (any timeframe, free text) + the duration, before Axon drafts the invite.
    internal class SchedulePrompt : Form
    {
        private static readonly int[] Durations = { 15, 30, 45, 60, 90, 120 };
        private readonly TextBox _when;
        private readonly ComboBox _dur;
        public string When { get; private set; }
        public int DurationMinutes { get; private set; }

        public SchedulePrompt(string subject)
        {
            Text = "Axon intelligence — Schedule";
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new System.Drawing.Size(500, 250);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false; ShowIcon = false;
            BackColor = System.Drawing.Color.White;

            var t1 = new Label { Left = 20, Top = 18, Width = 460, Height = 20, Text = "Meeting about: " + subject,
                Font = new System.Drawing.Font("Segoe UI", 10f, System.Drawing.FontStyle.Bold), AutoEllipsis = true };
            var t2 = new Label { Left = 20, Top = 48, Width = 460, Height = 36,
                Text = "When? e.g. \"next Tuesday 2pm\", \"week of 15 Sept\", \"in 3 months\" — or leave blank to suggest the soonest free time:",
                ForeColor = System.Drawing.Color.FromArgb(110, 110, 120) };
            _when = new TextBox { Left = 20, Top = 88, Width = 460, Height = 26, BorderStyle = BorderStyle.FixedSingle,
                Font = new System.Drawing.Font("Segoe UI", 9.75f) };
            var durLbl = new Label { Left = 20, Top = 132, Width = 64, Height = 22, Text = "Duration:",
                TextAlign = System.Drawing.ContentAlignment.MiddleLeft };
            _dur = new ComboBox { Left = 86, Top = 129, Width = 160, Height = 24, DropDownStyle = ComboBoxStyle.DropDownList };
            _dur.Items.AddRange(new object[] { "15 minutes", "30 minutes", "45 minutes", "1 hour", "1.5 hours", "2 hours" });
            _dur.SelectedIndex = 1;   // 30 minutes
            var create = new Button { Text = "Create invite", Left = 300, Top = 200, Width = 110, Height = 28 };
            var cancel = new Button { Text = "Cancel", Left = 415, Top = 200, Width = 65, Height = 28, DialogResult = DialogResult.Cancel };
            create.Click += (o, e) =>
            {
                When = _when.Text.Trim();
                int i = _dur.SelectedIndex;
                DurationMinutes = (i >= 0 && i < Durations.Length) ? Durations[i] : 30;
                DialogResult = DialogResult.OK; Close();
            };
            Controls.AddRange(new Control[] { t1, t2, _when, durLbl, _dur, create, cancel });
            AcceptButton = create;
            CancelButton = cancel;
        }
    }

    internal class FolderPicker : Form
    {
        public string Chosen;
        public string CreateFolder;
        private readonly ListBox _list;
        private readonly TextBox _search;
        private readonly TextBox _newFolderBox;
        private string[] _all;
        private readonly Panel _suggPanel;

        public FolderPicker(string subject, string[] folders, string winTitle = "Move email",
            string headTitle = "Move this email to a folder", string createVerb = "Create && Move",
            string primaryVerb = "Move", string cancelVerb = "Keep in Inbox")
        {
            _all = folders ?? new string[0];
            Text = "Axon intelligence — " + winTitle;
            ClientSize = new System.Drawing.Size(500, 648);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false; ShowInTaskbar = false;
            BackColor = System.Drawing.Color.White;
            Font = new System.Drawing.Font("Segoe UI", 9.5F);
            int W = ClientSize.Width, H = ClientSize.Height;

            Controls.Add(Ui.Title(headTitle, 22, 18, W - 44));
            Controls.Add(Ui.Sub(subject ?? "(no subject)", 22, 46, W - 44));

            _suggPanel = new Panel { Left = 20, Top = 74, Width = W - 40, Height = 176 };
            _suggPanel.Controls.Add(Ui.Hint("Finding best matches…", 2, 6, W - 60, 30));
            Controls.Add(_suggPanel);

            Controls.Add(Ui.Caption("OR CREATE A NEW FOLDER", 22, 256, 300));
            _newFolderBox = new TextBox { Left = 22, Top = 274, Width = W - 44 - 8 - 148, BorderStyle = BorderStyle.FixedSingle };
            Controls.Add(_newFolderBox);
            var createBtn = Ui.Accent(createVerb, W - 22 - 148, 272, 148, 28);
            createBtn.Click += (o, e) =>
            {
                string nm = _newFolderBox.Text.Trim();
                if (nm.Length == 0) { Ui.Notify("Type a name for the new folder.", "Axon intelligence"); return; }
                CreateFolder = nm; DialogResult = DialogResult.OK; Close();
            };
            Controls.Add(createBtn);

            Controls.Add(Ui.Caption("OR PICK AN EXISTING FOLDER", 22, 312, 320));
            _search = new TextBox { Left = 22, Top = 330, Width = W - 44, BorderStyle = BorderStyle.FixedSingle };
            _search.TextChanged += (o, e) => Filter();
            Controls.Add(_search);
            _list = new ListBox
            {
                Left = 22, Top = 360, Width = W - 44, Height = H - 360 - 66,
                BorderStyle = BorderStyle.FixedSingle, DrawMode = DrawMode.OwnerDrawFixed, ItemHeight = 46, IntegralHeight = false
            };
            _list.DrawItem += (o, e) => Ui.DrawFolderItem(_list, e);
            _list.DoubleClick += (o, e) => PickFromList();
            Controls.Add(_list);
            Filter();

            var keep = Ui.Subtle(cancelVerb, W - 22 - 130, H - 50, 130, 34);
            keep.DialogResult = DialogResult.Cancel;
            Controls.Add(keep);
            CancelButton = keep;
            var move = Ui.Accent(primaryVerb, keep.Left - 8 - 104, H - 50, 104, 34);
            move.Click += (o, e) => PickFromList();
            Controls.Add(move);
        }

        public void SetFolders(string[] folders)
        {
            if (IsDisposed || Disposing) return;
            if (InvokeRequired) { try { BeginInvoke(new Action(() => SetFolders(folders))); } catch { } return; }
            _all = folders ?? new string[0];
            Filter();
        }

        // Back-compat: the Move feature calls this without reasons.
        public void SetSuggestions(string[] matches, string newFolder) { SetSuggestions(matches, null, newFolder); }

        public void SetSuggestions(string[] matches, string[] reasons, string newFolder)
        {
            if (IsDisposed || Disposing) return;
            if (InvokeRequired) { try { BeginInvoke(new Action(() => SetSuggestions(matches, reasons, newFolder))); } catch { } return; }
            int w = _suggPanel.Width;
            _suggPanel.Controls.Clear();
            if (matches != null && matches.Length > 0)
            {
                _suggPanel.Controls.Add(Ui.Caption("SUGGESTED FOLDERS", 2, 0, 300));
                int y = 18;
                for (int i = 0; i < matches.Length; i++)
                {
                    string n = matches[i];
                    // A long path truncates at the RIGHT, hiding the part that matters (the category leaf,
                    // e.g. \MC \MS). Show the last two segments with a leading … so the leaf is always
                    // visible; the full path is still what gets saved and is shown in the list below.
                    string disp = n;
                    var segs = n.Split('\\');
                    if (segs.Length > 2) disp = "…\\" + segs[segs.Length - 2] + "\\" + segs[segs.Length - 1];
                    var b = Ui.RowBtn("→   " + disp, 2, y, w - 4, 28);
                    b.Click += (o, e) => { Chosen = n; DialogResult = DialogResult.OK; Close(); };
                    _suggPanel.Controls.Add(b); y += 31;
                }
                // #2 Pre-select the top suggestion in the list so the primary Save button uses it by default.
                for (int i = 0; i < _list.Items.Count; i++)
                    if (string.Equals(_list.Items[i].ToString(), matches[0], StringComparison.OrdinalIgnoreCase))
                    { _list.SelectedIndex = i; break; }
            }
            else
            {
                _suggPanel.Controls.Add(Ui.Hint("No existing folder clearly fits — create one above, or pick from the list.", 2, 6, w - 6, 40));
            }
            if (!string.IsNullOrEmpty(newFolder) && _newFolderBox.Text.Trim().Length == 0)
                _newFolderBox.Text = newFolder;
        }

        private void PickFromList()
        {
            if (_list.SelectedItem != null) { Chosen = _list.SelectedItem.ToString(); DialogResult = DialogResult.OK; Close(); }
        }

        private void Filter()
        {
            string q = _search.Text.Trim().ToLowerInvariant();
            _list.BeginUpdate();
            _list.Items.Clear();
            foreach (var f in _all)
                if (q == "" || f.ToLowerInvariant().Contains(q)) _list.Items.Add(f);
            if (_list.Items.Count > 0) _list.SelectedIndex = 0;
            _list.EndUpdate();
        }
    }

    // Download dialog — pick (or add) a disk folder to save the email into (no AI suggestion needed).
    internal class DownloadPicker : Form
    {
        public string Chosen;
        public System.Collections.Generic.List<string> Folders;
        private readonly ListBox _list;
        private readonly Label _emptyHint;

        public DownloadPicker(string subject, System.Collections.Generic.List<string> folders)
        {
            Folders = folders ?? new System.Collections.Generic.List<string>();
            Text = "Axon intelligence — Download email";
            ClientSize = new System.Drawing.Size(500, 404);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false; ShowInTaskbar = false;
            BackColor = System.Drawing.Color.White;
            Font = new System.Drawing.Font("Segoe UI", 9.5F);
            int W = ClientSize.Width, H = ClientSize.Height;

            Controls.Add(Ui.Title("Save this email to a folder", 22, 18, W - 44));
            Controls.Add(Ui.Sub(subject ?? "(no subject)", 22, 46, W - 44));

            Controls.Add(Ui.Caption("YOUR SAVE FOLDERS", 22, 82, 300));
            _list = new ListBox
            {
                Left = 22, Top = 102, Width = W - 44, Height = H - 102 - 66,
                BorderStyle = BorderStyle.FixedSingle, DrawMode = DrawMode.OwnerDrawFixed, ItemHeight = 46, IntegralHeight = false
            };
            foreach (var f in Folders) _list.Items.Add(f);
            if (_list.Items.Count > 0) _list.SelectedIndex = 0;
            _list.DrawItem += (o, e) => Ui.DrawFolderItem(_list, e);
            _list.DoubleClick += (o, e) => PickFromList();
            Controls.Add(_list);

            _emptyHint = Ui.Hint("No save folders yet — click “+ Add folder” to choose where emails are saved.", 30, 116, W - 88, 36);
            _emptyHint.Visible = Folders.Count == 0;
            Controls.Add(_emptyHint);
            _emptyHint.BringToFront();

            var add = Ui.Subtle("+  Add folder", 22, H - 50, 128, 34);
            add.Click += (o, e) => AddFolder();
            Controls.Add(add);
            var cancel = Ui.Subtle("Cancel", W - 22 - 100, H - 50, 100, 34);
            cancel.DialogResult = DialogResult.Cancel;
            Controls.Add(cancel);
            CancelButton = cancel;
            var save = Ui.Accent("Save here", cancel.Left - 8 - 108, H - 50, 108, 34);
            save.Click += (o, e) => PickFromList();
            Controls.Add(save);
        }

        private void AddFolder()
        {
            using (var fb = new FolderBrowserDialog())
            {
                fb.Description = "Choose a folder to save emails into";
                if (fb.ShowDialog() == DialogResult.OK && !string.IsNullOrEmpty(fb.SelectedPath))
                {
                    if (!Folders.Contains(fb.SelectedPath)) { Folders.Add(fb.SelectedPath); _list.Items.Add(fb.SelectedPath); }
                    _list.SelectedItem = fb.SelectedPath;
                    _emptyHint.Visible = false;
                }
            }
        }

        private void PickFromList()
        {
            if (_list.SelectedItem != null) { Chosen = _list.SelectedItem.ToString(); DialogResult = DialogResult.OK; Close(); }
            else Ui.Notify("Pick a folder, or click “+ Add folder”.");
        }
    }
}
