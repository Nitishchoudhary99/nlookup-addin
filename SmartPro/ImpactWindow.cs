using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace SmartPro
{
    // MY IMPACT: what DocBrisk has done for this user: time saved, the streak, the milestones and the bonus runs they gave,
    // the tools used most, and (for Free users) the Pro tools that fit the work they already do.
    internal static class ImpactWindow
    {
        public static void Show()
        {
            string run;
            using (ImpactForm f = new ImpactForm())
            {
                f.ShowDialog(new ExcelWin());
                run = f.RunId;
            }
            if (run != null) Runner.Run(run);
        }
    }

    internal sealed class ImpactForm : Form
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        static readonly Color Navy = Color.FromArgb(15, 23, 42), Slate = Color.FromArgb(71, 85, 105), Blue = Color.FromArgb(37, 99, 235),
                              Green = Color.FromArgb(21, 128, 61), Gold = Color.FromArgb(217, 119, 6), Soft = Color.FromArgb(241, 245, 249),
                              Line = Color.FromArgb(226, 232, 240);
        readonly float z = 1f;
        public string RunId;

        int S(int px) { return (int)Math.Round(px * z); }

        Label Lbl(Control parent, string text, int x, int y, int w, int h, float size, bool bold, Color color)
        {
            Label l = new Label();
            l.Text = text;
            l.SetBounds(S(x), S(y), S(w), S(h));
            l.Font = new Font("Segoe UI", size, bold ? FontStyle.Bold : FontStyle.Regular);
            l.ForeColor = color;
            l.BackColor = Color.Transparent;
            parent.Controls.Add(l);
            return l;
        }

        Panel Tile(int x, int y, string big, string small, Color color)
        {
            Panel p = new Panel();
            p.SetBounds(S(x), S(y), S(200), S(84));
            p.BackColor = Soft;
            Controls.Add(p);
            Lbl(p, big, 12, 8, 182, 40, 18f, true, color);
            Lbl(p, small, 12, 50, 182, 30, 8.5f, false, Slate);
            return p;
        }

        public ImpactForm()
        {
            try
            {
                using (Graphics g = Graphics.FromHwnd(IntPtr.Zero)) { z = g.DpiX / 96f; }
            }
            catch { z = 1f; }
            if (z < 1f) z = 1f;
            if (z > 3f) z = 3f;
            try
            {
                float fit = (Screen.PrimaryScreen.WorkingArea.Height - 48) / 620f;
                if (fit < z && fit > 0.6f) z = fit;
            }
            catch { }
            AutoScaleMode = AutoScaleMode.None;
            Text = Design.AppName + " - My Impact";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            BackColor = Color.White;
            Font = new Font("Segoe UI", 9.5f);
            ClientSize = new Size(S(880), S(620));

            bool pro = Licence.IsProNow();
            Panel head = new Panel();
            head.SetBounds(0, 0, S(880), S(70));
            head.BackColor = Navy;
            Controls.Add(head);
            Lbl(head, "My Impact", 18, 8, 400, 32, 16f, true, Color.White);
            Lbl(head, Licence.PlanText(), 20, 42, 840, 22, 9f, false, Color.FromArgb(203, 213, 225));

            // ---- four tiles ----
            double saved = Habit.MinutesSaved;
            Tile(18, 84, Habit.Hours(saved), "saved so far (estimate)" + (ProPlan.HourValueRupees > 0 && saved >= 60 ? ", worth about " + Habit.Rupees(saved) : ""), Green);
            Tile(234, 84, Habit.Hours(Habit.MinutesThisMonth), "saved this month", Blue);
            int st = Habit.Streak;
            Tile(450, 84, st + (st == 1 ? " day" : " days"), "streak of working days (best: " + Habit.BestStreak + ")", Gold);
            Tile(666, 84, Habit.Runs.ToString("#,##0", Inv), "tool runs, " + Habit.ToolsUsed().Count + " of " + Design.Tools.Length + " tools tried", Navy);

            // ---- left: milestones ----
            Lbl(this, "Milestones" + (pro ? "" : "   (each one adds bonus runs for any Pro tool)"), 18, 186, 420, 22, 10f, true, Navy);
            ListView miles = new ListView();
            miles.SetBounds(S(18), S(212), S(420), S(262));
            miles.View = View.Details;
            miles.FullRowSelect = true;
            miles.HeaderStyle = ColumnHeaderStyle.Nonclickable;
            miles.Columns.Add("", S(28));
            miles.Columns.Add("Milestone", S(230));
            miles.Columns.Add(pro ? "Progress" : "Reward", S(140));
            foreach (Habit.Milestone m in Habit.Milestones)
            {
                bool done = Habit.Reached(m.Id);
                int pct = (int)Math.Min(100, Math.Floor(Habit.Progress(m.Id) * 100));
                string right = done ? "Reached" : (pro ? pct + "%" : "+" + m.Bonus + (m.Bonus == 1 ? " run" : " runs") + "   (" + pct + "%)");
                ListViewItem it = new ListViewItem(new[] { done ? "✓" : "", m.Title, right });
                it.ForeColor = done ? Green : Slate;
                miles.Items.Add(it);
            }
            Controls.Add(miles);
            Habit.Milestone next = Habit.NextMilestone();
            string nextText = next == null ? "Every milestone reached. Thank you!" : "Next: " + next.Title + "  (" + (int)Math.Min(99, Math.Floor(Habit.Progress(next.Id) * 100)) + "% there)";
            if (!pro && Habit.Bonus > 0) nextText += "   |   Bonus runs to use: " + Habit.Bonus;
            Lbl(this, nextText, 18, 480, 420, 22, 9f, true, Gold);

            // ---- right: most used, and what Pro adds for this work ----
            Lbl(this, "Your most used tools", 458, 186, 404, 22, 10f, true, Navy);
            List<KeyValuePair<string, int>> top = Habit.TopTools(4);
            int y = 212;
            if (top.Count == 0) Lbl(this, "Nothing yet. Start with the Assistant or Smart Lookup.", 458, y, 404, 22, 9f, false, Slate);
            foreach (KeyValuePair<string, int> kv in top)
            {
                Lbl(this, Design.CaptionOf(kv.Key), 458, y, 250, 22, 9.5f, false, Navy);
                Lbl(this, kv.Value + (kv.Value == 1 ? " run" : " runs") + ", ~" + Habit.Hours(Math.Min(kv.Value, 1000) * Habit.MinutesOf(kv.Key)), 708, y, 160, 22, 9f, false, Slate);
                y += 24;
            }
            y = Math.Max(y + 10, 320);
            if (!pro)
            {
                Lbl(this, "Pro tools that fit your work", 458, y, 404, 22, 10f, true, Navy);
                y += 26;
                foreach (string id in Suggest(top))
                {
                    string pid = id;
                    LinkLabel l = new LinkLabel();
                    int left = Licence.TrialLeft(pid);
                    l.Text = Design.CaptionOf(pid) + (Habit.WelcomeDay > 0 ? "   (open in your Pro Week)" : left > 0 ? "   (" + left + " free " + (left == 1 ? "run" : "runs") + " left)" : "   (Pro)");
                    l.LinkArea = new LinkArea(0, Design.CaptionOf(pid).Length);
                    l.SetBounds(S(458), S(y), S(404), S(22));
                    l.LinkColor = Blue;
                    l.LinkClicked += delegate
                    {
                        RunId = pid;
                        Close();
                    };
                    Controls.Add(l);
                    y += 24;
                }
            }
            else
            {
                Lbl(this, "Pro tools have saved you about " + Habit.Hours(Habit.ProMinutes) + " in " + Habit.ProRuns + " runs.", 458, y, 404, 44, 9.5f, false, Green);
            }

            // ---- bottom ----
            Panel line = new Panel();
            line.SetBounds(S(18), S(516), S(844), 1);
            line.BackColor = Line;
            Controls.Add(line);
            if (!pro)
            {
                string val = ProPlan.HourValueRupees > 0 && saved >= 60
                    ? "Your time saved is worth about " + Habit.Rupees(saved) + ". Pro costs " + ProPlan.PerDay + "."
                    : "Pro unlocks all " + Design.ProIds.Length + " Pro tools, as often as you like, for " + ProPlan.PerDay + ".";
                Lbl(this, val, 18, 532, 600, 22, 9.5f, true, Navy);
                Button get = new Button();
                get.Text = "See Pro";
                get.FlatStyle = FlatStyle.Flat;
                get.FlatAppearance.BorderSize = 0;
                get.BackColor = Blue;
                get.ForeColor = Color.White;
                get.Font = new Font("Segoe UI", 10f, FontStyle.Bold);
                get.SetBounds(S(640), S(526), S(120), S(34));
                get.Click += delegate
                {
                    RunId = "licence";
                    Close();
                };
                Controls.Add(get);
            }
            CheckBox tips = new CheckBox();
            tips.Text = "Show tips after a tool runs";
            tips.Checked = !Nudge.TipsOff;
            tips.SetBounds(S(18), S(574), S(300), S(24));
            tips.CheckedChanged += delegate { Nudge.TipsOff = !tips.Checked; };
            Controls.Add(tips);
            Lbl(this, "Minutes saved are an estimate per tool; a tool counts at most 3 times a day. Nothing leaves this PC.", 18, 598, 700, 18, 8f, false, Slate);
            Button close = new Button();
            close.Text = "Close";
            close.SetBounds(S(776), S(572), S(86), S(30));
            close.DialogResult = DialogResult.Cancel;
            Controls.Add(close);
            CancelButton = close;
        }

        // Pro tools that go with the free tools used most (the same pairs as the tips), then the popular ones
        static List<string> Suggest(List<KeyValuePair<string, int>> top)
        {
            List<string> l = new List<string>();
            Dictionary<string, string> pair = new Dictionary<string, string>
            {
                { "clean", "analyst" }, { "report", "analyst" }, { "dash", "analyst" }, { "custom", "analyst" }, { "summary", "variance" },
                { "listcompare", "reconcile" }, { "smartlookup", "fuzzymerge" }, { "dups", "fuzzy" }, { "health", "duppay" },
                { "splitsheets", "mailparty" }, { "emailsheet", "mailparty" }, { "combsheets", "combfolder" }, { "explain", "audit" }
            };
            foreach (KeyValuePair<string, int> kv in top)
            {
                string p;
                if (pair.TryGetValue(kv.Key, out p) && !l.Contains(p)) l.Add(p);
            }
            foreach (string p in new[] { "analyst", "fuzzymerge", "reconcile", "duppay", "variance" })
            {
                if (l.Count >= 4) break;
                if (!l.Contains(p)) l.Add(p);
            }
            if (l.Count > 4) l.RemoveRange(4, l.Count - 4);
            return l;
        }
    }
}
