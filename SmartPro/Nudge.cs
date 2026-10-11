using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace SmartPro
{
    // NUDGE: a small card in the corner of the screen after a tool has run. It never blocks Excel, closes by itself
    // and is shown rarely (see the limits below), so it reads as a tip and not as an advert.
    //   - a milestone reached (with the bonus Pro runs it gives)
    //   - the Pro Week: first day, the day before it ends, the day after it ended (once each)
    //   - after a free tool: the Pro tool that takes the same job one step further
    //   - after a free run of a Pro tool: what is left, and what Pro members get
    //   - for Pro members: now and then, what Pro saved them (so the value is seen before the renewal)
    // "No tips" on the card switches the tips off (milestones and the Pro Week notes still show).
    internal static class Nudge
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        const int TipsPerDay = 2;            // cross-sell tips a day at most
        const int RunsBetweenTips = 3;       // tool runs between two tips at least
        const int TimesPerPair = 3;          // the same tip is shown this many times at most, ever

        // free tool -> the Pro tool that goes one step further, and why (in the words of the user's job)
        static readonly Dictionary<string, string[]> Next = new Dictionary<string, string[]>
        {
            { "clean", new[] { "analyst", "Your table is clean. Smart Analyst can now write the analysis of it for you: who drives the total, the trend and what to look at next." } },
            { "report", new[] { "analyst", "The report is a clean table now. Smart Analyst turns it into a written summary with charts, ready for an email." } },
            { "health", new[] { "duppay", "The data passed its health scan. Duplicate Payments checks the same rows for bills that may have been paid twice." } },
            { "dups", new[] { "fuzzy", "Exact repeats are coloured. Fuzzy Duplicates also finds look-alikes such as 'ABC Pvt Ltd' and 'A.B.C. Private Limited'." } },
            { "listcompare", new[] { "reconcile", "Comparing lists by hand? Reconcile matches two lists by reference AND amount, with a rounding tolerance: books against a statement in one click." } },
            { "smartlookup", new[] { "fuzzymerge", "Some keys were not found. Fuzzy Merge matches names that are spelled a little differently and shows how sure it is for every row." } },
            { "summary", new[] { "variance", "You have the totals. Why It Changed shows which customers or products made this month go up or down, in plain words." } },
            { "dash", new[] { "analyst", "Dashboard done. Smart Analyst adds the written commentary that goes with it: the drivers, the trend and the unusual rows." } },
            { "custom", new[] { "analyst", "Your dashboard shows the numbers; Smart Analyst writes what they mean, ready to paste into an email." } },
            { "webdash", new[] { "mailparty", "Sending reports by hand? Mail to Each Party sends every customer or vendor only their own rows, in one go." } },
            { "splitsheets", new[] { "mailparty", "One sheet per party is ready. Mail to Each Party can email each of them their own rows straight away." } },
            { "pdfs", new[] { "splitpdfs", "Need one PDF per party instead of one per sheet? Split to PDFs makes them from any column in one go." } },
            { "emailsheet", new[] { "mailparty", "Emailing one sheet at a time? Mail to Each Party prepares one email per party with only its own rows." } },
            { "combsheets", new[] { "combfolder", "Sheets combined. Combine Folder does the same for every Excel and CSV file in a folder." } },
            { "iferror", new[] { "audit", "Hiding errors is quick; Formula Risk Scan shows where they come from and which formulas are risky." } },
            { "outliers", new[] { "analyst", "Unusual numbers marked. Smart Analyst explains them together with the trend and the main drivers." } },
            { "neat", new[] { "analyst", "The table looks good. Smart Analyst can write what it says, with charts." } },
            { "fycol", new[] { "ageing", "Dates in place. Ageing Report shows what is outstanding by party in 0-30, 31-60, 61-90 and 90+ days." } },
            { "unpivot", new[] { "forecast", "Now that the months are in one column, Forecast can project the next 3, 6 or 12 months with a likely range." } },
            { "formulas", new[] { "aiformula", "Need a formula that is not in the list? AI Formula writes it from your own words, for your own columns, with a short explanation." } },
            { "explain", new[] { "aiformula", "Want a new formula instead of explaining one? AI Formula writes it from your own words." } },
            { "index", new[] { "fileorg", "Workbook tidy? File Organizer does the same for folders: list, rename in bulk, sort into sub-folders." } },
            { "webtable", new[] { "tablecode", "Need the same table in a database or in Python? Table to SQL / Python writes the code with the right column types." } },
            { "backup", new[] { "snap", "Backups keep the file; Take Snapshot and What Changed? show exactly which cells changed since." } }
        };

        // a specific tip set by the tool that just ran (for example Smart Lookup with the number of names Fuzzy Merge could match)
        static string ctxTool, ctxText;

        public static void SetContext(string proTool, string text)
        {
            ctxTool = proTool;
            ctxText = text;
        }

        public static bool TipsOff
        {
            get { return Settings.Get("TipsOff") == "1"; }
            set { Settings.Set("TipsOff", value ? "1" : "0"); }
        }

        // called by the Runner after a tool has run
        public static void AfterRun(string id, List<Habit.Milestone> reached)
        {
            string cTool = ctxTool, cText = ctxText;
            ctxTool = null;
            ctxText = null;
            try
            {
                bool pro = Licence.IsProNow();
                if (reached != null && reached.Count > 0)
                {
                    ShowMilestone(reached, pro);
                    return;
                }
                if (!pro && ProWeekNote()) return;
                if (!Habit.IsWork(id)) return;

                if (pro)
                {
                    ProValue(id);
                    return;
                }
                // a free run of a Pro tool: the moment to show what Pro members get
                if (Design.IsPro(id))
                {
                    AfterTrial(id);
                    return;
                }
                if (TipsOff) return;
                string[] tip = null;
                if (cTool != null) tip = new[] { cTool, cText };
                else Next.TryGetValue(id, out tip);
                if (tip == null || !MayTip(id + ">" + tip[0])) return;
                int left = Licence.TrialLeft(tip[0]);
                bool week = Habit.WelcomeDay > 0;
                string button = "Try " + Design.CaptionOf(tip[0]);
                string foot = tip[0] == "aiformula" ? "3 a day free with a DocBrisk key, 100 a day with Pro." : week ? "Open during your Pro Week." : left > 0 ? left + (left == 1 ? " free run" : " free runs") + " to try it." : "Pro tool  -  " + ProPlan.PerDay + ".";
                Toast.Show("Next step: " + Design.CaptionOf(tip[0]), tip[1], foot, button, tip[0], "See Pro", "licence", true);
                Counted(id + ">" + tip[0]);
            }
            catch { }
        }

        // ---- the limits ----
        static bool MayTip(string pair)
        {
            string today = DateTime.Now.ToString("yyyy-MM-dd", Inv);
            if (Settings.Get("N_Day") != today)
            {
                Settings.Set("N_Day", today);
                Settings.Set("N_DayCount", "0");
            }
            if (Settings.GetInt("N_DayCount", 0) >= TipsPerDay) return false;
            int lastAt = Settings.GetInt("N_LastRun", -100);
            if (Habit.Runs - lastAt < RunsBetweenTips) return false;
            if (Settings.GetInt("N_" + pair, 0) >= TimesPerPair) return false;
            return true;
        }

        static void Counted(string pair)
        {
            Settings.Set("N_DayCount", (Settings.GetInt("N_DayCount", 0) + 1).ToString(Inv));
            Settings.Set("N_LastRun", Habit.Runs.ToString(Inv));
            Settings.Set("N_" + pair, (Settings.GetInt("N_" + pair, 0) + 1).ToString(Inv));
        }

        static bool Once(string key)
        {
            if (Settings.Get(key) == "1") return false;
            Settings.Set(key, "1");
            return true;
        }

        // ---- the cards ----
        static void ShowMilestone(List<Habit.Milestone> reached, bool pro)
        {
            Habit.Milestone m = reached[reached.Count - 1];
            int bonus = 0;
            foreach (Habit.Milestone r in reached) bonus += r.Bonus;
            string body = "You reached '" + m.Title + "'. DocBrisk has saved you about " + Habit.Hours(Habit.MinutesSaved) + " so far.";
            string foot;
            if (pro) foot = "Thank you for being a Pro member.";
            else
            {
                body += "\n\nReward: +" + bonus + " bonus Pro " + (bonus == 1 ? "run" : "runs") + ", for any Pro tool.";
                foot = "You have " + Habit.Bonus + " bonus " + (Habit.Bonus == 1 ? "run" : "runs") + " now.";
            }
            Habit.Milestone next = Habit.NextMilestone();
            if (next != null) foot += "  Next: " + next.Title + ".";
            Toast.Show("Milestone reached", body, foot, "My Impact", "impact", pro ? null : "See Pro", pro ? null : "licence", false);
        }

        // first day, the day before the end, and the day after the end of the Pro Week
        static bool ProWeekNote()
        {
            int wd = Habit.WelcomeDay;
            int days = ProPlan.WelcomeProDays;
            if (days <= 0) return false;
            if (wd >= 1 && Once("PW_Start"))
            {
                Toast.Show("Your Pro Week has started",
                    "For " + days + " days every Pro tool is open: Smart Analyst, Why It Changed, Fuzzy Merge, Reconcile, Duplicate Payments, Mail to Each Party and more. No key and no payment needed.",
                    "Everything that is free stays free afterwards.", "See the Pro tools", "licence", null, null, false);
                return true;
            }
            if (wd > 0 && wd == days - 1 && Once("PW_Tomorrow"))
            {
                Toast.Show("Pro Week ends tomorrow",
                    "So far you used Pro tools " + Habit.ProRuns + (Habit.ProRuns == 1 ? " time" : " times") + (Habit.ProMinutes >= 5 ? " and they saved you about " + Habit.Hours(Habit.ProMinutes) : "") +
                    ". After tomorrow each Pro tool has " + Licence.FreeRuns + " free runs.",
                    "Keep every Pro tool: " + ProPlan.PerDay + " on the yearly plan.", "Keep Pro", "licence", null, null, false);
                return true;
            }
            if (Habit.WelcomeOver && Once("PW_End"))
            {
                string used = Habit.ProRuns > 0
                    ? "During it you used Pro tools " + Habit.ProRuns + (Habit.ProRuns == 1 ? " time" : " times") + (Habit.ProMinutes >= 5 ? " and saved about " + Habit.Hours(Habit.ProMinutes) : "") + ". "
                    : "";
                Toast.Show("Your Pro Week has ended",
                    used + "Every free tool keeps working. Each Pro tool now has " + Licence.FreeRuns + " free runs, plus the bonus runs you earn with milestones.",
                    "Pro: " + ProPlan.PerDay + " (" + ProPlan.Rupee + ProPlan.YearRupees + " a year), or " + ProPlan.Rupee + ProPlan.MonthRupees + " for one month.", "Keep Pro", "licence", null, null, false);
                return true;
            }
            return false;
        }

        // after a free run of a Pro tool: shown on the last two runs, and on a bonus run (status bar the other times)
        static void AfterTrial(string id)
        {
            Licence.TrialKind k = Licence.LastTrial;
            int left = Licence.LastLeft;
            string cap = Design.CaptionOf(id);
            if (k == Licence.TrialKind.FreeRun && left <= 1)
            {
                string title = left == 0 ? "That was the last free run of " + cap : "1 free run of " + cap + " left";
                string body = "Pro members run " + cap + " and all " + Design.ProIds.Length + " Pro tools as often as they like" +
                              (Habit.Bonus > 0 ? ". You also have " + Habit.Bonus + " bonus " + (Habit.Bonus == 1 ? "run" : "runs") + " for any Pro tool." : ", on every file, with new tools as they come.");
                string foot = "This run saved you about " + Habit.MinutesOf(id) + " minutes.  Pro: " + ProPlan.PerDay + ".";
                Toast.Show(title, body, foot, "Get Pro", "licence", "My Impact", "impact", true);
            }
            else if (k == Licence.TrialKind.Bonus && left == 0)
            {
                Toast.Show("Your bonus runs are used",
                    "Reach the next milestone to earn more, or unlock every Pro tool for good.",
                    "Pro: " + ProPlan.PerDay + ".", "Get Pro", "licence", "My Impact", "impact", true);
            }
        }

        // Pro members: at most once a week, the value Pro has brought (renewals depend on it)
        static void ProValue(string id)
        {
            if (!Design.IsPro(id) || TipsOff) return;
            string week = DateTime.Now.ToString("yyyy", Inv) + "-" + CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(DateTime.Now, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
            if (Settings.Get("N_ProWeek") == week) return;
            if (Habit.ProMinutes < 60) return;
            Settings.Set("N_ProWeek", week);
            Toast.Show("Your Pro plan at work",
                "Pro tools have saved you about " + Habit.Hours(Habit.ProMinutes) + " so far" + (ProPlan.HourValueRupees > 0 ? ", worth about " + Habit.Rupees(Habit.ProMinutes) + " of your time" : "") + ".",
                "Thank you for being a Pro member.", "My Impact", "impact", null, null, true);
        }
    }

    // The card itself: no frame, bottom right of the screen Excel is on, closes after a few seconds
    // unless the mouse is on it. Only one is open at a time. It does not take the focus away from Excel.
    internal sealed class Toast : Form
    {
        static Toast open;
        static readonly Color Navy = Color.FromArgb(15, 23, 42), Blue = Color.FromArgb(37, 99, 235), Slate = Color.FromArgb(71, 85, 105),
                              Gold = Color.FromArgb(217, 119, 6);
        readonly Timer timer = new Timer();
        int ticks;
        readonly float z = 1f;

        int S(int px) { return (int)Math.Round(px * z); }

        protected override bool ShowWithoutActivation { get { return true; } }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= 0x08000000 | 0x00000080;     // WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW: no focus, no taskbar button
                return cp;
            }
        }

        // tool ids: a button runs that tool (or opens a window: "licence", "impact"); null hides the button
        public static void Show(string title, string body, string foot, string primary, string primaryTool, string secondary, string secondaryTool, bool tipsLink)
        {
            try
            {
                if (open != null && !open.IsDisposed) open.Close();
                open = new Toast(title, body, foot, primary, primaryTool, secondary, secondaryTool, tipsLink);
                open.Show();
            }
            catch { }
        }

        Toast(string title, string body, string foot, string primary, string primaryTool, string secondary, string secondaryTool, bool tipsLink)
        {
            try
            {
                using (Graphics g = Graphics.FromHwnd(IntPtr.Zero)) { z = g.DpiX / 96f; }
            }
            catch { z = 1f; }
            if (z < 1f) z = 1f;
            if (z > 3f) z = 3f;
            AutoScaleMode = AutoScaleMode.None;
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.Manual;
            BackColor = Color.White;
            Font = new Font("Segoe UI", 9.5f);

            int w = 380;
            Size measure;
            using (Graphics g = CreateGraphics())
            {
                measure = TextRenderer.MeasureText(g, body, new Font("Segoe UI", 9.5f), new Size(S(w - 36), 1000), TextFormatFlags.WordBreak);
            }
            int bodyH = (int)Math.Ceiling(measure.Height / z) + 4;
            int h = 44 + bodyH + (string.IsNullOrEmpty(foot) ? 0 : 34) + 50;
            ClientSize = new Size(S(w), S(h));

            Panel strip = new Panel();
            strip.SetBounds(0, 0, S(5), S(h));
            strip.BackColor = title.StartsWith("Milestone", StringComparison.Ordinal) ? Gold : Blue;
            Controls.Add(strip);

            Label t = new Label();
            t.Text = title;
            t.Font = new Font("Segoe UI", 11f, FontStyle.Bold);
            t.ForeColor = Navy;
            t.SetBounds(S(18), S(12), S(w - 60), S(26));
            Controls.Add(t);

            Label x = new Label();
            x.Text = "×";
            x.Font = new Font("Segoe UI", 13f);
            x.ForeColor = Slate;
            x.Cursor = Cursors.Hand;
            x.SetBounds(S(w - 32), S(6), S(26), S(28));
            x.Click += delegate { Close(); };
            Controls.Add(x);

            Label b = new Label();
            b.Text = body;
            b.ForeColor = Color.FromArgb(30, 41, 59);
            b.SetBounds(S(18), S(42), S(w - 36), S(bodyH));
            Controls.Add(b);
            int y = 42 + bodyH;
            if (!string.IsNullOrEmpty(foot))
            {
                Label f = new Label();
                f.Text = foot;
                f.Font = new Font("Segoe UI", 8.5f);
                f.ForeColor = Slate;
                f.SetBounds(S(18), S(y + 2), S(w - 36), S(32));
                Controls.Add(f);
                y += 34;
            }
            int bx = 18;
            if (primary != null) bx = AddButton(primary, primaryTool, bx, y + 8, true);
            if (secondary != null) bx = AddButton(secondary, secondaryTool, bx, y + 8, false);
            if (tipsLink)
            {
                LinkLabel off = new LinkLabel();
                off.Text = "No tips";
                off.Font = new Font("Segoe UI", 8.5f);
                off.LinkColor = Slate;
                off.SetBounds(S(w - 70), S(y + 16), S(60), S(20));
                off.LinkClicked += delegate
                {
                    Nudge.TipsOff = true;
                    Close();
                };
                Controls.Add(off);
            }

            Paint += delegate (object s, PaintEventArgs e)
            {
                using (Pen p = new Pen(Color.FromArgb(203, 213, 225))) e.Graphics.DrawRectangle(p, 0, 0, ClientSize.Width - 1, ClientSize.Height - 1);
            };

            // bottom right of the screen that Excel is on
            Rectangle area;
            try { area = Screen.FromHandle(ExcelDnaUtil.WindowHandle).WorkingArea; }
            catch { area = Screen.PrimaryScreen.WorkingArea; }
            Location = new Point(area.Right - Width - S(16), area.Bottom - Height - S(16));

            timer.Interval = 1000;
            timer.Tick += delegate
            {
                if (Bounds.Contains(Cursor.Position)) { ticks = 0; return; }
                if (++ticks >= 14) Close();
            };
            timer.Start();
            FormClosed += delegate { timer.Stop(); timer.Dispose(); };
        }

        int AddButton(string text, string tool, int x, int y, bool main)
        {
            Button btn = new Button();
            btn.Text = text;
            btn.FlatStyle = FlatStyle.Flat;
            btn.Font = new Font("Segoe UI", 9f, main ? FontStyle.Bold : FontStyle.Regular);
            btn.Cursor = Cursors.Hand;
            if (main)
            {
                btn.BackColor = Blue;
                btn.ForeColor = Color.White;
                btn.FlatAppearance.BorderSize = 0;
            }
            else
            {
                btn.BackColor = Color.White;
                btn.ForeColor = Navy;
                btn.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
            }
            int w;
            using (Graphics g = CreateGraphics()) { w = (int)Math.Ceiling(TextRenderer.MeasureText(g, text, btn.Font).Width / z) + 28; }
            btn.SetBounds(S(x), S(y), S(w), S(32));
            btn.Click += delegate
            {
                Close();
                string id = tool;
                // run it as a macro: Excel may be busy (for example a cell in edit mode) when the card is clicked
                try { ExcelAsyncUtil.QueueAsMacro(delegate { Runner.Run(id); }); }
                catch { }
            };
            Controls.Add(btn);
            return x + w + 8;
        }
    }
}
