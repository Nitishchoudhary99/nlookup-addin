using System;
using System.Collections.Generic;
using System.Globalization;

namespace SmartPro
{
    // HABIT: what DocBrisk has done for this user so far, kept in the registry next to the other settings.
    //   - minutes saved: a careful estimate per tool (Minutes below), at most 3 runs of one tool counted per day
    //   - a streak of working days (Saturday and Sunday do not break it)
    //   - milestones that give bonus runs of the Pro tools (used by Licence.AllowPro when a tool's own free runs are gone)
    //   - the Pro Week: every Pro tool is open for ProPlan.WelcomeProDays days from the first day
    // Nothing leaves the PC. The window that shows it all is ImpactWindow.cs; the pop-ups are in Nudge.cs.
    internal static class Habit
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        const int CountPerToolPerDay = 3;

        // minutes a careful user would spend doing the same by hand, per run. Tools not listed count 2 minutes.
        static readonly Dictionary<string, int> Minutes = new Dictionary<string, int>
        {
            { "dash", 45 }, { "custom", 30 }, { "editdash", 10 }, { "webdash", 30 }, { "refresh", 15 }, { "biexport", 15 }, { "style", 3 },
            { "analyst", 40 }, { "variance", 30 }, { "forecast", 25 }, { "ask", 10 },
            { "fuzzymerge", 45 }, { "reconcile", 40 }, { "sumfinder", 30 }, { "duppay", 40 }, { "compare", 20 }, { "fuzzy", 20 },
            { "snap", 2 }, { "diff", 20 }, { "mailparty", 45 },
            { "smartlookup", 12 }, { "explain", 5 },
            { "clean", 10 }, { "report", 25 }, { "tonum", 5 }, { "todate", 5 }, { "trim", 3 }, { "blankrows", 3 }, { "filldown", 4 },
            { "unmerge", 4 }, { "tovalues", 2 }, { "health", 10 },
            { "neat", 5 }, { "summary", 8 }, { "listcompare", 10 },
            { "pastevisible", 5 }, { "delrows", 5 }, { "blankchange", 5 }, { "mergesame", 5 }, { "joincols", 3 }, { "texttools", 4 },
            { "mathcells", 4 }, { "iferror", 3 }, { "dollar", 3 },
            { "abc", 20 }, { "ageing", 30 }, { "audit", 25 }, { "outliers", 8 }, { "dups", 3 }, { "unpivot", 15 }, { "colourcount", 5 },
            { "splitsheets", 15 }, { "splitfiles", 25 }, { "combsheets", 15 }, { "combfolder", 30 }, { "index", 5 }, { "pdfs", 10 },
            { "emailsheet", 4 }, { "unhideall", 3 }, { "protectall", 4 }, { "unprotectall", 4 }, { "links", 8 }, { "diet", 10 },
            { "sortsheets", 3 }, { "sheetsfromlist", 5 }, { "valuescopy", 5 },
            { "inrfmt", 2 }, { "fycol", 5 }, { "words", 3 }, { "gstin", 8 }, { "backup", 1 }
        };

        // tools that only show something (no work done): not counted as runs
        static readonly string[] NotWork = { "assist", "find", "about", "licence", "funcs", "impact", "repeat", "openbackups" };

        // milestone id, how it is reached (shown in My Impact), bonus Pro runs it gives
        internal sealed class Milestone
        {
            public readonly string Id, Title;
            public readonly int Bonus;
            public Milestone(string id, string title, int bonus) { Id = id; Title = title; Bonus = bonus; }
        }

        public static readonly Milestone[] Milestones =
        {
            new Milestone("runs10", "10 tool runs", 1),
            new Milestone("tools5", "5 different tools tried", 1),
            new Milestone("streak3", "3 working days in a row", 1),
            new Milestone("hour1", "1 hour saved", 2),
            new Milestone("streak5", "5 working days in a row (a full week)", 2),
            new Milestone("tools12", "12 different tools tried", 2),
            new Milestone("hours5", "5 hours saved", 3),
            new Milestone("runs100", "100 tool runs", 3),
            new Milestone("streak20", "20 working days in a row (a month)", 5),
            new Milestone("hours20", "20 hours saved", 5)
        };

        static string Today { get { return DateTime.Now.ToString("yyyy-MM-dd", Inv); } }

        static DateTime? Day(string name)
        {
            DateTime d;
            string s = Settings.Get(name);
            if (s != null && DateTime.TryParseExact(s, "yyyy-MM-dd", Inv, DateTimeStyles.None, out d)) return d;
            return null;
        }

        static double GetDouble(string name)
        {
            double d;
            return double.TryParse(Settings.Get(name), NumberStyles.Float, Inv, out d) && d > 0 ? d : 0;
        }

        static void SetDouble(string name, double d) { Settings.Set(name, d.ToString("0.##", Inv)); }

        public static int MinutesOf(string id)
        {
            int m;
            return Minutes.TryGetValue(id, out m) ? m : 2;
        }

        public static bool IsWork(string id) { return Array.IndexOf(NotWork, id) < 0; }

        // ---- what is shown ----
        public static double MinutesSaved { get { return GetDouble("H_Min"); } }
        public static double MinutesThisMonth { get { return Settings.Get("H_Month") == DateTime.Now.ToString("yyyy-MM", Inv) ? GetDouble("H_MonthMin") : 0; } }
        public static double MinutesToday { get { return Settings.Get("H_Day") == Today ? GetDouble("H_DayMin") : 0; } }
        public static double ProMinutes { get { return GetDouble("H_ProMin"); } }
        public static int ProRuns { get { return Settings.GetInt("H_ProRuns", 0); } }
        public static int Runs { get { return Settings.GetInt("Runs", 0); } }
        public static int BestStreak { get { return Math.Max(Settings.GetInt("H_Best", 0), Streak); } }
        public static int Bonus { get { return Math.Max(0, Settings.GetInt("Bonus", 0)); } }

        // the streak as it stands today: it is still alive when the last day of use was the previous working day
        public static int Streak
        {
            get
            {
                DateTime? last = Day("H_Last");
                if (last == null) return 0;
                int gap = WorkDaysBetween(last.Value, DateTime.Now.Date);
                return gap <= 1 ? Settings.GetInt("H_Streak", 0) : 0;
            }
        }

        public static List<string> ToolsUsed()
        {
            List<string> l = new List<string>();
            string s = Settings.Get("H_Tools");
            if (!string.IsNullOrEmpty(s)) foreach (string t in s.Split(',')) if (t.Length > 0 && !l.Contains(t)) l.Add(t);
            return l;
        }

        public static bool Reached(string milestoneId)
        {
            string s = "," + (Settings.Get("H_Miles") ?? "") + ",";
            return s.IndexOf("," + milestoneId + ",", StringComparison.Ordinal) >= 0;
        }

        // how far the user is towards a milestone, 0 to 1
        public static double Progress(string id)
        {
            switch (id)
            {
                case "runs10": return Runs / 10.0;
                case "runs100": return Runs / 100.0;
                case "tools5": return ToolsUsed().Count / 5.0;
                case "tools12": return ToolsUsed().Count / 12.0;
                case "streak3": return BestStreak / 3.0;
                case "streak5": return BestStreak / 5.0;
                case "streak20": return BestStreak / 20.0;
                case "hour1": return MinutesSaved / 60.0;
                case "hours5": return MinutesSaved / 300.0;
                case "hours20": return MinutesSaved / 1200.0;
                default: return 0;
            }
        }

        public static Milestone NextMilestone()
        {
            Milestone best = null;
            double bp = -1;
            foreach (Milestone m in Milestones)
            {
                if (Reached(m.Id)) continue;
                double p = Math.Min(Progress(m.Id), 0.999);
                if (p > bp) { bp = p; best = m; }
            }
            return best;
        }

        // ---- Pro Week ----
        public static DateTime? FirstDay { get { return Day("H_First"); } }

        // 1 on the first day, WelcomeProDays on the last; 0 when the Pro Week has not started or is over
        public static int WelcomeDay
        {
            get
            {
                if (ProPlan.WelcomeProDays <= 0) return 0;
                DateTime? f = FirstDay;
                if (f == null) return 0;
                int d = (int)(DateTime.Now.Date - f.Value).TotalDays + 1;
                return d >= 1 && d <= ProPlan.WelcomeProDays ? d : 0;
            }
        }

        public static bool WelcomeOver
        {
            get { return ProPlan.WelcomeProDays > 0 && FirstDay != null && WelcomeDay == 0 && DateTime.Now.Date >= FirstDay.Value; }
        }

        // ---- bonus runs ----
        public static bool TakeBonus()
        {
            int b = Bonus;
            if (b <= 0) return false;
            Settings.Set("Bonus", (b - 1).ToString(Inv));
            return true;
        }

        // ---- recording one run ----
        // Returns the milestones reached by this run (usually none).
        public static List<Milestone> Record(string id)
        {
            List<Milestone> reached = new List<Milestone>();
            if (!IsWork(id)) return reached;
            string today = Today;
            if (Settings.Get("H_First") == null) Settings.Set("H_First", today);

            // streak of working days
            DateTime? last = Day("H_Last");
            if (last == null || last.Value.Date != DateTime.Now.Date)
            {
                int streak = 1;
                if (last != null && WorkDaysBetween(last.Value, DateTime.Now.Date) <= 1) streak = Settings.GetInt("H_Streak", 0) + 1;
                Settings.Set("H_Streak", streak.ToString(Inv));
                if (streak > Settings.GetInt("H_Best", 0)) Settings.Set("H_Best", streak.ToString(Inv));
                Settings.Set("H_Last", today);
                Settings.Set("H_Days", (Settings.GetInt("H_Days", 0) + 1).ToString(Inv));
            }

            // minutes saved: at most CountPerToolPerDay runs of one tool a day
            if (Settings.Get("H_Day") != today)
            {
                Settings.Set("H_Day", today);
                Settings.Set("H_DayMin", "0");
                Settings.Set("H_DayRuns", "");
            }
            Dictionary<string, int> dayRuns = ParseCounts(Settings.Get("H_DayRuns"));
            int n;
            dayRuns.TryGetValue(id, out n);
            dayRuns[id] = n + 1;
            Settings.Set("H_DayRuns", FormatCounts(dayRuns));
            if (n < CountPerToolPerDay)
            {
                double m = MinutesOf(id);
                SetDouble("H_Min", MinutesSaved + m);
                SetDouble("H_DayMin", GetDouble("H_DayMin") + m);
                string month = DateTime.Now.ToString("yyyy-MM", Inv);
                if (Settings.Get("H_Month") != month)
                {
                    Settings.Set("H_Month", month);
                    Settings.Set("H_MonthMin", "0");
                }
                SetDouble("H_MonthMin", GetDouble("H_MonthMin") + m);
                if (Design.IsPro(id)) SetDouble("H_ProMin", ProMinutes + m);
            }
            if (Design.IsPro(id)) Settings.Set("H_ProRuns", (ProRuns + 1).ToString(Inv));

            // tools tried
            List<string> tools = ToolsUsed();
            if (!tools.Contains(id))
            {
                tools.Add(id);
                Settings.Set("H_Tools", string.Join(",", tools.ToArray()));
            }
            Dictionary<string, int> all = ParseCounts(Settings.Get("H_ToolRuns"));
            all.TryGetValue(id, out n);
            all[id] = n + 1;
            Settings.Set("H_ToolRuns", FormatCounts(all));

            // milestones
            foreach (Milestone ms in Milestones)
            {
                if (Reached(ms.Id) || Progress(ms.Id) < 1) continue;
                string had = Settings.Get("H_Miles") ?? "";
                Settings.Set("H_Miles", had.Length == 0 ? ms.Id : had + "," + ms.Id);
                if (!Licence.IsProNow()) Settings.Set("Bonus", (Bonus + ms.Bonus).ToString(Inv));
                reached.Add(ms);
            }
            return reached;
        }

        // the tools used most, most used first
        public static List<KeyValuePair<string, int>> TopTools(int max)
        {
            List<KeyValuePair<string, int>> l = new List<KeyValuePair<string, int>>(ParseCounts(Settings.Get("H_ToolRuns")));
            l.Sort(delegate (KeyValuePair<string, int> a, KeyValuePair<string, int> b) { return b.Value.CompareTo(a.Value); });
            if (l.Count > max) l.RemoveRange(max, l.Count - max);
            return l;
        }

        // the line on the status bar when Excel starts (returning users only)
        public static string Greeting()
        {
            if (FirstDay == null) return null;
            List<string> parts = new List<string>();
            int st = Streak;
            if (st >= 2) parts.Add(st + "-day streak: use any tool today to keep it");
            double h = MinutesSaved;
            if (h >= 30) parts.Add("about " + Hours(h) + " saved so far");
            int wd = WelcomeDay;
            if (wd > 0 && !Licence.IsProNow()) parts.Add("Pro Week day " + wd + " of " + ProPlan.WelcomeProDays + ": every Pro tool is open");
            else if (Bonus > 0 && !Licence.IsProNow()) parts.Add(Bonus + " bonus Pro " + (Bonus == 1 ? "run" : "runs") + " waiting");
            if (parts.Count == 0) parts.Add("Ctrl+Shift+Q finds any tool");
            return string.Join("   |   ", parts.ToArray());
        }

        // 95 -> "1.6 hours", 40 -> "40 minutes"
        public static string Hours(double minutes)
        {
            if (minutes < 60) return Math.Round(minutes).ToString("0", Inv) + " minutes";
            double h = minutes / 60;
            return h.ToString(h < 10 ? "0.0" : "#,##0", Inv) + " hours";
        }

        public static string Rupees(double minutes)
        {
            double r = minutes / 60 * ProPlan.HourValueRupees;
            return ProPlan.Rupee + Logic.IndianGroup(Math.Round(r / 10) * 10, 0);
        }

        // working days from a to b (Saturdays and Sundays are skipped): Fri -> Mon is 1
        static int WorkDaysBetween(DateTime a, DateTime b)
        {
            a = a.Date;
            b = b.Date;
            if (b <= a) return 0;
            int n = 0;
            for (DateTime d = a.AddDays(1); d <= b && n < 60; d = d.AddDays(1))
            {
                if (d.DayOfWeek != DayOfWeek.Saturday && d.DayOfWeek != DayOfWeek.Sunday) n++;
            }
            // a day of use on a weekend still counts when the next working day follows
            if (n == 0) return 1;
            return n;
        }

        static Dictionary<string, int> ParseCounts(string s)
        {
            Dictionary<string, int> d = new Dictionary<string, int>();
            if (string.IsNullOrEmpty(s)) return d;
            foreach (string part in s.Split(';'))
            {
                int p = part.IndexOf(':');
                int n;
                if (p > 0 && int.TryParse(part.Substring(p + 1), NumberStyles.Integer, Inv, out n)) d[part.Substring(0, p)] = n;
            }
            return d;
        }

        static string FormatCounts(Dictionary<string, int> d)
        {
            List<string> l = new List<string>();
            foreach (KeyValuePair<string, int> kv in d) l.Add(kv.Key + ":" + kv.Value.ToString(Inv));
            return string.Join(";", l.ToArray());
        }
    }
}
