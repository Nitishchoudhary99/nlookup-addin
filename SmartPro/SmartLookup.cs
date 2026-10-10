using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Text;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace SmartPro
{
    // SMART LOOKUP: brings columns from a second table into this one, like VLOOKUP / XLOOKUP, but without the #N/A
    // that comes from invisible differences: spaces at the ends or in between, non-breaking spaces, capitals,
    // numbers stored as text ('00123' / 123), a leading apostrophe. Every row gets a Match column that says how it was
    // found, and the summary says how many rows VLOOKUP would have missed. Keys that appear twice in the second table
    // are counted (VLOOKUP silently takes the first). The worksheet function SMARTLOOKUP does the same in a formula.
    internal static class SmartLookup
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        // ---- the key: what VLOOKUP compares, and the forgiving version ----
        // VLOOKUP: text without regard to capitals, numbers as numbers; a number and a text never match.
        public static string StrictKey(object v)
        {
            if (v == null || v is ExcelEmpty || v is ExcelMissing) return null;
            if (v is double) return "n:" + ((double)v).ToString("R", Inv);
            if (v is DateTime) return "n:" + ((DateTime)v).ToOADate().ToString("R", Inv);
            if (v is bool) return "b:" + ((bool)v ? "1" : "0");
            string s = v as string;
            if (s != null) return s.Length == 0 ? null : "t:" + s.ToLowerInvariant();
            if (Xl.IsErr(v) || v is ExcelError) return null;
            return "t:" + Convert.ToString(v, Inv).ToLowerInvariant();
        }

        // Forgiving: spaces of every kind trimmed and squeezed, invisible characters and a leading apostrophe removed,
        // capitals ignored, a text that is a plain number compared as that number.
        public static string LooseKey(object v)
        {
            if (v == null || v is ExcelEmpty || v is ExcelMissing || v is ExcelError || Xl.IsErr(v)) return null;
            if (v is double) return "n:" + ((double)v).ToString("R", Inv);
            if (v is DateTime) return "n:" + ((DateTime)v).ToOADate().ToString("R", Inv);
            if (v is bool) return "b:" + ((bool)v ? "1" : "0");
            string s = Convert.ToString(v, Inv) ?? "";
            StringBuilder sb = new StringBuilder(s.Length);
            bool space = false;
            foreach (char ch in s)
            {
                if (ch == '​' || ch == '‌' || ch == '‍' || ch == '﻿' || ch == '­') continue;   // invisible
                if (char.IsWhiteSpace(ch) || ch == ' ')
                {
                    space = true;
                    continue;
                }
                if (space && sb.Length > 0) sb.Append(' ');
                space = false;
                sb.Append(char.ToLowerInvariant(ch));
            }
            string t = sb.ToString();
            if (t.StartsWith("'", StringComparison.Ordinal)) t = t.Substring(1).Trim();
            if (t.Length == 0) return null;
            double d;
            if (PlainNumber(t) && double.TryParse(t, NumberStyles.Float, Inv, out d)) return "n:" + d.ToString("R", Inv);
            return "t:" + t;
        }

        // digits with an optional sign and decimal point, nothing else (so codes like 1-2 or 1e5 stay text)
        static bool PlainNumber(string t)
        {
            int i = 0, digits = 0;
            bool dot = false;
            if (t.Length > 0 && (t[0] == '-' || t[0] == '+')) i++;
            for (; i < t.Length; i++)
            {
                char ch = t[i];
                if (ch >= '0' && ch <= '9') digits++;
                else if (ch == '.' && !dot) dot = true;
                else return false;
            }
            return digits > 0 && digits <= 15;
        }

        // ---- the tool ----
        public static void Run()
        {
            dynamic mine = Xl.PickRange("Smart Lookup (1 of 2)", "Select YOUR table, with its header row. The new columns are added on its right:");
            if (mine == null) return;
            if ((int)mine.Rows.Count < 2)
            {
                Ui.Info("Select a table with a header row and at least one row of data.");
                return;
            }
            object[,] mv = Xl.ReadTyped(mine);
            int mr = mv.GetLength(0), mc = mv.GetLength(1);
            string[] mh = Headers(mv);
            int myKey = Ui.PickColumn("Smart Lookup - your key column", "Which column of YOUR table holds the key to look up (code, invoice no., name ...)?", mh, null);
            if (myKey == 0) return;

            dynamic other = Xl.PickRange("Smart Lookup (2 of 2)", "Select the table to look the values up IN, with its header row. It can be on another sheet or in another open workbook:");
            if (other == null) return;
            object[,] ov = Xl.ReadTyped(other);
            int orows = ov.GetLength(0), oc = ov.GetLength(1);
            if (orows < 2)
            {
                Ui.Info("The second table needs a header row and at least one row of data.");
                return;
            }
            string[] oh = Headers(ov);
            int theirKey = Ui.PickColumn("Smart Lookup - key in the second table", "Which column of the SECOND table holds the same key?", oh, null);
            if (theirKey == 0) return;
            List<int> bring = MultiPick.Show("Smart Lookup - columns to bring",
                "Tick the columns of the second table to bring into your table:", oh, theirKey);
            if (bring == null || bring.Count == 0) return;

            // index of the second table: strict and forgiving keys, first row wins (like VLOOKUP), repeats counted
            Dictionary<string, int> strict = new Dictionary<string, int>(), loose = new Dictionary<string, int>();
            int repeated = 0;
            for (int r = 2; r <= orows; r++)
            {
                string sk = StrictKey(ov[r, theirKey]);
                if (sk != null && !strict.ContainsKey(sk)) strict[sk] = r;
                string lk = LooseKey(ov[r, theirKey]);
                if (lk == null) continue;
                if (loose.ContainsKey(lk)) repeated++;
                else loose[lk] = r;
            }

            int n = mr - 1, exact = 0, fixedUp = 0, missing = 0, blank = 0;
            object[,] outv = Xl.New2D(mr, bring.Count + 1);
            for (int j = 0; j < bring.Count; j++)
            {
                string hd = oh[bring[j]];
                outv[1, j + 1] = Array.IndexOf(mh, hd) > 0 ? hd + " (lookup)" : hd;
            }
            outv[1, bring.Count + 1] = "Match";
            List<long> paintFixed = new List<long>(), paintMissing = new List<long>();
            List<string> notFound = new List<string>();
            HashSet<string> notFoundSeen = new HashSet<string>();
            for (int r = 2; r <= mr; r++)
            {
                object key = mv[r, myKey];
                string sk = StrictKey(key), lk = LooseKey(key);
                int hit;
                string how;
                if (lk == null)
                {
                    blank++;
                    outv[r, bring.Count + 1] = "Blank key";
                    continue;
                }
                if (sk != null && strict.TryGetValue(sk, out hit))
                {
                    exact++;
                    how = "Exact";
                }
                else if (loose.TryGetValue(lk, out hit))
                {
                    fixedUp++;
                    how = "Found (" + Why(key, ov[hit, theirKey]) + ")";
                    paintFixed.Add(Xl.Cell(r, bring.Count + 1));
                }
                else
                {
                    missing++;
                    outv[r, bring.Count + 1] = "Not found";
                    paintMissing.Add(Xl.Cell(r, bring.Count + 1));
                    string show = Xl.Str(key).Trim();
                    if (notFoundSeen.Add(lk)) notFound.Add(show);
                    continue;
                }
                for (int j = 0; j < bring.Count; j++) outv[r, j + 1] = ov[hit, bring[j]];
                outv[r, bring.Count + 1] = how;
            }

            using (new Speed())
            {
                dynamic ws = mine.Worksheet;
                Xl.Backup(ws.Parent);
                dynamic target = mine.Cells[1, mc + 1].Resize[mr, bring.Count + 1];
                // never write over data: when the columns on the right are not empty, make room first
                double filled = 0;
                try { filled = Convert.ToDouble((object)Xl.App.WorksheetFunction.CountA(target), Inv); } catch { filled = 1; }
                if (filled > 0) target.EntireColumn.Insert();
                target = mine.Cells[1, mc + 1].Resize[mr, bring.Count + 1];
                Xl.WriteTable(target.Cells[1, 1], outv, 1);
                Xl.StyleHeader(target.Rows[1]);
                Xl.Paint(target, paintFixed, Design.Changed);
                Xl.Paint(target, paintMissing, Design.Bad);
                try { target.EntireColumn.AutoFit(); } catch { }
            }

            // how many of the missing keys Fuzzy Merge (Pro) would probably match: a quick estimate on a sample
            int likely = 0, tested = 0;
            if (notFound.Count > 0 && loose.Count <= 20000)
            {
                List<string> names = new List<string>();
                foreach (KeyValuePair<string, int> kv in loose) if (kv.Key.StartsWith("t:", StringComparison.Ordinal)) names.Add(Logic.NormName(kv.Key.Substring(2)));
                long budget = 3000000;
                foreach (string miss in notFound)
                {
                    if ((long)(tested + 1) * names.Count > budget) break;
                    tested++;
                    string a = Logic.NormName(miss);
                    if (a.Length < 3) continue;
                    foreach (string b in names)
                    {
                        if (b.Length == 0) continue;
                        if (Logic.JaroW(a, b) >= 0.9)
                        {
                            likely++;
                            break;
                        }
                    }
                }
                if (tested > 0 && tested < notFound.Count) likely = (int)Math.Round(likely * (double)notFound.Count / tested);
            }

            StringBuilder msg = new StringBuilder();
            msg.Append("Found ").Append(N(exact + fixedUp)).Append(" of ").Append(N(n)).Append(" rows.\n\n");
            msg.Append("Exact: ").Append(N(exact)).Append("\n");
            if (fixedUp > 0)
                msg.Append("Found after fixing spaces, capitals or numbers stored as text: ").Append(N(fixedUp))
                   .Append("   <- VLOOKUP / XLOOKUP would have shown #N/A for these (yellow)\n");
            if (missing > 0) msg.Append("Not found: ").Append(N(missing)).Append("   (red)\n");
            if (blank > 0) msg.Append("Rows with an empty key: ").Append(N(blank)).Append("\n");
            if (repeated > 0) msg.Append("\nNote: ").Append(N(repeated)).Append(" keys appear more than once in the second table. The first one was used, as VLOOKUP does.\n");
            msg.Append("\nThe columns were added on the right of your table, with a Match column. A backup was saved first.");
            msg.Append("\nTip: the formula =SMARTLOOKUP(key, key_column, return_column) does the same in a cell.");
            if (likely > 0)
            {
                msg.Append("\n\nAbout ").Append(N(likely)).Append(" of the ").Append(N(notFound.Count))
                   .Append(" keys not found look like names that are spelled a little differently. Fuzzy Merge can match them.");
                Nudge.SetContext("fuzzymerge", "About " + N(likely) + " of the keys that were not found look like the same names spelled differently " +
                                 "(Pvt / Private, typos, initials). Fuzzy Merge matches them and shows how sure it is for each row.");
            }
            Ui.Info(msg.ToString());
        }

        static string Why(object mine, object theirs)
        {
            bool num1 = mine is double, num2 = theirs is double;
            if (num1 != num2) return "number stored as text";
            string a = Xl.Str(mine), b = Xl.Str(theirs);
            if (a.Trim() != a || b.Trim() != b || a.Contains("  ") || b.Contains("  ") || a.IndexOf(' ') >= 0 || b.IndexOf(' ') >= 0) return "extra spaces";
            if (!string.Equals(a, b, StringComparison.Ordinal) && string.Equals(a, b, StringComparison.OrdinalIgnoreCase)) return "capitals";
            return "hidden characters";
        }

        static string[] Headers(object[,] v)
        {
            int nc = v.GetLength(1);
            string[] h = new string[nc + 1];
            h[0] = "";
            for (int c = 1; c <= nc; c++) h[c] = Xl.HeaderText(v[1, c], c);
            return h;
        }

        static string N(int n) { return n.ToString("#,##0", Inv); }
    }

    // A list with tick boxes; returns the 1-based columns ticked, or null when cancelled
    internal static class MultiPick
    {
        public static List<int> Show(string title, string hint, string[] headers, int skip)
        {
            using (Form f = new Form())
            {
                f.Text = title;
                f.StartPosition = FormStartPosition.CenterScreen;
                f.FormBorderStyle = FormBorderStyle.FixedDialog;
                f.MaximizeBox = false;
                f.MinimizeBox = false;
                f.ShowInTaskbar = false;
                f.Font = new Font("Segoe UI", 10f);
                f.ClientSize = new Size(480, 440);
                Label l = new Label();
                l.Text = hint;
                l.SetBounds(12, 10, 456, 40);
                f.Controls.Add(l);
                CheckedListBox list = new CheckedListBox();
                list.CheckOnClick = true;
                list.IntegralHeight = false;
                list.SetBounds(12, 52, 456, 330);
                List<int> map = new List<int>();
                for (int c = 1; c < headers.Length; c++)
                {
                    if (c == skip) continue;
                    list.Items.Add(Xl.ColLetter(c) + "   " + headers[c]);
                    map.Add(c);
                }
                f.Controls.Add(list);
                Button all = new Button();
                all.Text = "Tick all";
                all.SetBounds(12, 394, 100, 32);
                all.Click += delegate { for (int i = 0; i < list.Items.Count; i++) list.SetItemChecked(i, true); };
                f.Controls.Add(all);
                Button ok = new Button();
                ok.Text = "OK";
                ok.SetBounds(288, 394, 85, 32);
                ok.DialogResult = DialogResult.OK;
                f.Controls.Add(ok);
                Button cancel = new Button();
                cancel.Text = "Cancel";
                cancel.SetBounds(383, 394, 85, 32);
                cancel.DialogResult = DialogResult.Cancel;
                f.Controls.Add(cancel);
                f.AcceptButton = ok;
                f.CancelButton = cancel;
                if (f.ShowDialog(new ExcelWin()) != DialogResult.OK) return null;
                List<int> picked = new List<int>();
                foreach (int i in list.CheckedIndices) picked.Add(map[i]);
                if (picked.Count == 0) Ui.Info("Tick at least one column to bring.");
                return picked;
            }
        }
    }

    // The worksheet function. It appears in Insert Function under the DocBrisk category.
    public static class LookupFunctions
    {
        [ExcelFunction(Name = "SMARTLOOKUP", Category = Design.FunctionCategory,
                       Description = "Lookup without the #N/A from extra spaces, capitals or numbers stored as text. Optional: a value when nothing is found, and a fuzzy match for names.")]
        public static object SmartLookupFn(
            [ExcelArgument(Name = "lookup_value", Description = "the key to find")] object lookupValue,
            [ExcelArgument(Name = "lookup_range", Description = "one column (or row) to look in, for example Master!A:A")] object lookupRange,
            [ExcelArgument(Name = "return_range", Description = "the column (or row) to return from, the same size as lookup_range")] object returnRange,
            [ExcelArgument(Name = "if_not_found", Description = "what to show when nothing is found; #N/A if left out")] object ifNotFound,
            [ExcelArgument(Name = "fuzzy", Description = "TRUE: when nothing matches, take the closest name (90% alike or more)")] object fuzzy)
        {
            object notFound = (ifNotFound is ExcelMissing || ifNotFound is ExcelEmpty) ? (object)ExcelError.ExcelErrorNA : ifNotFound;
            object[,] look = lookupRange as object[,] ?? new object[,] { { lookupRange } };
            object[,] ret = returnRange as object[,] ?? new object[,] { { returnRange } };
            string want = SmartLookup.LooseKey(lookupValue);
            if (want == null) return notFound;
            int nr = look.GetLength(0), nc = look.GetLength(1);
            for (int i = 0; i < nr; i++)
            {
                for (int j = 0; j < nc; j++)
                {
                    if (SmartLookup.LooseKey(look[i, j]) == want) return Pick(ret, i, j, nr, nc);
                }
            }
            bool fz = fuzzy is bool && (bool)fuzzy || fuzzy is double && (double)fuzzy != 0;
            if (fz && want.StartsWith("t:", StringComparison.Ordinal))
            {
                string a = Logic.NormName(want.Substring(2));
                double best = 0;
                int bi = -1, bj = -1;
                for (int i = 0; i < nr; i++)
                {
                    for (int j = 0; j < nc; j++)
                    {
                        string t = look[i, j] as string;
                        if (t == null) continue;
                        t = Logic.NormName(t);
                        if (t.Length == 0) continue;
                        double sc = t == a ? 1 : Logic.JaroW(a, t);
                        if (sc > best) { best = sc; bi = i; bj = j; }
                    }
                }
                if (bi >= 0 && best >= 0.9) return Pick(ret, bi, bj, nr, nc);
            }
            return notFound;
        }

        // the cell of the return range that sits at the same place (a column returns by row, a row by column)
        static object Pick(object[,] ret, int i, int j, int nr, int nc)
        {
            int rr = ret.GetLength(0), rc = ret.GetLength(1);
            object v;
            if (nc == 1 && i < rr) v = ret[i, rc == 1 ? 0 : Math.Min(j, rc - 1)];
            else if (nr == 1 && j < rc) v = ret[rr == 1 ? 0 : Math.Min(i, rr - 1), j];
            else if (i < rr && j < rc) v = ret[i, j];
            else return ExcelError.ExcelErrorRef;
            if (v == null || v is ExcelEmpty) return "";
            return v;
        }
    }
}
