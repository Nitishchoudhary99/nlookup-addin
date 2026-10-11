using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;

namespace SmartPro
{
    // ABC analysis, ageing report, formula risk scan, snapshot / what changed
    internal static class ProTools
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        // headers (1-based) and which columns are mostly numbers / mostly dates
        static void Describe(object[,] v, int hr, out string[] heads, out bool[] numeric, out bool[] dates)
        {
            int nr = v.GetLength(0), nc = v.GetLength(1);
            heads = new string[nc + 1];
            numeric = new bool[nc + 1];
            dates = new bool[nc + 1];
            int last = Math.Min(nr, hr + 5000);
            for (int c = 1; c <= nc; c++)
            {
                heads[c] = Xl.HeaderText(v[hr, c], c);
                int nb = 0, nn = 0, nd = 0;
                for (int r = hr + 1; r <= last; r++)
                {
                    object x = v[r, c];
                    double d;
                    DateTime dt;
                    if (Xl.IsBlank(x) || Xl.IsErr(x)) continue;
                    nb++;
                    if (x is DateTime) nd++;
                    else if (Logic.TryNum(x, out d)) nn++;
                    else if (Logic.TryDate(x, out dt)) nd++;
                }
                numeric[c] = nb > 0 && nn >= 0.8 * nb;
                dates[c] = nb > 0 && nd >= 0.8 * nb;
            }
        }

        static bool Any(bool[] a)
        {
            foreach (bool b in a) if (b) return true;
            return false;
        }

        // =================================================================
        //  ABC ANALYSIS
        // =================================================================
        public static void AbcAnalysis()
        {
            dynamic src = Xl.PickRange("ABC Analysis");
            if (src == null) return;
            object[,] v = Xl.ReadTyped(src);
            if (v.GetLength(0) < 3 || v.GetLength(1) < 2)
            {
                Ui.Info("Select a table with a header row, an item column and a value column.");
                return;
            }
            int hr = Profiler.FindHeaderRow(v);
            string[] heads;
            bool[] numeric, dates;
            Describe(v, hr, out heads, out numeric, out dates);
            if (!Any(numeric))
            {
                Ui.Info("No number column was found. Run Text to Numbers first if the values are stored as text.");
                return;
            }
            int ic = Ui.PickColumn("ABC Analysis (1 of 2)", "Which column holds the ITEM (material, customer, SKU)?", heads, null);
            if (ic == 0) return;
            int vc = Ui.PickColumn("ABC Analysis (2 of 2)", "Which column holds the VALUE (amount, consumption value, quantity)?", heads, numeric);
            if (vc == 0) return;
            string msg;
            using (new Speed())
            {
                msg = AbcCore(src, v, hr, heads, ic, vc);
            }
            Ui.Info(msg);
        }

        static string AbcCore(dynamic src, object[,] v, int hr, string[] heads, int ic, int vc)
        {
            int nr = v.GetLength(0);
            Dictionary<string, double> sums = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            List<string> names = new List<string>();
            double total = 0;
            for (int r = hr + 1; r <= nr; r++)
            {
                double d;
                if (Xl.RowBlank(v, r) || !Logic.TryNum(v[r, vc], out d)) continue;
                string k = Xl.GroupKey(v[r, ic]);
                double cur;
                if (sums.TryGetValue(k, out cur)) sums[k] = cur + d;
                else
                {
                    sums[k] = d;
                    names.Add(k);
                }
                total += d;
            }
            int n = names.Count;
            if (n < 2 || total <= 0) return "There is not enough data: at least 2 different items with a positive total are needed.";
            names.Sort(delegate (string p, string q) { return sums[q].CompareTo(sums[p]); });

            object[,] t = Xl.New2D(n + 1, 6);
            t[1, 1] = "Rank";
            t[1, 2] = heads[ic];
            t[1, 3] = heads[vc];
            t[1, 4] = "Share";
            t[1, 5] = "Cumulative";
            t[1, 6] = "Class";
            int[] cnt = new int[3];
            double[] val = new double[3];
            double cum = 0;
            for (int i = 0; i < n; i++)
            {
                double x = sums[names[i]];
                int cls = cum < 0.8 * total ? 0 : (cum < 0.95 * total ? 1 : 2);   // the item that crosses 80% is still A
                cum += x;
                cnt[cls]++;
                val[cls] += x;
                t[i + 2, 1] = (double)(i + 1);
                t[i + 2, 2] = names[i];
                t[i + 2, 3] = x;
                t[i + 2, 4] = x / total;
                t[i + 2, 5] = cum / total;
                t[i + 2, 6] = cls == 0 ? "A" : (cls == 1 ? "B" : "C");
            }

            dynamic ws = Xl.AddSheet(src.Worksheet.Parent, "ABC Analysis");
            ws.Range["A1"].Value2 = "ABC Analysis: " + heads[vc] + " by " + heads[ic];
            ws.Range["A1"].Font.Size = 14;
            ws.Range["A1"].Font.Bold = true;
            object[,] s = Xl.New2D(4, 5);
            s[1, 1] = "Class";
            s[1, 2] = "Items";
            s[1, 3] = "Share of items";
            s[1, 4] = "Value";
            s[1, 5] = "Share of value";
            string[] label = { "A (first 80% of value)", "B (next 15%)", "C (last 5%)" };
            for (int k = 0; k < 3; k++)
            {
                s[k + 2, 1] = label[k];
                s[k + 2, 2] = (double)cnt[k];
                s[k + 2, 3] = (double)cnt[k] / n;
                s[k + 2, 4] = val[k];
                s[k + 2, 5] = val[k] / total;
            }
            Xl.WriteTable(ws.Range["A3"], s, 1);
            Xl.StyleHeader(ws.Range["A3:E3"]);
            ws.Range["C4:C6"].NumberFormat = "0.0%";
            ws.Range["E4:E6"].NumberFormat = "0.0%";
            ws.Range["D4:D6"].NumberFormat = "#,##0";

            Xl.WriteTable(ws.Range["A9"], t, 1);
            Xl.StyleHeader(ws.Range["A9:F9"]);
            ws.Range["C10"].Resize[n, 1].NumberFormat = "#,##0.00";
            ws.Range["D10"].Resize[n, 2].NumberFormat = "0.0%";
            int row = 10;
            int[] colors = { Design.Good, Design.Changed, Design.Bad };
            for (int k = 0; k < 3; k++)
            {
                if (cnt[k] == 0) continue;
                ws.Range["F" + row].Resize[cnt[k], 1].Interior.Color = colors[k];
                row += cnt[k];
            }
            ws.Columns["A:F"].AutoFit();

            // Pareto chart of the biggest items
            try
            {
                int top = Math.Min(n, 30);
                dynamic anchor = ws.Range["H3:R22"];
                dynamic co = ws.ChartObjects().Add((double)anchor.Left, (double)anchor.Top, (double)anchor.Width, (double)anchor.Height);
                dynamic ch = co.Chart;
                ch.ChartType = C.xlColumnClustered;
                dynamic s1 = ch.SeriesCollection().NewSeries();
                s1.Name = heads[vc];
                s1.Values = ws.Range["C10"].Resize[top, 1];
                s1.XValues = ws.Range["B10"].Resize[top, 1];
                dynamic s2 = ch.SeriesCollection().NewSeries();
                s2.Name = "Cumulative share";
                s2.Values = ws.Range["E10"].Resize[top, 1];
                s2.XValues = ws.Range["B10"].Resize[top, 1];
                s2.ChartType = C.xlLine;
                s2.AxisGroup = 2;
                ch.HasTitle = true;
                ch.ChartTitle.Text = "Pareto: top " + top + " by " + heads[vc];
                ch.HasLegend = false;
            }
            catch { }
            ws.Activate();
            return "ABC Analysis ready.\n\nA: " + cnt[0] + " items = " + (val[0] / total).ToString("0%", Inv) + " of the value\n" +
                   "B: " + cnt[1] + " items = " + (val[1] / total).ToString("0%", Inv) + "\n" +
                   "C: " + cnt[2] + " items = " + (val[2] / total).ToString("0%", Inv);
        }

        // =================================================================
        //  AGEING REPORT
        // =================================================================
        public static void AgeingReport()
        {
            dynamic src = Xl.PickRange("Ageing Report");
            if (src == null) return;
            object[,] v = Xl.ReadTyped(src);
            if (v.GetLength(0) < 2 || v.GetLength(1) < 3)
            {
                Ui.Info("Select a table with a party column, a date column and an amount column.");
                return;
            }
            int hr = Profiler.FindHeaderRow(v);
            string[] heads;
            bool[] numeric, dates;
            Describe(v, hr, out heads, out numeric, out dates);
            if (!Any(dates))
            {
                Ui.Info("No date column was found. Run Text to Dates first if the dates are stored as text.");
                return;
            }
            if (!Any(numeric))
            {
                Ui.Info("No number column was found. Run Text to Numbers first if the amounts are stored as text.");
                return;
            }
            int pc = Ui.PickColumn("Ageing Report (1 of 3)", "Which column holds the PARTY (customer, vendor, item)?", heads, null);
            if (pc == 0) return;
            int dc = Ui.PickColumn("Ageing Report (2 of 3)", "Which column holds the DATE (invoice date, due date, receipt date)?", heads, dates);
            if (dc == 0) return;
            int ac = Ui.PickColumn("Ageing Report (3 of 3)", "Which column holds the AMOUNT (outstanding, value, quantity)?", heads, numeric);
            if (ac == 0) return;
            string msg;
            using (new Speed())
            {
                msg = AgeingCore(src, v, hr, heads, pc, dc, ac);
            }
            Ui.Info(msg);
        }

        static string AgeingCore(dynamic src, object[,] v, int hr, string[] heads, int pc, int dc, int ac)
        {
            // per party: [0] not yet due, [1] 0-30, [2] 31-60, [3] 61-90, [4] 91-180, [5] over 180, [6] oldest age in days
            Dictionary<string, double[]> agg = new Dictionary<string, double[]>(StringComparer.OrdinalIgnoreCase);
            List<string> names = new List<string>();
            DateTime today = DateTime.Today;
            int nr = v.GetLength(0), skipped = 0;
            for (int r = hr + 1; r <= nr; r++)
            {
                if (Xl.RowBlank(v, r)) continue;
                double amt;
                DateTime dt;
                if (!Logic.TryNum(v[r, ac], out amt) || !Logic.TryDate(v[r, dc], out dt))
                {
                    skipped++;
                    continue;
                }
                string k = Xl.GroupKey(v[r, pc]);
                double[] b;
                if (!agg.TryGetValue(k, out b))
                {
                    b = new double[7];
                    agg[k] = b;
                    names.Add(k);
                }
                int days = (int)(today - dt.Date).TotalDays;
                int slot = days < 0 ? 0 : days <= 30 ? 1 : days <= 60 ? 2 : days <= 90 ? 3 : days <= 180 ? 4 : 5;
                b[slot] += amt;
                if (days > b[6]) b[6] = days;
            }
            int n = names.Count;
            if (n == 0) return "No rows had both a date and an amount.";
            Dictionary<string, double> tot = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            foreach (string k in names)
            {
                double[] b = agg[k];
                tot[k] = b[0] + b[1] + b[2] + b[3] + b[4] + b[5];
            }
            names.Sort(delegate (string p, string q) { return tot[q].CompareTo(tot[p]); });

            string[] cols = { heads[pc], "Not yet due", "0-30 days", "31-60 days", "61-90 days", "91-180 days", "Over 180 days", "Total", "Oldest (days)" };
            object[,] t = Xl.New2D(n + 2, 9);
            for (int j = 0; j < 9; j++) t[1, j + 1] = cols[j];
            double[] grand = new double[7];
            List<long> late = new List<long>();
            for (int i = 0; i < n; i++)
            {
                double[] b = agg[names[i]];
                t[i + 2, 1] = names[i];
                for (int j = 0; j < 6; j++)
                {
                    t[i + 2, j + 2] = b[j];
                    grand[j] += b[j];
                }
                t[i + 2, 8] = tot[names[i]];
                t[i + 2, 9] = b[6];
                if (b[6] > grand[6]) grand[6] = b[6];
                if (b[5] > 0) late.Add(Xl.Cell(i + 2, 7));
            }
            t[n + 2, 1] = "Total";
            double all = 0;
            for (int j = 0; j < 6; j++)
            {
                t[n + 2, j + 2] = grand[j];
                all += grand[j];
            }
            t[n + 2, 8] = all;
            t[n + 2, 9] = grand[6];

            dynamic ws = Xl.AddSheet(src.Worksheet.Parent, "Ageing Report");
            ws.Range["A1"].Value2 = "Ageing Report: " + heads[ac] + " by " + heads[pc];
            ws.Range["A1"].Font.Size = 14;
            ws.Range["A1"].Font.Bold = true;
            ws.Range["A2"].Value2 = "Age counted from " + heads[dc] + " to " + today.ToString("dd-MMM-yyyy", Inv) + ".";
            dynamic top = ws.Range["A4"];
            Xl.WriteTable(top, t, 1);
            dynamic tbl = top.Resize[n + 2, 9];
            Xl.StyleHeader(top.Resize[1, 9]);
            ws.Range["B5"].Resize[n + 1, 7].NumberFormat = "#,##0";
            ws.Range["I5"].Resize[n + 1, 1].NumberFormat = "0";
            Xl.Paint(tbl, late, Design.Bad);
            tbl.Rows[n + 2].Font.Bold = true;
            ws.Columns["A:I"].AutoFit();
            ws.Activate();
            double over = all == 0 ? 0 : grand[5] / all;
            return "Ageing Report ready for " + n + " parties.\n\nTotal: " + Logic.ShortINR(all) + "\nOver 180 days: " + Logic.ShortINR(grand[5]) + " (" + over.ToString("0%", Inv) + ")" +
                   (skipped > 0 ? "\n\n" + skipped + " rows were skipped because the date or the amount could not be read." : "");
        }

        // =================================================================
        //  FORMULA RISK SCAN
        // =================================================================
        static readonly Regex Quoted = new Regex("\"[^\"]*\"", RegexOptions.Compiled);
        static readonly Regex SheetRef = new Regex("'[^']*'!", RegexOptions.Compiled);
        static readonly Regex CellRef = new Regex(@"(?<![A-Za-z_])\$?[A-Za-z]{1,3}\$?\d{1,7}\b", RegexOptions.Compiled);
        static readonly Regex TypedNum = new Regex(@"(?<=[\*/\+\-\^])\s*(\d+(?:\.\d+)?)(?![\w\.\(])|(?<![\w\.])(\d+(?:\.\d+)?)\s*(?=[\*/\+\-\^%])", RegexOptions.Compiled);
        static readonly Regex VolatileFn = new Regex(@"\b(NOW|TODAY|RAND|RANDBETWEEN|OFFSET|INDIRECT)\s*\(", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        static readonly Regex ExtLink = new Regex(@"\[[^\[\]]+\.(xlsx|xlsm|xlsb|xls|csv)\]", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        static bool IsFormula(string s)
        {
            return s != null && s.Length > 1 && s[0] == '=';
        }

        static void AddIssue(List<object[]> rows, Dictionary<string, int> tally, string cell, string issue, string detail, string formula)
        {
            int k;
            tally[issue] = tally.TryGetValue(issue, out k) ? k + 1 : 1;
            if (rows.Count < 5000) rows.Add(new object[] { cell, issue + detail, formula });
        }

        public static void FormulaRiskScan()
        {
            dynamic ws = Xl.App.ActiveSheet;
            dynamic ur;
            long cells;
            try
            {
                ur = ws.UsedRange;
                cells = Xl.CellCount(ur);
            }
            catch
            {
                Ui.Info("Open a worksheet first.");
                return;
            }
            if (cells > 2000000)
            {
                Ui.Info("This sheet is very large (over 20 lakh cells). Scan a smaller sheet.");
                return;
            }
            string msg;
            using (new Speed())
            {
                msg = RiskCore(ws, ur);
            }
            Ui.Info(msg);
        }

        static string RiskCore(dynamic ws, dynamic ur)
        {
            Xl.Status("checking formulas...");
            object[,] v = Xl.ReadTyped(ur);
            object[,] f = Xl.ReadFormulas(ur);
            object[,] fr = Xl.ReadFormulasR1C1(ur);
            int nr = v.GetLength(0), nc = v.GetLength(1);
            int r0 = (int)ur.Row, c0 = (int)ur.Column;
            string sheet = (string)ws.Name;
            List<object[]> rows = new List<object[]>();
            Dictionary<string, int> tally = new Dictionary<string, int>();
            int nFormulas = 0;
            for (int c = 1; c <= nc; c++)
            {
                for (int r = 1; r <= nr; r++)
                {
                    string fa = f[r, c] as string;
                    string up = r > 1 ? fr[r - 1, c] as string : null;
                    string dn = r < nr ? fr[r + 1, c] as string : null;
                    bool neighbours = IsFormula(up) && IsFormula(dn) && up == dn;
                    string cell = Xl.A1(r0 + r - 1, c0 + c - 1);
                    if (IsFormula(fa))
                    {
                        nFormulas++;
                        if (Xl.IsErr(v[r, c])) AddIssue(rows, tally, cell, "Formula gives an error", ": " + Xl.Str(v[r, c]), fa);
                        if (ExtLink.IsMatch(fa)) AddIssue(rows, tally, cell, "Links to another workbook", "", fa);
                        string bare = Quoted.Replace(fa, "");
                        if (VolatileFn.IsMatch(bare)) AddIssue(rows, tally, cell, "Volatile function (recalculates on every change)", "", fa);
                        string noRefs = CellRef.Replace(SheetRef.Replace(bare, ""), "");
                        foreach (Match m in TypedNum.Matches(noRefs))
                        {
                            string num = m.Groups[1].Success ? m.Groups[1].Value : m.Groups[2].Value;
                            double d;
                            if (!double.TryParse(num, NumberStyles.Float, Inv, out d)) continue;
                            if (d == 0 || d == 1 || d == 100) continue;
                            AddIssue(rows, tally, cell, "Number typed inside the formula", ": " + num, fa);
                            break;
                        }
                        if (neighbours && (fr[r, c] as string) != up) AddIssue(rows, tally, cell, "Different from the formulas above and below", "", fa);
                    }
                    else if (neighbours && Xl.IsNum(v[r, c]))
                    {
                        AddIssue(rows, tally, cell, "Typed number inside a column of formulas (formula overwritten?)", "", Xl.Str(v[r, c]));
                    }
                }
            }
            if (nFormulas == 0) return "This sheet has no formulas to check.";
            int total = 0;
            foreach (KeyValuePair<string, int> kv in tally) total += kv.Value;
            if (total == 0) return nFormulas.ToString("#,##0", Inv) + " formulas checked on '" + sheet + "'. No risks found.";

            dynamic rep = Xl.AddSheet(ws.Parent, "Formula Risk");
            rep.Range["A1"].Value2 = "Formula Risk Scan: " + sheet;
            rep.Range["A1"].Font.Size = 14;
            rep.Range["A1"].Font.Bold = true;
            rep.Range["A2"].Value2 = nFormulas.ToString("#,##0", Inv) + " formulas checked, " + total.ToString("#,##0", Inv) + " points to review.";
            int row = 4;
            foreach (KeyValuePair<string, int> kv in tally)
            {
                rep.Cells[row, 1].Value2 = kv.Value;
                rep.Cells[row, 2].Value2 = kv.Key;
                row++;
            }
            row++;
            object[,] t = AnalyseTools.Table(new[] { "Cell", "What to review", "Formula / value" }, rows);
            Xl.WriteTable(rep.Cells[row, 1], t, 1);
            Xl.StyleHeader(rep.Cells[row, 1].Resize[1, 3]);
            string link = "'" + sheet.Replace("'", "''") + "'!";
            int links = Math.Min(rows.Count, 500);
            for (int i = 0; i < links; i++)
            {
                string addr = (string)rows[i][0];
                try { rep.Hyperlinks.Add(rep.Cells[row + 1 + i, 1], "", link + addr, Xl.M, addr); } catch { }
            }
            rep.Columns["A:B"].AutoFit();
            rep.Columns[3].ColumnWidth = 70;
            rep.Activate();
            return nFormulas.ToString("#,##0", Inv) + " formulas checked on '" + sheet + "'.\n" + total.ToString("#,##0", Inv) + " points to review are listed on the new 'Formula Risk' sheet.\nClick a cell address there to jump to it.";
        }

        // =================================================================
        //  SNAPSHOT / WHAT CHANGED
        // =================================================================
        static string PointerFile(string bookName)
        {
            return Path.Combine(Xl.BackupFolder(), "snapshot_" + Xl.SafeFileName(bookName) + ".txt");
        }

        public static void TakeSnapshot()
        {
            dynamic wb = Xl.App.ActiveWorkbook;
            string f = Xl.Backup(wb);
            if (f == null)
            {
                Ui.Warn("The snapshot could not be saved. Save the workbook once, then try again.");
                return;
            }
            File.WriteAllText(PointerFile((string)wb.Name), f);
            Ui.Info("Snapshot taken at " + DateTime.Now.ToString("dd-MMM-yyyy HH:mm", Inv) + ".\n\nWork as usual, then click What Changed? to see every cell that is different.");
        }

        // every formula / value of a sheet from A1 to the last used cell; null when the sheet is too large
        static object[,] ReadAll(dynamic sh)
        {
            dynamic ur = sh.UsedRange;
            int rows = (int)ur.Row + (int)ur.Rows.Count - 1;
            int cols = (int)ur.Column + (int)ur.Columns.Count - 1;
            if ((long)rows * cols > 3000000) return null;
            return Xl.ReadFormulas(sh.Range[sh.Cells[1, 1], sh.Cells[rows, cols]]);
        }

        public static void WhatChanged()
        {
            dynamic app = Xl.App;
            dynamic wb = app.ActiveWorkbook;
            string ptr = PointerFile((string)wb.Name);
            string snap = File.Exists(ptr) ? File.ReadAllText(ptr).Trim() : null;
            if (snap == null || !File.Exists(snap))
            {
                Ui.Info("Take a snapshot of this workbook first (Trust > Take Snapshot).");
                return;
            }
            string msg;
            using (new Speed())
            {
                msg = DiffCore(app, wb, snap);
            }
            Ui.Info(msg);
        }

        static string DiffCore(dynamic app, dynamic wb, string snap)
        {
            Xl.Status("comparing with the snapshot...");
            List<object[]> rows = new List<object[]>();
            int total = 0;
            object sec = null;
            try
            {
                sec = app.AutomationSecurity;
                app.AutomationSecurity = 3;
            }
            catch { }
            dynamic old = null;
            try
            {
                old = app.Workbooks.Open(snap, 0, true);
                Dictionary<string, object[,]> before = new Dictionary<string, object[,]>(StringComparer.OrdinalIgnoreCase);
                foreach (dynamic sh in old.Worksheets)
                {
                    string nm = (string)sh.Name;
                    object[,] arr = ReadAll(sh);
                    before[nm] = arr;
                }
                HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (dynamic sh in wb.Worksheets)
                {
                    string nm = (string)sh.Name;
                    if (nm.StartsWith("What Changed", StringComparison.OrdinalIgnoreCase)) continue;
                    seen.Add(nm);
                    object[,] was;
                    if (!before.TryGetValue(nm, out was))
                    {
                        total++;
                        rows.Add(new object[] { nm, "(whole sheet)", "", "New sheet" });
                        continue;
                    }
                    object[,] now = ReadAll(sh);
                    if (was == null || now == null)
                    {
                        rows.Add(new object[] { nm, "(whole sheet)", "", "Too large to compare" });
                        continue;
                    }
                    int ra = was.GetLength(0), ca = was.GetLength(1), rb = now.GetLength(0), cb = now.GetLength(1);
                    int mr = Math.Max(ra, rb), mc = Math.Max(ca, cb);
                    for (int r = 1; r <= mr; r++)
                    {
                        for (int c = 1; c <= mc; c++)
                        {
                            string x = (r <= ra && c <= ca) ? Xl.Str(was[r, c]) : "";
                            string y = (r <= rb && c <= cb) ? Xl.Str(now[r, c]) : "";
                            if (x == y) continue;
                            total++;
                            if (rows.Count < 20000) rows.Add(new object[] { nm, Xl.A1(r, c), x, y });
                        }
                    }
                }
                foreach (KeyValuePair<string, object[,]> kv in before)
                {
                    if (seen.Contains(kv.Key)) continue;
                    total++;
                    rows.Add(new object[] { kv.Key, "(whole sheet)", "Sheet existed", "Sheet deleted" });
                }
            }
            finally
            {
                if (old != null)
                {
                    try { old.Close(false); } catch { }
                }
                try { if (sec != null) app.AutomationSecurity = sec; } catch { }
            }
            string stamp = File.GetLastWriteTime(snap).ToString("dd-MMM-yyyy HH:mm", Inv);
            if (rows.Count == 0) return "Nothing has changed since the snapshot of " + stamp + ".";
            dynamic ws = Xl.AddSheet(wb, "What Changed");
            ws.Range["A1"].Value2 = "What changed since the snapshot of " + stamp;
            ws.Range["A1"].Font.Size = 14;
            ws.Range["A1"].Font.Bold = true;
            ws.Range["A2"].Value2 = total.ToString("#,##0", Inv) + " differences" + (total > rows.Count ? " (first " + rows.Count.ToString("#,##0", Inv) + " listed)" : "") + ". Formulas are compared as formulas, typed values as values.";
            object[,] t = AnalyseTools.Table(new[] { "Sheet", "Cell", "Before (snapshot)", "Now" }, rows);
            Xl.WriteTable(ws.Range["A4"], t, 1);
            Xl.StyleHeader(ws.Range["A4:D4"]);
            ws.Columns["A:B"].AutoFit();
            ws.Columns[3].ColumnWidth = 45;
            ws.Columns[4].ColumnWidth = 45;
            ws.Activate();
            return total.ToString("#,##0", Inv) + " cells are different from the snapshot of " + stamp + ".\nThey are listed on the new 'What Changed' sheet.";
        }
    }
}
