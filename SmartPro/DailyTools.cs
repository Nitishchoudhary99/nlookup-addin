using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Windows.Forms;

namespace SmartPro
{
    // Everyday jobs that formulas and Power Query do not do easily:
    //   Sum Finder ....... which of these amounts add up to this total?
    //   Reconcile ........ match two lists by reference and amount, with a rounding tolerance
    //   Report to Table .. turn a printed-style report (repeated headers, totals, gaps) into one clean table
    internal static class DailyTools
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        // =================================================================
        //  SUM FINDER
        // =================================================================
        sealed class Finder
        {
            readonly long[] a, suf;
            readonly int maxSol;
            readonly long maxMs;
            readonly int[] stack;
            readonly Stopwatch sw = new Stopwatch();
            long nodes;
            public readonly List<int[]> Found = new List<int[]>();
            public bool TimedOut;

            public Finder(long[] values, long[] suffix, int maxSol, long maxMs)
            {
                a = values;
                suf = suffix;
                this.maxSol = maxSol;
                this.maxMs = maxMs;
                stack = new int[values.Length];
            }

            public void Go(long target)
            {
                sw.Start();
                Dfs(0, target, 0);
            }

            // values are sorted big to small; returns true when the search should stop
            bool Dfs(int start, long remain, int depth)
            {
                if (remain == 0)
                {
                    int[] s = new int[depth];
                    Array.Copy(stack, s, depth);
                    Found.Add(s);
                    return Found.Count >= maxSol;
                }
                if ((++nodes & 4095) == 0 && sw.ElapsedMilliseconds > maxMs)
                {
                    TimedOut = true;
                    return true;
                }
                for (int i = start; i < a.Length; i++)
                {
                    if (suf[i] < remain) break;                        // everything that is left is not enough
                    if (a[i] > remain) continue;                       // too big for what is still missing
                    if (i > start && a[i] == a[i - 1]) continue;       // an equal amount at this step gives the same combination
                    stack[depth] = i;
                    if (Dfs(i + 1, remain - a[i], depth + 1)) return true;
                }
                return false;
            }
        }

        public static void SumFinder()
        {
            dynamic src = Xl.PickRange("Sum Finder", "Select the amounts (a column or any block that contains the numbers):");
            if (src == null) return;
            object[,] v = Xl.ReadTyped(src);
            string ans = Ui.AskText("Which total are you looking for?\nExample: a payment of 248560.50 that covers some of these invoices.", Design.AppName + " - Sum Finder", "");
            if (ans == null) return;
            double target;
            if (!Logic.TryNumText(ans.Trim(), out target) || target == 0)
            {
                Ui.Info("Type the total as a number, for example 248560.50");
                return;
            }
            bool neg = target < 0;
            long t = (long)Math.Round(Math.Abs(target) * 100.0);
            int nr = v.GetLength(0), nc = v.GetLength(1);
            List<long> val = new List<long>();
            List<int> rr = new List<int>(), cc = new List<int>();
            for (int c = 1; c <= nc; c++)
            {
                for (int r = 1; r <= nr; r++)
                {
                    if (!Xl.IsNum(v[r, c])) continue;
                    double d = Xl.Num(v[r, c]);
                    if (neg) d = -d;
                    long p = (long)Math.Round(d * 100.0);
                    if (p <= 0 || p > t) continue;                    // cannot be part of the total
                    val.Add(p);
                    rr.Add(r);
                    cc.Add(c);
                }
            }
            int n = val.Count;
            if (n == 0)
            {
                Ui.Info("The selection has no amounts that could be part of that total.");
                return;
            }
            if (n > 5000)
            {
                Ui.Info("The selection has " + n.ToString("#,##0", Inv) + " usable amounts. Select up to 5,000 at a time.");
                return;
            }
            // sort big to small, keeping track of the cells
            long[] a = val.ToArray();
            int[] order = new int[n];
            for (int i = 0; i < n; i++) order[i] = i;
            Array.Sort(a, order);
            Array.Reverse(a);
            Array.Reverse(order);
            long[] suf = new long[n + 1];
            for (int i = n - 1; i >= 0; i--) suf[i] = suf[i + 1] + a[i];
            if (suf[0] < t)
            {
                Ui.Info("Even all the amounts together (" + Logic.IndianGroup(suf[0] / 100.0, 2) + ") are less than the total.");
                return;
            }
            Xl.Status("looking for combinations...");
            Finder f = new Finder(a, suf, 5, 6000);
            f.Go(t);
            try { Xl.App.StatusBar = false; } catch { }
            if (f.Found.Count == 0)
            {
                Ui.Info(f.TimedOut
                    ? "No combination was found in 6 seconds. With this many amounts there are too many possibilities to try them all: select fewer amounts (for example one party or one month) and try again."
                    : "No combination of these amounts adds up to exactly " + Logic.IndianGroup(Math.Abs(target), 2) + ".");
                return;
            }
            string msg;
            using (new Speed())
            {
                msg = WriteSums(src, f, a, order, rr, cc, target);
            }
            Ui.Info(msg);
        }

        static string WriteSums(dynamic src, Finder f, long[] a, int[] order, List<int> rr, List<int> cc, double target)
        {
            int r0 = (int)src.Row, c0 = (int)src.Column;
            // the first combination is marked on the sheet
            List<long> mark = new List<long>();
            foreach (int i in f.Found[0]) mark.Add(Xl.Cell(rr[order[i]], cc[order[i]]));
            Xl.Paint(src, mark, Design.Good);
            List<object[]> rows = new List<object[]>();
            for (int k = 0; k < f.Found.Count; k++)
            {
                foreach (int i in f.Found[k])
                {
                    int o = order[i];
                    rows.Add(new object[] { (double)(k + 1), Xl.A1(r0 + rr[o] - 1, c0 + cc[o] - 1), a[i] / 100.0 * (target < 0 ? -1 : 1) });
                }
            }
            dynamic ws = Xl.AddSheet(src.Worksheet.Parent, "Sum Finder");
            ws.Range["A1"].Value2 = "Sum Finder: amounts that add up to " + Logic.IndianGroup(target, 2);
            ws.Range["A1"].Font.Size = 14;
            ws.Range["A1"].Font.Bold = true;
            ws.Range["A2"].Value2 = f.Found.Count + " combination(s) listed" + (f.Found.Count >= 5 ? " (the search stops at 5)" : (f.TimedOut ? " (the search stopped after 6 seconds, there may be more)" : "")) +
                                    ". Combination 1 is marked green on your sheet.";
            object[,] t = AnalyseTools.Table(new[] { "Combination", "Cell", "Amount" }, rows);
            Xl.WriteTable(ws.Range["A4"], t, 1);
            Xl.StyleHeader(ws.Range["A4:C4"]);
            ws.Range["C5"].Resize[rows.Count, 1].NumberFormat = "#,##0.00";
            ws.Columns["A:C"].AutoFit();
            ws.Activate();
            return f.Found.Count + " combination(s) found for " + Logic.IndianGroup(target, 2) + ".\n\nCombination 1 uses " + f.Found[0].Length + " amounts and is marked green on your sheet.\nAll combinations are listed on the new 'Sum Finder' sheet.";
        }

        // =================================================================
        //  RECONCILE
        // =================================================================
        static string[] Heads(object[,] v)
        {
            int nc = v.GetLength(1);
            string[] h = new string[nc + 1];
            for (int c = 1; c <= nc; c++) h[c] = Xl.HeaderText(v[1, c], c);
            return h;
        }

        static bool[] NumericCols(object[,] v)
        {
            int nr = v.GetLength(0), nc = v.GetLength(1);
            bool[] ok = new bool[nc + 1];
            int last = Math.Min(nr, 3000);
            for (int c = 1; c <= nc; c++)
            {
                int nb = 0, nn = 0;
                for (int r = 2; r <= last; r++)
                {
                    double d;
                    object x = v[r, c];
                    if (Xl.IsBlank(x) || Xl.IsErr(x) || x is DateTime) continue;
                    nb++;
                    if (Logic.TryNum(x, out d)) nn++;
                }
                ok[c] = nb > 0 && nn >= 0.8 * nb;
            }
            return ok;
        }

        // reference without spaces, punctuation, case and leading zeros: "INV/0012-A" -> "INV12A"
        static string CleanRef(object x)
        {
            if (Xl.IsBlank(x) || Xl.IsErr(x)) return "";
            string s = Xl.Str(x);
            StringBuilder sb = new StringBuilder(s.Length);
            foreach (char raw in s)
            {
                char ch = char.ToUpperInvariant(raw);
                if ((ch >= 'A' && ch <= 'Z') || (ch >= '0' && ch <= '9')) sb.Append(ch);
            }
            string k = sb.ToString().TrimStart('0');
            return k.Length == 0 && sb.Length > 0 ? "0" : k;
        }

        public static void Reconcile()
        {
            dynamic ra = Xl.PickRange("Reconcile (1 of 2)", "Select the FIRST table with its header row (for example your books):");
            if (ra == null) return;
            object[,] av = Xl.ReadTyped(ra);
            if (av.GetLength(0) < 2 || av.GetLength(1) < 2)
            {
                Ui.Info("Select a table with a header row, a reference column and an amount column.");
                return;
            }
            string[] ah = Heads(av);
            int aRef = Ui.PickColumn("Reconcile: first table", "Which column holds the reference (invoice no., document no., cheque no.)?", ah, null);
            if (aRef == 0) return;
            int aAmt = Ui.PickColumn("Reconcile: first table", "Which column holds the amount?", ah, NumericCols(av));
            if (aAmt == 0) return;
            dynamic rb = Xl.PickRange("Reconcile (2 of 2)", "Select the SECOND table with its header row (for example the party statement, bank or GST portal data):");
            if (rb == null) return;
            object[,] bv = Xl.ReadTyped(rb);
            if (bv.GetLength(0) < 2 || bv.GetLength(1) < 2)
            {
                Ui.Info("The second table needs a header row, a reference column and an amount column.");
                return;
            }
            string[] bh = Heads(bv);
            int bRef = Ui.PickColumn("Reconcile: second table", "Which column holds the reference?", bh, null);
            if (bRef == 0) return;
            int bAmt = Ui.PickColumn("Reconcile: second table", "Which column holds the amount?", bh, NumericCols(bv));
            if (bAmt == 0) return;
            string ans = Ui.AskText("Allowed difference in the amount (for rounding).\nType 0 for exact, 1 for up to 1 rupee.", Design.AppName + " - Reconcile", "1");
            if (ans == null) return;
            double tol;
            if (!Logic.TryNumText(ans.Trim(), out tol) || tol < 0) tol = 0;
            string msg;
            using (new Speed())
            {
                msg = ReconcileCore(ra, rb, av, bv, aRef, aAmt, bRef, bAmt, tol);
            }
            Ui.Info(msg);
        }

        static string ReconcileCore(dynamic ra, dynamic rb, object[,] av, object[,] bv, int aRef, int aAmt, int bRef, int bAmt, double tol)
        {
            Xl.Status("reconciling...");
            int na = av.GetLength(0), nb = bv.GetLength(0);
            int a0 = (int)ra.Row, b0 = (int)rb.Row;
            string[] aKey = new string[na + 1], bKey = new string[nb + 1];
            double[] aVal = new double[na + 1], bVal = new double[nb + 1];
            bool[] aHas = new bool[na + 1], bHas = new bool[nb + 1], aRow = new bool[na + 1], bRow = new bool[nb + 1];
            int[] aTo = new int[na + 1], bTo = new int[nb + 1];           // matched partner row, 0 = none
            int[] aHow = new int[na + 1];                                 // 1 matched, 2 reference only, 3 amount only
            for (int i = 2; i <= na; i++)
            {
                if (Xl.RowBlank(av, i)) continue;
                aRow[i] = true;
                aKey[i] = CleanRef(av[i, aRef]);
                aHas[i] = Logic.TryNum(av[i, aAmt], out aVal[i]);
            }
            Dictionary<string, List<int>> byRef = new Dictionary<string, List<int>>(StringComparer.Ordinal);
            for (int j = 2; j <= nb; j++)
            {
                if (Xl.RowBlank(bv, j)) continue;
                bRow[j] = true;
                bKey[j] = CleanRef(bv[j, bRef]);
                bHas[j] = Logic.TryNum(bv[j, bAmt], out bVal[j]);
                if (bKey[j].Length == 0) continue;
                List<int> list;
                if (!byRef.TryGetValue(bKey[j], out list))
                {
                    list = new List<int>();
                    byRef[bKey[j]] = list;
                }
                list.Add(j);
            }
            double eps = tol + 0.0000001;
            // pass 1: same reference and same amount
            for (int i = 2; i <= na; i++)
            {
                List<int> list;
                if (!aRow[i] || aKey[i].Length == 0 || !aHas[i] || !byRef.TryGetValue(aKey[i], out list)) continue;
                foreach (int j in list)
                {
                    if (bTo[j] != 0 || !bHas[j] || Math.Abs(aVal[i] - bVal[j]) > eps) continue;
                    aTo[i] = j;
                    bTo[j] = i;
                    aHow[i] = 1;
                    break;
                }
            }
            // pass 2: same reference, different amount
            for (int i = 2; i <= na; i++)
            {
                List<int> list;
                if (!aRow[i] || aTo[i] != 0 || aKey[i].Length == 0 || !byRef.TryGetValue(aKey[i], out list)) continue;
                foreach (int j in list)
                {
                    if (bTo[j] != 0) continue;
                    aTo[i] = j;
                    bTo[j] = i;
                    aHow[i] = 2;
                    break;
                }
            }
            // pass 3: same amount, different reference (closest amount within the tolerance)
            List<int> free = new List<int>();
            for (int j = 2; j <= nb; j++) if (bRow[j] && bTo[j] == 0 && bHas[j]) free.Add(j);
            int nf = free.Count;
            double[] fv = new double[nf];
            int[] fj = free.ToArray();
            for (int k = 0; k < nf; k++) fv[k] = bVal[fj[k]];
            Array.Sort(fv, fj);
            for (int i = 2; i <= na; i++)
            {
                if (!aRow[i] || aTo[i] != 0 || !aHas[i] || nf == 0) continue;
                int lo = Array.BinarySearch(fv, aVal[i] - eps);
                if (lo < 0) lo = ~lo;
                while (lo > 0 && fv[lo - 1] >= aVal[i] - eps) lo--;
                int best = -1;
                double bestDiff = double.MaxValue;
                for (int k = lo; k < nf && fv[k] <= aVal[i] + eps; k++)
                {
                    if (bTo[fj[k]] != 0) continue;
                    double diff = Math.Abs(fv[k] - aVal[i]);
                    if (diff < bestDiff)
                    {
                        bestDiff = diff;
                        best = k;
                    }
                }
                if (best < 0) continue;
                aTo[i] = fj[best];
                bTo[fj[best]] = i;
                aHow[i] = 3;
            }

            // ---- report: problems first, matched rows last ----
            string[] label = { "", "Matched", "Reference matches, amount differs", "Amount matches, reference differs", "Only in first table", "Only in second table" };
            int[] cnt = new int[6];
            double[] sumA = new double[6], sumB = new double[6];
            List<object[]> rows = new List<object[]>();
            int[] sequence = { 2, 3, 4, 5, 1 };
            foreach (int st in sequence)
            {
                if (st == 5)
                {
                    for (int j = 2; j <= nb; j++)
                    {
                        if (!bRow[j] || bTo[j] != 0) continue;
                        cnt[5]++;
                        if (bHas[j]) sumB[5] += bVal[j];
                        rows.Add(new object[] { label[5], null, null, AnalyseTools.Disp(bv[j, bRef]), bHas[j] ? (object)bVal[j] : null, null, null, (double)(b0 + j - 1) });
                    }
                    continue;
                }
                for (int i = 2; i <= na; i++)
                {
                    if (!aRow[i]) continue;
                    int how = aTo[i] == 0 ? 4 : aHow[i];
                    if (how != st) continue;
                    cnt[st]++;
                    if (aHas[i]) sumA[st] += aVal[i];
                    int j = aTo[i];
                    if (j == 0)
                    {
                        rows.Add(new object[] { label[4], AnalyseTools.Disp(av[i, aRef]), aHas[i] ? (object)aVal[i] : null, null, null, null, (double)(a0 + i - 1), null });
                        continue;
                    }
                    if (bHas[j]) sumB[st] += bVal[j];
                    object diff = aHas[i] && bHas[j] ? (object)(aVal[i] - bVal[j]) : null;
                    rows.Add(new object[] { label[st], AnalyseTools.Disp(av[i, aRef]), aHas[i] ? (object)aVal[i] : null, AnalyseTools.Disp(bv[j, bRef]), bHas[j] ? (object)bVal[j] : null, diff, (double)(a0 + i - 1), (double)(b0 + j - 1) });
                }
            }
            if (rows.Count == 0) return "Both tables are empty under their headers.";

            dynamic ws = Xl.AddSheet(rb.Worksheet.Parent, "Reconcile");
            ws.Range["A1"].Value2 = "Reconcile: first table against second table";
            ws.Range["A1"].Font.Size = 14;
            ws.Range["A1"].Font.Bold = true;
            object[,] sm = Xl.New2D(6, 4);
            sm[1, 1] = "Result";
            sm[1, 2] = "Rows";
            sm[1, 3] = "Amount in first table";
            sm[1, 4] = "Amount in second table";
            for (int k = 1; k <= 5; k++)
            {
                sm[k + 1, 1] = label[k];
                sm[k + 1, 2] = (double)cnt[k];
                sm[k + 1, 3] = sumA[k];
                sm[k + 1, 4] = sumB[k];
            }
            Xl.WriteTable(ws.Range["A3"], sm, 1);
            Xl.StyleHeader(ws.Range["A3:D3"]);
            ws.Range["C4:D8"].NumberFormat = "#,##0.00";

            string[] heads = { "Result", "First: reference", "First: amount", "Second: reference", "Second: amount", "Difference", "First: sheet row", "Second: sheet row" };
            object[,] t = Xl.New2D(rows.Count + 1, 8);
            for (int c = 0; c < 8; c++) t[1, c + 1] = heads[c];
            List<long> yellow = new List<long>(), red = new List<long>();
            for (int r = 0; r < rows.Count; r++)
            {
                for (int c = 0; c < 8; c++) t[r + 2, c + 1] = rows[r][c];
                string st = (string)rows[r][0];
                if (st == label[2] || st == label[3]) yellow.Add(Xl.Cell(r + 2, 1));
                else if (st == label[4] || st == label[5]) red.Add(Xl.Cell(r + 2, 1));
            }
            dynamic top = ws.Range["A11"];
            Xl.WriteTable(top, t, 1);
            dynamic tbl = top.Resize[rows.Count + 1, 8];
            Xl.StyleHeader(top.Resize[1, 8]);
            tbl.Columns[3].NumberFormat = "#,##0.00";
            tbl.Columns[5].NumberFormat = "#,##0.00";
            tbl.Columns[6].NumberFormat = "#,##0.00";
            Xl.Paint(tbl, yellow, Design.Changed);
            Xl.Paint(tbl, red, Design.Bad);
            ws.Columns["A:H"].AutoFit();
            ws.Activate();
            return "Reconcile finished.\n\nMatched: " + cnt[1].ToString("#,##0", Inv) +
                   "\nReference matches, amount differs: " + cnt[2].ToString("#,##0", Inv) +
                   "\nAmount matches, reference differs: " + cnt[3].ToString("#,##0", Inv) +
                   "\nOnly in first table: " + cnt[4].ToString("#,##0", Inv) +
                   "\nOnly in second table: " + cnt[5].ToString("#,##0", Inv) +
                   "\n\nThe problems are listed first on the new 'Reconcile' sheet.";
        }

        // =================================================================
        //  REPORT TO TABLE
        // =================================================================
        static bool IsRule(string s)
        {
            if (s.Length == 0) return false;
            foreach (char ch in s) if (ch != '-' && ch != '=' && ch != '_' && ch != '*' && ch != '.' && ch != ' ') return false;
            return true;
        }

        public static void ReportToTable()
        {
            dynamic src = Xl.PickRange("Report to Table", "Select the whole report: titles, repeated headers, totals and all.");
            if (src == null) return;
            object[,] v = Xl.ReadTyped(src);
            if (v.GetLength(0) < 3 || v.GetLength(1) < 2)
            {
                Ui.Info("Select a report with a header row and data rows.");
                return;
            }
            DialogResult fill = Ui.YesNoCancel("Reports often print a customer or material only on its first line.\n\nFill those blank cells with the value above?\n\nYes = fill down\nNo = leave them blank");
            if (fill == DialogResult.Cancel) return;
            string msg;
            using (new Speed())
            {
                msg = ReportCore(src, v, fill == DialogResult.Yes);
            }
            Ui.Info(msg);
        }

        static string ReportCore(dynamic src, object[,] v, bool fillDown)
        {
            Xl.Status("cleaning the report...");
            int nr = v.GetLength(0), nc = v.GetLength(1);
            int hr = Profiler.FindHeaderRow(v);
            string[] h = Profiler.MakeHeaders(v, hr);
            ColInfo[] ci = Profiler.Profile(v, hr, h, 50000);
            string[] rawHead = new string[nc + 1];
            int headCells = 0;
            for (int c = 1; c <= nc; c++)
            {
                rawHead[c] = Xl.IsBlank(v[hr, c]) ? "" : Xl.Str(v[hr, c]).Trim();
                if (rawHead[c].Length > 0) headCells++;
            }
            List<int> cols = new List<int>();
            int firstNum = nc + 1;
            for (int c = 1; c <= nc; c++)
            {
                if (ci[c].Kind == "empty") continue;
                cols.Add(c);
                if (ci[c].Kind == "num" && c < firstNum) firstNum = c;
            }
            if (cols.Count == 0) return "The selection has no data under the header.";

            // rows to keep
            List<int> keep = new List<int>();
            int nHead = 0, nTotal = 0, nRule = 0;
            for (int r = hr + 1; r <= nr; r++)
            {
                if (Xl.RowBlank(v, r)) continue;
                int filled = 0, same = 0, rule = 0;
                for (int c = 1; c <= nc; c++)
                {
                    if (Xl.IsBlank(v[r, c])) continue;
                    filled++;
                    string s = Xl.Str(v[r, c]).Trim();
                    if (rawHead[c].Length > 0 && string.Equals(s, rawHead[c], StringComparison.OrdinalIgnoreCase)) same++;
                    if (IsRule(s)) rule++;
                }
                if (rule == filled)
                {
                    nRule++;
                    continue;
                }
                if (headCells > 0 && same >= 2 && same >= 0.6 * filled)
                {
                    nHead++;                                           // the header printed again
                    continue;
                }
                if (!Profiler.RowIsData(v, r, nc))
                {
                    nTotal++;                                          // total / subtotal line
                    continue;
                }
                keep.Add(r);
            }
            if (keep.Count == 0) return "No data rows were left after removing headers, totals and separator lines.";

            object[,] o = Xl.New2D(keep.Count + 1, cols.Count);
            int nFilled = 0;
            for (int k = 0; k < cols.Count; k++)
            {
                int c = cols[k];
                string kind = ci[c].Kind;
                o[1, k + 1] = h[c];
                // label columns on the left of the first number column that are often blank are filled down
                bool fillThis = fillDown && c < firstNum && kind != "num" && kind != "date" && ci[c].Blank >= 0.2 * (ci[c].Blank + ci[c].NonBlank);
                object lastSeen = null;
                for (int i = 0; i < keep.Count; i++)
                {
                    object x = v[keep[i], c];
                    object outv = null;
                    double d;
                    DateTime dt;
                    if (Xl.IsBlank(x))
                    {
                        if (fillThis && lastSeen != null)
                        {
                            outv = lastSeen;
                            nFilled++;
                        }
                    }
                    else if (kind == "num" && Logic.TryNum(x, out d)) outv = d;
                    else if (kind == "date" && Logic.TryDate(x, out dt)) outv = dt;
                    else if (x is string) outv = Logic.CleanText((string)x);
                    else outv = x;
                    if (!Xl.IsBlank(x)) lastSeen = outv;
                    o[i + 2, k + 1] = outv;
                }
            }
            dynamic ws = Xl.AddSheet(src.Worksheet.Parent, "Clean Table");
            Xl.WriteTable(ws.Range["A1"], o, 1);
            Xl.StyleHeader(ws.Range["A1"].Resize[1, cols.Count]);
            try { ws.Range["A1"].Resize[1, cols.Count].EntireColumn.AutoFit(); } catch { }
            ws.Activate();
            return "Clean table ready: " + keep.Count.ToString("#,##0", Inv) + " rows, " + cols.Count + " columns, on the new 'Clean Table' sheet.\n\nRemoved:\n" +
                   nHead.ToString("#,##0", Inv) + " repeated header rows\n" + nTotal.ToString("#,##0", Inv) + " total / subtotal rows\n" + nRule.ToString("#,##0", Inv) + " separator lines\n" +
                   (nc - cols.Count) + " empty columns" + (fillDown ? "\n\nFilled down: " + nFilled.ToString("#,##0", Inv) + " cells" : "") +
                   "\n\nYour report itself was not changed.";
        }
    }
}
