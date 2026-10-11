using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows.Forms;

namespace SmartPro
{
    internal static class AnalyseTools
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        // header row + data rows -> 1-based table for Xl.WriteTable
        public static object[,] Table(string[] heads, List<object[]> rows)
        {
            object[,] t = Xl.New2D(rows.Count + 1, heads.Length);
            for (int j = 0; j < heads.Length; j++) t[1, j + 1] = heads[j];
            for (int i = 0; i < rows.Count; i++)
                for (int j = 0; j < heads.Length; j++) t[i + 2, j + 1] = rows[i][j];
            return t;
        }

        // value shown in a report: dates and errors as text, everything else as it is
        public static object Disp(object v)
        {
            if (v is DateTime || Xl.IsErr(v)) return Xl.Str(v);
            return v;
        }

        static bool Same(object x, object y)
        {
            if (x == null || y == null) return Xl.Str(x).Trim() == Xl.Str(y).Trim();
            if (Xl.IsNum(x) && Xl.IsNum(y)) return Math.Abs(Xl.Num(x) - Xl.Num(y)) < 0.0000001;
            if (x is DateTime && y is DateTime) return (DateTime)x == (DateTime)y;
            return Xl.Str(x).Trim() == Xl.Str(y).Trim();
        }

        // ---------------------------------------------------------------
        public static void CompareByKey()
        {
            dynamic ra = Xl.PickRange("Compare by Key (1 of 2)", "Select the OLD table with its header row. The first column is the key (ID, invoice no., code):");
            if (ra == null) return;
            dynamic rb = Xl.PickRange("Compare by Key (2 of 2)", "Now select the NEW table with its header row. You can click another sheet or workbook:");
            if (rb == null) return;
            if ((int)ra.Rows.Count < 2 || (int)rb.Rows.Count < 2)
            {
                Ui.Info("Both tables need a header row and at least one data row.");
                return;
            }
            DialogResult hl = Ui.YesNoCancel("Also colour the differences in the NEW table?\nYellow = changed cell, green = new row.");
            if (hl == DialogResult.Cancel) return;
            bool paint = hl == DialogResult.Yes;
            if (paint) Xl.Backup(rb.Worksheet.Parent);
            string msg;
            using (new Speed())
            {
                msg = CompareCore(ra, rb, paint);
            }
            Ui.Info(msg);
        }

        static string CompareCore(dynamic ra, dynamic rb, bool paint)
        {
            Xl.Status("comparing...");
            object[,] a = Xl.ReadTyped(ra);
            object[,] b = Xl.ReadTyped(rb);
            int na = a.GetLength(0), ca = a.GetLength(1), nb = b.GetLength(0), cb = b.GetLength(1);

            Dictionary<string, int> ha = new Dictionary<string, int>();
            for (int j = 1; j <= ca; j++)
            {
                string k = Xl.Key(a[1, j]);
                if (k.Length > 0 && !ha.ContainsKey(k)) ha[k] = j;
            }
            int[] colMap = new int[cb + 1];
            for (int j = 1; j <= cb; j++)
            {
                int cj;
                if (ha.TryGetValue(Xl.Key(b[1, j]), out cj)) colMap[j] = cj;
            }
            Dictionary<string, int> da = new Dictionary<string, int>();
            int dupA = 0, dupB = 0, nAdd = 0, nRem = 0, nChg = 0, nSame = 0;
            for (int i = 2; i <= na; i++)
            {
                string k = Xl.Key(a[i, 1]);
                if (k.Length == 0) continue;
                if (da.ContainsKey(k)) dupA++;
                else da[k] = i;
            }
            HashSet<string> db = new HashSet<string>();
            List<object[]> rows = new List<object[]>();
            List<long> yellow = new List<long>(), green = new List<long>();
            for (int i = 2; i <= nb; i++)
            {
                string k = Xl.Key(b[i, 1]);
                if (k.Length == 0) continue;
                if (!db.Add(k))
                {
                    dupB++;
                    continue;
                }
                int ia;
                if (!da.TryGetValue(k, out ia))
                {
                    nAdd++;
                    if (rows.Count < 500000) rows.Add(new object[] { Disp(b[i, 1]), "Added (only in NEW)", null, null, null });
                    if (paint) for (int c = 1; c <= cb; c++) green.Add(Xl.Cell(i, c));
                    continue;
                }
                bool changed = false;
                for (int j = 2; j <= cb; j++)
                {
                    if (colMap[j] == 0) continue;
                    object oldV = a[ia, colMap[j]], newV = b[i, j];
                    if (Same(oldV, newV)) continue;
                    changed = true;
                    nChg++;
                    if (rows.Count < 500000) rows.Add(new object[] { Disp(b[i, 1]), "Changed", Xl.HeaderText(b[1, j], j), Disp(oldV), Disp(newV) });
                    if (paint) yellow.Add(Xl.Cell(i, j));
                }
                if (!changed) nSame++;
            }
            foreach (KeyValuePair<string, int> kv in da)
            {
                if (db.Contains(kv.Key)) continue;
                nRem++;
                if (rows.Count < 500000) rows.Add(new object[] { Disp(a[kv.Value, 1]), "Removed (only in OLD)", null, null, null });
            }
            if (paint)
            {
                Xl.Paint(rb, green, Design.Good);
                Xl.Paint(rb, yellow, Design.Changed);
            }

            dynamic wb = rb.Worksheet.Parent;
            dynamic ws = Xl.AddSheet(wb, "Compare Report");
            ws.Range["A1"].Value2 = "Compare by Key";
            ws.Range["A1"].Font.Size = 14;
            ws.Range["A1"].Font.Bold = true;
            ws.Range["A3"].Value2 = "Rows only in NEW";
            ws.Range["B3"].Value2 = nAdd;
            ws.Range["A4"].Value2 = "Rows only in OLD";
            ws.Range["B4"].Value2 = nRem;
            ws.Range["A5"].Value2 = "Changed cells";
            ws.Range["B5"].Value2 = nChg;
            ws.Range["A6"].Value2 = "Rows exactly the same";
            ws.Range["B6"].Value2 = nSame;
            ws.Range["A7"].Value2 = "Repeated keys ignored (OLD / NEW)";
            ws.Range["B7"].Value2 = "'" + dupA + " / " + dupB;
            if (rows.Count > 0)
            {
                object[,] t = Table(new[] { "Key", "What happened", "Column", "Old value", "New value" }, rows);
                Xl.WriteTable(ws.Range["A9"], t, 1);
                Xl.StyleHeader(ws.Range["A9:E9"]);
            }
            else ws.Range["A9"].Value2 = "No differences found.";
            ws.Columns["A:E"].AutoFit();
            ws.Activate();
            return "Compare finished.\n\n" + nAdd + " rows only in NEW\n" + nRem + " rows only in OLD\n" + nChg + " changed cells\n" + nSame + " rows exactly the same";
        }

        // ---------------------------------------------------------------
        public static void FuzzyDuplicates()
        {
            dynamic src = Xl.PickRange("Fuzzy Duplicates", "Select the column of names to check (customer, vendor, item names):");
            if (src == null) return;
            string ans = Ui.AskText("How similar must two names be to count as a match?\nType 70 to 99.  88 is a good start; lower finds more.", Design.AppName, "88");
            if (ans == null) return;
            double thr;
            if (!double.TryParse(ans, NumberStyles.Any, Inv, out thr)) thr = 88;
            thr /= 100;
            if (thr < 0.5 || thr > 0.999) thr = 0.88;
            string msg;
            using (new Speed())
            {
                msg = FuzzyCore(src, thr);
            }
            Ui.Info(msg);
        }

        static string FuzzyCore(dynamic src, double thr)
        {
            object[,] v = Xl.ReadTyped(src);
            Dictionary<string, int> counts = new Dictionary<string, int>(StringComparer.Ordinal);
            List<string> names = new List<string>();
            foreach (object x in v)
            {
                string s = x as string;
                if (s == null) continue;
                s = s.Trim();
                if (s.Length < 2) continue;
                int k;
                if (counts.TryGetValue(s, out k)) counts[s] = k + 1;
                else
                {
                    counts[s] = 1;
                    names.Add(s);
                }
            }
            int n = names.Count;
            if (n < 2) return "There are not enough text values in the selection.";
            if (n > 8000) return "The selection has " + n.ToString("#,##0", Inv) + " different names. Check up to 8,000 different names at a time.";
            string[] norm = new string[n];
            for (int i = 0; i < n; i++) norm[i] = Logic.NormName(names[i]);
            List<object[]> hits = new List<object[]>();
            for (int i = 0; i < n - 1; i++)
            {
                int li = norm[i].Length;
                if (li == 0) continue;
                for (int j = i + 1; j < n; j++)
                {
                    int lj = norm[j].Length;
                    if (lj == 0) continue;
                    int mx = Math.Max(li, lj);
                    if (Math.Abs(li - lj) > 0.35 * mx + 1) continue;
                    double sim = norm[i] == norm[j] ? 1.0 : Logic.JaroW(norm[i], norm[j]);
                    if (sim >= thr && hits.Count < 200000)
                        hits.Add(new object[] { names[i], names[j], Math.Round(sim * 100), (double)counts[names[i]], (double)counts[names[j]] });
                }
                if (i % 200 == 0) Xl.Status("comparing names " + (100 * i / n) + "%");
            }
            if (hits.Count == 0) return "No look-alike names found at " + Math.Round(thr * 100) + "% similarity.";
            hits.Sort(delegate (object[] p, object[] q) { return ((double)q[2]).CompareTo((double)p[2]); });
            dynamic ws = Xl.AddSheet(src.Worksheet.Parent, "Fuzzy Duplicates");
            object[,] t = Table(new[] { "Name A", "Name B", "Similarity %", "Rows with A", "Rows with B" }, hits);
            Xl.WriteTable(ws.Range["A1"], t, 1);
            Xl.StyleHeader(ws.Range["A1:E1"]);
            ws.Columns["A:E"].AutoFit();
            ws.Activate();
            return hits.Count + " look-alike pairs found among " + n + " different names.\n100% means the names differ only by case, punctuation or Pvt/Ltd.";
        }

        // ---------------------------------------------------------------
        public static void FindOutliers()
        {
            dynamic src = Xl.PickRange("Find Outliers", "Select the number column to check:");
            if (src == null) return;
            object[,] v = Xl.ReadTyped(src);
            int nr = v.GetLength(0), nc = v.GetLength(1);
            List<double> vals = new List<double>();
            foreach (object x in v) if (Xl.IsNum(x)) vals.Add(Xl.Num(x));
            if (vals.Count < 8)
            {
                Ui.Info("At least 8 numbers are needed to judge what is unusual.");
                return;
            }
            double[] sorted = vals.ToArray();
            Array.Sort(sorted);
            double q1 = Logic.Percentile(sorted, 0.25), q3 = Logic.Percentile(sorted, 0.75);
            double lo = q1 - 1.5 * (q3 - q1), hi = q3 + 1.5 * (q3 - q1);
            List<long> hits = new List<long>();
            for (int c = 1; c <= nc; c++)
                for (int r = 1; r <= nr; r++)
                {
                    if (!Xl.IsNum(v[r, c])) continue;
                    double d = Xl.Num(v[r, c]);
                    if (d < lo || d > hi) hits.Add(Xl.Cell(r, c));
                }
            using (new Speed())
            {
                Xl.Paint(src, hits, Design.Unusual);
            }
            Ui.Info(hits.Count + " unusual values highlighted in orange.\n\nNormal range for this data: " + Logic.ShortINR(lo) + " to " + Logic.ShortINR(hi) +
                    "\n(values far outside the middle 50% of the numbers)");
        }

        public static void HighlightDuplicates()
        {
            dynamic src = Xl.PickRange("Highlight Duplicates", "Select the column(s) to check for repeated values:");
            if (src == null) return;
            object[,] v = Xl.ReadTyped(src);
            int nr = v.GetLength(0), nc = v.GetLength(1);
            Dictionary<string, int> counts = new Dictionary<string, int>();
            foreach (object x in v)
            {
                if (Xl.IsBlank(x)) continue;
                string k = Xl.Key(x);
                int n;
                counts[k] = counts.TryGetValue(k, out n) ? n + 1 : 1;
            }
            List<long> hits = new List<long>();
            for (int c = 1; c <= nc; c++)
                for (int r = 1; r <= nr; r++)
                {
                    if (Xl.IsBlank(v[r, c])) continue;
                    if (counts[Xl.Key(v[r, c])] > 1) hits.Add(Xl.Cell(r, c));
                }
            int repeated = 0;
            foreach (KeyValuePair<string, int> kv in counts) if (kv.Value > 1) repeated++;
            using (new Speed())
            {
                Xl.Paint(src, hits, Design.Bad);
            }
            Ui.Info(hits.Count.ToString("#,##0", Inv) + " cells highlighted (" + repeated.ToString("#,##0", Inv) + " values appear more than once).");
        }

        public static void Unpivot()
        {
            dynamic src = Xl.PickRange("Unpivot Table", "Select the cross-tab table with its header row:");
            if (src == null) return;
            string ans = Ui.AskText("How many columns on the left are labels (kept as they are)?\nAll other columns are stacked into rows.", Design.AppName, "1");
            if (ans == null) return;
            int k;
            if (!int.TryParse(ans.Trim(), NumberStyles.Integer, Inv, out k)) k = 0;
            object[,] v = Xl.ReadTyped(src);
            int nr = v.GetLength(0), nc = v.GetLength(1);
            if (k < 1 || k >= nc || nr < 2)
            {
                Ui.Info("The number of label columns must be between 1 and " + (nc - 1) + ", and the table needs data rows.");
                return;
            }
            long n = 0;
            for (int r = 2; r <= nr; r++)
                for (int c = k + 1; c <= nc; c++)
                    if (!Xl.IsBlank(v[r, c])) n++;
            if (n == 0 || n > 1048000)
            {
                Ui.Info("The result would have " + n.ToString("#,##0", Inv) + " rows, which cannot be written to one sheet.");
                return;
            }
            using (new Speed())
            {
                object[,] o = Xl.New2D((int)n + 1, k + 2);
                for (int j = 1; j <= k; j++) o[1, j] = Xl.HeaderText(v[1, j], j);
                o[1, k + 1] = "Attribute";
                o[1, k + 2] = "Value";
                int i = 1;
                for (int r = 2; r <= nr; r++)
                    for (int c = k + 1; c <= nc; c++)
                    {
                        if (Xl.IsBlank(v[r, c])) continue;
                        i++;
                        for (int j = 1; j <= k; j++) o[i, j] = v[r, j];
                        o[i, k + 1] = Disp(v[1, c]);
                        o[i, k + 2] = v[r, c];
                    }
                dynamic ws = Xl.AddSheet(src.Worksheet.Parent, "Unpivoted");
                Xl.WriteTable(ws.Range["A1"], o, 1);
                Xl.StyleHeader(ws.Range["A1"].Resize[1, k + 2]);
                ws.Range["A1"].Resize[1, k + 2].EntireColumn.AutoFit();
                ws.Activate();
            }
            Ui.Info(n.ToString("#,##0", Inv) + " rows written to the new 'Unpivoted' sheet.");
        }
    }
}
