using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Threading.Tasks;

namespace SmartPro
{
    // FUZZY MERGE: joins two tables on names that are nearly the same and shows a similarity percentage.
    // It does not use Power Query, so it works the same in every Excel version.
    //
    // Speed: names are cleaned once, names that are equal after cleaning are matched by a dictionary lookup,
    // and the rest are found through an index of three-letter pieces (trigrams), so a name is only compared
    // with the few master names that share pieces with it. The matching runs on all CPU cores.
    internal static class FuzzyMerge
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        const int K = 24;                 // candidates that get the full comparison, per name
        const double MinDice = 0.2;       // candidates must share at least this much of their trigrams
        const double NumberCap = 0.79;    // names that differ in a number (Unit 1 / Unit 2) never score above this

        struct Hit
        {
            public int Id;
            public double Score;
        }

        // the distinct cleaned names of the master table
        sealed class Index
        {
            public string[] Norm, Sorted, Digits;
            public int[] GramCount, FirstRow, Dup;
            public Dictionary<string, int> Exact;
            public Dictionary<int, int[]> Post;      // trigram -> ids of the names that contain it
        }

        // per-thread scratch space
        sealed class Work
        {
            public int[] Counts, Touched;
            public int[] CandId = new int[K];
            public double[] CandDice = new double[K];
            public Hit[] Buf = new Hit[8];
            public Work(int n)
            {
                Counts = new int[n];
                Touched = new int[n];
            }
        }

        static string Clean(string raw)
        {
            string n = Logic.NormName(raw);
            return n.Length > 0 ? n : raw.Trim().ToLowerInvariant();
        }

        // distinct trigrams of a cleaned name, sorted
        static int[] Grams(string s)
        {
            string p = " " + s + " ";
            int n = p.Length - 2;
            if (n <= 0) return new int[0];
            int[] g = new int[n];
            for (int i = 0; i < n; i++) g[i] = (p[i] << 16) ^ (p[i + 1] << 8) ^ p[i + 2];
            Array.Sort(g);
            int k = 0;
            for (int i = 0; i < n; i++)
            {
                if (i == 0 || g[i] != g[i - 1]) g[k++] = g[i];
            }
            if (k < n) Array.Resize(ref g, k);
            return g;
        }

        // words in alphabetical order, so "Steel Tata" and "Tata Steel" compare as equal
        static string SortTokens(string s)
        {
            if (s.IndexOf(' ') < 0) return s;
            string[] w = s.Split(' ');
            Array.Sort(w, StringComparer.Ordinal);
            return string.Join(" ", w);
        }

        // the numbers inside a name, sorted: "unit 2 line 10" -> "10,2"
        static string DigitsOf(string s)
        {
            List<string> runs = null;
            int i = 0;
            while (i < s.Length)
            {
                if (s[i] >= '0' && s[i] <= '9')
                {
                    int j = i;
                    while (j < s.Length && s[j] >= '0' && s[j] <= '9') j++;
                    if (runs == null) runs = new List<string>();
                    runs.Add(s.Substring(i, j - i));
                    i = j;
                }
                else i++;
            }
            if (runs == null) return "";
            runs.Sort(StringComparer.Ordinal);
            return string.Join(",", runs.ToArray());
        }

        // rowNorm[r] = cleaned name of master row r ("" when blank); rows start at 'from'
        static Index Build(string[] rowNorm, int from)
        {
            Index ix = new Index();
            ix.Exact = new Dictionary<string, int>(StringComparer.Ordinal);
            List<string> names = new List<string>();
            List<int> firstRow = new List<int>(), dup = new List<int>();
            for (int r = from; r < rowNorm.Length; r++)
            {
                string s = rowNorm[r];
                if (s.Length == 0) continue;
                int id;
                if (ix.Exact.TryGetValue(s, out id)) dup[id] = dup[id] + 1;
                else
                {
                    ix.Exact[s] = names.Count;
                    names.Add(s);
                    firstRow.Add(r);
                    dup.Add(1);
                }
            }
            int n = names.Count;
            ix.Norm = names.ToArray();
            ix.FirstRow = firstRow.ToArray();
            ix.Dup = dup.ToArray();
            ix.Sorted = new string[n];
            ix.Digits = new string[n];
            ix.GramCount = new int[n];
            Dictionary<int, List<int>> post = new Dictionary<int, List<int>>();
            for (int i = 0; i < n; i++)
            {
                ix.Sorted[i] = SortTokens(ix.Norm[i]);
                ix.Digits[i] = DigitsOf(ix.Norm[i]);
                int[] g = Grams(ix.Norm[i]);
                ix.GramCount[i] = g.Length;
                for (int a = 0; a < g.Length; a++)
                {
                    List<int> list;
                    if (!post.TryGetValue(g[a], out list))
                    {
                        list = new List<int>();
                        post[g[a]] = list;
                    }
                    list.Add(i);
                }
            }
            ix.Post = new Dictionary<int, int[]>(post.Count);
            foreach (KeyValuePair<int, List<int>> kv in post) ix.Post[kv.Key] = kv.Value.ToArray();
            return ix;
        }

        // 0..1: 60% Jaro-Winkler (also tried with the words sorted) + 40% shared trigrams
        static double Score(Index ix, string q, string qs, string qd, int id, double dice)
        {
            string b = ix.Norm[id], bs = ix.Sorted[id];
            double s = Logic.JaroW(q, b);
            if (!string.Equals(qs, q, StringComparison.Ordinal) || !string.Equals(bs, b, StringComparison.Ordinal))
            {
                double s2 = Logic.JaroW(qs, bs);
                if (s2 > s) s = s2;
            }
            double f = 0.6 * s + 0.4 * dice;
            if (f > NumberCap && qd.Length > 0 && ix.Digits[id].Length > 0 && !string.Equals(qd, ix.Digits[id], StringComparison.Ordinal)) f = NumberCap;
            return f;
        }

        // best matches for one cleaned name, best first; returns how many were written to hits
        static int Query(Index ix, string q, Work w, int want, double threshold, Hit[] hits)
        {
            int[] g = Grams(q);
            if (g.Length == 0) return 0;
            string qs = SortTokens(q), qd = DigitsOf(q);
            int nt = 0;
            for (int a = 0; a < g.Length; a++)
            {
                int[] list;
                if (!ix.Post.TryGetValue(g[a], out list)) continue;
                for (int b = 0; b < list.Length; b++)
                {
                    int id = list[b];
                    if (w.Counts[id]++ == 0) w.Touched[nt++] = id;
                }
            }
            // keep the K names that share the most trigrams
            int nc = 0, minPos = -1;
            double minVal = 0;
            for (int t = 0; t < nt; t++)
            {
                int id = w.Touched[t];
                double dice = 2.0 * w.Counts[id] / (g.Length + ix.GramCount[id]);
                if (dice < MinDice) continue;
                if (nc < K)
                {
                    w.CandId[nc] = id;
                    w.CandDice[nc] = dice;
                    nc++;
                    continue;
                }
                if (minPos < 0)
                {
                    minPos = 0;
                    minVal = w.CandDice[0];
                    for (int z = 1; z < K; z++)
                    {
                        if (w.CandDice[z] < minVal)
                        {
                            minVal = w.CandDice[z];
                            minPos = z;
                        }
                    }
                }
                if (dice > minVal)
                {
                    w.CandId[minPos] = id;
                    w.CandDice[minPos] = dice;
                    minPos = -1;
                }
            }
            int found = 0;
            for (int z = 0; z < nc; z++)
            {
                double f = Score(ix, q, qs, qd, w.CandId[z], w.CandDice[z]);
                if (f < threshold) continue;
                if (found >= want && f <= hits[want - 1].Score) continue;
                int pos = found < want ? found : want - 1;
                while (pos > 0 && hits[pos - 1].Score < f)
                {
                    hits[pos] = hits[pos - 1];
                    pos--;
                }
                hits[pos].Id = w.CandId[z];
                hits[pos].Score = f;
                if (found < want) found++;
            }
            for (int t = 0; t < nt; t++) w.Counts[w.Touched[t]] = 0;
            return found;
        }

        static string[] Heads(object[,] v)
        {
            int nc = v.GetLength(1);
            string[] h = new string[nc + 1];
            for (int c = 1; c <= nc; c++) h[c] = Xl.HeaderText(v[1, c], c);
            return h;
        }

        static string RawName(object x)
        {
            return Xl.IsBlank(x) || Xl.IsErr(x) ? "" : Xl.Str(x).Trim();
        }

        // =================================================================
        public static void Run()
        {
            dynamic left = Xl.PickRange("Fuzzy Merge (1 of 2)", "Select the table you want to fill in, with its header row:");
            if (left == null) return;
            object[,] lv = Xl.ReadTyped(left);
            if (lv.GetLength(0) < 2)
            {
                Ui.Info("Select a table with a header row and data rows.");
                return;
            }
            string[] lh = Heads(lv);
            int lk = Ui.PickColumn("Fuzzy Merge: your table", "Which column holds the names to match?", lh, null);
            if (lk == 0) return;
            dynamic right = Xl.PickRange("Fuzzy Merge (2 of 2)", "Select the master table to look in, with its header row. You can click another sheet or workbook:");
            if (right == null) return;
            object[,] rv = Xl.ReadTyped(right);
            if (rv.GetLength(0) < 2)
            {
                Ui.Info("The master table needs a header row and data rows.");
                return;
            }
            string[] rh = Heads(rv);
            int rk = Ui.PickColumn("Fuzzy Merge: master table", "Which column holds the names to match against?", rh, null);
            if (rk == 0) return;
            string ans = Ui.AskText("How similar must two names be to count as a match?\nType 50 to 100.  80 is a good start; a higher number is stricter.", Design.AppName + " - Fuzzy Merge", "80");
            if (ans == null) return;
            double thr;
            if (!double.TryParse(ans.Trim().TrimEnd('%'), NumberStyles.Any, Inv, out thr)) thr = 80;
            if (thr > 1) thr /= 100;
            if (thr < 0.5) thr = 0.5;
            if (thr > 1) thr = 1;
            int mode = Ui.Pick("Fuzzy Merge: matches per row", "For each row of your table, bring:", new[] { "The best match only", "Up to 3 matches (one output row each)", "Up to 5 matches (one output row each)" }, null);
            if (mode < 0) return;
            int want = mode == 0 ? 1 : (mode == 1 ? 3 : 5);
            string msg;
            using (new Speed())
            {
                msg = Core(left, lv, lk, lh, rv, rk, rh, thr, want);
            }
            Ui.Info(msg);
        }

        static string Core(dynamic left, object[,] lv, int lk, string[] lh, object[,] rv, int rk, string[] rh, double thr, int want)
        {
            int nl = lv.GetLength(0), lc = lv.GetLength(1), nrr = rv.GetLength(0), rc = rv.GetLength(1);
            Xl.Status("preparing the names...");
            Stopwatch sw = Stopwatch.StartNew();

            // master side
            string[] rRaw = new string[nrr + 1], rNorm = new string[nrr + 1];
            rRaw[0] = rRaw[1] = rNorm[0] = rNorm[1] = "";
            for (int r = 2; r <= nrr; r++)
            {
                rRaw[r] = RawName(rv[r, rk]);
                rNorm[r] = rRaw[r].Length == 0 ? "" : Clean(rRaw[r]);
            }
            Index ix = Build(rNorm, 2);
            int nR = ix.Norm.Length;
            if (nR == 0) return "The master table has no names in that column.";

            // your side: each different name is matched once
            Dictionary<string, int> seen = new Dictionary<string, int>(StringComparer.Ordinal);
            List<string> names = new List<string>();
            string[] lRaw = new string[nl + 1];
            int[] lId = new int[nl + 1];
            for (int r = 2; r <= nl; r++)
            {
                lRaw[r] = RawName(lv[r, lk]);
                if (lRaw[r].Length == 0)
                {
                    lId[r] = -1;
                    continue;
                }
                string n = Clean(lRaw[r]);
                int id;
                if (!seen.TryGetValue(n, out id))
                {
                    id = names.Count;
                    seen[n] = id;
                    names.Add(n);
                }
                lId[r] = id;
            }
            string[] q = names.ToArray();
            int nq = q.Length;
            Hit[][] res = new Hit[nq][];
            Xl.Status("matching " + nq.ToString("#,##0", Inv) + " names against " + nR.ToString("#,##0", Inv) + "...");
            Parallel.For(0, nq,
                () => new Work(nR),
                (i, state, w) =>
                {
                    int e;
                    if (ix.Exact.TryGetValue(q[i], out e))
                    {
                        res[i] = new[] { new Hit { Id = e, Score = 1.0 } };
                        return w;
                    }
                    int n = Query(ix, q[i], w, want, thr, w.Buf);
                    Hit[] got = new Hit[n];
                    Array.Copy(w.Buf, got, n);
                    res[i] = got;
                    return w;
                },
                w => { });
            double seconds = sw.Elapsed.TotalSeconds;

            // ---- output ----
            Xl.Status("writing the result...");
            long total = 0;
            for (int r = 2; r <= nl; r++)
            {
                int id = lId[r];
                int k = id < 0 ? 0 : res[id].Length;
                total += k == 0 ? 1 : k;
            }
            if (total > 1048000) return "The result would have " + total.ToString("#,##0", Inv) + " rows, more than one sheet can hold. Choose 'The best match only'.";
            int extra = want > 1 ? 3 : 2;                       // Match %, Match type, (Rank)
            int cols = lc + extra + rc;
            object[,] o = Xl.New2D((int)total + 1, cols);
            HashSet<string> used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int c = 1; c <= lc; c++)
            {
                o[1, c] = lh[c];
                used.Add(lh[c]);
            }
            o[1, lc + 1] = "Match %";
            o[1, lc + 2] = "Match type";
            if (want > 1) o[1, lc + 3] = "Rank";
            for (int c = 1; c <= rc; c++) o[1, lc + extra + c] = used.Contains(rh[c]) ? rh[c] + " (master)" : rh[c];

            int nExact = 0, nCleaned = 0, nFuzzy = 0, nNone = 0, nBlank = 0, row = 1, dupUsed = 0;
            List<long> review = new List<long>(), none = new List<long>();
            double reviewBelow = Math.Min(1.0, thr + 0.08);
            for (int r = 2; r <= nl; r++)
            {
                int id = lId[r];
                Hit[] hits = id < 0 ? null : res[id];
                int k = hits == null ? 0 : hits.Length;
                if (k == 0)
                {
                    row++;
                    for (int c = 1; c <= lc; c++) o[row, c] = lv[r, c];
                    if (id < 0)
                    {
                        o[row, lc + 2] = "Blank name";
                        nBlank++;
                    }
                    else
                    {
                        o[row, lc + 2] = "No match";
                        none.Add(Xl.Cell(row, lc + 2));
                        nNone++;
                    }
                    continue;
                }
                for (int j = 0; j < k; j++)
                {
                    row++;
                    int rid = hits[j].Id, mr = ix.FirstRow[rid];
                    for (int c = 1; c <= lc; c++) o[row, c] = lv[r, c];
                    o[row, lc + 1] = hits[j].Score;
                    string type;
                    if (hits[j].Score >= 1.0)
                    {
                        if (string.Equals(lRaw[r], rRaw[mr], StringComparison.OrdinalIgnoreCase)) type = "Exact";
                        else type = "Same after cleaning";
                    }
                    else type = "Fuzzy";
                    o[row, lc + 2] = type;
                    if (want > 1) o[row, lc + 3] = (double)(j + 1);
                    for (int c = 1; c <= rc; c++) o[row, lc + extra + c] = rv[mr, c];
                    if (j == 0)
                    {
                        if (type == "Exact") nExact++;
                        else if (type == "Fuzzy") nFuzzy++;
                        else nCleaned++;
                        if (ix.Dup[rid] > 1) dupUsed++;
                    }
                    if (type == "Fuzzy" && hits[j].Score < reviewBelow) review.Add(Xl.Cell(row, lc + 1));
                }
            }

            dynamic ws = Xl.AddSheet(left.Worksheet.Parent, "Fuzzy Merge");
            ws.Range["A1"].Value2 = "Fuzzy Merge: " + lh[lk] + " matched to " + rh[rk];
            ws.Range["A1"].Font.Size = 14;
            ws.Range["A1"].Font.Bold = true;
            ws.Range["A2"].Value2 = "Exact " + nExact.ToString("#,##0", Inv) + "   |   Same after cleaning " + nCleaned.ToString("#,##0", Inv) + "   |   Fuzzy " + nFuzzy.ToString("#,##0", Inv) +
                                    "   |   No match " + nNone.ToString("#,##0", Inv) + (nBlank > 0 ? "   |   Blank " + nBlank.ToString("#,##0", Inv) : "") +
                                    "   |   Similarity from " + thr.ToString("0%", Inv) + "   |   Yellow = fuzzy match close to the limit, please check";
            dynamic top = ws.Range["A4"];
            Xl.WriteTable(top, o, 1);
            dynamic tbl = top.Resize[(int)total + 1, cols];
            Xl.StyleHeader(top.Resize[1, cols]);
            tbl.Columns[lc + 1].NumberFormat = "0%";
            Xl.Paint(tbl, review, Design.Changed);
            Xl.Paint(tbl, none, Design.Bad);
            try { top.Resize[1, cols].EntireColumn.AutoFit(); } catch { }
            ws.Activate();
            return "Fuzzy Merge finished in " + seconds.ToString("0.0", Inv) + " seconds for the matching.\n\n" +
                   "Exact: " + nExact.ToString("#,##0", Inv) + "\nSame after cleaning: " + nCleaned.ToString("#,##0", Inv) + "\nFuzzy: " + nFuzzy.ToString("#,##0", Inv) +
                   "\nNo match: " + nNone.ToString("#,##0", Inv) +
                   (dupUsed > 0 ? "\n\n" + dupUsed.ToString("#,##0", Inv) + " rows matched a name that appears more than once in the master table; the first of those master rows was used." : "") +
                   "\n\nYellow cells are fuzzy matches close to your limit: check those first.";
        }
    }
}
