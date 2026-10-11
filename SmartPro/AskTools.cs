using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace SmartPro
{
    // "Ask Your Data": type a short request in plain words and get a table and a chart.
    // It works by matching your words to the column names. It runs on this PC and sends nothing anywhere.
    internal static class AskTools
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        const int Sum = 0, Avg = 1, Count = 2, Max = 3, Min = 4;

        public static void Ask()
        {
            dynamic src = Xl.PickRange("Ask Your Data");
            if (src == null) return;
            object[,] v = Xl.ReadTyped(src);
            if (v.GetLength(0) < 3 || v.GetLength(1) < 2)
            {
                Ui.Info("Select a table with a header row, at least 2 columns and 2 data rows.");
                return;
            }
            string q = Ui.AskText("What do you want to see? Examples:\n  amount by customer top 10\n  average rate by plant\n  qty by month\n  count by grade", Design.AppName + " - Ask Your Data", "");
            if (q == null || q.Trim().Length == 0) return;
            string msg;
            using (new Speed())
            {
                msg = AskCore(src, v, q.Trim());
            }
            Ui.Info(msg);
        }

        static string Norm(string s)
        {
            return " " + Regex.Replace(Logic.PadName(s ?? ""), @"\s+", " ").Trim() + " ";
        }

        static bool Has(string text, params string[] words)
        {
            foreach (string w in words) if (text.IndexOf(" " + w + " ", StringComparison.Ordinal) >= 0) return true;
            return false;
        }

        static bool SameWord(string a, string b)
        {
            if (a == b) return true;
            string a1 = a.TrimEnd('s'), b1 = b.TrimEnd('s');
            if (a1.Length >= 3 && a1 == b1) return true;
            if (a.Length >= 4 && b.StartsWith(a, StringComparison.Ordinal)) return true;
            if (b.Length >= 4 && a.StartsWith(b, StringComparison.Ordinal)) return true;
            return false;
        }

        // the column whose header matches the words best; 0 when nothing matches
        static int BestColumn(string text, string[] h, ColInfo[] ci, bool wantNumber, int exclude)
        {
            string[] words = text.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            int best = 0, bestScore = 0;
            for (int c = 1; c < h.Length; c++)
            {
                if (c == exclude || ci[c].NonBlank == 0) continue;
                bool numeric = ci[c].NumCount >= 0.8 * ci[c].NonBlank;
                if (wantNumber && !numeric) continue;
                if (!wantNumber && (ci[c].Kind == "date" || ci[c].Distinct < 2)) continue;
                string hn = Norm(h[c]);
                if (hn.Trim().Length == 0) continue;
                int score = 0;
                if (text.IndexOf(hn, StringComparison.Ordinal) >= 0) score = 100 + hn.Length;
                else
                {
                    foreach (string hw in hn.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        if (hw.Length < 2) continue;
                        foreach (string w in words)
                        {
                            if (w.Length < 2 || !SameWord(w, hw)) continue;
                            score += 10 + hw.Length;
                            break;
                        }
                    }
                }
                if (score > bestScore)
                {
                    bestScore = score;
                    best = c;
                }
            }
            return best;
        }

        static double ValueOf(double[] a, int agg)
        {
            switch (agg)
            {
                case Avg: return a[1] > 0 ? a[0] / a[1] : 0;
                case Count: return a[1];
                case Max: return a[2];
                case Min: return a[3];
                default: return a[0];
            }
        }

        static string AskCore(dynamic src, object[,] v, string question)
        {
            int hr = Profiler.FindHeaderRow(v);
            string[] h = Profiler.MakeHeaders(v, hr);
            ColInfo[] ci = Profiler.Profile(v, hr, h, 20000);
            int nr = v.GetLength(0), nc = v.GetLength(1);
            string q = Norm(question);

            // ---- what to calculate ----
            int agg = Sum;
            if (Has(q, "average", "avg", "mean")) agg = Avg;
            else if (Has(q, "count", "how many", "number of")) agg = Count;
            else if (Has(q, "max", "maximum", "highest", "largest")) agg = Max;
            else if (Has(q, "min", "minimum", "lowest", "smallest")) agg = Min;
            int topN = 0;
            Match tm = Regex.Match(q, @" top (\d{1,3}) ");
            if (tm.Success) topN = int.Parse(tm.Groups[1].Value, Inv);
            bool byMonth = Has(q, "month", "monthly", "monthwise", "trend");
            bool byYear = !byMonth && Has(q, "year", "yearly", "yearwise", "fy", "fiscal");
            bool wantPie = Has(q, "pie", "share", "donut", "doughnut");

            int dateCol = 0;
            for (int c = 1; c <= nc; c++) if (ci[c].Kind == "date" && dateCol == 0) dateCol = c;

            // ---- "<number> by <group>": the left part names the number, the right part names the group ----
            string left = q, right = "";
            int byPos = q.IndexOf(" by ", StringComparison.Ordinal);
            if (byPos >= 0)
            {
                left = q.Substring(0, byPos + 1);
                right = q.Substring(byPos + 3);
            }
            int mCol = 0;
            if (agg != Count)
            {
                mCol = BestColumn(left, h, ci, true, 0);
                if (mCol == 0)
                {
                    // nothing named: take the most "amount-like" number column
                    for (int rank = 1; rank <= 3 && mCol == 0; rank++)
                        for (int c = 1; c <= nc && mCol == 0; c++)
                            if (ci[c].Kind == "num" && Logic.MeasureRank(ci[c].Hdr) == rank) mCol = c;
                }
                if (mCol == 0) agg = Count;
            }
            int dCol = 0;
            if (!byMonth && !byYear)
            {
                dCol = BestColumn(right.Trim().Length > 0 ? right : q, h, ci, false, mCol);
                if (dCol == 0 && right.Trim().Length == 0)
                {
                    if (dateCol > 0) byMonth = true;
                    else
                        for (int c = 1; c <= nc && dCol == 0; c++)
                            if (c != mCol && ci[c].Kind == "cat") dCol = c;
                }
                if (dCol == 0 && !byMonth)
                {
                    List<string> names = new List<string>();
                    for (int c = 1; c <= nc; c++) if (ci[c].NonBlank > 0) names.Add(h[c]);
                    return "I could not match \"" + right.Trim() + "\" to a column.\n\nColumns in this table:\n" + string.Join(", ", names.ToArray());
                }
            }
            if ((byMonth || byYear) && dateCol == 0) return "This table has no date column, so it cannot be grouped by month or year.\nIf the dates are stored as text, run Text to Dates first.";

            // ---- add it up ----
            Dictionary<string, double[]> acc = new Dictionary<string, double[]>(StringComparer.OrdinalIgnoreCase);   // sum, rows, max, min
            List<string> keys = new List<string>();
            for (int r = hr + 1; r <= nr; r++)
            {
                if (!Profiler.RowIsData(v, r, nc)) continue;
                string k;
                if (byMonth || byYear)
                {
                    DateTime dt;
                    if (!Logic.TryDate(v[r, dateCol], out dt)) continue;
                    k = byMonth ? dt.ToString("yyyy-MM", Inv) + " (" + dt.ToString("MMM", Inv) + ")" : Logic.FYLabel(dt, 4);
                }
                else k = Xl.GroupKey(v[r, dCol]);
                double x = 1;
                if (agg != Count && !Logic.TryNum(v[r, mCol], out x)) continue;
                double[] a;
                if (!acc.TryGetValue(k, out a))
                {
                    a = new double[] { 0, 0, double.MinValue, double.MaxValue };
                    acc[k] = a;
                    keys.Add(k);
                }
                a[0] += x;
                a[1] += 1;
                if (x > a[2]) a[2] = x;
                if (x < a[3]) a[3] = x;
            }
            if (keys.Count == 0) return "No rows could be used for this request.";
            bool timeline = byMonth || byYear;
            int aggF = agg;
            if (timeline) keys.Sort(StringComparer.Ordinal);
            else keys.Sort(delegate (string p, string s) { return ValueOf(acc[s], aggF).CompareTo(ValueOf(acc[p], aggF)); });
            int total = keys.Count;
            if (topN > 0 && !timeline && keys.Count > topN) keys.RemoveRange(topN, keys.Count - topN);
            if (keys.Count > 1000) keys.RemoveRange(1000, keys.Count - 1000);

            string groupName = byMonth ? "Month" : byYear ? "Fiscal Year" : h[dCol];
            string mName = mCol > 0 ? h[mCol] : "";
            string valueName = agg == Avg ? "Average " + mName : agg == Count ? "Rows" : agg == Max ? "Highest " + mName : agg == Min ? "Lowest " + mName : "Total " + mName;

            // ---- write the answer ----
            int n = keys.Count;
            object[,] t = Xl.New2D(n + 1, 2);
            t[1, 1] = groupName;
            t[1, 2] = valueName;
            double all = 0;
            bool whole = true;
            for (int i = 0; i < n; i++)
            {
                double val = ValueOf(acc[keys[i]], agg);
                t[i + 2, 1] = keys[i];
                t[i + 2, 2] = val;
                all += val;
                if (val != Math.Floor(val)) whole = false;
            }
            string line;
            if (timeline && n >= 2)
            {
                double last = ValueOf(acc[keys[n - 1]], agg), prev = ValueOf(acc[keys[n - 2]], agg);
                line = "Latest " + keys[n - 1] + ": " + Logic.ShortINR(last) + (prev != 0 ? " (" + ((last - prev) / Math.Abs(prev)).ToString("+0%;-0%", Inv) + " vs " + keys[n - 2] + ")" : "");
            }
            else
            {
                double first = ValueOf(acc[keys[0]], agg);
                line = "Top: " + keys[0] + " with " + Logic.ShortINR(first) + ((agg == Sum || agg == Count) && all > 0 ? " (" + (first / all).ToString("0%", Inv) + " of the " + n + " shown)" : "");
            }

            dynamic ws = Xl.AddSheet(src.Worksheet.Parent, "Answer");
            ws.Range["A1"].Value2 = valueName + " by " + groupName + (topN > 0 && !timeline ? " (top " + n + " of " + total + ")" : "");
            ws.Range["A1"].Font.Size = 14;
            ws.Range["A1"].Font.Bold = true;
            ws.Range["A2"].Value2 = line;
            Xl.WriteTable(ws.Range["A4"], t, 1);
            Xl.StyleHeader(ws.Range["A4:B4"]);
            ws.Range["B5"].Resize[n, 1].NumberFormat = whole ? "#,##0" : "#,##0.00";
            ws.Columns["A:B"].AutoFit();
            try
            {
                int type = timeline ? C.xlLineMarkers : (wantPie && n <= 8) ? C.xlDoughnut : n <= 8 ? C.xlColumnClustered : C.xlBarClustered;
                dynamic anchor = ws.Range["D4:M24"];
                dynamic co = ws.ChartObjects().Add((double)anchor.Left, (double)anchor.Top, (double)anchor.Width, (double)anchor.Height);
                dynamic ch = co.Chart;
                ch.SetSourceData(ws.Range["A4"].Resize[Math.Min(n, 40) + 1, 2]);
                ch.ChartType = type;
                ch.HasTitle = true;
                ch.ChartTitle.Text = valueName + " by " + groupName;
                ch.HasLegend = type == C.xlDoughnut;
            }
            catch { }
            ws.Activate();
            return "Answer ready: " + valueName + " by " + groupName + ", " + n + " rows.\n\n" + line;
        }
    }
}
