using System;
using System.Collections.Generic;
using System.Globalization;

namespace SmartPro
{
    // Pro insight tools that go past a dashboard:
    //   Why It Changed .. compares two periods and shows what caused the rise or fall
    //   Forecast ........ the next months from the monthly history, with a likely range
    internal static class InsightTools
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        sealed class Cols
        {
            public string[] H;
            public bool[] Num, Date, Group;
            public int Hr, NDate, FirstDate;
        }

        static Cols Describe(object[,] v)
        {
            Cols k = new Cols();
            k.Hr = Profiler.FindHeaderRow(v);
            k.H = Profiler.MakeHeaders(v, k.Hr);
            ColInfo[] ci = Profiler.Profile(v, k.Hr, k.H, 20000);
            int nc = v.GetLength(1);
            k.Num = new bool[nc + 1];
            k.Date = new bool[nc + 1];
            k.Group = new bool[nc + 1];
            for (int c = 1; c <= nc; c++)
            {
                string kind = ci[c].Kind;
                k.Date[c] = kind == "date";
                k.Num[c] = kind != "date" && ci[c].NonBlank > 0 && ci[c].NumCount >= 0.8 * ci[c].NonBlank;
                k.Group[c] = kind != "date" && kind != "num" && ci[c].NonBlank > 0 && ci[c].Distinct >= 2;
                if (k.Date[c])
                {
                    k.NDate++;
                    if (k.FirstDate == 0) k.FirstDate = c;
                }
            }
            return k;
        }

        static bool Any(bool[] a)
        {
            foreach (bool b in a) if (b) return true;
            return false;
        }

        static string MonthName(int ym)
        {
            return new DateTime(ym / 12, ym % 12 + 1, 1).ToString("MMM-yyyy", Inv);
        }

        static string Signed(double d)
        {
            return (d >= 0 ? "+" : "-") + Logic.ShortINR(Math.Abs(d));
        }

        // =================================================================
        //  WHY IT CHANGED
        // =================================================================
        public static void Variance()
        {
            dynamic src = Xl.PickRange("Why It Changed");
            if (src == null) return;
            object[,] v = Xl.ReadTyped(src);
            if (v.GetLength(0) < 3 || v.GetLength(1) < 3)
            {
                Ui.Info("Select a table with a date column, a number column and a column to explain it by (customer, product, plant).");
                return;
            }
            Cols k = Describe(v);
            if (k.NDate == 0)
            {
                Ui.Info("No date column was found. If the dates are stored as text, run Smart Clean first.");
                return;
            }
            if (!Any(k.Num) || !Any(k.Group))
            {
                Ui.Info("This tool needs a number column and a column to explain it by (customer, product, plant).");
                return;
            }
            int dc = k.NDate == 1 ? k.FirstDate : Ui.PickColumn("Why It Changed", "Which column holds the date?", k.H, k.Date);
            if (dc == 0) return;
            int mc = Ui.PickColumn("Why It Changed (1 of 2)", "Which number do you want explained?", k.H, k.Num);
            if (mc == 0) return;
            int gc = Ui.PickColumn("Why It Changed (2 of 2)", "Explain the change by which column (customer, product, plant ...)?", k.H, k.Group);
            if (gc == 0) return;
            int mode = Ui.Pick("Why It Changed: which periods", "Compare:", new[]
            {
                "The latest month with the month before",
                "The latest month with the same month last year",
                "The latest fiscal year with the fiscal year before (April to March)"
            }, null);
            if (mode < 0) return;
            string msg;
            using (new Speed())
            {
                msg = VarianceCore(src, v, k, dc, mc, gc, mode);
            }
            Ui.Info(msg);
        }

        static string VarianceCore(dynamic src, object[,] v, Cols k, int dc, int mc, int gc, int mode)
        {
            int nr = v.GetLength(0), nc = v.GetLength(1);
            DateTime latest = DateTime.MinValue;
            for (int r = k.Hr + 1; r <= nr; r++)
            {
                DateTime d;
                if (Logic.TryDate(v[r, dc], out d) && d > latest) latest = d;
            }
            if (latest == DateTime.MinValue) return "No dates could be read in that column.";
            string curKey, prevKey, curName, prevName;
            if (mode == 2)
            {
                curKey = curName = Logic.FYLabel(latest, 4);
                prevKey = prevName = Logic.FYLabel(latest.AddYears(-1), 4);
            }
            else
            {
                DateTime p = latest.AddMonths(mode == 0 ? -1 : -12);
                curKey = latest.ToString("yyyy-MM", Inv);
                prevKey = p.ToString("yyyy-MM", Inv);
                curName = latest.ToString("MMM-yyyy", Inv);
                prevName = p.ToString("MMM-yyyy", Inv);
            }
            Dictionary<string, double[]> agg = new Dictionary<string, double[]>(StringComparer.OrdinalIgnoreCase);   // previous, current
            List<string> names = new List<string>();
            double totPrev = 0, totCur = 0;
            int rowsPrev = 0, rowsCur = 0;
            for (int r = k.Hr + 1; r <= nr; r++)
            {
                DateTime d;
                double x;
                if (!Profiler.RowIsData(v, r, nc) || !Logic.TryDate(v[r, dc], out d) || !Logic.TryNum(v[r, mc], out x)) continue;
                string key = mode == 2 ? Logic.FYLabel(d, 4) : d.ToString("yyyy-MM", Inv);
                int slot = key == curKey ? 1 : (key == prevKey ? 0 : -1);
                if (slot < 0) continue;
                string g = Xl.GroupKey(v[r, gc]);
                double[] a;
                if (!agg.TryGetValue(g, out a))
                {
                    a = new double[2];
                    agg[g] = a;
                    names.Add(g);
                }
                a[slot] += x;
                if (slot == 1)
                {
                    totCur += x;
                    rowsCur++;
                }
                else
                {
                    totPrev += x;
                    rowsPrev++;
                }
            }
            if (rowsPrev == 0) return "There are no rows for " + prevName + ", so " + curName + " has nothing to be compared with.";
            if (rowsCur == 0) return "There are no rows for " + curName + ".";
            names.Sort(delegate (string p, string q)
            {
                return Math.Abs(agg[q][1] - agg[q][0]).CompareTo(Math.Abs(agg[p][1] - agg[p][0]));
            });
            int n = names.Count;
            double change = totCur - totPrev;
            string mName = k.H[mc], gName = k.H[gc];

            object[,] t = Xl.New2D(n + 1, 6);
            t[1, 1] = gName;
            t[1, 2] = prevName;
            t[1, 3] = curName;
            t[1, 4] = "Change";
            t[1, 5] = "Change %";
            t[1, 6] = "Share of the total change";
            List<long> up = new List<long>(), down = new List<long>();
            string falls = "", rises = "";
            int nFalls = 0, nRises = 0;
            for (int i = 0; i < n; i++)
            {
                double[] a = agg[names[i]];
                double d = a[1] - a[0];
                t[i + 2, 1] = names[i];
                t[i + 2, 2] = a[0];
                t[i + 2, 3] = a[1];
                t[i + 2, 4] = d;
                if (a[0] != 0) t[i + 2, 5] = d / Math.Abs(a[0]);
                if (change != 0) t[i + 2, 6] = d / change;
                if (d > 0)
                {
                    up.Add(Xl.Cell(i + 2, 4));
                    if (nRises < 3)
                    {
                        rises += (nRises > 0 ? ", " : "") + names[i] + " (" + Signed(d) + ")";
                        nRises++;
                    }
                }
                else if (d < 0)
                {
                    down.Add(Xl.Cell(i + 2, 4));
                    if (nFalls < 3)
                    {
                        falls += (nFalls > 0 ? ", " : "") + names[i] + " (" + Signed(d) + ")";
                        nFalls++;
                    }
                }
            }
            string headline = mName + (change > 0 ? " went up by " : change < 0 ? " went down by " : " did not change: ") +
                              (change != 0 ? Logic.ShortINR(Math.Abs(change)) : "") +
                              (totPrev != 0 && change != 0 ? " (" + (change / Math.Abs(totPrev)).ToString("+0%;-0%", Inv) + ")" : "") +
                              ", from " + Logic.ShortINR(totPrev) + " in " + prevName + " to " + Logic.ShortINR(totCur) + " in " + curName + ".";

            dynamic ws = Xl.AddSheet(src.Worksheet.Parent, "Why It Changed");
            ws.Range["A1"].Value2 = "Why it changed: " + mName + ", " + curName + " against " + prevName;
            ws.Range["A1"].Font.Size = 14;
            ws.Range["A1"].Font.Bold = true;
            ws.Range["A2"].Value2 = headline;
            ws.Range["A3"].Value2 = nFalls > 0 ? "Biggest falls: " + falls + "." : "No " + gName + " fell.";
            ws.Range["A4"].Value2 = nRises > 0 ? "Biggest rises: " + rises + "." : "No " + gName + " rose.";
            dynamic top = ws.Range["A6"];
            Xl.WriteTable(top, t, 1);
            dynamic tbl = top.Resize[n + 1, 6];
            Xl.StyleHeader(top.Resize[1, 6]);
            ws.Range["B7"].Resize[n, 3].NumberFormat = "#,##0.00";
            ws.Range["E7"].Resize[n, 2].NumberFormat = "0%";
            Xl.Paint(tbl, up, Design.Good);
            Xl.Paint(tbl, down, Design.Bad);
            ws.Columns["A:F"].AutoFit();
            try
            {
                int show = Math.Min(n, 10);
                dynamic anchor = ws.Range["H6:R26"];
                dynamic co = ws.ChartObjects().Add((double)anchor.Left, (double)anchor.Top, (double)anchor.Width, (double)anchor.Height);
                dynamic ch = co.Chart;
                ch.ChartType = C.xlBarClustered;
                dynamic s1 = ch.SeriesCollection().NewSeries();
                s1.Name = "Change";
                s1.Values = ws.Range["D7"].Resize[show, 1];
                s1.XValues = ws.Range["A7"].Resize[show, 1];
                ch.HasTitle = true;
                ch.ChartTitle.Text = "Biggest changes in " + mName + " by " + gName;
                ch.HasLegend = false;
            }
            catch { }
            ws.Activate();
            return headline + "\n\n" + (nFalls > 0 ? "Biggest falls: " + falls + ".\n" : "") + (nRises > 0 ? "Biggest rises: " + rises + ".\n" : "") +
                   "\nThe full list is on the new 'Why It Changed' sheet. Green = went up, red = went down.";
        }

        // =================================================================
        //  FORECAST
        // =================================================================
        public static void Forecast()
        {
            dynamic src = Xl.PickRange("Forecast");
            if (src == null) return;
            object[,] v = Xl.ReadTyped(src);
            if (v.GetLength(0) < 3 || v.GetLength(1) < 2)
            {
                Ui.Info("Select a table with a date column and a number column.");
                return;
            }
            Cols k = Describe(v);
            if (k.NDate == 0)
            {
                Ui.Info("No date column was found. If the dates are stored as text, run Smart Clean first.");
                return;
            }
            if (!Any(k.Num))
            {
                Ui.Info("No number column was found to forecast.");
                return;
            }
            int dc = k.NDate == 1 ? k.FirstDate : Ui.PickColumn("Forecast", "Which column holds the date?", k.H, k.Date);
            if (dc == 0) return;
            int mc = Ui.PickColumn("Forecast", "Which number do you want to forecast?", k.H, k.Num);
            if (mc == 0) return;
            int hz = Ui.Pick("Forecast: how far ahead", "Forecast the next:", new[] { "3 months", "6 months", "12 months" }, null);
            if (hz < 0) return;
            string msg;
            using (new Speed())
            {
                msg = ForecastCore(src, v, k, dc, mc, hz == 0 ? 3 : (hz == 1 ? 6 : 12));
            }
            Ui.Info(msg);
        }

        static string ForecastCore(dynamic src, object[,] v, Cols k, int dc, int mc, int horizon)
        {
            int nr = v.GetLength(0), nc = v.GetLength(1);
            Dictionary<int, double> sums = new Dictionary<int, double>();
            int lo = int.MaxValue, hi = int.MinValue;
            for (int r = k.Hr + 1; r <= nr; r++)
            {
                DateTime d;
                double x;
                if (!Profiler.RowIsData(v, r, nc) || !Logic.TryDate(v[r, dc], out d) || !Logic.TryNum(v[r, mc], out x)) continue;
                int ym = d.Year * 12 + d.Month - 1;
                double cur;
                sums[ym] = sums.TryGetValue(ym, out cur) ? cur + x : x;
                if (ym < lo) lo = ym;
                if (ym > hi) hi = ym;
            }
            if (sums.Count == 0) return "No rows had both a date and a number.";
            // a month that is still running would pull the forecast down, so it is left out
            DateTime today = DateTime.Today;
            int thisMonth = today.Year * 12 + today.Month - 1;
            bool dropped = false;
            if (hi == thisMonth && today.Day < DateTime.DaysInMonth(today.Year, today.Month))
            {
                hi--;
                dropped = true;
            }
            int n = hi - lo + 1;
            if (n < 6) return "At least 6 complete months of history are needed. This table has " + Math.Max(n, 0) + ".";
            double[] y = new double[n];
            bool nonNeg = true;
            for (int i = 0; i < n; i++)
            {
                double val;
                y[i] = sums.TryGetValue(lo + i, out val) ? val : 0;
                if (y[i] < 0) nonNeg = false;
            }
            // fit on the last 36 months at most: straight-line trend
            int s0 = Math.Max(0, n - 36), m = n - s0;
            double sx = 0, sy = 0, sxx = 0, sxy = 0;
            for (int i = s0; i < n; i++)
            {
                sx += i;
                sy += y[i];
                sxx += (double)i * i;
                sxy += i * y[i];
            }
            double den = m * sxx - sx * sx;
            double slope = den == 0 ? 0 : (m * sxy - sx * sy) / den;
            double icpt = (sy - slope * sx) / m;
            // with two years or more, each calendar month gets its own seasonal factor
            double[] factor = new double[12];
            for (int j = 0; j < 12; j++) factor[j] = 1;
            bool seasonal = m >= 24;
            if (seasonal)
            {
                double[] fs = new double[12];
                int[] fn = new int[12];
                for (int i = s0; i < n; i++)
                {
                    double tr = icpt + slope * i;
                    if (tr <= 0) continue;
                    int cm = (lo + i) % 12;
                    fs[cm] += y[i] / tr;
                    fn[cm]++;
                }
                double mean = 0;
                int have = 0;
                for (int j = 0; j < 12; j++)
                {
                    if (fn[j] == 0) continue;
                    factor[j] = fs[j] / fn[j];
                    mean += factor[j];
                    have++;
                }
                if (have > 0 && mean > 0)
                {
                    mean /= have;
                    for (int j = 0; j < 12; j++) factor[j] = fn[j] == 0 ? 1 : factor[j] / mean;
                }
                else seasonal = false;
            }
            double sse = 0;
            for (int i = s0; i < n; i++)
            {
                double fit = (icpt + slope * i) * factor[(lo + i) % 12];
                sse += (y[i] - fit) * (y[i] - fit);
            }
            double sd = m > 2 ? Math.Sqrt(sse / (m - 2)) : 0;

            int shown = Math.Min(n, 36), first = n - shown;
            object[,] t = Xl.New2D(shown + horizon + 1, 5);
            t[1, 1] = "Month";
            t[1, 2] = "Actual";
            t[1, 3] = "Forecast";
            t[1, 4] = "Likely low";
            t[1, 5] = "Likely high";
            for (int i = first; i < n; i++)
            {
                t[i - first + 2, 1] = MonthName(lo + i);
                t[i - first + 2, 2] = y[i];
            }
            t[shown + 1, 3] = y[n - 1];                               // joins the two lines on the chart
            double total = 0, next = 0;
            for (int j = 1; j <= horizon; j++)
            {
                int i = n - 1 + j;
                double f = (icpt + slope * i) * factor[(lo + i) % 12];
                double low = f - 1.28 * sd, high = f + 1.28 * sd;
                if (nonNeg)
                {
                    if (f < 0) f = 0;
                    if (low < 0) low = 0;
                }
                t[shown + 1 + j, 1] = MonthName(lo + i);
                t[shown + 1 + j, 3] = f;
                t[shown + 1 + j, 4] = low;
                t[shown + 1 + j, 5] = high;
                total += f;
                if (j == 1) next = f;
            }
            string mName = k.H[mc];
            dynamic ws = Xl.AddSheet(src.Worksheet.Parent, "Forecast");
            ws.Range["A1"].Value2 = "Forecast: " + mName + ", next " + horizon + " months";
            ws.Range["A1"].Font.Size = 14;
            ws.Range["A1"].Font.Bold = true;
            ws.Range["A2"].Value2 = "Method: straight-line trend over the last " + m + " months" + (seasonal ? ", adjusted for each calendar month's usual high or low." : ". With 24 months or more of history, seasonality is added.");
            ws.Range["A3"].Value2 = "Likely low / high is the range the number stayed inside about 8 times out of 10 in the past. A forecast is a guide, not a promise." +
                                    (dropped ? "  " + MonthName(thisMonth) + " is not complete yet, so it was left out." : "");
            dynamic top = ws.Range["A5"];
            Xl.WriteTable(top, t, 1);
            Xl.StyleHeader(top.Resize[1, 5]);
            ws.Range["B6"].Resize[shown + horizon, 4].NumberFormat = "#,##0";
            ws.Range["A6"].Resize[shown + horizon, 1].EntireColumn.AutoFit();
            ws.Range["C" + (6 + shown)].Resize[horizon, 3].Interior.Color = Design.Changed;
            try
            {
                dynamic anchor = ws.Range["H5:S26"];
                dynamic co = ws.ChartObjects().Add((double)anchor.Left, (double)anchor.Top, (double)anchor.Width, (double)anchor.Height);
                dynamic ch = co.Chart;
                ch.SetSourceData(top.Resize[shown + horizon + 1, 3]);
                ch.ChartType = C.xlLineMarkers;
                ch.HasTitle = true;
                ch.ChartTitle.Text = mName + ": actual and forecast";
                ch.HasLegend = true;
            }
            catch { }
            ws.Activate();
            return "Forecast ready on the new 'Forecast' sheet.\n\nNext month (" + MonthName(hi + 1) + "): about " + Logic.ShortINR(next) +
                   "\nNext " + horizon + " months together: about " + Logic.ShortINR(total) +
                   "\n\nBased on " + m + " months of history" + (seasonal ? ", with seasonality." : ".") +
                   (dropped ? "\n" + MonthName(thisMonth) + " is not complete yet, so it was left out." : "") +
                   "\nA forecast is a guide, not a promise.";
        }
    }
}
