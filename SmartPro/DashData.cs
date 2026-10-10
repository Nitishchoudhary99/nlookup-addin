using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace SmartPro
{
    // =====================================================================
    //  THE TABLE BEHIND A DASHBOARD, read once into typed columns, and the arithmetic on it.
    //  Dashboard Studio uses it for the live preview, the builder for the helper sheet and the insights,
    //  the Web Dashboard for the data it carries. No Excel calls in this file.
    // =====================================================================
    internal sealed class DataCol
    {
        public string Name = "";
        public string Kind = "empty";     // what the profiler said: date, num, cat, id, text
        public int Index;                 // 1-based column in the source range
        public bool IsNum, IsCat, IsDate;
        public int Distinct;
        public double[] Num;              // one number for each data row, NaN = no number      (when IsNum)
        public int[] Code;                // one code for each data row: its position in Items  (when IsCat)
        public List<string> Items;        // the different values, in order of first appearance
        public int[] Ymd;                 // year * 10000 + month * 100 + day, 0 = no date      (when IsDate)
        public bool[] Filled;             // true where the cell has something in it            (when it is none of the three)
    }

    internal sealed class DataTab
    {
        public const int MaxItems = 20000;       // a column with more different values than this cannot be grouped by
        public const int MaxNumItems = 200;      // the same limit for a column of measures
        public const string NoDate = "(no date)", Blank = "(blank)";

        public int Rows, Skipped, HeaderRow;
        public bool Sampled;                     // true when only the first rows were read
        public List<DataCol> Cols = new List<DataCol>();
        public int DateIndex = -1;               // position in Cols of the date column behind month / quarter / year
        readonly Dictionary<string, int[]> keyCache = new Dictionary<string, int[]>(StringComparer.Ordinal);
        readonly Dictionary<string, List<string>> labelCache = new Dictionary<string, List<string>>(StringComparer.Ordinal);

        static readonly string[] MonthAbbr = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
        static readonly string[] DayAbbr = { "Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun" };

        // v is the block read from Excel (1-based, header row included). maxRows = 0 reads every row.
        public static DataTab Load(object[,] v, int maxRows)
        {
            DataTab t = new DataTab();
            int hr = Profiler.FindHeaderRow(v);
            string[] h = Profiler.MakeHeaders(v, hr);
            ColInfo[] ci = Profiler.Profile(v, hr, h, 50000);
            int nr = v.GetLength(0), nc = v.GetLength(1);
            int last = nr;
            if (maxRows > 0 && nr - hr > maxRows)
            {
                last = hr + maxRows;
                t.Sampled = true;
            }
            int[] rowAt = new int[Math.Max(last - hr, 0)];
            int n = 0;
            for (int r = hr + 1; r <= last; r++)
            {
                if (Profiler.RowIsData(v, r, nc)) rowAt[n++] = r;
                else t.Skipped++;
            }
            t.Rows = n;
            t.HeaderRow = hr;
            for (int c = 1; c <= nc; c++)
            {
                ColInfo k = ci[c];
                if (k.Kind == "empty" || k.NonBlank == 0) continue;
                DataCol col = new DataCol();
                col.Name = h[c];
                col.Kind = k.Kind;
                col.Index = c;
                col.Distinct = k.Distinct;
                if (k.Kind == "date")
                {
                    col.IsDate = true;
                    col.Ymd = new int[n];
                    for (int i = 0; i < n; i++)
                    {
                        DateTime dt;
                        if (Logic.TryDate(v[rowAt[i], c], out dt)) col.Ymd[i] = dt.Year * 10000 + dt.Month * 100 + dt.Day;
                    }
                }
                else
                {
                    if (k.NumCount >= 0.8 * k.NonBlank)
                    {
                        col.IsNum = true;
                        col.Num = new double[n];
                        for (int i = 0; i < n; i++)
                        {
                            double d;
                            col.Num[i] = Logic.TryNum(v[rowAt[i], c], out d) ? d : double.NaN;
                        }
                    }
                    // a true measure (amounts, weights) is only worth grouping by when it has few different values
                    if (k.Distinct <= (k.Kind == "num" ? MaxNumItems : MaxItems))
                    {
                        // pivot tables do not tell "abc" from "ABC", so the groups here do not either
                        Dictionary<string, int> map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                        List<string> items = new List<string>();
                        int[] code = new int[n];
                        bool tooMany = false;
                        for (int i = 0; i < n; i++)
                        {
                            object x = v[rowAt[i], c];
                            string s = Xl.IsBlank(x) ? Blank : Xl.Str(x).Trim();
                            int id;
                            if (!map.TryGetValue(s, out id))
                            {
                                if (items.Count >= MaxItems)
                                {
                                    tooMany = true;
                                    break;
                                }
                                id = items.Count;
                                map[s] = id;
                                items.Add(s);
                            }
                            code[i] = id;
                        }
                        if (!tooMany)
                        {
                            col.IsCat = true;
                            col.Code = code;
                            col.Items = items;
                            col.Distinct = items.Count;
                        }
                    }
                    // a column of text with too many different values to group by can still be counted
                    if (!col.IsNum && !col.IsCat)
                    {
                        col.Filled = new bool[n];
                        for (int i = 0; i < n; i++) col.Filled[i] = !Xl.IsBlank(v[rowAt[i], c]);
                    }
                }
                t.Cols.Add(col);
            }
            t.UseDate("");
            return t;
        }

        // picks the date column behind month / quarter / year: the named one, or the first date column
        public void UseDate(string name)
        {
            int found = -1;
            for (int i = 0; i < Cols.Count; i++)
            {
                if (!Cols[i].IsDate) continue;
                if (found < 0) found = i;
                if (name != null && name.Length > 0 && string.Equals(Cols[i].Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    found = i;
                    break;
                }
            }
            if (found != DateIndex)
            {
                DateIndex = found;
                keyCache.Clear();
                labelCache.Clear();
            }
        }

        public bool HasDate { get { return DateIndex >= 0; } }

        public string DateName { get { return DateIndex >= 0 ? Cols[DateIndex].Name : ""; } }

        public DataCol Find(string name)
        {
            if (name == null || name.Length == 0) return null;
            foreach (DataCol c in Cols)
            {
                if (string.Equals(c.Name, name, StringComparison.OrdinalIgnoreCase)) return c;
            }
            return null;
        }

        // can the rows be grouped by this column or date part?
        public bool CanGroup(string dim)
        {
            if (DatePart.Is(dim)) return HasDate && Array.IndexOf(DatePart.All, dim) >= 0;
            DataCol c = Find(dim);
            return c != null && c.IsCat;
        }

        public static string FiscalYear(int year, int month, int fyStart)
        {
            int y = month < fyStart ? year - 1 : year;
            if (fyStart == 1) return "FY " + y.ToString(CultureInfo.InvariantCulture);
            return "FY " + y.ToString(CultureInfo.InvariantCulture) + "-" + ((y + 1) % 100).ToString("00", CultureInfo.InvariantCulture);
        }

        // the name of the date part of one date; these names sort in time order
        public static string PartLabel(string part, int ymd, int fyStart)
        {
            if (ymd <= 0) return NoDate;
            int y = ymd / 10000, m = (ymd / 100) % 100, d = ymd % 100;
            if (part == DatePart.Month)
                return y.ToString("0000", CultureInfo.InvariantCulture) + "-" + m.ToString("00", CultureInfo.InvariantCulture) + " (" + MonthAbbr[m - 1] + ")";
            if (part == DatePart.Quarter)
            {
                int q = ((m - fyStart + 12) % 12) / 3 + 1;
                if (fyStart == 1) return y.ToString(CultureInfo.InvariantCulture) + " Q" + q.ToString(CultureInfo.InvariantCulture);
                return FiscalYear(y, m, fyStart) + " Q" + q.ToString(CultureInfo.InvariantCulture);
            }
            if (part == DatePart.FiscalYear) return FiscalYear(y, m, fyStart);
            if (part == DatePart.Year) return y.ToString("0000", CultureInfo.InvariantCulture);
            if (part == DatePart.Weekday)
            {
                int dow = ((int)new DateTime(y, m, d).DayOfWeek + 6) % 7;        // Monday = 0
                return (dow + 1).ToString(CultureInfo.InvariantCulture) + " " + DayAbbr[dow];
            }
            return y.ToString("0000", CultureInfo.InvariantCulture) + "-" + m.ToString("00", CultureInfo.InvariantCulture) + "-" + d.ToString("00", CultureInfo.InvariantCulture);
        }

        // The short name of a date group for a chart axis: "2026-04 (Apr)" becomes "Apr 26", "FY 2026-27 Q1" becomes
        // "Q1 26-27", "2 Tue" becomes "Tue", "2026-04-17" becomes "17 Apr 26". Other names stay as they are.
        // Short names do not sort in time order, so they are only used where the order is set by hand.
        public static string ShortLabel(string part, string label)
        {
            if (label == null || label == NoDate) return label;
            if (part == DatePart.Month)
            {
                if (label.Length == 13 && label[4] == '-' && label[8] == '(' && label[12] == ')') return label.Substring(9, 3) + " " + label.Substring(2, 2);
                return label;
            }
            if (part == DatePart.Quarter)
            {
                int q = label.LastIndexOf(" Q", StringComparison.Ordinal);
                if (q < 0 || q + 3 != label.Length) return label;
                string head = label.Substring(0, q), tail = label.Substring(q + 1);
                if (head.StartsWith("FY ", StringComparison.Ordinal) && head.Length == 10) return tail + " " + head.Substring(5);
                return tail + " " + head;
            }
            if (part == DatePart.Weekday)
            {
                if (label.Length == 5 && label[1] == ' ' && char.IsDigit(label[0])) return label.Substring(2);
                return label;
            }
            if (part == DatePart.Day)
            {
                int m;
                if (label.Length == 10 && label[4] == '-' && label[7] == '-' &&
                    int.TryParse(label.Substring(5, 2), NumberStyles.Integer, CultureInfo.InvariantCulture, out m) && m >= 1 && m <= 12)
                    return label.Substring(8, 2) + " " + MonthAbbr[m - 1] + " " + label.Substring(2, 2);
                return label;
            }
            return label;
        }

        // true when ShortLabel gives this date part other names than the ones that sort in time order
        public static bool HasShortLabels(string part)
        {
            return part == DatePart.Month || part == DatePart.Quarter || part == DatePart.Weekday || part == DatePart.Day;
        }

        // One group number for every data row and the names of the groups. null when the rows cannot be grouped this way.
        public int[] Keys(string dim, int fyStart, out List<string> labels)
        {
            labels = null;
            if (!CanGroup(dim)) return null;
            if (!DatePart.Is(dim))
            {
                DataCol c = Find(dim);
                labels = c.Items;
                return c.Code;
            }
            string ck = dim + "|" + fyStart.ToString(CultureInfo.InvariantCulture);
            int[] key;
            if (keyCache.TryGetValue(ck, out key))
            {
                labels = labelCache[ck];
                return key;
            }
            int[] ymd = Cols[DateIndex].Ymd;
            key = new int[Rows];
            List<string> names = new List<string>();
            Dictionary<int, int> byNumber = new Dictionary<int, int>();
            for (int i = 0; i < Rows; i++)
            {
                int x = ymd[i];
                int num;
                if (x <= 0) num = 0;
                else if (dim == DatePart.Day) num = x;
                else if (dim == DatePart.Weekday) num = 1 + ((int)new DateTime(x / 10000, (x / 100) % 100, x % 100).DayOfWeek + 6) % 7;
                else if (dim == DatePart.Year) num = x / 10000;
                else if (dim == DatePart.Month) num = x / 100;
                else
                {
                    int y = x / 10000, m = (x / 100) % 100;
                    int fy = m < fyStart ? y - 1 : y;
                    num = dim == DatePart.FiscalYear ? fy : fy * 10 + ((m - fyStart + 12) % 12) / 3 + 1;
                }
                int id;
                if (!byNumber.TryGetValue(num, out id))
                {
                    id = names.Count;
                    byNumber[num] = id;
                    names.Add(PartLabel(dim, x, fyStart));
                }
                key[i] = id;
            }
            keyCache[ck] = key;
            labelCache[ck] = names;
            labels = names;
            return key;
        }
    }

    internal sealed class Series
    {
        public string Name = "";
        public double[] V = new double[0];
        public bool Percent;              // the values are shares, 0 to 1
    }

    // What one chart shows, worked out from the table
    internal sealed class AggResult
    {
        public List<string> Cats = new List<string>();
        public List<Series> S = new List<Series>();
        public bool Time;                 // the groups are in time order
        public int AllCats;               // how many groups there are before the top N cut
        public int TopN;                  // > 0: only the N largest are shown
        public bool Others;               // the rest is added up as one group called OthersName
        public string OthersName = "Others";
        public int SortMode;              // 0 by name (time order), 1 largest first, 2 smallest first, 3 A to Z
        public List<string> Keep = new List<string>();          // the groups that are shown, without Others
        public List<string> SplitKeep = new List<string>();     // the series of a stacked chart, without Others
        public bool SplitOthers;
        public string SplitOthersName = "Others";
        public double MaxAbs;             // largest absolute value of the first series (for axis units)

        public double[] First { get { return S.Count > 0 ? S[0].V : new double[0]; } }
    }

    internal static class DashEngine
    {
        const int MaxSeries = 8;          // colours in a stacked chart

        // ---------- one number: a KPI card ----------
        public static double Scalar(DataTab t, string col, int agg)
        {
            int[] one = new int[t.Rows];
            double[] r = Reduce(t, one, 1, col, agg);
            return r.Length > 0 ? r[0] : 0;
        }

        // Adds up one column inside every group. key[row] = group number, or -1 to leave the row out.
        // col "" counts rows.
        public static double[] Reduce(DataTab t, int[] key, int groups, string col, int agg)
        {
            double[] res = new double[groups];
            if (groups == 0) return res;
            int n = t.Rows;
            DataCol c = t.Find(col);
            if (col == null || col.Length == 0 || c == null)
            {
                for (int i = 0; i < n; i++) if (key[i] >= 0) res[key[i]] += 1;
                return res;
            }
            if (agg == AggKind.Distinct)
            {
                HashSet<long> seen = new HashSet<long>();
                if (c.IsCat)
                {
                    for (int i = 0; i < n; i++)
                    {
                        if (key[i] < 0) continue;
                        if (c.Items[c.Code[i]] == DataTab.Blank) continue;
                        if (seen.Add((long)key[i] * (DataTab.MaxItems + 1) + c.Code[i])) res[key[i]] += 1;
                    }
                }
                else if (c.IsNum)
                {
                    Dictionary<double, int> ids = new Dictionary<double, int>();
                    for (int i = 0; i < n; i++)
                    {
                        if (key[i] < 0 || double.IsNaN(c.Num[i])) continue;
                        int id;
                        if (!ids.TryGetValue(c.Num[i], out id))
                        {
                            id = ids.Count;
                            ids[c.Num[i]] = id;
                        }
                        if (seen.Add((long)key[i] * 2000000000L + id)) res[key[i]] += 1;
                    }
                }
                return res;
            }
            if (agg == AggKind.Count)
            {
                for (int i = 0; i < n; i++)
                {
                    if (key[i] < 0) continue;
                    bool filled;
                    if (c.IsNum) filled = !double.IsNaN(c.Num[i]);
                    else if (c.IsCat) filled = c.Items[c.Code[i]] != DataTab.Blank;
                    else if (c.IsDate) filled = c.Ymd[i] > 0;
                    else filled = c.Filled != null && c.Filled[i];
                    if (filled) res[key[i]] += 1;
                }
                return res;
            }
            if (!c.IsNum) return res;
            double[] sum = new double[groups];
            int[] cnt = new int[groups];
            double[] mx = new double[groups], mn = new double[groups];
            for (int i = 0; i < n; i++)
            {
                int g = key[i];
                if (g < 0) continue;
                double x = c.Num[i];
                if (double.IsNaN(x)) continue;
                if (cnt[g] == 0)
                {
                    mx[g] = x;
                    mn[g] = x;
                }
                else
                {
                    if (x > mx[g]) mx[g] = x;
                    if (x < mn[g]) mn[g] = x;
                }
                sum[g] += x;
                cnt[g]++;
            }
            for (int g = 0; g < groups; g++)
            {
                if (agg == AggKind.Avg) res[g] = cnt[g] > 0 ? sum[g] / cnt[g] : 0;
                else if (agg == AggKind.Max) res[g] = cnt[g] > 0 ? mx[g] : 0;
                else if (agg == AggKind.Min) res[g] = cnt[g] > 0 ? mn[g] : 0;
                else res[g] = sum[g];
            }
            return res;
        }

        // the chart kind decides what "automatic" means for top N and for the order
        public static int AutoTop(int kind, int groups)
        {
            if (ChartKind.IsRound(kind)) return groups > 7 ? 6 : 0;
            if (kind == ChartKind.Bar) return groups > 12 ? 10 : 0;
            if (kind == ChartKind.Pareto) return groups > 15 ? 15 : 0;
            if (kind == ChartKind.Table) return groups > 10 ? 10 : 0;
            if (ChartKind.IsLine(kind) || kind == ChartKind.Combo) return groups > 24 ? 20 : 0;
            return groups > 12 ? 10 : 0;
        }

        // true when the chart groups by a date part (its own groups or its colours): rows without a date are then left out
        public static bool UsesDate(DataTab t, ChartSpec c)
        {
            if (!t.HasDate) return false;
            if (DatePart.Is(c.Dim)) return true;
            return ChartKind.IsStacked(c.Kind) && DatePart.Is(c.Split) && c.Split != c.Dim;
        }

        // how many data rows have no date in the date column
        public static int Undated(DataTab t)
        {
            if (!t.HasDate) return 0;
            int n = 0;
            foreach (int x in t.Cols[t.DateIndex].Ymd) if (x <= 0) n++;
            return n;
        }

        // "% of total" and "running total" are for charts with one series
        public static bool CanShowAs(int kind)
        {
            return kind == ChartKind.Column || kind == ChartKind.Bar || kind == ChartKind.Line || kind == ChartKind.Area;
        }

        static string FreeLabel(List<string> labels, string want)
        {
            bool clash = true;
            while (clash)
            {
                clash = false;
                foreach (string s in labels)
                {
                    if (string.Equals(s, want, StringComparison.OrdinalIgnoreCase))
                    {
                        clash = true;
                        break;
                    }
                }
                if (clash) want += " (rest)";
            }
            return want;
        }

        // order of the groups: positions into 'ids', sorted. A stable sort, so equal values keep their first order.
        static void SortIds(List<int> ids, double[] val, List<string> labels, int mode)
        {
            int n = ids.Count;
            int[] a = ids.ToArray();
            int[] b = new int[n];
            // bottom-up merge sort
            for (int width = 1; width < n; width *= 2)
            {
                for (int lo = 0; lo < n; lo += 2 * width)
                {
                    int mid = Math.Min(lo + width, n), hi = Math.Min(lo + 2 * width, n);
                    int i = lo, j = mid, k = lo;
                    while (i < mid && j < hi)
                    {
                        if (Before(a[j], a[i], val, labels, mode)) b[k++] = a[j++];
                        else b[k++] = a[i++];
                    }
                    while (i < mid) b[k++] = a[i++];
                    while (j < hi) b[k++] = a[j++];
                }
                int[] tmp = a;
                a = b;
                b = tmp;
            }
            ids.Clear();
            for (int i = 0; i < n; i++) ids.Add(a[i]);
        }

        // true when group x must come strictly before group y
        static bool Before(int x, int y, double[] val, List<string> labels, int mode)
        {
            if (mode == 1) return val[x] > val[y];
            if (mode == 2) return val[x] < val[y];
            // by name, the way a pivot table does it: capitals and small letters are the same
            int c = string.Compare(labels[x], labels[y], StringComparison.OrdinalIgnoreCase);
            return c != 0 ? c < 0 : string.CompareOrdinal(labels[x], labels[y]) < 0;
        }

        // By value, and groups with the same value by name (a pivot table leaves equal values in name order).
        static void SortByValue(List<int> ids, double[] val, List<string> labels, int mode)
        {
            SortIds(ids, val, labels, 3);
            SortIds(ids, val, labels, mode);
        }

        // ---------- many numbers: a chart ----------
        public static AggResult Group(DataTab t, ChartSpec c, DashSpec s)
        {
            AggResult res = new AggResult();
            List<string> labels;
            int[] key = t.Keys(c.Dim, s.FyStart, out labels);
            if (key == null) return res;
            int ng = labels.Count;
            res.Time = DatePart.Is(c.Dim);
            // a chart over time shows only the rows that have a date
            if (UsesDate(t, c))
            {
                int[] ymd = t.Cols[t.DateIndex].Ymd;
                int[] dated = new int[t.Rows];
                for (int i = 0; i < t.Rows; i++) dated[i] = ymd[i] > 0 ? key[i] : -1;
                key = dated;
            }

            // which groups have rows, and their value
            double[] rows = Reduce(t, key, ng, "", AggKind.Count);
            double[] val = Reduce(t, key, ng, c.Col, c.Agg);
            List<int> ids = new List<int>();
            for (int g = 0; g < ng; g++) if (rows[g] > 0) ids.Add(g);
            res.AllCats = ids.Count;

            // the order
            int mode = c.Sort;
            if (mode == 0)
            {
                if (res.Time || ChartKind.IsLine(c.Kind) || c.Kind == ChartKind.Combo) mode = res.Time ? 0 : 3;
                else mode = 1;
            }
            if (c.Kind == ChartKind.Pareto) mode = 1;
            res.SortMode = mode;

            // top N: always the N largest, then shown in the chosen order
            int top = c.Top;
            if (top == 0) top = res.Time ? 0 : AutoTop(c.Kind, ids.Count);
            if (res.Time && c.Top <= 0) top = 0;
            // a running total needs every group, so it is never cut
            if (c.ShowAs == 2 && CanShowAs(c.Kind)) top = 0;
            bool cut = top > 0 && ids.Count > top;
            List<int> keep = new List<int>(ids);
            if (cut)
            {
                SortByValue(keep, val, labels, 1);
                keep.RemoveRange(top, keep.Count - top);
                res.TopN = top;
                res.Others = c.Others == 1 || (c.Others == 0 && ChartKind.IsRound(c.Kind));
                if (c.Kind == ChartKind.Pareto) res.Others = false;
            }
            if (mode == 1 || mode == 2) SortByValue(keep, val, labels, mode);
            else SortIds(keep, val, labels, mode);

            // new group numbers in display order; the rest becomes Others or is left out
            int nk = keep.Count;
            int[] place = new int[ng];
            for (int g = 0; g < ng; g++) place[g] = -1;
            for (int i = 0; i < nk; i++)
            {
                place[keep[i]] = i;
                res.Keep.Add(labels[keep[i]]);
                res.Cats.Add(labels[keep[i]]);
            }
            int total = nk;
            if (cut && res.Others)
            {
                res.OthersName = FreeLabel(labels, "Others");
                for (int g = 0; g < ng; g++) if (place[g] < 0 && rows[g] > 0) place[g] = nk;
                res.Cats.Add(res.OthersName);
                total = nk + 1;
            }
            int[] k2 = new int[t.Rows];
            for (int i = 0; i < t.Rows; i++) k2[i] = key[i] < 0 ? -1 : place[key[i]];

            if (ChartKind.IsStacked(c.Kind) && t.CanGroup(c.Split) && c.Split != c.Dim)
            {
                AddSplitSeries(t, c, s, res, k2, total);
            }
            else
            {
                Series s1 = new Series();
                s1.Name = DashText.Measure(c.Col, c.Agg);
                s1.V = Reduce(t, k2, total, c.Col, c.Agg);
                res.S.Add(s1);
                if (c.Kind == ChartKind.Combo || (c.Kind == ChartKind.Table && c.Col2.Length > 0))
                {
                    Series s2 = new Series();
                    s2.Name = DashText.Measure(c.Col2, c.Agg2);
                    s2.V = Reduce(t, k2, total, c.Col2, c.Agg2);
                    res.S.Add(s2);
                }
                if (c.Kind == ChartKind.Pareto)
                {
                    Series cum = new Series();
                    cum.Name = "Cumulative share";
                    cum.Percent = true;
                    cum.V = new double[total];
                    double all = 0, run = 0;
                    for (int i = 0; i < total; i++) all += s1.V[i];
                    for (int i = 0; i < total; i++)
                    {
                        run += s1.V[i];
                        cum.V[i] = all != 0 ? run / all : 0;
                    }
                    res.S.Add(cum);
                }
                else if (c.Kind == ChartKind.Table)
                {
                    Series share = new Series();
                    share.Name = "Share";
                    share.Percent = true;
                    share.V = new double[total];
                    double all = 0;
                    for (int i = 0; i < total; i++) all += s1.V[i];
                    for (int i = 0; i < total; i++) share.V[i] = all != 0 ? s1.V[i] / all : 0;
                    res.S.Add(share);
                }
                else if (c.ShowAs == 1 && CanShowAs(c.Kind))
                {
                    double all = 0;
                    for (int i = 0; i < total; i++) all += s1.V[i];
                    for (int i = 0; i < total; i++) s1.V[i] = all != 0 ? s1.V[i] / all : 0;
                    s1.Percent = true;
                }
                else if (c.ShowAs == 2 && CanShowAs(c.Kind))
                {
                    double run = 0;
                    for (int i = 0; i < total; i++)
                    {
                        run += s1.V[i];
                        s1.V[i] = run;
                    }
                }
            }
            double mxa = 0;
            if (res.S.Count > 0)
            {
                foreach (double x in res.S[0].V) if (Math.Abs(x) > mxa) mxa = Math.Abs(x);
            }
            res.MaxAbs = mxa;
            return res;
        }

        // one series for each value of the split column (the largest ones; the rest as Others)
        static void AddSplitSeries(DataTab t, ChartSpec c, DashSpec s, AggResult res, int[] k2, int total)
        {
            List<string> sl;
            int[] sk = t.Keys(c.Split, s.FyStart, out sl);
            int ns = sl.Count;
            // only the rows the chart shows decide which colours there are
            int[] shown = new int[t.Rows];
            for (int i = 0; i < t.Rows; i++) shown[i] = k2[i] < 0 ? -1 : sk[i];
            double[] srows = Reduce(t, shown, ns, "", AggKind.Count);
            double[] sval = Reduce(t, shown, ns, c.Col, c.Agg);
            List<int> sid = new List<int>();
            for (int g = 0; g < ns; g++) if (srows[g] > 0) sid.Add(g);
            bool splitTime = DatePart.Is(c.Split);
            bool scut = sid.Count > MaxSeries + 1;
            if (scut)
            {
                SortByValue(sid, sval, sl, 1);
                sid.RemoveRange(MaxSeries, sid.Count - MaxSeries);
            }
            if (splitTime) SortIds(sid, sval, sl, 0);
            else SortByValue(sid, sval, sl, 1);
            int nsk = sid.Count;
            int[] splace = new int[ns];
            for (int g = 0; g < ns; g++) splace[g] = scut ? nsk : -1;
            for (int i = 0; i < nsk; i++) splace[sid[i]] = i;
            int stotal = scut ? nsk + 1 : nsk;
            res.SplitOthers = scut;
            if (scut) res.SplitOthersName = FreeLabel(sl, "Others");
            // one pass over a combined group number
            int[] both = new int[t.Rows];
            for (int i = 0; i < t.Rows; i++)
            {
                int a = k2[i], b = sk[i] < 0 ? -1 : splace[sk[i]];
                both[i] = (a < 0 || b < 0) ? -1 : a * stotal + b;
            }
            double[] cell = Reduce(t, both, total * stotal, c.Col, c.Agg);
            for (int j = 0; j < stotal; j++)
            {
                Series se = new Series();
                se.Name = j < nsk ? sl[sid[j]] : res.SplitOthersName;
                if (j < nsk) res.SplitKeep.Add(se.Name);
                se.V = new double[total];
                for (int i = 0; i < total; i++) se.V[i] = cell[i * stotal + j];
                res.S.Add(se);
            }
            if (c.Kind == ChartKind.Stacked100)
            {
                for (int i = 0; i < total; i++)
                {
                    double all = 0;
                    for (int j = 0; j < stotal; j++) all += res.S[j].V[i];
                    for (int j = 0; j < stotal; j++) res.S[j].V[i] = all != 0 ? res.S[j].V[i] / all : 0;
                }
                foreach (Series se in res.S) se.Percent = true;
            }
        }

        // the months of one number, oldest first, for the small trend line and the comparison on a card
        public static AggResult Months(DataTab t, string col, int agg, DashSpec s)
        {
            ChartSpec c = new ChartSpec();
            c.Kind = ChartKind.Line;
            c.Dim = DatePart.Month;
            c.Col = col;
            c.Agg = agg;
            AggResult r = Group(t, c, s);
            // rows without a date are not a month
            if (r.Cats.Count > 0 && r.Cats[0] == DataTab.NoDate && r.S.Count > 0)
            {
                r.Cats.RemoveAt(0);
                double[] v = new double[r.S[0].V.Length - 1];
                for (int i = 0; i < v.Length; i++) v[i] = r.S[0].V[i + 1];
                r.S[0].V = v;
            }
            return r;
        }

        // change of the latest month: mode 1 against the month before, mode 2 against the same month last year.
        // NaN when it cannot be worked out.
        public static double Change(AggResult months, int mode)
        {
            int n = months.Cats.Count;
            if (n < 2 || months.S.Count == 0) return double.NaN;
            double[] v = months.S[0].V;
            double last = v[n - 1], prev = double.NaN;
            if (mode == 2)
            {
                string cur = months.Cats[n - 1];
                if (cur.Length < 7) return double.NaN;
                int y;
                if (!int.TryParse(cur.Substring(0, 4), NumberStyles.Integer, CultureInfo.InvariantCulture, out y)) return double.NaN;
                string want = (y - 1).ToString("0000", CultureInfo.InvariantCulture) + cur.Substring(4, 3);
                for (int i = 0; i < n - 1; i++)
                {
                    if (months.Cats[i].StartsWith(want, StringComparison.Ordinal)) prev = v[i];
                }
            }
            else prev = v[n - 2];
            if (double.IsNaN(prev) || prev == 0) return double.NaN;
            return (last - prev) / Math.Abs(prev);
        }
    }

    // =====================================================================
    //  Numbers as text: 12.35 Cr, 1.2 M, 12,34,567
    // =====================================================================
    internal sealed class CardFormat
    {
        public int Mode;              // 0 short Indian (L / Cr), 1 short international (K / M / B), 2 full number, 3 percent
        public int Dec = -1;          // -1 = automatic
        public bool Indian;           // digit grouping 12,34,567
        public string Prefix = "", Suffix = "";
    }

    internal static class NumFmt
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        public static string Zeros(int dec)
        {
            if (dec <= 0) return "0";
            return "0." + new string('0', dec > 6 ? 6 : dec);
        }

        public static string Fixed(double v, int dec)
        {
            return v.ToString(Zeros(dec), Inv);
        }

        public static bool Whole(double v)
        {
            return v == Math.Floor(v);
        }

        public static string Grouped(double v, int dec, bool indian)
        {
            if (dec < 0) dec = Whole(v) ? 0 : 2;
            if (indian) return Logic.IndianGroup(v, dec);
            return v.ToString("#,##0" + (dec > 0 ? "." + new string('0', dec > 6 ? 6 : dec) : ""), Inv);
        }

        // unit: 0 Indian (L / Cr), 1 international (K / M / B), 2 full numbers
        public static string Short(double v, int unit, int dec)
        {
            double a = Math.Abs(v);
            int d = dec < 0 ? 2 : dec;
            if (unit == 0)
            {
                if (a >= 1e7) return Fixed(v / 1e7, d) + " Cr";
                if (a >= 1e5) return Fixed(v / 1e5, d) + " L";
                return Grouped(v, dec, true);
            }
            if (unit == 1)
            {
                if (a >= 1e9) return Fixed(v / 1e9, d) + " B";
                if (a >= 1e6) return Fixed(v / 1e6, d) + " M";
                if (a >= 1e3) return Fixed(v / 1e3, dec < 0 ? 1 : dec) + " K";
                return Grouped(v, dec, false);
            }
            return Grouped(v, dec, false);
        }

        public static string Percent(double v, int dec)
        {
            return Fixed(v * 100.0, dec < 0 ? 1 : dec) + "%";
        }

        // how a card writes its number
        public static CardFormat Resolve(CardSpec c, DashSpec s)
        {
            CardFormat f = new CardFormat();
            f.Indian = s.Unit == 0;
            f.Dec = c.Dec;
            if (c.Fmt == 3) f.Mode = 3;
            else if (c.Fmt == 2 || s.Unit == 2) f.Mode = 2;
            else f.Mode = s.Unit == 1 ? 1 : 0;
            f.Prefix = c.Prefix;
            f.Suffix = c.Suffix;
            bool money = c.Col.Length > 0 && c.Agg != AggKind.Count && c.Agg != AggKind.Distinct && Logic.MeasureRank(c.Col) == 1;
            if (f.Prefix.Length == 0 && s.Currency.Length > 0 && money && f.Mode != 3) f.Prefix = s.Currency + " ";
            return f;
        }

        public static string Apply(double v, CardFormat f)
        {
            string core;
            if (f.Mode == 3) core = Percent(v, f.Dec);
            else if (f.Mode == 2) core = Grouped(v, f.Dec, f.Indian);
            else core = Short(v, f.Mode, f.Dec);
            return f.Prefix + core + f.Suffix;
        }

        // short text for chart axes, data labels and tables
        public static string Compact(double v, int unit, bool percent)
        {
            if (percent) return Fixed(v * 100.0, Math.Abs(v) < 0.1 && v != 0 ? 1 : 0) + "%";
            double a = Math.Abs(v);
            if (unit == 0)
            {
                if (a >= 1e7) return Trim0(Fixed(v / 1e7, 2)) + " Cr";
                if (a >= 1e5) return Trim0(Fixed(v / 1e5, 2)) + " L";
                return Logic.IndianGroup(v, a < 100 && !Whole(v) ? 1 : 0);
            }
            if (unit == 1)
            {
                if (a >= 1e9) return Trim0(Fixed(v / 1e9, 2)) + "B";
                if (a >= 1e6) return Trim0(Fixed(v / 1e6, 2)) + "M";
                if (a >= 1e3) return Trim0(Fixed(v / 1e3, 1)) + "K";
            }
            return v.ToString(a < 100 && !Whole(v) ? "#,##0.0" : "#,##0", Inv);
        }

        // 12.50 -> 12.5, 12.00 -> 12
        static string Trim0(string s)
        {
            if (s.IndexOf('.') < 0) return s;
            s = s.TrimEnd('0');
            if (s.EndsWith(".", StringComparison.Ordinal)) s = s.Substring(0, s.Length - 1);
            return s;
        }
    }

    // =====================================================================
    //  The words on a dashboard: automatic card labels, chart titles, key insights
    // =====================================================================
    internal static class DashText
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        // "Total Amount", "Average Rate", "Rows"
        public static string Measure(string col, int agg)
        {
            if (col == null || col.Length == 0) return "Rows";
            if (agg == AggKind.Distinct) return "Different " + col;
            if (agg == AggKind.Count) return "Count of " + col;
            return AggKind.Words[agg] + " " + col;
        }

        public static string CardLabel(CardSpec c)
        {
            return c.Label.Trim().Length > 0 ? c.Label.Trim() : Measure(c.Col, c.Agg);
        }

        public static string DimName(string dim)
        {
            return DatePart.Is(dim) ? DatePart.Label(dim).ToLowerInvariant() : dim;
        }

        // in a chart title "Total" is left out: "Amount by Customer"
        static string Short(string col, int agg)
        {
            if (col == null || col.Length == 0) return "Rows";
            if (agg == AggKind.Sum) return col;
            return Measure(col, agg);
        }

        public static string ChartTitle(ChartSpec c, AggResult r)
        {
            if (c.Title.Trim().Length > 0) return c.Title.Trim();
            string m = Short(c.Col, c.Agg), d = DimName(c.Dim);
            string top = r != null && r.TopN > 0 ? "Top " + r.TopN.ToString(Inv) + ": " : "";
            if (c.Kind == ChartKind.Pareto) return "Pareto: " + m + " by " + d;
            if (c.Kind == ChartKind.Combo) return m + " and " + Short(c.Col2, c.Agg2) + " by " + d;
            if (ChartKind.IsStacked(c.Kind) && c.Split.Length > 0) return top + m + " by " + d + " and " + DimName(c.Split);
            string t = top + m + " by " + d;
            if (c.ShowAs == 1) t += " (% of total)";
            if (c.ShowAs == 2) t += " (running total)";
            return t;
        }

        // a title for a new dashboard: the name of the sheet, or the main number when the sheet has no real name
        public static string AutoTitle(string sheetName, DashSpec s)
        {
            string nm = (sheetName ?? "").Trim();
            bool plain = nm.Length == 0 || nm.StartsWith("Sheet", StringComparison.OrdinalIgnoreCase) || nm.StartsWith("Table", StringComparison.OrdinalIgnoreCase);
            if (!plain) return nm.IndexOf("dashboard", StringComparison.OrdinalIgnoreCase) >= 0 ? nm : nm + " Dashboard";
            foreach (CardSpec c in s.Cards)
            {
                if (c.Col.Length > 0) return c.Col + " Dashboard";
            }
            return "Dashboard";
        }

        // the line under the title when the design has none of its own
        public static string AutoSubtitle(string source, int rows)
        {
            string r = rows.ToString("#,##0", Inv) + " rows";
            return source != null && source.Length > 0 ? "Source " + source + "   |   " + r : r;
        }

        // the note under the number of a card, when the card compares nothing
        public static string CardNote(CardSpec c)
        {
            if (c.Col.Length == 0) return "rows after filters";
            if (c.Agg == AggKind.Avg) return "average of all rows";
            if (c.Agg == AggKind.Max) return "largest value";
            if (c.Agg == AggKind.Min) return "smallest value";
            if (c.Agg == AggKind.Count) return "filled cells";
            if (c.Agg == AggKind.Distinct) return "different values";
            return "";
        }

        public static string CompareWords(int mode)
        {
            return mode == 2 ? "vs same month last year" : "vs previous month";
        }

        // What the key insights are about: the first number that is added up on a card ("" = the rows are counted)
        // and the first column that a chart or a filter groups by ("" = none).
        public static void InsightBasis(DataTab t, DashSpec s, out string col, out string dim)
        {
            col = "";
            dim = "";
            foreach (CardSpec c in s.Cards)
            {
                if (c.Agg == AggKind.Sum && c.Col.Length > 0 && t.Find(c.Col) != null && t.Find(c.Col).IsNum)
                {
                    col = c.Col;
                    break;
                }
            }
            foreach (ChartSpec c in s.Charts)
            {
                if (!DatePart.Is(c.Dim) && t.CanGroup(c.Dim))
                {
                    dim = c.Dim;
                    return;
                }
            }
            foreach (string f in s.Slicers)
            {
                if (!DatePart.Is(f) && t.CanGroup(f))
                {
                    dim = f;
                    return;
                }
            }
        }

        // A few plain sentences about the table, from the first number and the first group of the design.
        public static List<string> Insights(DataTab t, DashSpec s)
        {
            List<string> res = new List<string>();
            string col, dim;
            InsightBasis(t, s, out col, out dim);
            int agg = col.Length > 0 ? AggKind.Sum : AggKind.Count;
            string mName = col.Length > 0 ? col : "rows";
            if (dim.Length > 0)
            {
                ChartSpec q = new ChartSpec();
                q.Kind = ChartKind.Column;
                q.Dim = dim;
                q.Col = col;
                q.Agg = agg;
                q.Top = -1;
                q.Sort = 1;
                AggResult r = DashEngine.Group(t, q, s);
                double[] v = r.First;
                double tot = 0;
                foreach (double x in v) tot += x;
                int n = v.Length;
                if (n > 1 && tot > 0)
                {
                    res.Add(r.Cats[0] + " leads " + dim + " with " + NumFmt.Short(v[0], s.Unit, -1) + " (" + (v[0] / tot).ToString("0%", Inv) + " of " + mName + ").");
                    if (n >= 5)
                    {
                        double top3 = v[0] + v[1] + v[2];
                        res.Add("The top 3 " + dim + " values together make " + (top3 / tot).ToString("0%", Inv) + " of " + mName + ".");
                        double cum = 0;
                        int cnt = 0;
                        foreach (double x in v)
                        {
                            cum += x;
                            cnt++;
                            if (cum >= 0.8 * tot) break;
                        }
                        res.Add(cnt.ToString(Inv) + " of " + n.ToString(Inv) + " " + dim + " values give 80% of " + mName + ".");
                    }
                    res.Add("Smallest " + dim + ": " + r.Cats[n - 1] + " (" + NumFmt.Short(v[n - 1], s.Unit, -1) + ").");
                }
            }
            if (t.HasDate)
            {
                AggResult m = DashEngine.Months(t, col, agg, s);
                int n = m.Cats.Count;
                if (n > 1)
                {
                    double[] v = m.First;
                    int best = 0;
                    for (int i = 1; i < n; i++) if (v[i] > v[best]) best = i;
                    res.Add("Best month: " + m.Cats[best] + " with " + NumFmt.Short(v[best], s.Unit, -1) + ".");
                    if (v[n - 2] != 0)
                    {
                        double ch = (v[n - 1] - v[n - 2]) / Math.Abs(v[n - 2]);
                        res.Add("Latest month " + m.Cats[n - 1] + " is " + Math.Abs(ch).ToString("0%", Inv) + (ch >= 0 ? " up on " : " down on ") + m.Cats[n - 2] + ".");
                    }
                }
            }
            return res;
        }
    }
}
