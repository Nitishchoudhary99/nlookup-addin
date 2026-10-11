using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace SmartPro
{
    // =====================================================================
    //  THE BUILDING PLAN OF A DASHBOARD
    //  Works out, without touching Excel, what the builder has to make: the helper table, the pivot tables behind
    //  the cards and charts, and the formulas of the cards. The builder (Dashboard.cs) then only carries it out.
    //  No Excel calls in this file.
    // =====================================================================

    // one number in a pivot table
    internal sealed class PlanField
    {
        public string Source = "";        // column of the helper table
        public string Caption = "";       // name of the number in the pivot table
        public int Function = C.xlSum;
        public int Calc;                  // 0 = as it is, or C.xlPercentOfTotal, C.xlRunningTotal, C.xlPercentRunningTotal
        public string Format = "";        // number format ("" = leave)
    }

    internal sealed class PlanPivot
    {
        public string Name = "", Purpose = "";
        public string RowField = "", ColField = "";
        public bool DatedOnly;            // only the rows that have a date
        public List<PlanField> Data = new List<PlanField>();
        public int Sort;                  // rows: 0 names A to Z, 1 names Z to A, 2 largest number first, 3 smallest number first
        public List<string> Order;        // the rows in the order they must have, set by hand (null = sorted as Sort says)
        public List<string> ColOrder;     // the same for the columns of a stacked chart
        public string PlainRow = "";      // the column to use instead when the order cannot be set by hand
        public string PlainCol = "";
        public int TopN;                  // only the N largest rows (0 = all)
        public int Columns = 3;           // room on the helper sheet

        // filled in by the builder once the pivot table exists
        public object Table;
        public int FirstRow, LabelCol, ValueCol;      // the first row of numbers, the column of the names, the first column of numbers
        public bool ByHand;               // the order was set by hand and it held
    }

    internal sealed class PlanCard
    {
        public string Label = "";
        public PlanField Field;           // its number in the KPI pivot table (null when the card counts different values)
        public PlanField Field2;          // the second number of a card that calculates
        public int DistinctPivot = -1;    // the pivot table that lists the different values
        public int MonthPivot = -1;       // the pivot table with the months, for the comparison and the small trend line
        public CardFormat Format;
    }

    internal sealed class PlanChart
    {
        public int Pivot = -1;
        public AggResult Result;          // what the chart shows when no filter is set
        public string Title = "";
        public List<string> Names = new List<string>();        // the groups as the chart names them, first to last
        public List<string> SeriesNames = new List<string>();  // the colours of a stacked chart
        public string OthersName = "", SplitOthersName = "";   // "" = there is no Others group
        public int DisplayUnit, DisplayUnit2;                  // 0, C.xlHundredThousands or C.xlTenMillions
        public string AxisFormat = "", LabelFormat = "", Axis2Format = "";
        public double UnitDivisor = 1;    // table charts: the numbers are shown divided by this
        public string UnitFormat = "", Unit2Format = "";
        public double Unit2Divisor = 1;
    }

    internal sealed class PlanSlicer
    {
        public string Field = "", Caption = "";
        public int Items;
    }

    internal sealed class DashPlan
    {
        public object[,] Data;                                // the helper table, row 1 = headers (1-based like every Excel block)
        public List<string> Headers = new List<string>();
        public List<bool> IsText = new List<bool>();
        public int Rows;
        public List<PlanPivot> Pivots = new List<PlanPivot>();
        public int KpiPivot = -1;
        public PlanField RowsField;                           // the number of rows in the KPI pivot table
        public List<PlanCard> Cards = new List<PlanCard>();
        public List<PlanChart> Charts = new List<PlanChart>();
        public List<PlanSlicer> Slicers = new List<PlanSlicer>();
        public string DatedField = "";                        // the column that says Yes for rows with a date ("" = every row has one)
        public int Undated;

        public const string Yes = "Yes", No = "No";
        public const int MonthRows = 600;                     // rows read under a pivot table of months

        readonly DataTab t;
        readonly DashSpec s;
        readonly List<string[]> textVals = new List<string[]>();
        readonly List<double[]> numVals = new List<double[]>();
        readonly Dictionary<string, int> colOf = new Dictionary<string, int>(StringComparer.Ordinal);
        readonly HashSet<string> names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        readonly Dictionary<string, int> monthPivotOf = new Dictionary<string, int>(StringComparer.Ordinal);
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        DashPlan(DataTab table, DashSpec spec)
        {
            t = table;
            s = spec;
        }

        // The design must have been repaired for this table first (DashAuto.Fix).
        public static DashPlan Build(DataTab t, DashSpec s)
        {
            DashPlan p = new DashPlan(t, s);
            p.Rows = t.Rows;
            p.Undated = DashEngine.Undated(t);
            p.PlanCards();
            p.PlanCharts();
            p.PlanSlicers();
            p.FixCaptions();
            p.FillData();
            return p;
        }

        // ------------------------------------------------------------------ columns of the helper table
        string Unique(string want)
        {
            want = (want ?? "").Trim();
            if (want.Length == 0) want = "Column";
            if (want.Length > 200) want = want.Substring(0, 200);
            string nm = want;
            int i = 1;
            while (names.Contains(nm))
            {
                i++;
                nm = want + " (" + i.ToString(Inv) + ")";
            }
            names.Add(nm);
            return nm;
        }

        int AddText(string key, string header, string[] vals)
        {
            Headers.Add(header);
            IsText.Add(true);
            textVals.Add(vals);
            numVals.Add(null);
            colOf[key] = Headers.Count - 1;
            return Headers.Count - 1;
        }

        int AddNum(string key, string header, double[] vals)
        {
            Headers.Add(header);
            IsText.Add(false);
            textVals.Add(null);
            numVals.Add(vals);
            colOf[key] = Headers.Count - 1;
            return Headers.Count - 1;
        }

        // the names of the groups of a column or a date part, one for each row
        int DimColumn(string dim)
        {
            bool part = DatePart.Is(dim);
            DataCol c = part ? null : t.Find(dim);
            string key = "dim:" + (part ? dim : c.Name.ToLowerInvariant());
            int at;
            if (colOf.TryGetValue(key, out at)) return at;
            List<string> labels;
            int[] k = t.Keys(dim, s.FyStart, out labels);
            string[] vals = new string[t.Rows];
            for (int i = 0; i < t.Rows; i++) vals[i] = labels[k[i]];
            return AddText(key, Unique(part ? DatePart.Label(dim) : c.Name), vals);
        }

        // a column as numbers
        int NumColumn(DataCol c)
        {
            string key = "num:" + c.Name.ToLowerInvariant();
            int at;
            if (colOf.TryGetValue(key, out at)) return at;
            double[] vals = new double[t.Rows];
            for (int i = 0; i < t.Rows; i++) vals[i] = c.IsNum ? c.Num[i] : double.NaN;
            string header = names.Contains(c.Name) ? Unique(c.Name + " (number)") : Unique(c.Name);
            return AddNum(key, header, vals);
        }

        // 1 where a column is filled: adding these up counts the filled cells
        int FlagColumn(DataCol c)
        {
            string key = "flag:" + c.Name.ToLowerInvariant();
            int at;
            if (colOf.TryGetValue(key, out at)) return at;
            double[] vals = new double[t.Rows];
            for (int i = 0; i < t.Rows; i++)
            {
                bool filled;
                if (c.IsNum) filled = !double.IsNaN(c.Num[i]);
                else if (c.IsCat) filled = c.Items[c.Code[i]] != DataTab.Blank;
                else if (c.IsDate) filled = c.Ymd[i] > 0;
                else filled = c.Filled != null && c.Filled[i];
                vals[i] = filled ? 1.0 : double.NaN;
            }
            return AddNum(key, Unique(c.Name + " (filled)"), vals);
        }

        int RecordsColumn()
        {
            int at;
            if (colOf.TryGetValue("records", out at)) return at;
            double[] vals = new double[t.Rows];
            for (int i = 0; i < t.Rows; i++) vals[i] = 1.0;
            return AddNum("records", Unique("Records"), vals);
        }

        void NeedDated()
        {
            if (DatedField.Length > 0 || Undated == 0 || !t.HasDate) return;
            int[] ymd = t.Cols[t.DateIndex].Ymd;
            string[] vals = new string[t.Rows];
            for (int i = 0; i < t.Rows; i++) vals[i] = ymd[i] > 0 ? Yes : No;
            DatedField = Headers[AddText("dated", Unique("Has Date"), vals)];
        }

        // The names a chart shows: the groups it keeps (date groups under their short names) and the rest as Others.
        int LabelColumn(string dim, List<string> keep, bool others, string othersName, bool shortNames)
        {
            StringBuilder sig = new StringBuilder("lab:");
            sig.Append(DatePart.Is(dim) ? dim : t.Find(dim).Name.ToLowerInvariant()).Append(shortNames ? "|s|" : "|-|");
            if (others)
            {
                sig.Append(othersName);
                foreach (string k in keep) sig.Append('\u001E').Append(k);
            }
            string key = sig.ToString();
            int at;
            if (colOf.TryGetValue(key, out at)) return at;
            List<string> labels;
            int[] gk = t.Keys(dim, s.FyStart, out labels);
            HashSet<string> kept = new HashSet<string>(keep, StringComparer.Ordinal);
            string[] shown = new string[labels.Count];
            for (int g = 0; g < labels.Count; g++)
            {
                if (others && !kept.Contains(labels[g])) shown[g] = othersName;
                else shown[g] = shortNames ? DataTab.ShortLabel(dim, labels[g]) : labels[g];
            }
            string[] vals = new string[t.Rows];
            for (int i = 0; i < t.Rows; i++) vals[i] = shown[gk[i]];
            string baseName = DatePart.Is(dim) ? DatePart.Label(dim) : t.Find(dim).Name;
            string header = others ? baseName + " (top " + keep.Count.ToString(Inv) + ")" : baseName + " (short)";
            return AddText(key, Unique(header), vals);
        }

        // true when the short names of these date groups are all different
        bool ShortNamesDiffer(string part, List<string> labels)
        {
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string l in labels)
            {
                if (!seen.Add(DataTab.ShortLabel(part, l))) return false;
            }
            return true;
        }

        // every group of a column or date part that has rows, by name
        List<string> AllGroups(string dim, bool datedOnly)
        {
            List<string> labels;
            int[] k = t.Keys(dim, s.FyStart, out labels);
            bool[] has = new bool[labels.Count];
            int[] ymd = datedOnly && t.HasDate ? t.Cols[t.DateIndex].Ymd : null;
            for (int i = 0; i < t.Rows; i++)
            {
                if (ymd != null && ymd[i] <= 0) continue;
                has[k[i]] = true;
            }
            List<string> res = new List<string>();
            for (int g = 0; g < labels.Count; g++) if (has[g]) res.Add(labels[g]);
            res.Sort(StringComparer.Ordinal);
            return res;
        }

        // ------------------------------------------------------------------ numbers
        // the pivot field for a number of a card or chart
        PlanField Field(string col, int agg, string caption)
        {
            PlanField f = new PlanField();
            f.Caption = caption;
            DataCol c = t.Find(col);
            if (c == null)
            {
                f.Source = Headers[RecordsColumn()];
                return f;
            }
            if (agg == AggKind.Count || agg == AggKind.Distinct || !c.IsNum)
            {
                // different values cannot be counted by a pivot table that also follows the filters, so a chart counts filled cells
                f.Source = Headers[FlagColumn(c)];
                return f;
            }
            f.Source = Headers[NumColumn(c)];
            if (agg == AggKind.Avg) f.Function = C.xlAverage;
            else if (agg == AggKind.Max) f.Function = C.xlMax;
            else if (agg == AggKind.Min) f.Function = C.xlMin;
            return f;
        }

        // ------------------------------------------------------------------ cards
        void PlanCards()
        {
            PlanPivot kpi = new PlanPivot();
            kpi.Name = "DB_KPI";
            kpi.Purpose = "The numbers of the KPI cards";
            Pivots.Add(kpi);
            KpiPivot = Pivots.Count - 1;
            for (int i = 0; i < s.Cards.Count; i++)
            {
                CardSpec c = s.Cards[i];
                PlanCard pc = new PlanCard();
                pc.Label = DashText.CardLabel(c);
                pc.Format = NumFmt.Resolve(c, s);
                DataCol col = t.Find(c.Col);
                if (c.Agg == AggKind.Distinct && col != null)
                {
                    PlanPivot d = new PlanPivot();
                    d.Name = "DB_Count" + (i + 1).ToString(Inv);
                    d.Purpose = "Card " + (i + 1).ToString(Inv) + ": the different values of " + col.Name;
                    d.RowField = Headers[col.IsCat ? DimColumn(col.Name) : NumColumn(col)];
                    d.Data.Add(Field("", AggKind.Count, "Rows"));
                    d.Columns = 3;
                    Pivots.Add(d);
                    pc.DistinctPivot = Pivots.Count - 1;
                }
                else if (c.Op > 0)
                {
                    pc.Field = Field(c.Col, c.Agg, "K" + (i + 1).ToString(Inv));
                    pc.Field2 = Field(c.Col2, c.Agg2, "K" + (i + 1).ToString(Inv) + "b");
                    kpi.Data.Add(pc.Field);
                    kpi.Data.Add(pc.Field2);
                }
                else
                {
                    pc.Field = Field(c.Col, c.Agg, "K" + (i + 1).ToString(Inv));
                    kpi.Data.Add(pc.Field);
                    if ((c.Compare > 0 || c.Spark) && t.HasDate) pc.MonthPivot = MonthPivot(c.Col, c.Agg);
                }
                Cards.Add(pc);
            }
            RowsField = Field("", AggKind.Count, "KR");
            kpi.Data.Add(RowsField);
            kpi.Columns = kpi.Data.Count + 2;
        }

        // one pivot table of months for each number that a card compares or draws
        int MonthPivot(string col, int agg)
        {
            DataCol c = t.Find(col);
            string key = (c == null ? "" : c.Name.ToLowerInvariant()) + "|" + agg.ToString(Inv);
            int at;
            if (monthPivotOf.TryGetValue(key, out at)) return at;
            PlanPivot m = new PlanPivot();
            m.Name = "DB_Month" + (monthPivotOf.Count + 1).ToString(Inv);
            m.Purpose = "Months of " + DashText.Measure(col, agg) + " (comparison and trend line of a card)";
            m.RowField = Headers[DimColumn(DatePart.Month)];
            if (Undated > 0)
            {
                NeedDated();
                m.DatedOnly = true;
            }
            m.Data.Add(Field(col, agg, DashText.Measure(col, agg)));
            m.Sort = 0;
            m.Columns = 4;
            Pivots.Add(m);
            monthPivotOf[key] = Pivots.Count - 1;
            return Pivots.Count - 1;
        }

        // ------------------------------------------------------------------ charts
        void PlanCharts()
        {
            for (int i = 0; i < s.Charts.Count; i++)
            {
                ChartSpec c = s.Charts[i];
                AggResult r = DashEngine.Group(t, c, s);
                PlanChart pc = new PlanChart();
                pc.Result = r;
                pc.Title = DashText.ChartTitle(c, r);
                PlanPivot p = new PlanPivot();
                p.Name = "DB_Chart" + (i + 1).ToString(Inv);
                p.Purpose = "Chart " + (i + 1).ToString(Inv) + ": " + pc.Title;
                if (Undated > 0 && DashEngine.UsesDate(t, c))
                {
                    NeedDated();
                    p.DatedOnly = true;
                }
                bool bar = c.Kind == ChartKind.Bar || c.Kind == ChartKind.StackedBar;
                bool byName = r.SortMode == 0 || r.SortMode == 3;

                // ---- the groups (rows of the pivot table) ----
                bool shortNames = DatePart.Is(c.Dim) && DataTab.HasShortLabels(c.Dim) && ShortNamesDiffer(c.Dim, AllGroups(c.Dim, p.DatedOnly));
                foreach (string cat in r.Cats) pc.Names.Add(shortNames ? DataTab.ShortLabel(c.Dim, cat) : cat);
                if (r.Others) pc.OthersName = r.OthersName;
                if (r.Others || shortNames) p.RowField = Headers[LabelColumn(c.Dim, r.Keep, r.Others, r.OthersName, shortNames)];
                else p.RowField = Headers[DimColumn(c.Dim)];
                if (r.Others)
                {
                    // Others stays last, so the order is fixed by hand
                    p.Order = new List<string>(pc.Names);
                }
                else if (shortNames && byName)
                {
                    // short names do not sort in time order
                    p.Order = new List<string>();
                    foreach (string g in AllGroups(c.Dim, p.DatedOnly)) p.Order.Add(DataTab.ShortLabel(c.Dim, g));
                    p.PlainRow = Headers[DimColumn(c.Dim)];
                }
                if (p.Order != null && bar) p.Order.Reverse();           // a bar chart draws its first row at the bottom
                if (r.SortMode == 1) p.Sort = bar ? 3 : 2;
                else if (r.SortMode == 2) p.Sort = bar ? 2 : 3;
                else p.Sort = bar ? 1 : 0;
                if (r.TopN > 0 && !r.Others) p.TopN = r.TopN;

                // ---- the colours of a stacked chart (columns of the pivot table) ----
                bool split = ChartKind.IsStacked(c.Kind) && t.CanGroup(c.Split) && c.Split != c.Dim && r.S.Count > 0;
                if (split)
                {
                    bool sShort = DatePart.Is(c.Split) && DataTab.HasShortLabels(c.Split) && ShortNamesDiffer(c.Split, AllGroups(c.Split, p.DatedOnly));
                    for (int j = 0; j < r.S.Count; j++)
                    {
                        bool isOthers = r.SplitOthers && j == r.S.Count - 1;
                        pc.SeriesNames.Add(sShort && !isOthers ? DataTab.ShortLabel(c.Split, r.S[j].Name) : r.S[j].Name);
                    }
                    if (r.SplitOthers) pc.SplitOthersName = r.SplitOthersName;
                    if (r.SplitOthers || sShort) p.ColField = Headers[LabelColumn(c.Split, r.SplitKeep, r.SplitOthers, r.SplitOthersName, sShort)];
                    else p.ColField = Headers[DimColumn(c.Split)];
                    p.ColOrder = new List<string>(pc.SeriesNames);
                    if (sShort && !r.SplitOthers) p.PlainCol = Headers[DimColumn(c.Split)];
                }
                else if (r.S.Count > 0)
                {
                    foreach (Series se in r.S) pc.SeriesNames.Add(se.Name);
                }

                // ---- the numbers ----
                PlanField f1 = Field(c.Col, c.Agg, DashText.Measure(c.Col, c.Agg));
                if (!split && DashEngine.CanShowAs(c.Kind))
                {
                    if (c.ShowAs == 1) f1.Calc = C.xlPercentOfTotal;
                    else if (c.ShowAs == 2) f1.Calc = C.xlRunningTotal;
                }
                p.Data.Add(f1);
                if (!split && (c.Kind == ChartKind.Combo || (c.Kind == ChartKind.Table && c.Col2.Length > 0)))
                    p.Data.Add(Field(c.Col2, c.Agg2, DashText.Measure(c.Col2, c.Agg2)));
                if (!split && c.Kind == ChartKind.Pareto)
                {
                    PlanField cum = Field(c.Col, c.Agg, "Cumulative share");
                    cum.Calc = C.xlPercentRunningTotal;
                    cum.Format = "0%";
                    p.Data.Add(cum);
                }
                p.Columns = 1 + (split ? r.S.Count : p.Data.Count) + 2;

                // ---- units on the axes ----
                if (r.S.Count > 0)
                {
                    double top = 0;
                    if (split)
                    {
                        for (int k = 0; k < r.Cats.Count; k++)
                        {
                            double sum = 0;
                            foreach (Series se in r.S) if (se.V[k] > 0) sum += se.V[k];
                            if (sum > top) top = sum;
                        }
                    }
                    else top = MaxAbs(r.S[0].V);
                    string lf;
                    DashFormula.Units(s.Unit, top, r.S[0].Percent, out pc.DisplayUnit, out pc.AxisFormat, out pc.LabelFormat);
                    if (!split && r.S.Count > 1 && (c.Kind == ChartKind.Combo || c.Kind == ChartKind.Pareto))
                        DashFormula.Units(s.Unit, MaxAbs(r.S[1].V), r.S[1].Percent, out pc.DisplayUnit2, out pc.Axis2Format, out lf);
                    if (c.Kind == ChartKind.Table)
                    {
                        DashFormula.TableUnit(s.Unit, MaxAbs(r.S[0].V), out pc.UnitDivisor, out pc.UnitFormat);
                        if (r.S.Count > 2) DashFormula.TableUnit(s.Unit, MaxAbs(r.S[1].V), out pc.Unit2Divisor, out pc.Unit2Format);
                    }
                }
                Pivots.Add(p);
                pc.Pivot = Pivots.Count - 1;
                Charts.Add(pc);
            }
        }

        static double MaxAbs(double[] v)
        {
            double m = 0;
            foreach (double x in v) if (Math.Abs(x) > m) m = Math.Abs(x);
            return m;
        }

        // ------------------------------------------------------------------ filters
        void PlanSlicers()
        {
            if (s.SlicerPos == 3) return;
            foreach (string f in s.Slicers)
            {
                if (!t.CanGroup(f)) continue;
                PlanSlicer sl = new PlanSlicer();
                sl.Field = Headers[DimColumn(f)];
                sl.Caption = DatePart.Is(f) ? DatePart.Label(f) : t.Find(f).Name;
                List<string> labels;
                t.Keys(f, s.FyStart, out labels);
                sl.Items = labels.Count;
                Slicers.Add(sl);
            }
        }

        // A number in a pivot table may not have the name of a column, nor the name of another number of the same table.
        void FixCaptions()
        {
            foreach (PlanPivot p in Pivots)
            {
                HashSet<string> taken = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (PlanField f in p.Data)
                {
                    if (f.Caption.Trim().Length == 0) f.Caption = "Value";
                    while (names.Contains(f.Caption) || taken.Contains(f.Caption)) f.Caption += " ";
                    taken.Add(f.Caption);
                }
            }
        }

        // Excel takes an apostrophe at the start of a text as a mark ("this is text") and drops it.
        // A text that really starts with an apostrophe is therefore written with two.
        public static string CellText(string s)
        {
            return s != null && s.Length > 0 && s[0] == '\'' ? "'" + s : s;
        }

        void FillData()
        {
            int nc = Headers.Count;
            Data = Xl.New2D(Rows + 1, nc);
            for (int c = 0; c < nc; c++)
            {
                Data[1, c + 1] = CellText(Headers[c]);
                if (IsText[c])
                {
                    string[] v = textVals[c];
                    for (int i = 0; i < Rows; i++) Data[i + 2, c + 1] = CellText(v[i]);
                }
                else
                {
                    double[] v = numVals[c];
                    for (int i = 0; i < Rows; i++) if (!double.IsNaN(v[i])) Data[i + 2, c + 1] = v[i];
                }
            }
        }

        // how many rows of names to read under the pivot table of a card that counts different values
        public int DistinctRows(int pivot)
        {
            string field = Pivots[pivot].RowField;
            for (int c = 0; c < Headers.Count; c++)
            {
                if (Headers[c] != field) continue;
                if (IsText[c])
                {
                    HashSet<string> seen = new HashSet<string>(textVals[c], StringComparer.OrdinalIgnoreCase);
                    return seen.Count + 20;
                }
                HashSet<double> nums = new HashSet<double>(numVals[c]);
                return nums.Count + 20;
            }
            return DataTab.MaxItems + 20;
        }
    }

    // The design of a dashboard is kept in hidden cells of the dashboard sheet, so Refresh and Edit can read it back.
    // It is stored as base64 in pieces: a cell takes any such text, whatever characters the design holds.
    internal static class DashMeta
    {
        public const int Piece = 8000;
        public const string Mark = "~";

        public static List<string> Pack(string text)
        {
            string b64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(text ?? ""));
            List<string> res = new List<string>();
            for (int i = 0; i < b64.Length; i += Piece) res.Add(Mark + b64.Substring(i, Math.Min(Piece, b64.Length - i)));
            if (res.Count == 0) res.Add(Mark);
            return res;
        }

        // null when the pieces are not what Pack wrote
        public static string Unpack(List<string> pieces)
        {
            StringBuilder sb = new StringBuilder();
            foreach (string p in pieces)
            {
                if (p == null || !p.StartsWith(Mark, StringComparison.Ordinal)) return null;
                sb.Append(p, Mark.Length, p.Length - Mark.Length);
            }
            try
            {
                return Encoding.UTF8.GetString(Convert.FromBase64String(sb.ToString()));
            }
            catch (FormatException)
            {
                return null;
            }
        }
    }

    // =====================================================================
    //  FORMULAS AND NUMBER FORMATS the builder writes. All of them are plain text here, so they can be checked
    //  without Excel. Cell addresses are passed in as text ($B$3, or 'Sheet'!$B$3).
    // =====================================================================
    internal static class DashFormula
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        public const string Up = "\u25B2 ", Down = "\u25BC ";        // the arrows in front of a rise and a fall

        // "text" as it is written inside a formula
        public static string Q(string s)
        {
            return "\"" + (s ?? "").Replace("\"", "\"\"") + "\"";
        }

        // 'Sheet name'!
        public static string Sheet(string name)
        {
            return "'" + name.Replace("'", "''") + "'!";
        }

        public static string Num(double d)
        {
            return d.ToString("R", Inv);
        }

        // 12,34,567.00 (Indian) or 1,234,567.00, as an Excel number format.
        // A number format takes two conditions at most, so the Indian one is right up to 99,99,99,999;
        // IndianFormatLarge gives the formats for larger numbers.
        public static string GroupFormat(int dec, bool indian)
        {
            string z = dec > 0 ? "." + new string('0', dec > 6 ? 6 : dec) : "";
            if (!indian) return "#,##0" + z;
            return "[>=10000000]##\\,##\\,##\\,##0" + z + ";[>=100000]##\\,##\\,##0" + z + ";##,##0" + z;
        }

        // Indian digit grouping for numbers of 10 or 11 digits (groups = 4) and of 12 or 13 digits (groups = 5)
        public static string IndianFormatLarge(int dec, int groups)
        {
            string z = dec > 0 ? "." + new string('0', dec > 6 ? 6 : dec) : "";
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < groups; i++) sb.Append("##\\,");
            return sb.Append("##0").Append(z).ToString();
        }

        // a number with digit grouping, as text
        static string G(string v, int dec, bool indian)
        {
            if (!indian) return "TEXT(" + v + "," + Q(GroupFormat(dec, false)) + ")";
            string a = "ABS(" + v + ")";
            return "IF(" + v + "<0,\"-\",\"\")&TEXT(" + a + ",IF(" + a + ">=1E11," + Q(IndianFormatLarge(dec, 5)) + ",IF(" + a + ">=1E9," + Q(IndianFormatLarge(dec, 4)) + "," +
                   Q(GroupFormat(dec, true)) + ")))";
        }

        static string Grouped(string v, int dec, bool indian)
        {
            if (dec >= 0) return G(v, dec, indian);
            return "IF(" + v + "=INT(" + v + ")," + G(v, 0, indian) + "," + G(v, 2, indian) + ")";
        }

        static string Fixed(string expr, int dec)
        {
            return "TEXT(" + expr + "," + Q(NumFmt.Zeros(dec)) + ")";
        }

        // the number of a card: =IFERROR(GETPIVOTDATA("K1",$L$6),0)
        public static string KpiValue(string caption, string pivotCell)
        {
            return "=IFERROR(GETPIVOTDATA(" + Q(caption) + "," + pivotCell + "),0)";
        }

        // a card that calculates: both numbers come from the KPI pivot table, so the filters change it too
        public static string KpiCalc(int op, string caption1, string caption2, string pivotCell)
        {
            string a = "IFERROR(GETPIVOTDATA(" + Q(caption1) + "," + pivotCell + "),0)";
            string b = "IFERROR(GETPIVOTDATA(" + Q(caption2) + "," + pivotCell + "),0)";
            switch (op)
            {
                case CardOp.Percent:
                case CardOp.Divide: return "=IFERROR(" + a + "/" + b + ",0)";
                case CardOp.Minus: return "=" + a + "-" + b;
                default: return "=" + a + "+" + b;
            }
        }

        // how many different values a pivot table lists (blank cells are not a value)
        public static string DistinctCount(string labelRange)
        {
            return "=COUNTA(" + labelRange + ")-COUNTIF(" + labelRange + "," + Q(DataTab.Blank) + ")";
        }

        // The number of a card as text, the same way NumFmt.Apply writes it: 12.35 Cr, 1.2 M, 12,34,567 or 12.3%
        public static string CardText(string v, CardFormat f)
        {
            string core;
            int d = f.Dec < 0 ? 2 : f.Dec;
            if (f.Mode == 3) core = "TEXT(" + v + "," + Q(NumFmt.Zeros(f.Dec < 0 ? 1 : f.Dec) + "%") + ")";
            else if (f.Mode == 2) core = Grouped(v, f.Dec, f.Indian);
            else if (f.Mode == 1)
            {
                core = "IF(ABS(" + v + ")>=1E9," + Fixed(v + "/1E9", d) + "&\" B\",IF(ABS(" + v + ")>=1E6," + Fixed(v + "/1E6", d) + "&\" M\",IF(ABS(" + v +
                       ")>=1E3," + Fixed(v + "/1E3", f.Dec < 0 ? 1 : f.Dec) + "&\" K\"," + Grouped(v, f.Dec, false) + ")))";
            }
            else
            {
                core = "IF(ABS(" + v + ")>=1E7," + Fixed(v + "/1E7", d) + "&\" Cr\",IF(ABS(" + v + ")>=1E5," + Fixed(v + "/1E5", d) + "&\" L\"," +
                       Grouped(v, f.Dec, true) + "))";
            }
            string res = "=";
            if (f.Prefix.Length > 0) res += Q(f.Prefix) + "&";
            res += core;
            if (f.Suffix.Length > 0) res += "&" + Q(f.Suffix);
            return res;
        }

        // ---- the comparison of a card: the latest month against an earlier one ----
        public static string MonthCount(string labels)
        {
            return "=COUNTA(" + labels + ")";
        }

        // the number to compare with: mode 1 the month before, mode 2 the same month one year back ("" when there is none)
        public static string Earlier(int mode, string labels, string values, string n)
        {
            if (mode == 2)
            {
                string last = "INDEX(" + labels + "," + n + ")";
                return "=IF(" + n + "<2,\"\",IFERROR(INDEX(" + values + ",MATCH((LEFT(" + last + ",4)-1)&MID(" + last + ",5,20)," + labels + ",0)),\"\"))";
            }
            return "=IF(" + n + "<2,\"\",INDEX(" + values + "," + n + "-1))";
        }

        // the change as a share ("" when it cannot be worked out)
        public static string Change(string values, string n, string earlier)
        {
            return "=IF(OR(" + earlier + "=\"\"," + earlier + "=0),\"\",(INDEX(" + values + "," + n + ")-" + earlier + ")/ABS(" + earlier + "))";
        }

        // sign > 0: the text for a rise only, sign < 0: for a fall only, sign = 0: for both.
        // When there is nothing to say the result is one space, not an empty text: a text box that is linked to
        // the cell then always holds a character and keeps its font.
        public static string ChangeNote(string change, int mode, int sign)
        {
            string words = Q(" " + DashText.CompareWords(mode));
            string pct = Q("0.0%");
            string up = Q(Up), down = Q(Down), none = Q(" ");
            if (sign > 0) return "=IF(ISNUMBER(" + change + "),IF(" + change + ">=0," + up + "&TEXT(" + change + "," + pct + ")&" + words + "," + none + ")," + none + ")";
            if (sign < 0) return "=IF(ISNUMBER(" + change + "),IF(" + change + "<0," + down + "&TEXT(-" + change + "," + pct + ")&" + words + "," + none + ")," + none + ")";
            return "=IF(ISNUMBER(" + change + "),IF(" + change + ">=0," + up + "," + down + ")&TEXT(ABS(" + change + ")," + pct + ")&" + words + "," + none + ")";
        }

        // a number format that writes the change with an arrow; a light card shows a rise in green and a fall in red
        public static string ChangeFormat(int mode, bool coloured, bool dark)
        {
            string words = "\" " + DashText.CompareWords(mode) + "\"";
            string up = (coloured ? (dark ? "[Color4]" : "[Color10]") : "") + "\"" + Up + "\"0.0%" + words;
            string down = (coloured ? "[Color3]" : "") + "\"" + Down + "\"0.0%" + words;
            return up + ";" + down + ";" + up + ";@";
        }

        // "85% of target 5.00 Cr"
        public static string TargetNote(string v, double target, int unit)
        {
            return "=TEXT(" + v + "/" + Num(target) + ",\"0%\")&" + Q(" of target " + NumFmt.Short(target, unit, -1));
        }

        public static string PerRow(string v, string rows)
        {
            return "=IFERROR(" + v + "/" + rows + ",0)";
        }

        public static string PerRowNote(string perRow, bool indian)
        {
            return "=\"avg per row: \"&" + G(perRow, 2, indian);
        }

        // ---- charts ----
        // The unit of a value axis. Lakh and crore need Excel's display unit, K and M are done by the number format.
        public static void Units(int unit, double max, bool percent, out int displayUnit, out string axisFormat, out string labelFormat)
        {
            displayUnit = 0;
            if (percent)
            {
                axisFormat = "0%";
                labelFormat = "0%";
                return;
            }
            if (max < 100)
            {
                axisFormat = "General";
                labelFormat = "#,##0.0";
                return;
            }
            axisFormat = "#,##0";
            labelFormat = "#,##0";
            if (unit == 0)
            {
                if (max >= 1e7)
                {
                    displayUnit = C.xlTenMillions;
                    axisFormat = "General\" Cr\"";
                    labelFormat = "0.00\" Cr\"";
                }
                else if (max >= 1e5)
                {
                    displayUnit = C.xlHundredThousands;
                    axisFormat = "General\" L\"";
                    labelFormat = "0.00\" L\"";
                }
            }
            else if (unit == 1 && max >= 1000)
            {
                axisFormat = "[>=1000000]0.0,,\"M\";[>=1000]0,\"K\";0";
                labelFormat = "[>=1000000]0.00,,\"M\";[>=1000]0.0,\"K\";0";
            }
        }

        // The unit of a column of a table chart: the numbers are divided once and the format names the unit.
        public static void TableUnit(int unit, double max, out double divisor, out string format)
        {
            divisor = 1;
            format = max < 100 ? "#,##0.0" : "#,##0";
            if (unit == 0)
            {
                if (max >= 1e7)
                {
                    divisor = 1e7;
                    format = "0.00\" Cr\"";
                }
                else if (max >= 1e5)
                {
                    divisor = 1e5;
                    format = "0.00\" L\"";
                }
            }
            else if (unit == 1)
            {
                if (max >= 1e9)
                {
                    divisor = 1e9;
                    format = "0.00\" B\"";
                }
                else if (max >= 1e6)
                {
                    divisor = 1e6;
                    format = "0.00\" M\"";
                }
                else if (max >= 1e3)
                {
                    divisor = 1e3;
                    format = "0.0\" K\"";
                }
            }
        }

        // a table chart reads row k of the pivot table: the name, a number, and the share of the first number
        public static string TableName(string labels, int k)
        {
            return "=IFERROR(INDEX(" + labels + "," + k.ToString(Inv) + ")&\"\",\"\")";
        }

        public static string TableValue(string nameCell, string values, int k, double divisor)
        {
            string v = "INDEX(" + values + "," + k.ToString(Inv) + ")";
            if (divisor != 1) v += "/" + Num(divisor);
            return "=IF(" + nameCell + "=\"\",\"\"," + v + ")";
        }

        public static string TableShare(string nameCell, string values, int k)
        {
            return "=IF(" + nameCell + "=\"\",\"\",IFERROR(INDEX(" + values + "," + k.ToString(Inv) + ")/SUM(" + values + "),0))";
        }
    }
}
