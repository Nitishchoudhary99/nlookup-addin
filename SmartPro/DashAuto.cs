using System;
using System.Collections.Generic;
using System.Globalization;

namespace SmartPro
{
    // One ready-made look in the design gallery
    internal sealed class DesignIdea
    {
        public string Name = "";
        public int Theme, CardStyle, Arrange, SlicerPos;
        public bool Rounded;
    }

    // =====================================================================
    //  SUGGESTIONS
    //  Works out a sensible dashboard for a table (One-Click Dashboard, and the starting point in Dashboard Studio),
    //  repairs a design whose columns are no longer in the table, and lists the ready-made looks.
    //  No Excel calls in this file.
    // =====================================================================
    internal static class DashAuto
    {
        // how many different months / fiscal years the date column has
        public static int CountParts(DataTab t, string part, int fyStart)
        {
            List<string> labels;
            int[] key = t.Keys(part, fyStart, out labels);
            if (key == null) return 0;
            int n = 0;
            foreach (string s in labels) if (s != DataTab.NoDate) n++;
            return n;
        }

        // the number columns that look most like amounts come first
        public static List<DataCol> Measures(DataTab t, int max)
        {
            List<DataCol> meas = new List<DataCol>();
            for (int rank = 1; rank <= 3; rank++)
            {
                foreach (DataCol c in t.Cols)
                {
                    if (c.Kind == "num" && c.IsNum && meas.Count < max && Logic.MeasureRank(c.Name) == rank) meas.Add(c);
                }
            }
            return meas;
        }

        public static DashSpec Suggest(DataTab t, DashSpec look)
        {
            DashSpec s = new DashSpec();
            if (look != null)
            {
                s.CopyLook(look);
                s.Insights = look.Insights;
                s.FyStart = look.FyStart;
            }
            int months = t.HasDate ? CountParts(t, DatePart.Month, s.FyStart) : 0;
            int years = t.HasDate ? CountParts(t, DatePart.FiscalYear, s.FyStart) : 0;

            // ---- KPI cards ----
            List<DataCol> meas = Measures(t, 4);
            for (int i = 0; i < meas.Count; i++)
            {
                CardSpec c = new CardSpec();
                c.Col = meas[i].Name;
                c.Agg = Logic.MeasureIsAvg(meas[i].Name) ? AggKind.Avg : AggKind.Sum;
                if (months >= 2 && c.Agg == AggKind.Sum) c.Compare = 1;
                if (months >= 3 && i < 2) c.Spark = true;
                s.Cards.Add(c);
            }
            CardSpec rows = new CardSpec();
            rows.Agg = AggKind.Count;
            s.Cards.Add(rows);
            string main = meas.Count > 0 ? meas[0].Name : "";
            int mainAgg = meas.Count > 0 ? s.Cards[0].Agg : AggKind.Count;

            // ---- charts: the trend first, then one chart for each group-by column ----
            if (months >= 2)
            {
                ChartSpec c = new ChartSpec();
                c.Kind = ChartKind.Line;
                c.Dim = DatePart.Month;
                c.Col = main;
                c.Agg = mainAgg;
                s.Charts.Add(c);
            }
            bool round = false;
            foreach (DataCol d in t.Cols)
            {
                if (s.Charts.Count >= 4) break;
                if (d.Kind != "cat" || !d.IsCat || d.Distinct < 2) continue;
                ChartSpec c = new ChartSpec();
                if (d.Distinct <= 6 && !round)
                {
                    c.Kind = ChartKind.Doughnut;
                    round = true;
                }
                else c.Kind = d.Distinct > 8 ? ChartKind.Bar : ChartKind.Column;
                c.Dim = d.Name;
                c.Col = main;
                c.Agg = mainAgg;
                s.Charts.Add(c);
            }

            // ---- filters ----
            if (years > 1) s.Slicers.Add(DatePart.FiscalYear);
            else if (months > 1 && months <= 24) s.Slicers.Add(DatePart.Month);
            foreach (DataCol d in t.Cols)
            {
                if (s.Slicers.Count >= 4) break;
                if (d.Kind == "cat" && d.IsCat && d.Distinct >= 2 && d.Distinct <= 30) s.Slicers.Add(d.Name);
            }
            return s;
        }

        // A dashboard made by a version before 3.0 only remembered which columns it used (their positions in the
        // source range). This turns that choice into a design, so such a dashboard can be refreshed and edited.
        public static DashSpec FromPicks(DataTab t, DashSpec look, List<int> measureCols, List<int> dimCols)
        {
            DashSpec auto = Suggest(t, look);
            List<DataCol> meas = new List<DataCol>(), dims = new List<DataCol>();
            foreach (int idx in measureCols)
            {
                foreach (DataCol c in t.Cols) if (c.Index == idx && c.IsNum && !meas.Contains(c)) meas.Add(c);
            }
            foreach (int idx in dimCols)
            {
                foreach (DataCol c in t.Cols) if (c.Index == idx && c.IsCat && !dims.Contains(c)) dims.Add(c);
            }
            if (meas.Count == 0 && dims.Count == 0) return auto;
            DashSpec s = auto.Clone();
            int months = t.HasDate ? CountParts(t, DatePart.Month, s.FyStart) : 0;
            if (meas.Count > 0)
            {
                s.Cards.Clear();
                for (int i = 0; i < meas.Count && i < 4; i++)
                {
                    CardSpec c = new CardSpec();
                    c.Col = meas[i].Name;
                    c.Agg = Logic.MeasureIsAvg(meas[i].Name) ? AggKind.Avg : AggKind.Sum;
                    if (months >= 2 && c.Agg == AggKind.Sum) c.Compare = 1;
                    if (months >= 3 && i < 2) c.Spark = true;
                    s.Cards.Add(c);
                }
                CardSpec rows = new CardSpec();
                rows.Agg = AggKind.Count;
                s.Cards.Add(rows);
            }
            if (dims.Count > 0)
            {
                string main = s.Cards.Count > 0 ? s.Cards[0].Col : "";
                int mainAgg = main.Length > 0 ? s.Cards[0].Agg : AggKind.Count;
                s.Charts.Clear();
                if (months >= 2)
                {
                    ChartSpec c = new ChartSpec();
                    c.Kind = ChartKind.Line;
                    c.Dim = DatePart.Month;
                    c.Col = main;
                    c.Agg = mainAgg;
                    s.Charts.Add(c);
                }
                bool round = false;
                foreach (DataCol d in dims)
                {
                    if (s.Charts.Count >= 4) break;
                    ChartSpec c = new ChartSpec();
                    if (d.Distinct <= 6 && !round)
                    {
                        c.Kind = ChartKind.Doughnut;
                        round = true;
                    }
                    else c.Kind = d.Distinct > 8 ? ChartKind.Bar : ChartKind.Column;
                    c.Dim = d.Name;
                    c.Col = main;
                    c.Agg = mainAgg;
                    s.Charts.Add(c);
                }
                List<string> keep = new List<string>();
                foreach (string f in s.Slicers) if (DatePart.Is(f)) keep.Add(f);
                s.Slicers.Clear();
                foreach (string f in keep) s.Slicers.Add(f);
                foreach (DataCol d in dims)
                {
                    if (s.Slicers.Count < 4 && d.Distinct <= 30) s.Slicers.Add(d.Name);
                }
            }
            return s;
        }

        // true when the table has something to group by, so a dashboard can be built at all
        public static bool CanBuild(DataTab t)
        {
            if (t.HasDate) return true;
            foreach (DataCol c in t.Cols) if (c.IsCat && c.Distinct >= 2) return true;
            return false;
        }

        static bool NumberOk(DataTab t, string col, int agg)
        {
            if (col.Length == 0) return true;
            DataCol c = t.Find(col);
            if (c == null) return false;
            if (agg == AggKind.Count) return true;
            if (agg == AggKind.Distinct) return c.IsCat || c.IsNum;
            return c.IsNum;
        }

        // Takes out what the table cannot show any more (a renamed or deleted column) and says what was taken out.
        public static List<string> Fix(DashSpec s, DataTab t)
        {
            List<string> notes = new List<string>();
            t.UseDate(s.DateCol);
            if (s.DateCol.Length > 0 && !string.Equals(t.DateName, s.DateCol, StringComparison.OrdinalIgnoreCase)) s.DateCol = "";
            for (int i = s.Cards.Count - 1; i >= 0; i--)
            {
                CardSpec c = s.Cards[i];
                if (c.Col.Length == 0) c.Agg = AggKind.Count;
                if (!NumberOk(t, c.Col, c.Agg))
                {
                    notes.Add("Card '" + DashText.CardLabel(c) + "' was left out: column '" + c.Col + "' is not a " + (t.Find(c.Col) == null ? "column of this table" : "number column") + ".");
                    s.Cards.RemoveAt(i);
                    continue;
                }
                // a card that counts different values has no months to compare
                if (!t.HasDate || c.Agg == AggKind.Distinct || (c.Col.Length > 0 && c.Agg != AggKind.Count && !t.Find(c.Col).IsNum))
                {
                    c.Compare = 0;
                    c.Spark = false;
                }
            }
            for (int i = s.Charts.Count - 1; i >= 0; i--)
            {
                ChartSpec c = s.Charts[i];
                if (c.Col.Length == 0) c.Agg = AggKind.Count;
                if (!t.CanGroup(c.Dim))
                {
                    notes.Add("A chart by '" + DashText.DimName(c.Dim) + "' was left out: the table cannot be grouped by it.");
                    s.Charts.RemoveAt(i);
                    continue;
                }
                if (!NumberOk(t, c.Col, c.Agg))
                {
                    notes.Add("A chart of '" + c.Col + "' was left out: it is not a number column of this table.");
                    s.Charts.RemoveAt(i);
                    continue;
                }
                // a chart cannot count different values and still follow the filters in Excel: it counts filled cells
                if (c.Agg == AggKind.Distinct) c.Agg = AggKind.Count;
                if (c.Agg2 == AggKind.Distinct) c.Agg2 = AggKind.Count;
                if (c.Split.Length > 0 && (!t.CanGroup(c.Split) || string.Equals(c.Split, c.Dim, StringComparison.OrdinalIgnoreCase))) c.Split = "";
                if (ChartKind.IsStacked(c.Kind) && c.Split.Length == 0) c.Kind = c.Kind == ChartKind.StackedBar ? ChartKind.Bar : ChartKind.Column;
                if (c.Col2.Length > 0 && !NumberOk(t, c.Col2, c.Agg2)) c.Col2 = "";
                if (c.Col2.Length == 0 && c.Kind == ChartKind.Combo) c.Agg2 = AggKind.Count;
                if (c.Kind == ChartKind.Pareto && c.Agg != AggKind.Sum && c.Agg != AggKind.Count) c.Kind = ChartKind.Column;
            }
            for (int i = s.Slicers.Count - 1; i >= 0; i--)
            {
                if (!t.CanGroup(s.Slicers[i])) s.Slicers.RemoveAt(i);
            }
            while (s.Cards.Count > DashSpec.MaxCards) s.Cards.RemoveAt(s.Cards.Count - 1);
            while (s.Charts.Count > DashSpec.MaxCharts) s.Charts.RemoveAt(s.Charts.Count - 1);
            while (s.Slicers.Count > DashSpec.MaxSlicers) s.Slicers.RemoveAt(s.Slicers.Count - 1);
            return notes;
        }

        // The ready-made looks: every theme several times, each time with another card style and arrangement.
        public static List<DesignIdea> Ideas(bool hasDate)
        {
            List<DesignIdea> list = new List<DesignIdea>();
            int[] cardCycle = { 0, 1, 4, 2, 5, 3 };
            int[] arrDate = { 1, 0, 4, 5, 7, 3, 2, 6 };
            int[] arrPlain = { 0, 3, 4, 5, 6, 2, 7, 1 };
            int[] arr = hasDate ? arrDate : arrPlain;
            int nThemes = DashThemes.Count;
            int total = nThemes * 3;
            for (int i = 0; i < total; i++)
            {
                DesignIdea d = new DesignIdea();
                int round = i / nThemes;
                d.Theme = i % nThemes;
                d.CardStyle = cardCycle[(i + round * 2) % cardCycle.Length];
                d.Arrange = arr[(i + round * 3) % arr.Length];
                d.Rounded = (i + round) % 4 != 3;
                d.SlicerPos = i % 5 == 4 ? 1 : (i % 7 == 6 ? 2 : 0);
                d.Name = DashThemes.NameOf(d.Theme) + "  |  " + DashThemes.CardStyles[d.CardStyle] + "  |  " + ShortArrange(d.Arrange);
                list.Add(d);
            }
            return list;
        }

        static string ShortArrange(int a)
        {
            string s = DashLayout.ArrangeNames[a < 0 || a >= DashLayout.ArrangeNames.Length ? 0 : a];
            int p = s.IndexOf(" (", StringComparison.Ordinal);
            return p > 0 ? s.Substring(0, p) : s;
        }

        // puts a ready-made look on a design; the cards, charts and filters stay as they are
        public static void Apply(DashSpec s, DesignIdea d)
        {
            string font = s.Theme.Font;
            s.Theme = DashThemes.Get(d.Theme);
            s.Theme.Font = font;
            s.CardStyle = d.CardStyle;
            s.Arrange = d.Arrange;
            s.Rounded = d.Rounded;
            s.SlicerPos = d.SlicerPos;
            foreach (ChartSpec c in s.Charts)
            {
                c.Span = 0;
                c.Height = 0;
            }
        }

        // the look a user had last (kept in the registry), or the look of the older Style / Layout / Cards settings
        public static DashSpec SavedLook()
        {
            DashSpec look = DashSpec.FromText(Settings.Get("Look"));
            if (look != null) return look;
            look = new DashSpec();
            int style = Settings.GetInt("Style", 0);
            look.Theme = DashThemes.Get(style < 0 || style >= 12 ? 0 : style);
            int layout = Settings.GetInt("Layout", 0);
            int[] map = { 0, 1, 2, 6, 4, 3 };
            look.Arrange = layout >= 0 && layout < map.Length ? map[layout] : 0;
            look.CardStyle = Settings.GetInt("Cards", 0) == 1 ? 1 : 0;
            return look;
        }

        public static void SaveLook(DashSpec s)
        {
            DashSpec look = new DashSpec();
            look.CopyLook(s);
            look.Insights = s.Insights;
            look.FyStart = s.FyStart;
            Settings.Set("Look", look.ToText());
            int ti = DashThemes.IndexOf(s.Theme.Name);
            if (ti >= 0 && ti < 12) Settings.Set("Style", ti.ToString(CultureInfo.InvariantCulture));
        }
    }
}
