using System;
using System.Collections.Generic;

namespace SmartPro
{
    // The colours of one KPI card, worked out from the card style and the card's colour
    internal sealed class CardLook
    {
        public int Bg, Bg2;               // background; Bg2 is the second colour of a gradient
        public int Label, Value, Note;    // text colours
        public int Line;                  // colour of the top line, side bar or outline
        public int LineKind;              // 0 none, 1 line on top, 2 bar on the left, 3 outline
        public bool Gradient;
        public bool Plain;                // light card: rise and fall can be shown in green and red
    }

    // the numbers on a value axis
    internal sealed class AxisScale
    {
        public double Min, Max = 1, Step = 0.25;

        public int Lines { get { return (int)Math.Round((Max - Min) / Step); } }

        // 0 for the lowest value on the axis, 1 for the highest
        public double Where(double v)
        {
            return Max > Min ? (v - Min) / (Max - Min) : 0;
        }
    }

    // =====================================================================
    //  LOOK: sizes and colour rules shared by the Excel builder, the preview and the web page,
    //  so all three show the same dashboard. No Excel calls in this file.
    // =====================================================================
    internal static class DashLook
    {
        // font sizes on the sheet, in points
        public const double TitlePt = 20, SubtitlePt = 9, CardLabelPt = 8.5, CardValuePt = 20, CardValueSmallPt = 16, CardNotePt = 8,
                            ChartTitlePt = 11, AxisPt = 8, LegendPt = 8, InsightPt = 10, InsightHeadPt = 9, FooterPt = 8, TablePt = 9;

        public static readonly int Up = ColorMath.Rgb(22, 163, 74), Down = ColorMath.Rgb(220, 38, 38);
        public static readonly int UpDark = ColorMath.Rgb(74, 222, 128), DownDark = ColorMath.Rgb(248, 113, 113);

        // a colour that stands out from the card background: the wanted colour, or the text colour when that is too close
        static int Readable(ThemeSpec t, int want, int background)
        {
            return Math.Abs(ColorMath.Lum(want) - ColorMath.Lum(background)) < 0.2 ? t.Text : want;
        }

        public static CardLook Card(DashSpec s, int accent)
        {
            ThemeSpec t = s.Theme;
            CardLook k = new CardLook();
            int style = s.CardStyle;
            if (style == 0 || style == 5)
            {
                int ink = ColorMath.Ink(accent);
                k.Bg = accent;
                k.Bg2 = style == 5 ? ColorMath.Mix(accent, ink == 0xFFFFFF ? 0 : 0xFFFFFF, 0.38) : accent;
                k.Gradient = style == 5;
                k.Value = ink;
                k.Label = ColorMath.Mix(accent, ink, 0.88);
                k.Note = ColorMath.Mix(accent, ink, 0.82);
                k.Line = accent;
                return k;
            }
            k.Plain = true;
            k.Bg = t.Card;
            k.Bg2 = t.Card;
            k.Label = t.Muted;
            k.Note = t.Muted;
            k.Value = Readable(t, accent, t.Card);
            k.Line = accent;
            if (style == 1) k.LineKind = 1;
            else if (style == 2) k.LineKind = 2;
            else if (style == 3)
            {
                k.LineKind = 3;
                k.Value = t.Text;
            }
            else
            {
                k.Bg = ColorMath.Mix(t.Card, accent, t.Dark ? 0.24 : 0.12);
                k.Bg2 = k.Bg;
                k.Value = Readable(t, t.Dark ? ColorMath.Lighter(accent, 0.25) : ColorMath.Darker(accent, 0.2), k.Bg);
                k.Label = ColorMath.Mix(t.Muted, k.Value, 0.35);
            }
            return k;
        }

        // The largest font size, not above pt, at which a number fits a width in points
        // (a bold digit is a little over half as wide as the size of the font).
        public static double FitFont(string text, double widthPt, double pt)
        {
            if (text == null || text.Length == 0 || widthPt <= 0) return pt;
            if (text.Length * pt * 0.57 <= widthPt) return pt;
            double fit = Math.Floor(widthPt / (text.Length * 0.57) * 2.0) / 2.0;
            return fit < 9 ? 9 : fit;
        }

        // data labels on a line or on columns: every k-th point gets one, so that they do not run into each other
        public static int LabelEvery(int points, double widthPt, int labelChars)
        {
            if (points < 1 || widthPt <= 0) return 1;
            double slot = widthPt / points;
            int k = (int)Math.Ceiling(labelChars * 4.6 / slot);
            return k < 1 ? 1 : k;
        }

        public static int UpColor(DashSpec s) { return s.Theme.Dark ? UpDark : Up; }
        public static int DownColor(DashSpec s) { return s.Theme.Dark ? DownDark : Down; }

        // rows of a card box: where the label, the number and the note sit. Returns the first row of each part,
        // counted from the top of the box: label, value, note, end.
        public static int[] CardRows(int boxRows)
        {
            int pad = boxRows - 7;
            if (pad < 0) pad = 0;
            int top = (pad + 1) / 2;
            return new[] { top, top + 2, top + 5, top + 7 };
        }

        // the colour of the j-th slice, stack or line of a chart
        public static int PointColor(DashSpec s, int j)
        {
            int c = s.Theme.Accent[((j % 6) + 6) % 6];
            int round = j / 6;
            if (round == 1) c = ColorMath.Mix(c, s.Theme.Card, 0.4);
            else if (round >= 2) c = ColorMath.Mix(c, s.Theme.Text, 0.35);
            return c;
        }

        public static int OthersColor(DashSpec s)
        {
            return ColorMath.Mix(s.Theme.Muted, s.Theme.Card, 0.45);
        }

        // the colour of a chart with one series
        public static int ChartColor(DashSpec s, ChartSpec c, int chartIndex)
        {
            return ColorMath.Pick(s.Theme, c.Color, chartIndex);
        }

        // the colour of series j of a chart with several series
        public static int SeriesColor(DashSpec s, ChartSpec c, int chartIndex, int j)
        {
            if (c.Kind == ChartKind.Combo || c.Kind == ChartKind.Pareto)
            {
                int first = ColorMath.Pick(s.Theme, c.Color, chartIndex);
                if (j == 0) return first;
                int second = s.Theme.Accent[2];
                return second == first ? s.Theme.Accent[4] : second;
            }
            return PointColor(s, j);
        }

        // legend: 1 off, 2 right, 3 bottom, 4 top
        public static int LegendPos(ChartSpec c, int series)
        {
            if (c.Legend > 0) return c.Legend;
            if (c.Kind == ChartKind.Table) return 1;
            if (ChartKind.IsRound(c.Kind)) return 2;
            return series > 1 ? 3 : 1;
        }

        // data labels: 1 off, 2 values, 3 percent
        public static int LabelMode(ChartSpec c)
        {
            if (c.Labels > 0)
            {
                if (c.Labels == 3 && !ChartKind.IsRound(c.Kind)) return 2;
                return c.Labels;
            }
            return ChartKind.IsRound(c.Kind) ? 3 : 1;
        }

        // 1, 2, 2.5, 5 or 10 times a power of ten: a round number at or above v
        public static double NiceMax(double v)
        {
            if (!(v > 0) || double.IsInfinity(v)) return 1;
            double p = Math.Pow(10, Math.Floor(Math.Log10(v)));
            double f = v / p;
            double n = f <= 1 ? 1 : (f <= 2 ? 2 : (f <= 2.5 ? 2.5 : (f <= 5 ? 5 : 10)));
            return n * p;
        }

        // A value axis with round numbers that holds everything from lo to hi (zero is always on it).
        public static AxisScale Axis(double lo, double hi, bool percent)
        {
            AxisScale a = new AxisScale();
            if (double.IsNaN(lo) || double.IsInfinity(lo) || lo > 0) lo = 0;
            if (double.IsNaN(hi) || double.IsInfinity(hi) || hi < 0) hi = 0;
            if (percent && lo == 0 && hi <= 1.0000001 && hi > 0.5)
            {
                a.Min = 0;
                a.Max = 1;
                a.Step = 0.25;
                return a;
            }
            double span = hi - lo;
            if (!(span > 0))
            {
                a.Min = 0;
                a.Max = 1;
                a.Step = 0.25;
                return a;
            }
            double step = NiceMax(span / 5.0);
            a.Step = step;
            a.Max = Math.Ceiling(hi / step - 1e-9) * step;
            a.Min = -Math.Ceiling(-lo / step - 1e-9) * step;
            if (a.Max == a.Min) a.Max = a.Min + step;
            return a;
        }

        // ---- filters ----
        // A filter button that is switched on is a soft tint of the first theme colour; one that is switched off is plain.
        // With no filter set every button is on, so the panel stays calm.
        public static int SlicerOn(DashSpec s) { return ColorMath.Mix(s.Theme.Card, s.Theme.Accent[0], s.Theme.Dark ? 0.42 : 0.2); }
        public static int SlicerOnText(DashSpec s) { return s.Theme.Text; }
        public static int SlicerOff(DashSpec s) { return ColorMath.Mix(s.Theme.Card, s.Theme.Muted, 0.08); }
        public static int SlicerOffText(DashSpec s) { return s.Theme.Muted; }
        public static int SlicerHover(DashSpec s) { return ColorMath.Mix(s.Theme.Card, s.Theme.Accent[0], s.Theme.Dark ? 0.6 : 0.36); }
        public static int SlicerEdge(DashSpec s) { return ColorMath.Mix(s.Theme.Card, s.Theme.Muted, 0.3); }

        // how many buttons a filter in the strip on top puts side by side; widthCells is the width of the filter in cells
        public static int SlicerColumns(int items, float widthCells)
        {
            int fit = (int)Math.Floor((widthCells - 1f) / 7f);
            if (fit < 1) fit = 1;
            if (items < 1) items = 1;
            return items < fit ? items : fit;
        }

        // The slicer style that comes with Excel and is nearest to the theme, for when an own style cannot be made.
        // Light1 to Light6 and Dark1 to Dark6 follow the six colours of the workbook theme: blue, orange, grey, gold, light blue, green.
        public static string BuiltInSlicerStyle(DashSpec s)
        {
            int c = s.Theme.Accent[0];
            int r = ColorMath.R(c), g = ColorMath.G(c), b = ColorMath.B(c);
            int mx = Math.Max(r, Math.Max(g, b)), mn = Math.Min(r, Math.Min(g, b));
            int pick = 1;
            if (mx - mn < 40) pick = 3;                               // grey
            else
            {
                double hue;
                double d = mx - mn;
                if (mx == r) hue = 60.0 * (((g - b) / d) % 6.0);
                else if (mx == g) hue = 60.0 * ((b - r) / d + 2.0);
                else hue = 60.0 * ((r - g) / d + 4.0);
                if (hue < 0) hue += 360.0;
                if (hue < 38 || hue >= 320) pick = 2;                 // red and orange
                else if (hue < 70) pick = 4;                          // gold
                else if (hue < 170) pick = 6;                         // green
                else if (hue < 205) pick = 5;                         // light blue
                else pick = 1;                                        // blue and purple
            }
            return (s.Theme.Dark ? "SlicerStyleDark" : "SlicerStyleLight") + pick.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
    }

    // What the preview has already worked out. Changing a colour redraws from here without adding up the table again.
    internal sealed class DashCache
    {
        readonly Dictionary<string, AggResult> charts = new Dictionary<string, AggResult>(StringComparer.Ordinal);
        readonly Dictionary<string, AggResult> months = new Dictionary<string, AggResult>(StringComparer.Ordinal);
        readonly Dictionary<string, double> scalars = new Dictionary<string, double>(StringComparer.Ordinal);
        readonly Dictionary<string, int> items = new Dictionary<string, int>(StringComparer.Ordinal);
        List<string> insights;
        string insightKey = "";
        public string Source = "";        // where the table comes from, for the line under the title

        public void Clear()
        {
            charts.Clear();
            months.Clear();
            scalars.Clear();
            items.Clear();
            insights = null;
            insightKey = "";
        }

        // Results are kept by what they were asked for, so they stay right when the design changes.
        // During a long session in Dashboard Studio the store is emptied now and then to keep the memory in check.
        public void Trim(int max)
        {
            if (charts.Count + months.Count + scalars.Count > max) Clear();
        }

        static string Base(DataTab t, DashSpec s)
        {
            return t.DateIndex.ToString(System.Globalization.CultureInfo.InvariantCulture) + "/" + s.FyStart.ToString(System.Globalization.CultureInfo.InvariantCulture) + "/";
        }

        static string I(int n) { return n.ToString(System.Globalization.CultureInfo.InvariantCulture); }

        public AggResult Chart(DataTab t, ChartSpec c, DashSpec s)
        {
            string key = Base(t, s) + I(c.Kind) + "\t" + c.Dim + "\t" + c.Col + "\t" + I(c.Agg) + "\t" + c.Split + "\t" + c.Col2 + "\t" + I(c.Agg2) + "\t" +
                         I(c.Top) + "\t" + I(c.Others) + "\t" + I(c.Sort) + "\t" + I(c.ShowAs);
            AggResult r;
            if (!charts.TryGetValue(key, out r))
            {
                r = DashEngine.Group(t, c, s);
                charts[key] = r;
            }
            return r;
        }

        public AggResult Months(DataTab t, string col, int agg, DashSpec s)
        {
            string key = Base(t, s) + col + "\t" + I(agg);
            AggResult r;
            if (!months.TryGetValue(key, out r))
            {
                r = DashEngine.Months(t, col, agg, s);
                months[key] = r;
            }
            return r;
        }

        public double Scalar(DataTab t, string col, int agg)
        {
            string key = col + "\t" + I(agg);
            double v;
            if (!scalars.TryGetValue(key, out v))
            {
                v = DashEngine.Scalar(t, col, agg);
                scalars[key] = v;
            }
            return v;
        }

        // how many buttons a filter has
        public int ItemCount(DataTab t, string field, DashSpec s)
        {
            string key = Base(t, s) + field;
            int n;
            if (!items.TryGetValue(key, out n))
            {
                List<string> labels;
                int[] k = t.Keys(field, s.FyStart, out labels);
                n = k == null ? 0 : labels.Count;
                items[key] = n;
            }
            return n;
        }

        public List<string> Insights(DataTab t, DashSpec s)
        {
            if (!s.Insights) return new List<string>();
            string firstCard = "", firstDim = "";
            foreach (CardSpec c in s.Cards)
            {
                if (c.Agg == AggKind.Sum && c.Col.Length > 0)
                {
                    firstCard = c.Col;
                    break;
                }
            }
            foreach (ChartSpec c in s.Charts)
            {
                if (!DatePart.Is(c.Dim))
                {
                    firstDim = c.Dim;
                    break;
                }
            }
            string key = Base(t, s) + firstCard + "\t" + firstDim + "\t" + I(s.Unit) + "\t" + string.Join("|", s.Slicers.ToArray());
            if (insights == null || key != insightKey)
            {
                insights = DashText.Insights(t, s);
                insightKey = key;
            }
            return insights;
        }

        // the sizes of the filters of a design, for the layout
        public int[] SlicerItems(DataTab t, DashSpec s)
        {
            if (s.SlicerPos == 3) return new int[0];
            List<int> list = new List<int>();
            foreach (string f in s.Slicers)
            {
                if (!t.CanGroup(f)) continue;
                list.Add(ItemCount(t, f, s));
            }
            return list.ToArray();
        }

        public DashGrid Grid(DataTab t, DashSpec s)
        {
            return DashLayout.Compute(s, SlicerItems(t, s), Insights(t, s).Count);
        }
    }
}
