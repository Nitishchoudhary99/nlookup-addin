using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace SmartPro
{
    // =====================================================================
    //  THE DESIGN OF ONE DASHBOARD
    //  Everything a dashboard shows is described here: its cards, charts, filters, colours and layout.
    //  Dashboard Studio edits a DashSpec, the builder draws it in Excel, the Web Dashboard draws it in a browser,
    //  and the text form (ToText / FromText) is what gets saved in the dashboard sheet and in design files.
    //  No Excel calls in this file.
    // =====================================================================

    internal static class AggKind
    {
        public const int Sum = 0, Avg = 1, Count = 2, Max = 3, Min = 4, Distinct = 5;
        public static readonly string[] Names = { "Total (sum)", "Average", "Count (filled cells)", "Largest", "Smallest", "Different values" };
        public static readonly string[] Words = { "Total", "Average", "Count of", "Largest", "Smallest", "Different" };
    }

    internal static class ChartKind
    {
        public const int Column = 0, Bar = 1, Line = 2, Area = 3, Pie = 4, Doughnut = 5, StackedColumn = 6, StackedBar = 7,
                         Stacked100 = 8, Combo = 9, Pareto = 10, Table = 11;
        public static readonly string[] Names =
        {
            "Column", "Bar", "Line", "Area", "Pie", "Doughnut", "Stacked column", "Stacked bar", "100% stacked column",
            "Combo (columns + line)", "Pareto (80 / 20)", "Table with bars"
        };

        public static bool IsStacked(int k) { return k == StackedColumn || k == StackedBar || k == Stacked100; }
        public static bool IsRound(int k) { return k == Pie || k == Doughnut; }
        public static bool HasSecond(int k) { return k == Combo || k == Table; }
        public static bool IsLine(int k) { return k == Line || k == Area; }
    }

    // A KPI card can calculate with a second number: Profit % of Sales, Gross - Discount ...
    internal static class CardOp
    {
        public const int None = 0, Percent = 1, Divide = 2, Minus = 3, Plus = 4;
        public static readonly string[] Signs = { "", "% of", "\u00F7", "\u2212", "+" };
        public static readonly string[] Names = { "No (just this number)", "% of the second number", "\u00F7 the second number", "\u2212 the second number", "+ the second number" };

        public static double Apply(int op, double a, double b)
        {
            switch (op)
            {
                case Percent:
                case Divide: return b == 0 || double.IsNaN(b) ? 0 : a / b;
                case Minus: return a - b;
                case Plus: return a + b;
                default: return a;
            }
        }
    }

    // Date parts a chart or a filter can group by. They are written like column names that start with @.
    internal static class DatePart
    {
        public const string Month = "@month", Quarter = "@quarter", FiscalYear = "@fy", Year = "@year", Weekday = "@weekday", Day = "@day";
        public static readonly string[] All = { Month, Quarter, FiscalYear, Year, Weekday, Day };
        public static readonly string[] Labels = { "Month", "Quarter", "Fiscal Year", "Year", "Weekday", "Day" };

        public static bool Is(string name)
        {
            return name != null && name.Length > 1 && name[0] == '@';
        }

        public static string Label(string part)
        {
            for (int i = 0; i < All.Length; i++) if (All[i] == part) return Labels[i];
            return part;
        }
    }

    // Colours are Excel colour numbers: red + 256 * green + 65536 * blue.
    internal sealed class ThemeSpec
    {
        public string Name = "Navy Classic";
        public bool Dark;
        public int Page = 0xF9F5F1, Band = 0x2A170F, BandText = 0xFFFFFF, Card = 0xFFFFFF, Text = 0x2A170F, Muted = 0x8B7464, Grid = 0xF0E8E2;
        public int[] Accent = { 0xEB6325, 0x88940D, 0x0677D9, 0xED3A7C, 0x481DE1, 0xB29108 };
        public string Font = "Segoe UI";

        public ThemeSpec Clone()
        {
            ThemeSpec t = new ThemeSpec();
            t.Name = Name;
            t.Dark = Dark;
            t.Page = Page;
            t.Band = Band;
            t.BandText = BandText;
            t.Card = Card;
            t.Text = Text;
            t.Muted = Muted;
            t.Grid = Grid;
            t.Accent = new int[6];
            for (int i = 0; i < 6; i++) t.Accent[i] = Accent[i];
            t.Font = Font;
            return t;
        }
    }

    internal sealed class CardSpec
    {
        public const int CustomColor = 0x1000000;      // Color >= CustomColor: an own colour (the Excel colour number + CustomColor)

        public string Col = "";        // column header; "" = the number of rows
        public int Agg;                // AggKind
        public string Label = "";      // "" = automatic
        public int Fmt;                // 0 automatic, 1 short with units (L / Cr or K / M), 2 full number, 3 percent
        public int Dec = -1;           // decimals; -1 = automatic
        public string Prefix = "", Suffix = "";
        public int Color = -1;         // -1 = automatic, 0..5 = theme colour 1..6, >= CustomColor = own colour
        public int Compare;            // 0 none, 1 latest month with the month before, 2 latest month with the same month last year
        public bool Spark;             // small trend line
        public double Target;          // 0 = no target
        public int Op;                 // CardOp: calculate with a second number
        public string Col2 = "";       // the second number ("" = rows)
        public int Agg2;

        public CardSpec Clone()
        {
            CardSpec c = new CardSpec();
            c.Col = Col;
            c.Agg = Agg;
            c.Label = Label;
            c.Fmt = Fmt;
            c.Dec = Dec;
            c.Prefix = Prefix;
            c.Suffix = Suffix;
            c.Color = Color;
            c.Compare = Compare;
            c.Spark = Spark;
            c.Target = Target;
            c.Op = Op;
            c.Col2 = Col2;
            c.Agg2 = Agg2;
            return c;
        }
    }

    internal sealed class ChartSpec
    {
        public int Kind;               // ChartKind
        public string Dim = "";        // column to group by, or a DatePart
        public string Col = "";        // number column; "" = rows
        public int Agg;                // AggKind
        public string Split = "";      // second column to group by: one colour for each of its values (stacked charts)
        public string Col2 = "";       // second number (combo chart: the line; table: a second column)
        public int Agg2;
        public int Top;                // 0 = automatic, -1 = all, otherwise the N largest
        public int Others;             // 0 = automatic, 1 = add up the rest as "Others", 2 = leave the rest out
        public int Sort;               // 0 automatic, 1 largest first, 2 smallest first, 3 A to Z
        public string Title = "";      // "" = automatic
        public int Labels;             // 0 automatic, 1 off, 2 values, 3 percent
        public int Legend;             // 0 automatic, 1 off, 2 right, 3 bottom, 4 top
        public int Span;               // 0 automatic, otherwise twelfths of the width: 3, 4, 6, 8, 9, 12
        public int Height;             // 0 automatic, 1 short, 2 standard, 3 tall, 4 extra tall
        public int Color = -1;         // like CardSpec.Color
        public bool Smooth;            // curved line
        public int Trend;              // 0 none, 1 straight trend line, 2 moving average of 3
        public int ShowAs;             // 0 values, 1 % of total, 2 running total

        public ChartSpec Clone()
        {
            ChartSpec c = new ChartSpec();
            c.Kind = Kind;
            c.Dim = Dim;
            c.Col = Col;
            c.Agg = Agg;
            c.Split = Split;
            c.Col2 = Col2;
            c.Agg2 = Agg2;
            c.Top = Top;
            c.Others = Others;
            c.Sort = Sort;
            c.Title = Title;
            c.Labels = Labels;
            c.Legend = Legend;
            c.Span = Span;
            c.Height = Height;
            c.Color = Color;
            c.Smooth = Smooth;
            c.Trend = Trend;
            c.ShowAs = ShowAs;
            return c;
        }
    }

    internal sealed class DashSpec
    {
        public const int MaxCards = 8, MaxCharts = 10, MaxSlicers = 6;
        public const string Magic = "DOCBRISK-DASH 3";

        public string Title = "", Subtitle = "", Footer = "";
        public string LogoName = "", LogoData = "";    // the logo picture as PNG bytes in base64 ("" = no logo)
        public List<CardSpec> Cards = new List<CardSpec>();
        public List<ChartSpec> Charts = new List<ChartSpec>();
        public List<string> Slicers = new List<string>();     // column headers and DateParts
        public int SlicerPos;          // 0 left panel, 1 strip on top, 2 right panel, 3 no filters
        public string DateCol = "";    // the date column behind month / quarter / year ("" = the first date column)
        public int FyStart = 4;        // first month of the fiscal year
        public ThemeSpec Theme = new ThemeSpec();
        public int CardStyle;          // 0 filled, 1 light with a top line, 2 light with a side bar, 3 outline, 4 soft tint, 5 gradient
        public bool Rounded = true, Shadow = true;
        public int Density = 1;        // 0 compact, 1 comfortable, 2 spacious
        public int Width;              // 0 standard (fits a laptop screen), 1 wide
        public int Arrange;            // how charts without an own size are arranged (DashLayout.ArrangeNames)
        public int CardsPerRow;        // 0 = automatic
        public int Unit;               // 0 Indian (L / Cr), 1 international (K / M / B), 2 full numbers
        public string Currency = "";   // put in front of amounts on cards, for example a rupee sign
        public bool Insights = true;
        public bool HideHeadings = true, PrintSetup = true;

        // ---------- copies ----------
        public DashSpec Clone()
        {
            DashSpec d = FromText(ToText());
            return d ?? new DashSpec();
        }

        // takes the look (colours, card style, layout) of another design and keeps the own cards, charts and filters
        public void CopyLook(DashSpec other)
        {
            Theme = other.Theme.Clone();
            CardStyle = other.CardStyle;
            Rounded = other.Rounded;
            Shadow = other.Shadow;
            Density = other.Density;
            Width = other.Width;
            Arrange = other.Arrange;
            SlicerPos = other.SlicerPos;
            Unit = other.Unit;
            Currency = other.Currency;
            HideHeadings = other.HideHeadings;
            PrintSetup = other.PrintSetup;
        }

        // ---------- text form ----------
        public string ToText()
        {
            StringBuilder sb = new StringBuilder();
            sb.Append(Magic).Append('\n');
            Rec r = new Rec("dash");
            r.Set("title", Title);
            r.Set("sub", Subtitle);
            r.Set("foot", Footer);
            r.Set("slpos", SlicerPos);
            r.Set("date", DateCol);
            r.Set("fy", FyStart);
            r.Set("cards", CardStyle);
            r.Set("round", Rounded);
            r.Set("shadow", Shadow);
            r.Set("dens", Density);
            r.Set("width", Width);
            r.Set("arr", Arrange);
            r.Set("cpr", CardsPerRow);
            r.Set("unit", Unit);
            r.Set("cur", Currency);
            r.Set("ins", Insights);
            r.Set("hidehead", HideHeadings);
            r.Set("print", PrintSetup);
            sb.Append(r.ToLine()).Append('\n');

            r = new Rec("theme");
            r.Set("name", Theme.Name);
            r.Set("dark", Theme.Dark);
            r.Set("page", Theme.Page);
            r.Set("band", Theme.Band);
            r.Set("bandtext", Theme.BandText);
            r.Set("card", Theme.Card);
            r.Set("text", Theme.Text);
            r.Set("muted", Theme.Muted);
            r.Set("grid", Theme.Grid);
            for (int i = 0; i < 6; i++) r.Set("a" + i.ToString(CultureInfo.InvariantCulture), Theme.Accent[i]);
            r.Set("font", Theme.Font);
            sb.Append(r.ToLine()).Append('\n');

            foreach (CardSpec c in Cards)
            {
                r = new Rec("card");
                r.Set("col", c.Col);
                r.Set("agg", c.Agg);
                r.Set("label", c.Label);
                r.Set("fmt", c.Fmt);
                r.Set("dec", c.Dec);
                r.Set("pre", c.Prefix);
                r.Set("suf", c.Suffix);
                r.Set("color", c.Color);
                r.Set("cmp", c.Compare);
                r.Set("spark", c.Spark);
                r.Set("target", c.Target);
                if (c.Op > 0)
                {
                    r.Set("op", c.Op);
                    r.Set("col2", c.Col2);
                    r.Set("agg2", c.Agg2);
                }
                sb.Append(r.ToLine()).Append('\n');
            }
            foreach (ChartSpec c in Charts)
            {
                r = new Rec("chart");
                r.Set("kind", c.Kind);
                r.Set("dim", c.Dim);
                r.Set("col", c.Col);
                r.Set("agg", c.Agg);
                r.Set("split", c.Split);
                r.Set("col2", c.Col2);
                r.Set("agg2", c.Agg2);
                r.Set("top", c.Top);
                r.Set("others", c.Others);
                r.Set("sort", c.Sort);
                r.Set("title", c.Title);
                r.Set("labels", c.Labels);
                r.Set("legend", c.Legend);
                r.Set("span", c.Span);
                r.Set("height", c.Height);
                r.Set("color", c.Color);
                r.Set("smooth", c.Smooth);
                r.Set("trend", c.Trend);
                r.Set("showas", c.ShowAs);
                sb.Append(r.ToLine()).Append('\n');
            }
            foreach (string s in Slicers)
            {
                r = new Rec("slicer");
                r.Set("f", s);
                sb.Append(r.ToLine()).Append('\n');
            }
            if (LogoData.Length > 0)
            {
                r = new Rec("logo");
                r.Set("name", LogoName);
                r.Set("data", LogoData);
                sb.Append(r.ToLine()).Append('\n');
            }
            return sb.ToString();
        }

        // null when the text is not a dashboard design
        public static DashSpec FromText(string text)
        {
            if (text == null) return null;
            string[] lines = text.Replace("\r\n", "\n").Split('\n');
            if (lines.Length == 0 || !lines[0].Trim().StartsWith("DOCBRISK-DASH", StringComparison.Ordinal)) return null;
            DashSpec d = new DashSpec();
            for (int i = 1; i < lines.Length; i++)
            {
                if (lines[i].Length == 0) continue;
                Rec r = Rec.Parse(lines[i]);
                if (r.Kind == "dash")
                {
                    d.Title = r.Get("title", "");
                    d.Subtitle = r.Get("sub", "");
                    d.Footer = r.Get("foot", "");
                    d.SlicerPos = Clamp(r.GetInt("slpos", 0), 0, 3);
                    d.DateCol = r.Get("date", "");
                    d.FyStart = Clamp(r.GetInt("fy", 4), 1, 12);
                    d.CardStyle = Clamp(r.GetInt("cards", 0), 0, 5);
                    d.Rounded = r.GetBool("round", true);
                    d.Shadow = r.GetBool("shadow", true);
                    d.Density = Clamp(r.GetInt("dens", 1), 0, 2);
                    d.Width = Clamp(r.GetInt("width", 0), 0, 1);
                    d.Arrange = Clamp(r.GetInt("arr", 0), 0, 7);
                    d.CardsPerRow = Clamp(r.GetInt("cpr", 0), 0, 8);
                    d.Unit = Clamp(r.GetInt("unit", 0), 0, 2);
                    d.Currency = r.Get("cur", "");
                    d.Insights = r.GetBool("ins", true);
                    d.HideHeadings = r.GetBool("hidehead", true);
                    d.PrintSetup = r.GetBool("print", true);
                }
                else if (r.Kind == "theme")
                {
                    ThemeSpec t = d.Theme;
                    t.Name = r.Get("name", t.Name);
                    t.Dark = r.GetBool("dark", false);
                    t.Page = r.GetInt("page", t.Page) & 0xFFFFFF;
                    t.Band = r.GetInt("band", t.Band) & 0xFFFFFF;
                    t.BandText = r.GetInt("bandtext", t.BandText) & 0xFFFFFF;
                    t.Card = r.GetInt("card", t.Card) & 0xFFFFFF;
                    t.Text = r.GetInt("text", t.Text) & 0xFFFFFF;
                    t.Muted = r.GetInt("muted", t.Muted) & 0xFFFFFF;
                    t.Grid = r.GetInt("grid", t.Grid) & 0xFFFFFF;
                    for (int k = 0; k < 6; k++) t.Accent[k] = r.GetInt("a" + k.ToString(CultureInfo.InvariantCulture), t.Accent[k]) & 0xFFFFFF;
                    t.Font = r.Get("font", t.Font);
                    if (t.Font.Trim().Length == 0) t.Font = "Segoe UI";
                }
                else if (r.Kind == "card")
                {
                    if (d.Cards.Count >= MaxCards) continue;
                    CardSpec c = new CardSpec();
                    c.Col = r.Get("col", "");
                    c.Agg = Clamp(r.GetInt("agg", 0), 0, 5);
                    c.Label = r.Get("label", "");
                    c.Fmt = Clamp(r.GetInt("fmt", 0), 0, 3);
                    c.Dec = Clamp(r.GetInt("dec", -1), -1, 4);
                    c.Prefix = r.Get("pre", "");
                    c.Suffix = r.Get("suf", "");
                    c.Color = r.GetInt("color", -1);
                    c.Compare = Clamp(r.GetInt("cmp", 0), 0, 2);
                    c.Spark = r.GetBool("spark", false);
                    c.Target = r.GetNum("target", 0);
                    c.Op = Clamp(r.GetInt("op", 0), 0, 4);
                    c.Col2 = r.Get("col2", "");
                    c.Agg2 = Clamp(r.GetInt("agg2", 0), 0, 4);
                    d.Cards.Add(c);
                }
                else if (r.Kind == "chart")
                {
                    if (d.Charts.Count >= MaxCharts) continue;
                    ChartSpec c = new ChartSpec();
                    c.Kind = Clamp(r.GetInt("kind", 0), 0, 11);
                    c.Dim = r.Get("dim", "");
                    c.Col = r.Get("col", "");
                    c.Agg = Clamp(r.GetInt("agg", 0), 0, 5);
                    c.Split = r.Get("split", "");
                    c.Col2 = r.Get("col2", "");
                    c.Agg2 = Clamp(r.GetInt("agg2", 0), 0, 5);
                    c.Top = Clamp(r.GetInt("top", 0), -1, 100);
                    c.Others = Clamp(r.GetInt("others", 0), 0, 2);
                    c.Sort = Clamp(r.GetInt("sort", 0), 0, 3);
                    c.Title = r.Get("title", "");
                    c.Labels = Clamp(r.GetInt("labels", 0), 0, 3);
                    c.Legend = Clamp(r.GetInt("legend", 0), 0, 4);
                    c.Span = Clamp(r.GetInt("span", 0), 0, 12);
                    c.Height = Clamp(r.GetInt("height", 0), 0, 4);
                    c.Color = r.GetInt("color", -1);
                    c.Smooth = r.GetBool("smooth", false);
                    c.Trend = Clamp(r.GetInt("trend", 0), 0, 2);
                    c.ShowAs = Clamp(r.GetInt("showas", 0), 0, 2);
                    if (c.Dim.Length > 0) d.Charts.Add(c);
                }
                else if (r.Kind == "slicer")
                {
                    string f = r.Get("f", "");
                    if (f.Length > 0 && d.Slicers.Count < MaxSlicers && !d.Slicers.Contains(f)) d.Slicers.Add(f);
                }
                else if (r.Kind == "logo")
                {
                    d.LogoName = r.Get("name", "");
                    d.LogoData = r.Get("data", "");
                }
            }
            return d;
        }

        static int Clamp(int v, int lo, int hi)
        {
            return v < lo ? lo : (v > hi ? hi : v);
        }
    }

    // One line of the text form: a kind and named values, separated by tabs. Unknown names are ignored when read,
    // so a design saved by a newer version still opens in an older one.
    internal sealed class Rec
    {
        public string Kind = "";
        readonly List<string> keys = new List<string>();
        readonly List<string> vals = new List<string>();

        public Rec(string kind)
        {
            Kind = kind;
        }

        public void Set(string key, string value)
        {
            keys.Add(key);
            vals.Add(value ?? "");
        }

        public void Set(string key, int value)
        {
            Set(key, value.ToString(CultureInfo.InvariantCulture));
        }

        public void Set(string key, bool value)
        {
            Set(key, value ? "1" : "0");
        }

        public void Set(string key, double value)
        {
            Set(key, value.ToString("R", CultureInfo.InvariantCulture));
        }

        public string Get(string key, string def)
        {
            for (int i = 0; i < keys.Count; i++) if (keys[i] == key) return vals[i];
            return def;
        }

        public int GetInt(string key, int def)
        {
            int n;
            string s = Get(key, null);
            if (s == null) return def;
            return int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out n) ? n : def;
        }

        public bool GetBool(string key, bool def)
        {
            string s = Get(key, null);
            if (s == null) return def;
            return s == "1" || s == "true" || s == "True";
        }

        public double GetNum(string key, double def)
        {
            double d;
            string s = Get(key, null);
            if (s == null) return def;
            if (!double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out d)) return def;
            if (double.IsNaN(d) || double.IsInfinity(d)) return def;
            return d;
        }

        static string Esc(string s)
        {
            if (s.IndexOf('\\') < 0 && s.IndexOf('\t') < 0 && s.IndexOf('\n') < 0 && s.IndexOf('\r') < 0) return s;
            StringBuilder sb = new StringBuilder(s.Length + 8);
            foreach (char ch in s)
            {
                if (ch == '\\') sb.Append("\\\\");
                else if (ch == '\t') sb.Append("\\t");
                else if (ch == '\n') sb.Append("\\n");
                else if (ch == '\r') sb.Append("\\r");
                else sb.Append(ch);
            }
            return sb.ToString();
        }

        static string Unesc(string s)
        {
            if (s.IndexOf('\\') < 0) return s;
            StringBuilder sb = new StringBuilder(s.Length);
            int i = 0;
            while (i < s.Length)
            {
                char ch = s[i];
                if (ch == '\\' && i + 1 < s.Length)
                {
                    char nx = s[i + 1];
                    if (nx == 't') sb.Append('\t');
                    else if (nx == 'n') sb.Append('\n');
                    else if (nx == 'r') sb.Append('\r');
                    else sb.Append(nx);
                    i += 2;
                }
                else
                {
                    sb.Append(ch);
                    i++;
                }
            }
            return sb.ToString();
        }

        public string ToLine()
        {
            StringBuilder sb = new StringBuilder(Kind);
            for (int i = 0; i < keys.Count; i++) sb.Append('\t').Append(keys[i]).Append('=').Append(Esc(vals[i]));
            return sb.ToString();
        }

        public static Rec Parse(string line)
        {
            string[] parts = line.Split('\t');
            Rec r = new Rec(parts[0].Trim());
            for (int i = 1; i < parts.Length; i++)
            {
                int p = parts[i].IndexOf('=');
                if (p <= 0) continue;
                r.keys.Add(parts[i].Substring(0, p));
                r.vals.Add(Unesc(parts[i].Substring(p + 1)));
            }
            return r;
        }
    }
}
