using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace SmartPro
{
    // =====================================================================
    //  WEB DASHBOARD
    //  The rows and the design of a dashboard as one block of JSON inside a web page (WebTemplate.Html).
    //  The page adds the rows up again in the browser, so its filters change every card and every chart,
    //  and whoever opens it can switch theme, card style and chart kinds. No Excel calls in this file.
    // =====================================================================
    internal static class WebDash
    {
        public const int MaxRows = 100000;          // more rows than this are left out, to keep the file small
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        const int Numbers = 1, Groups = 2, Flags = 4;

        // The whole page as text. rows = data rows in the page; left = rows of the table that did not fit.
        public static string Page(DataTab t, DashSpec s, string source, string built, out int rows, out int left)
        {
            string data = Payload(t, s, source, built, MaxRows, out rows, out left);
            return WebTemplate.Html.Replace("__TITLE__", HtmlText(DashRender.TitleText(s))).Replace("/*DATA*/null", data);
        }

        public static string HtmlText(string s)
        {
            return (s ?? "").Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
        }

        // ---------- JSON ----------
        static void Str(StringBuilder sb, string s)
        {
            sb.Append('"');
            if (s != null)
            {
                foreach (char ch in s)
                {
                    switch (ch)
                    {
                        case '"': sb.Append("\\\""); break;
                        case '\\': sb.Append("\\\\"); break;
                        case '\n': sb.Append("\\n"); break;
                        case '\r': sb.Append("\\r"); break;
                        case '\t': sb.Append("\\t"); break;
                        default:
                            // < > & are written as codes, so that no text of the table can end the script block of the page
                            if (ch < ' ' || ch == '<' || ch == '>' || ch == '&' || ch == '\u2028' || ch == '\u2029')
                                sb.Append("\\u").Append(((int)ch).ToString("x4", Inv));
                            else sb.Append(ch);
                            break;
                    }
                }
            }
            sb.Append('"');
        }

        static void Num(StringBuilder sb, double d)
        {
            if (double.IsNaN(d) || double.IsInfinity(d)) sb.Append("null");
            else sb.Append(d.ToString("R", Inv));
        }

        static void Key(StringBuilder sb, string name)
        {
            sb.Append(",\"").Append(name).Append("\":");
        }

        static void Int(StringBuilder sb, string name, int v)
        {
            Key(sb, name);
            sb.Append(v.ToString(Inv));
        }

        static void Text(StringBuilder sb, string name, string v)
        {
            Key(sb, name);
            Str(sb, v);
        }

        static void Flag(StringBuilder sb, string name, bool v)
        {
            Key(sb, name);
            sb.Append(v ? "true" : "false");
        }

        // a colour of a card or chart: -1 automatic, 0 to 5 a theme colour, or "#rrggbb"
        static void Colour(StringBuilder sb, string name, int code)
        {
            Key(sb, name);
            if (code >= CardSpec.CustomColor) Str(sb, ColorMath.Hex((code - CardSpec.CustomColor) & 0xFFFFFF));
            else sb.Append((code >= 0 && code < 6 ? code : -1).ToString(Inv));
        }

        static void Theme(StringBuilder sb, ThemeSpec th)
        {
            sb.Append("{\"n\":");
            Str(sb, th.Name);
            Flag(sb, "d", th.Dark);
            Text(sb, "page", ColorMath.Hex(th.Page));
            Text(sb, "band", ColorMath.Hex(th.Band));
            Text(sb, "bandText", ColorMath.Hex(th.BandText));
            Text(sb, "card", ColorMath.Hex(th.Card));
            Text(sb, "text", ColorMath.Hex(th.Text));
            Text(sb, "muted", ColorMath.Hex(th.Muted));
            Text(sb, "grid", ColorMath.Hex(th.Grid));
            Key(sb, "a");
            sb.Append('[');
            for (int i = 0; i < 6; i++)
            {
                if (i > 0) sb.Append(',');
                Str(sb, ColorMath.Hex(th.Accent[i]));
            }
            sb.Append("]}");
        }

        // ---------- which columns the page needs, and what of them ----------
        static void Want(Dictionary<string, int> need, string name, int what)
        {
            int had;
            need.TryGetValue(name, out had);
            need[name] = had | what;
        }

        static void WantNumber(DataTab t, Dictionary<string, int> need, string col, int agg)
        {
            DataCol c = t.Find(col);
            if (c == null) return;
            // the same order of questions as DashEngine.Reduce
            if (agg == AggKind.Distinct) Want(need, c.Name, c.IsCat ? Groups : Numbers);
            else if (agg == AggKind.Count) Want(need, c.Name, c.IsNum ? Numbers : (c.IsCat ? Groups : Flags));
            else Want(need, c.Name, Numbers);
        }

        static void WantGroup(DataTab t, Dictionary<string, int> need, string field)
        {
            if (field == null || field.Length == 0 || DatePart.Is(field)) return;
            DataCol c = t.Find(field);
            if (c != null && c.IsCat) Want(need, c.Name, Groups);
        }

        // the name of a column as the table writes it (a design may write it in other capitals)
        static string Canon(DataTab t, string field)
        {
            if (field == null || field.Length == 0 || DatePart.Is(field)) return field ?? "";
            DataCol c = t.Find(field);
            return c == null ? field : c.Name;
        }

        // ---------- the block of JSON ----------
        public static string Payload(DataTab t, DashSpec s, string source, string built, int maxRows, out int rows, out int left)
        {
            t.UseDate(s.DateCol);
            int n = t.Rows > maxRows ? maxRows : t.Rows;
            rows = n;
            left = t.Rows - n;
            DashCache cache = new DashCache();
            cache.Source = source;
            DashGrid grid = cache.Grid(t, s);

            StringBuilder sb = new StringBuilder(1 << 16);
            sb.Append("{\"v\":3");
            Text(sb, "app", Design.AppName);
            Text(sb, "title", DashRender.TitleText(s));
            string sub = s.Subtitle.Trim().Length > 0 ? s.Subtitle.Trim() : DashText.AutoSubtitle(source, n);
            if (built != null && built.Length > 0) sub += "   |   " + built;
            Text(sb, "sub", sub);
            Text(sb, "foot", DashRender.FooterText(s));
            Text(sb, "logo", s.LogoData);
            Int(sb, "n", n);
            Int(sb, "left", left);
            Int(sb, "fy", s.FyStart);
            Int(sb, "unit", s.Unit);

            // ---- the columns ----
            Dictionary<string, int> need = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (CardSpec c in s.Cards)
            {
                WantNumber(t, need, c.Col, c.Agg);
                if (c.Op > 0) WantNumber(t, need, c.Col2, c.Agg2);
            }
            foreach (ChartSpec c in s.Charts)
            {
                WantGroup(t, need, c.Dim);
                WantNumber(t, need, c.Col, c.Agg);
                if (ChartKind.IsStacked(c.Kind)) WantGroup(t, need, c.Split);
                if (ChartKind.HasSecond(c.Kind)) WantNumber(t, need, c.Col2, c.Agg2);
            }
            foreach (string f in s.Slicers) WantGroup(t, need, f);
            string insCol, insDim;
            DashText.InsightBasis(t, s, out insCol, out insDim);
            if (s.Insights)
            {
                WantNumber(t, need, insCol, AggKind.Sum);
                WantGroup(t, need, insDim);
            }
            Key(sb, "cols");
            sb.Append('[');
            bool first = true;
            foreach (DataCol c in t.Cols)
            {
                int what;
                if (!need.TryGetValue(c.Name, out what)) continue;
                if (!first) sb.Append(',');
                first = false;
                sb.Append("{\"name\":");
                Str(sb, c.Name);
                if ((what & Numbers) != 0 && c.IsNum)
                {
                    Key(sb, "num");
                    sb.Append('[');
                    for (int i = 0; i < n; i++)
                    {
                        if (i > 0) sb.Append(',');
                        Num(sb, c.Num[i]);
                    }
                    sb.Append(']');
                }
                if ((what & Groups) != 0 && c.IsCat)
                {
                    // only the values that are in the rows of the page, numbered in the order they first appear
                    int[] place = new int[c.Items.Count];
                    for (int i = 0; i < place.Length; i++) place[i] = -1;
                    List<string> items = new List<string>();
                    StringBuilder codes = new StringBuilder(n * 3);
                    for (int i = 0; i < n; i++)
                    {
                        int k = c.Code[i];
                        if (place[k] < 0)
                        {
                            place[k] = items.Count;
                            items.Add(c.Items[k]);
                        }
                        if (i > 0) codes.Append(',');
                        codes.Append(place[k].ToString(Inv));
                    }
                    Key(sb, "items");
                    sb.Append('[');
                    for (int i = 0; i < items.Count; i++)
                    {
                        if (i > 0) sb.Append(',');
                        Str(sb, items[i]);
                    }
                    sb.Append(']');
                    Key(sb, "code");
                    sb.Append('[').Append(codes).Append(']');
                }
                if ((what & Flags) != 0)
                {
                    Key(sb, "filled");
                    sb.Append('"');
                    for (int i = 0; i < n; i++)
                    {
                        bool filled = c.IsDate ? c.Ymd[i] > 0 : (c.Filled != null && c.Filled[i]);
                        sb.Append(filled ? '1' : '0');
                    }
                    sb.Append('"');
                }
                sb.Append('}');
            }
            sb.Append(']');
            Key(sb, "ymd");
            if (t.HasDate)
            {
                int[] ymd = t.Cols[t.DateIndex].Ymd;
                sb.Append('[');
                for (int i = 0; i < n; i++)
                {
                    if (i > 0) sb.Append(',');
                    sb.Append(ymd[i].ToString(Inv));
                }
                sb.Append(']');
            }
            else sb.Append("null");

            // ---- KPI cards ----
            Key(sb, "cards");
            sb.Append('[');
            for (int i = 0; i < s.Cards.Count; i++)
            {
                CardSpec c = s.Cards[i];
                CardFormat f = NumFmt.Resolve(c, s);
                if (i > 0) sb.Append(',');
                sb.Append("{\"col\":");
                Str(sb, c.Col);
                Int(sb, "agg", c.Agg);
                Text(sb, "label", DashText.CardLabel(c));
                Int(sb, "mode", f.Mode);
                Int(sb, "dec", f.Dec);
                Flag(sb, "indian", f.Indian);
                Text(sb, "pre", f.Prefix);
                Text(sb, "suf", f.Suffix);
                Colour(sb, "color", c.Color);
                Int(sb, "cmp", t.HasDate ? c.Compare : 0);
                Flag(sb, "spark", c.Spark && t.HasDate);
                Key(sb, "target");
                Num(sb, c.Target > 0 ? c.Target : 0);
                if (c.Op > 0)
                {
                    Int(sb, "op", c.Op);
                    Text(sb, "col2", c.Col2);
                    Int(sb, "agg2", c.Agg2);
                    Text(sb, "opnote", DashText.CalcNote(c));
                }
                sb.Append('}');
            }
            sb.Append(']');

            // ---- charts: the width as twelfths of the row and the height in pixels come from the layout of the sheet ----
            int minX = int.MaxValue, maxRight = 0;
            foreach (Box b in grid.Boxes)
            {
                if (b.Kind != Box.Chart) continue;
                if (b.X < minX) minX = b.X;
                if (b.X + b.W > maxRight) maxRight = b.X + b.W;
            }
            int areaW = maxRight > minX ? maxRight - minX : 1;
            Key(sb, "charts");
            sb.Append('[');
            for (int i = 0; i < s.Charts.Count; i++)
            {
                ChartSpec c = s.Charts[i];
                Box b = grid.Find(Box.Chart, i);
                int span = 6, height = 22 * DashLayout.CellPx;
                if (b != null)
                {
                    span = (int)Math.Round((b.W + 1) * 12.0 / (areaW + 1));
                    if (span < 2) span = 2;
                    if (span > 12) span = 12;
                    height = b.H * DashLayout.CellPx;
                }
                if (i > 0) sb.Append(',');
                sb.Append("{\"kind\":").Append(c.Kind.ToString(Inv));
                Text(sb, "dim", Canon(t, c.Dim));
                Text(sb, "dimName", DashText.DimName(c.Dim));
                Text(sb, "col", c.Col);
                Int(sb, "agg", c.Agg);
                Text(sb, "split", ChartKind.IsStacked(c.Kind) ? Canon(t, c.Split) : "");
                Text(sb, "splitName", DashText.DimName(c.Split));
                Text(sb, "col2", ChartKind.HasSecond(c.Kind) ? c.Col2 : "");
                Int(sb, "agg2", c.Agg2);
                Int(sb, "top", c.Top);
                Int(sb, "others", c.Others);
                Int(sb, "sort", c.Sort);
                Text(sb, "title", c.Title.Trim());
                Int(sb, "labels", c.Labels);
                Int(sb, "legend", c.Legend);
                Colour(sb, "color", c.Color);
                Flag(sb, "smooth", c.Smooth);
                Int(sb, "trend", c.Trend);
                Int(sb, "showAs", c.ShowAs);
                Int(sb, "w", span);
                Int(sb, "h", height);
                sb.Append('}');
            }
            sb.Append(']');

            // ---- filters ----
            Key(sb, "slicers");
            sb.Append('[');
            first = true;
            if (s.SlicerPos != 3)
            {
                foreach (string f in s.Slicers)
                {
                    if (!t.CanGroup(f)) continue;
                    if (!first) sb.Append(',');
                    first = false;
                    sb.Append("{\"f\":");
                    Str(sb, DatePart.Is(f) ? f : t.Find(f).Name);
                    Text(sb, "name", DatePart.Is(f) ? DatePart.Label(f) : t.Find(f).Name);
                    sb.Append('}');
                }
            }
            sb.Append(']');

            // ---- the look ----
            int perRow = 0, firstY = -1;
            foreach (Box b in grid.Boxes)
            {
                if (b.Kind != Box.Card) continue;
                if (firstY < 0) firstY = b.Y;
                if (b.Y == firstY) perRow++;
            }
            Key(sb, "look");
            sb.Append("{\"pos\":").Append(s.SlicerPos.ToString(Inv));
            Int(sb, "cards", s.CardStyle);
            Flag(sb, "round", s.Rounded);
            Flag(sb, "shadow", s.Shadow);
            Int(sb, "dens", s.Density);
            Int(sb, "per", perRow);
            Flag(sb, "wide", s.Width == 1);
            Flag(sb, "insights", s.Insights);
            Text(sb, "font", s.Theme.Font);
            sb.Append('}');
            Key(sb, "ins");
            sb.Append("{\"col\":");
            Str(sb, insCol);
            Text(sb, "dim", Canon(t, insDim));
            sb.Append('}');
            Int(sb, "theme", DashThemes.IndexOf(s.Theme.Name));
            Key(sb, "own");
            Theme(sb, s.Theme);
            Key(sb, "themes");
            sb.Append('[');
            for (int i = 0; i < DashThemes.Count; i++)
            {
                if (i > 0) sb.Append(',');
                Theme(sb, DashThemes.Get(i));
            }
            sb.Append("]}");
            return sb.ToString();
        }
    }
}
