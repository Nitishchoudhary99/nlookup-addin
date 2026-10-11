using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Globalization;

namespace SmartPro
{
    // a part of the preview that can be clicked
    internal sealed class HitBox
    {
        public int Kind, Index;           // Box.Card, Box.Chart, Box.Slicer, Box.Header, Box.Insights
        public RectangleF R;

        public HitBox(int kind, int index, RectangleF r)
        {
            Kind = kind;
            Index = index;
            R = r;
        }
    }

    // =====================================================================
    //  PREVIEW
    //  Draws a design with the real numbers of the table, in the same boxes the Excel builder uses.
    //  Dashboard Studio shows it live; the design gallery uses it for its small pictures.
    // =====================================================================
    internal static class DashRender
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        static Color C(int rgb)
        {
            return Color.FromArgb(ColorMath.R(rgb), ColorMath.G(rgb), ColorMath.B(rgb));
        }

        static Color C(int rgb, int alpha)
        {
            return Color.FromArgb(alpha, ColorMath.R(rgb), ColorMath.G(rgb), ColorMath.B(rgb));
        }

        static GraphicsPath RoundPath(RectangleF r, float rad)
        {
            GraphicsPath p = new GraphicsPath();
            float d = rad * 2;
            if (d > r.Width) d = r.Width;
            if (d > r.Height) d = r.Height;
            if (d < 1f)
            {
                p.AddRectangle(r);
                return p;
            }
            p.AddArc(r.X, r.Y, d, d, 180f, 90f);
            p.AddArc(r.Right - d, r.Y, d, d, 270f, 90f);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0f, 90f);
            p.AddArc(r.X, r.Bottom - d, d, d, 90f, 90f);
            p.CloseFigure();
            return p;
        }

        static void FillBox(Graphics g, Color c, RectangleF r, float rad)
        {
            using (SolidBrush b = new SolidBrush(c))
            {
                if (rad < 1f)
                {
                    g.FillRectangle(b, r);
                    return;
                }
                using (GraphicsPath p = RoundPath(r, rad))
                {
                    g.FillPath(b, p);
                }
            }
        }

        static void StrokeBox(Graphics g, Color c, RectangleF r, float rad, float width)
        {
            using (Pen pen = new Pen(c, width))
            {
                if (rad < 1f)
                {
                    g.DrawRectangle(pen, r.X, r.Y, r.Width, r.Height);
                    return;
                }
                using (GraphicsPath p = RoundPath(r, rad))
                {
                    g.DrawPath(pen, p);
                }
            }
        }

        // text in a box; too long text ends in three dots. Below 4.5 pixels a grey bar stands in for the text.
        static void Text(Graphics g, string s, string family, float px, bool bold, Color c, RectangleF r, StringAlignment h, StringAlignment v)
        {
            if (s == null || s.Length == 0 || r.Width < 2f || r.Height < 2f) return;
            if (px < 4.5f)
            {
                float w = Math.Min(r.Width, s.Length * px * 0.55f);
                float bh = Math.Max(1f, px * 0.55f);
                float x = h == StringAlignment.Far ? r.Right - w : (h == StringAlignment.Center ? r.X + (r.Width - w) / 2f : r.X);
                float y = v == StringAlignment.Far ? r.Bottom - bh : (v == StringAlignment.Center ? r.Y + (r.Height - bh) / 2f : r.Y);
                using (SolidBrush b = new SolidBrush(Color.FromArgb(120, c)))
                {
                    g.FillRectangle(b, x, y, w, bh);
                }
                return;
            }
            using (Font f = new Font(family, px, bold ? FontStyle.Bold : FontStyle.Regular, GraphicsUnit.Pixel))
            using (SolidBrush b = new SolidBrush(c))
            using (StringFormat sf = new StringFormat(StringFormatFlags.NoWrap))
            {
                sf.Alignment = h;
                sf.LineAlignment = v;
                sf.Trimming = StringTrimming.EllipsisCharacter;
                g.DrawString(s, f, b, r, sf);
            }
        }

        static float TextWidth(Graphics g, string s, string family, float px, bool bold)
        {
            if (s == null || s.Length == 0) return 0f;
            if (px < 4.5f) return s.Length * px * 0.55f;
            using (Font f = new Font(family, px, bold ? FontStyle.Bold : FontStyle.Regular, GraphicsUnit.Pixel))
            {
                return g.MeasureString(s, f).Width;
            }
        }

        // ------------------------------------------------------------------
        // Draws the whole dashboard into 'area', scaled to its width. Returns the height it needs.
        // hits (may be null) receives the boxes that can be clicked. sel = the selected box, or -1.
        public static float Paint(Graphics g, RectangleF area, DataTab t, DashSpec s, DashCache cache, int selKind, int selIndex, List<HitBox> hits)
        {
            DashGrid grid = cache.Grid(t, s);
            float u = area.Width / grid.Cols;
            float total = grid.Rows * u;
            ThemeSpec th = s.Theme;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            using (SolidBrush b = new SolidBrush(C(th.Page)))
            {
                g.FillRectangle(b, area.X, area.Y, area.Width, Math.Max(total, area.Height));
            }
            float rad = s.Rounded ? u * 0.75f : 0f;
            List<string> insights = cache.Insights(t, s);
            List<string> fields = new List<string>();
            foreach (string f in s.Slicers) if (t.CanGroup(f)) fields.Add(f);
            RectangleF selRect = RectangleF.Empty;
            bool hasSel = false;
            foreach (Box bx in grid.Boxes)
            {
                RectangleF r = new RectangleF(area.X + bx.X * u, area.Y + bx.Y * u, bx.W * u, bx.H * u);
                if (bx.Kind == Box.Header) Header(g, r, s, grid, u, t, cache.Source);
                else if (bx.Kind == Box.Card && bx.Index < s.Cards.Count) Card(g, r, t, s, cache, bx.Index, u, rad);
                else if (bx.Kind == Box.Chart && bx.Index < s.Charts.Count) Chart(g, r, t, s, cache, bx.Index, u, rad);
                else if (bx.Kind == Box.Slicer && bx.Index < fields.Count) Slicer(g, r, t, s, fields[bx.Index], u, rad);
                else if (bx.Kind == Box.Insights) Insights(g, r, s, insights, u, rad);
                else if (bx.Kind == Box.Footer) Text(g, FooterText(s), th.Font, Px(DashLook.FooterPt, u), false, C(th.Muted), r, StringAlignment.Near, StringAlignment.Center);
                if (hits != null && bx.Kind != Box.Footer) hits.Add(new HitBox(bx.Kind, bx.Index, r));
                if (bx.Kind == selKind && bx.Index == selIndex && bx.Kind != Box.Footer)
                {
                    selRect = r;
                    hasSel = true;
                }
            }
            if (hasSel)
            {
                RectangleF o = new RectangleF(selRect.X - 2f, selRect.Y - 2f, selRect.Width + 4f, selRect.Height + 4f);
                int mark = th.Dark ? 0xFFFFFF : ColorMath.Rgb(15, 23, 42);
                using (Pen pen = new Pen(C(mark), 2f))
                {
                    pen.DashStyle = DashStyle.Dash;
                    g.DrawRectangle(pen, o.X, o.Y, o.Width, o.Height);
                }
            }
            return total;
        }

        // points on the sheet to pixels in the preview: one grid cell is 9 points
        static float Px(double points, float u)
        {
            return (float)(points * u / 9.0);
        }

        public static string FooterText(DashSpec s)
        {
            return s.Footer.Trim().Length > 0 ? s.Footer.Trim() : "Made with " + Design.AppName + " for Excel";
        }

        public static string TitleText(DashSpec s)
        {
            return s.Title.Trim().Length > 0 ? s.Title.Trim() : "Dashboard";
        }

        static void Header(Graphics g, RectangleF r, DashSpec s, DashGrid grid, float u, DataTab t, string source)
        {
            ThemeSpec th = s.Theme;
            using (SolidBrush b = new SolidBrush(C(th.Band)))
            {
                g.FillRectangle(b, r);
            }
            if (ColorMath.Lum(th.Band) > 0.9 && ColorMath.Lum(th.Page) > 0.9)
            {
                using (Pen pen = new Pen(C(th.Grid), 1f))
                {
                    g.DrawLine(pen, r.X, r.Bottom - 0.5f, r.Right, r.Bottom - 0.5f);
                }
            }
            float left = r.X + grid.Margin * u;
            float w = grid.Width * u;
            float logoW = 0f;
            if (s.LogoData.Length > 0)
            {
                // the picture itself is drawn by Dashboard Studio; here its place is kept free
                logoW = r.Height * 2.2f;
            }
            float split = r.Height * 0.6f;
            RectangleF tr = new RectangleF(left, r.Y, w - logoW, split);
            Text(g, TitleText(s), th.Font, Px(DashLook.TitlePt, u), true, C(th.BandText), tr, StringAlignment.Near, StringAlignment.Far);
            string sub = s.Subtitle.Trim().Length > 0 ? s.Subtitle.Trim() : DashText.AutoSubtitle(source, t.Rows);
            RectangleF sr = new RectangleF(left, r.Y + split, w - logoW, r.Height - split);
            Text(g, sub, th.Font, Px(DashLook.SubtitlePt, u), false, C(ColorMath.Mix(th.Band, th.BandText, 0.72)), sr, StringAlignment.Near, StringAlignment.Near);
        }

        // where the logo goes inside the header box
        public static RectangleF LogoBox(RectangleF header, DashGrid grid, float u)
        {
            float h = header.Height - 2f * u;
            if (h < 4f) h = header.Height;
            float w = header.Height * 2.2f;
            float right = header.X + (grid.Margin + grid.Width) * u;
            return new RectangleF(right - w, header.Y + (header.Height - h) / 2f, w, h);
        }

        // ------------------------------------------------------------------ KPI card
        public static double CardValue(DataTab t, DashCache cache, CardSpec c)
        {
            double v = cache.Scalar(t, c.Col, c.Agg);
            return c.Op > 0 ? CardOp.Apply(c.Op, v, cache.Scalar(t, c.Col2, c.Agg2)) : v;
        }

        public static string CardValueText(DataTab t, DashSpec s, DashCache cache, CardSpec c)
        {
            return NumFmt.Apply(CardValue(t, cache, c), NumFmt.Resolve(c, s));
        }

        // the line under the number; sign says whether it is a rise (1), a fall (-1) or neither (0)
        public static string CardNoteText(DataTab t, DashSpec s, DashCache cache, CardSpec c, out int sign)
        {
            sign = 0;
            if (c.Compare > 0 && t.HasDate)
            {
                double ch = DashEngine.Change(cache.Months(t, c.Col, c.Agg, s), c.Compare);
                if (double.IsNaN(ch)) return "";
                sign = ch >= 0 ? 1 : -1;
                return (ch >= 0 ? "\u25B2 " : "\u25BC ") + Math.Abs(ch).ToString("0.0%", Inv) + " " + DashText.CompareWords(c.Compare);
            }
            if (c.Op > 0 && c.Target <= 0) return DashText.CalcNote(c);
            if (c.Target > 0)
            {
                double v = CardValue(t, cache, c);
                return (v / c.Target).ToString("0%", Inv) + " of target " + NumFmt.Short(c.Target, s.Unit, -1);
            }
            if (c.Agg == AggKind.Sum && c.Col.Length > 0)
            {
                double v = cache.Scalar(t, c.Col, c.Agg);
                double n = cache.Scalar(t, "", AggKind.Count);
                return "avg per row: " + NumFmt.Grouped(n > 0 ? v / n : 0, 2, s.Unit == 0);
            }
            return DashText.CardNote(c);
        }

        static void Card(Graphics g, RectangleF r, DataTab t, DashSpec s, DashCache cache, int index, float u, float rad)
        {
            CardSpec c = s.Cards[index];
            ThemeSpec th = s.Theme;
            int accent = ColorMath.Pick(th, c.Color, index);
            CardLook k = DashLook.Card(s, accent);
            if (s.Shadow && s.Rounded)
            {
                RectangleF sh = new RectangleF(r.X, r.Y + u * 0.18f, r.Width, r.Height);
                FillBox(g, Color.FromArgb(th.Dark ? 70 : 26, 0, 0, 0), sh, rad);
            }
            if (k.Gradient)
            {
                using (LinearGradientBrush lb = new LinearGradientBrush(new RectangleF(r.X, r.Y, r.Width + 1f, r.Height + 1f), C(k.Bg), C(k.Bg2), 0f))
                using (GraphicsPath p = RoundPath(r, rad))
                {
                    g.FillPath(lb, p);
                }
            }
            else FillBox(g, C(k.Bg), r, rad);
            if (k.LineKind == 1)
            {
                RectangleF line = rad >= 1f ? new RectangleF(r.X + u, r.Y + u * 0.45f, u * 3f, Math.Max(2f, u * 0.3f)) : new RectangleF(r.X, r.Y, r.Width, Math.Max(2f, u * 0.25f));
                FillBox(g, C(k.Line), line, rad >= 1f ? u * 0.15f : 0f);
            }
            else if (k.LineKind == 2)
            {
                RectangleF bar = rad >= 1f ? new RectangleF(r.X + u * 0.5f, r.Y + u, Math.Max(2f, u * 0.3f), r.Height - 2f * u) : new RectangleF(r.X, r.Y, Math.Max(2f, u * 0.3f), r.Height);
                FillBox(g, C(k.Line), bar, rad >= 1f ? u * 0.15f : 0f);
            }
            else if (k.LineKind == 3) StrokeBox(g, C(ColorMath.Mix(th.Card, k.Line, 0.6)), r, rad, Math.Max(1f, u * 0.12f));

            int rowsInBox = (int)Math.Round(r.Height / u);
            int[] rows = DashLook.CardRows(rowsInBox);
            float pad = u * (k.LineKind == 2 ? 1.4f : 1f);
            float inner = r.Width - pad - u;
            float sparkW = c.Spark && t.HasDate ? inner * 0.38f : 0f;
            RectangleF lr = new RectangleF(r.X + pad, r.Y + rows[0] * u, inner, 2f * u);
            RectangleF vr = new RectangleF(r.X + pad, r.Y + rows[1] * u, inner - sparkW, 3f * u);
            RectangleF nr = new RectangleF(r.X + pad, r.Y + rows[2] * u, inner, 2f * u);
            Text(g, DashText.CardLabel(c).ToUpperInvariant(), th.Font, Px(DashLook.CardLabelPt, u), true, C(k.Label), lr, StringAlignment.Near, StringAlignment.Center);
            string valueText = CardValueText(t, s, cache, c);
            double valuePt = DashLook.FitFont(valueText, vr.Width / u * DashLayout.CellPt, r.Width / u < 16f ? DashLook.CardValueSmallPt : DashLook.CardValuePt);
            Text(g, valueText, th.Font, Px(valuePt, u), true, C(k.Value), vr, StringAlignment.Near, StringAlignment.Center);
            int sign;
            string note = CardNoteText(t, s, cache, c, out sign);
            int noteColor = k.Note;
            if (k.Plain && sign != 0) noteColor = sign > 0 ? DashLook.UpColor(s) : DashLook.DownColor(s);
            Text(g, note, th.Font, Px(DashLook.CardNotePt, u), false, C(noteColor), nr, StringAlignment.Near, StringAlignment.Near);
            if (sparkW > 0f)
            {
                AggResult m = cache.Months(t, c.Col, c.Agg, s);
                RectangleF sr = new RectangleF(vr.Right + u * 0.3f, vr.Y + u * 0.3f, sparkW - u * 0.3f, vr.Height - u * 0.6f);
                Spark(g, sr, m.First, C(k.Plain ? k.Line : k.Value), Math.Max(1f, u * 0.14f));
            }
        }

        static void Spark(Graphics g, RectangleF r, double[] v, Color c, float width)
        {
            int n = v.Length;
            if (n < 2 || r.Width < 4f || r.Height < 3f) return;
            double mn = v[0], mx = v[0];
            foreach (double x in v)
            {
                if (x < mn) mn = x;
                if (x > mx) mx = x;
            }
            double span = mx - mn;
            PointF[] pts = new PointF[n];
            for (int i = 0; i < n; i++)
            {
                float x = r.X + r.Width * i / (n - 1);
                float y = span > 0 ? r.Bottom - (float)((v[i] - mn) / span) * r.Height : r.Y + r.Height / 2f;
                pts[i] = new PointF(x, y);
            }
            using (Pen pen = new Pen(c, width))
            {
                pen.LineJoin = LineJoin.Round;
                g.DrawLines(pen, pts);
            }
        }

        // ------------------------------------------------------------------ filter
        // Every button is drawn switched on, the way a slicer looks before anything is filtered.
        static void Slicer(Graphics g, RectangleF r, DataTab t, DashSpec s, string field, float u, float rad)
        {
            ThemeSpec th = s.Theme;
            FillBox(g, C(th.Card), r, 0f);
            StrokeBox(g, C(DashLook.SlicerEdge(s)), r, 0f, 1f);
            float pad = u * 0.6f;
            RectangleF hr = new RectangleF(r.X + pad, r.Y + u * 0.2f, r.Width - 2f * pad, u * 2f);
            string caption = DatePart.Is(field) ? DatePart.Label(field) : field;
            Text(g, caption, th.Font, Px(9, u), true, C(th.Text), hr, StringAlignment.Near, StringAlignment.Center);
            List<string> labels;
            int[] key = t.Keys(field, s.FyStart, out labels);
            if (key == null) return;
            List<string> shown = new List<string>(labels);
            shown.Sort(StringComparer.OrdinalIgnoreCase);
            float top = r.Y + u * 2.4f;
            float bh = u * 1.75f, gap = u * 0.3f;
            Color tint = C(DashLook.SlicerOn(s)), ink = C(DashLook.SlicerOnText(s));
            if (r.Width > r.Height * 2.5f)
            {
                // strip on top: buttons side by side
                int cols = DashLook.SlicerColumns(shown.Count, r.Width / u);
                float bw = (r.Width - 2f * pad - (cols - 1) * gap) / cols;
                for (int i = 0; i < shown.Count; i++)
                {
                    int row = i / cols, col = i % cols;
                    float y = top + row * (bh + gap);
                    if (y + bh > r.Bottom - u * 0.2f) break;
                    RectangleF b = new RectangleF(r.X + pad + col * (bw + gap), y, bw, bh);
                    FillBox(g, tint, b, u * 0.2f);
                    Text(g, shown[i], th.Font, Px(8, u), false, ink, new RectangleF(b.X + u * 0.4f, b.Y, b.Width - u * 0.6f, b.Height), StringAlignment.Near, StringAlignment.Center);
                }
                return;
            }
            for (int i = 0; i < shown.Count; i++)
            {
                float y = top + i * (bh + gap);
                if (y + bh > r.Bottom - u * 0.3f) break;
                RectangleF b = new RectangleF(r.X + pad, y, r.Width - 2f * pad, bh);
                FillBox(g, tint, b, u * 0.2f);
                Text(g, shown[i], th.Font, Px(8, u), false, ink, new RectangleF(b.X + u * 0.5f, b.Y, b.Width - u * 0.8f, b.Height), StringAlignment.Near, StringAlignment.Center);
            }
        }

        static void Insights(Graphics g, RectangleF r, DashSpec s, List<string> lines, float u, float rad)
        {
            ThemeSpec th = s.Theme;
            Panel(g, r, s, u, rad);
            RectangleF hr = new RectangleF(r.X + u, r.Y + u * 0.6f, r.Width - 2f * u, u * 2f);
            Text(g, "KEY INSIGHTS", th.Font, Px(DashLook.InsightHeadPt, u), true, C(th.Muted), hr, StringAlignment.Near, StringAlignment.Center);
            for (int i = 0; i < lines.Count; i++)
            {
                RectangleF lr = new RectangleF(r.X + u, r.Y + u * (2.8f + 2f * i), r.Width - 2f * u, u * 2f);
                Text(g, "\u2022  " + lines[i], th.Font, Px(DashLook.InsightPt, u), false, C(th.Text), lr, StringAlignment.Near, StringAlignment.Center);
            }
        }

        // the card behind a chart, a filter or the insights
        static void Panel(Graphics g, RectangleF r, DashSpec s, float u, float rad)
        {
            ThemeSpec th = s.Theme;
            if (s.Shadow && s.Rounded)
            {
                RectangleF sh = new RectangleF(r.X, r.Y + u * 0.18f, r.Width, r.Height);
                FillBox(g, Color.FromArgb(th.Dark ? 70 : 22, 0, 0, 0), sh, rad);
            }
            FillBox(g, C(th.Card), r, rad);
            if (!s.Shadow && ColorMath.Lum(th.Card) > 0.93 && ColorMath.Lum(th.Page) > 0.93) StrokeBox(g, C(th.Grid), r, rad, 1f);
        }

        // ------------------------------------------------------------------ chart
        static void Chart(Graphics g, RectangleF r, DataTab t, DashSpec s, DashCache cache, int index, float u, float rad)
        {
            ChartSpec c = s.Charts[index];
            ThemeSpec th = s.Theme;
            AggResult a = cache.Chart(t, c, s);
            Panel(g, r, s, u, rad);
            float pad = u * 0.9f;
            RectangleF tr = new RectangleF(r.X + pad, r.Y + u * 0.5f, r.Width - 2f * pad, u * 2f);
            Text(g, DashText.ChartTitle(c, a), th.Font, Px(DashLook.ChartTitlePt, u), true, C(th.Text), tr, StringAlignment.Near, StringAlignment.Center);
            RectangleF body = new RectangleF(r.X + pad, tr.Bottom + u * 0.3f, r.Width - 2f * pad, r.Bottom - tr.Bottom - u * 1.1f);
            if (a.Cats.Count == 0 || a.S.Count == 0 || body.Width < 8f || body.Height < 8f)
            {
                Text(g, "Nothing to show", th.Font, Px(9, u), false, C(th.Muted), body, StringAlignment.Center, StringAlignment.Center);
                return;
            }
            if (c.Kind == ChartKind.Table)
            {
                Table(g, body, s, c, a, index, u);
                return;
            }
            int legend = DashLook.LegendPos(c, ChartKind.IsRound(c.Kind) ? a.Cats.Count : a.S.Count);
            List<string> names = new List<string>();
            List<int> colors = new List<int>();
            if (ChartKind.IsRound(c.Kind))
            {
                for (int i = 0; i < a.Cats.Count; i++)
                {
                    names.Add(a.Cats[i]);
                    colors.Add(a.Others && i == a.Cats.Count - 1 ? DashLook.OthersColor(s) : DashLook.PointColor(s, i));
                }
            }
            else if (a.S.Count > 1)
            {
                for (int j = 0; j < a.S.Count; j++)
                {
                    names.Add(a.S[j].Name);
                    colors.Add(a.SplitOthers && j == a.S.Count - 1 ? DashLook.OthersColor(s) : DashLook.SeriesColor(s, c, index, j));
                }
            }
            else
            {
                names.Add(a.S[0].Name);
                colors.Add(DashLook.ChartColor(s, c, index));
            }
            if (legend == 2 && body.Width > u * 22f)
            {
                float lw = Math.Min(body.Width * 0.42f, u * 16f);
                Legend(g, new RectangleF(body.Right - lw, body.Y, lw, body.Height), s, names, colors, u, true);
                body = new RectangleF(body.X, body.Y, body.Width - lw - u * 0.5f, body.Height);
            }
            else if (legend == 3 || legend == 2)
            {
                Legend(g, new RectangleF(body.X, body.Bottom - u * 1.6f, body.Width, u * 1.6f), s, names, colors, u, false);
                body = new RectangleF(body.X, body.Y, body.Width, body.Height - u * 2f);
            }
            else if (legend == 4)
            {
                Legend(g, new RectangleF(body.X, body.Y, body.Width, u * 1.6f), s, names, colors, u, false);
                body = new RectangleF(body.X, body.Y + u * 2f, body.Width, body.Height - u * 2f);
            }
            if (ChartKind.IsRound(c.Kind)) Round(g, body, s, c, a, colors, u);
            else if (c.Kind == ChartKind.Bar || c.Kind == ChartKind.StackedBar) Bars(g, body, s, c, a, colors, u);
            else Columns(g, body, s, c, a, colors, u);
        }

        static void Legend(Graphics g, RectangleF r, DashSpec s, List<string> names, List<int> colors, float u, bool vertical)
        {
            ThemeSpec th = s.Theme;
            float px = Px(DashLook.LegendPt, u);
            float box = u * 0.8f;
            if (vertical)
            {
                float lh = u * 1.7f;
                float y = r.Y + Math.Max(0f, (r.Height - lh * names.Count) / 2f);
                for (int i = 0; i < names.Count; i++)
                {
                    if (y + lh > r.Bottom + 1f) break;
                    FillBox(g, C(colors[i]), new RectangleF(r.X, y + (lh - box) / 2f, box, box), box * 0.25f);
                    Text(g, names[i], th.Font, px, false, C(th.Muted), new RectangleF(r.X + box + u * 0.5f, y, r.Width - box - u * 0.5f, lh), StringAlignment.Near, StringAlignment.Center);
                    y += lh;
                }
                return;
            }
            float total = 0f;
            float[] w = new float[names.Count];
            for (int i = 0; i < names.Count; i++)
            {
                w[i] = Math.Min(u * 12f, TextWidth(g, names[i], th.Font, px, false) + box + u * 1.4f);
                total += w[i];
            }
            if (total > r.Width && total > 0f)
            {
                // too many names for one line: every name gets less room
                float k = r.Width / total;
                for (int i = 0; i < w.Length; i++) w[i] *= k;
                total = r.Width;
            }
            float x0 = total < r.Width ? r.X + (r.Width - total) / 2f : r.X;
            for (int i = 0; i < names.Count; i++)
            {
                if (x0 + w[i] > r.Right + 1f) break;
                FillBox(g, C(colors[i]), new RectangleF(x0, r.Y + (r.Height - box) / 2f, box, box), box * 0.25f);
                Text(g, names[i], th.Font, px, false, C(th.Muted), new RectangleF(x0 + box + u * 0.4f, r.Y, w[i] - box - u * 0.6f, r.Height), StringAlignment.Near, StringAlignment.Center);
                x0 += w[i];
            }
        }

        static void Round(Graphics g, RectangleF r, DashSpec s, ChartSpec c, AggResult a, List<int> colors, float u)
        {
            ThemeSpec th = s.Theme;
            double[] v = a.First;
            double tot = 0;
            foreach (double x in v) tot += Math.Abs(x);
            float d = Math.Min(r.Width, r.Height) - u * 0.4f;
            if (d < 6f || tot <= 0) return;
            float cx = r.X + r.Width / 2f, cy = r.Y + r.Height / 2f;
            float start = -90f;
            int mode = DashLook.LabelMode(c);
            for (int i = 0; i < v.Length; i++)
            {
                float sweep = (float)(Math.Abs(v[i]) / tot * 360.0);
                if (sweep <= 0f) continue;
                using (SolidBrush b = new SolidBrush(C(colors[i % colors.Count])))
                {
                    g.FillPie(b, cx - d / 2f, cy - d / 2f, d, d, start, sweep);
                }
                start += sweep;
            }
            // thin gaps between the slices
            using (Pen gapPen = new Pen(C(th.Card), Math.Max(1f, u * 0.12f)))
            {
                float ang = -90f;
                for (int i = 0; i < v.Length; i++)
                {
                    double rr = ang * Math.PI / 180.0;
                    g.DrawLine(gapPen, cx, cy, cx + (float)(Math.Cos(rr) * d / 2.0), cy + (float)(Math.Sin(rr) * d / 2.0));
                    ang += (float)(Math.Abs(v[i]) / tot * 360.0);
                }
            }
            if (c.Kind == ChartKind.Doughnut)
            {
                float hole = d * 0.58f;
                using (SolidBrush b = new SolidBrush(C(th.Card)))
                {
                    g.FillEllipse(b, cx - hole / 2f, cy - hole / 2f, hole, hole);
                }
                RectangleF mid = new RectangleF(cx - hole / 2f, cy - hole / 4f, hole, hole / 2f);
                Text(g, NumFmt.Compact(tot, s.Unit, false), th.Font, Math.Min(Px(12, u), hole * 0.2f), true, C(th.Text), mid, StringAlignment.Center, StringAlignment.Center);
            }
            if (mode == 1) return;
            // labels on the slices that are large enough
            float ang2 = -90f;
            float lr = c.Kind == ChartKind.Doughnut ? d * 0.395f : d * 0.32f;
            for (int i = 0; i < v.Length; i++)
            {
                float sweep = (float)(Math.Abs(v[i]) / tot * 360.0);
                if (sweep >= 22f)
                {
                    double rr = (ang2 + sweep / 2f) * Math.PI / 180.0;
                    float lx = cx + (float)(Math.Cos(rr) * lr), ly = cy + (float)(Math.Sin(rr) * lr);
                    string txt = mode == 3 ? (Math.Abs(v[i]) / tot).ToString("0%", Inv) : NumFmt.Compact(v[i], s.Unit, false);
                    RectangleF tr = new RectangleF(lx - u * 3f, ly - u, u * 6f, u * 2f);
                    Text(g, txt, th.Font, Px(8, u), true, C(ColorMath.Ink(colors[i % colors.Count])), tr, StringAlignment.Center, StringAlignment.Center);
                }
                ang2 += sweep;
            }
        }

        // data labels: every k-th point gets one, so that they do not run into each other
        static int LabelEvery(Graphics g, double[] v, DashSpec s, bool percent, float px, float slot)
        {
            float widest = 0f;
            int probe = Math.Min(v.Length, 24);
            for (int i = 0; i < probe; i++) widest = Math.Max(widest, TextWidth(g, NumFmt.Compact(v[i], s.Unit, percent), s.Theme.Font, px, false));
            if (slot <= 0f) return 1;
            return Math.Max(1, (int)Math.Ceiling(widest / slot));
        }

        // the value axis for some of the series of a chart (stacked or side by side)
        static AxisScale AxisFor(ChartSpec c, AggResult a, int firstSeries, int lastSeries)
        {
            double mx = 0, mn = 0;
            bool stacked = ChartKind.IsStacked(c.Kind);
            int n = a.Cats.Count;
            for (int i = 0; i < n; i++)
            {
                double pos = 0, neg = 0;
                for (int j = firstSeries; j <= lastSeries; j++)
                {
                    double x = a.S[j].V[i];
                    if (stacked)
                    {
                        if (x >= 0) pos += x;
                        else neg += x;
                    }
                    else
                    {
                        if (x > pos) pos = x;
                        if (x < neg) neg = x;
                    }
                }
                if (pos > mx) mx = pos;
                if (neg < mn) mn = neg;
            }
            return DashLook.Axis(mn, mx, a.S[firstSeries].Percent);
        }

        // the widest number on an axis, in pixels
        static float AxisTextWidth(Graphics g, AxisScale ax, DashSpec s, bool percent, float px)
        {
            float w = 0f;
            for (int i = 0; i <= ax.Lines; i++)
                w = Math.Max(w, TextWidth(g, NumFmt.Compact(ax.Min + ax.Step * i, s.Unit, percent), s.Theme.Font, px, false));
            return w;
        }

        static void Columns(Graphics g, RectangleF r, DashSpec s, ChartSpec c, AggResult a, List<int> colors, float u)
        {
            ThemeSpec th = s.Theme;
            int n = a.Cats.Count;
            bool combo = c.Kind == ChartKind.Combo || c.Kind == ChartKind.Pareto;
            bool line = ChartKind.IsLine(c.Kind);
            int barSeries = combo ? 1 : a.S.Count;
            bool percent = a.S[0].Percent;
            AxisScale ax = AxisFor(c, a, 0, barSeries - 1);
            float axisPx = Px(DashLook.AxisPt, u);
            float gutter = AxisTextWidth(g, ax, s, percent, axisPx) + u * 0.7f;
            float right = 0f;
            AxisScale ax2 = null;
            if (combo && a.S.Count > 1)
            {
                ax2 = AxisFor(c, a, 1, 1);
                right = AxisTextWidth(g, ax2, s, a.S[1].Percent, axisPx) + u * 0.7f;
            }
            float labelH = u * 1.9f;
            RectangleF plot = new RectangleF(r.X + gutter, r.Y + u * 0.8f, r.Width - gutter - right - u * 0.2f, r.Height - labelH - u * 0.8f);
            if (plot.Width < 6f || plot.Height < 6f) return;
            // grid lines and axis numbers
            using (Pen gp = new Pen(C(th.Grid), 1f))
            {
                for (int i = 0; i <= ax.Lines; i++)
                {
                    float y = plot.Bottom - plot.Height * i / ax.Lines;
                    g.DrawLine(gp, plot.X, y, plot.Right, y);
                    Text(g, NumFmt.Compact(ax.Min + ax.Step * i, s.Unit, percent), th.Font, axisPx, false, C(th.Muted), new RectangleF(r.X, y - u, gutter - u * 0.4f, u * 2f), StringAlignment.Far, StringAlignment.Center);
                }
                if (ax2 != null)
                {
                    for (int i = 0; i <= ax2.Lines; i++)
                    {
                        float y = plot.Bottom - plot.Height * i / ax2.Lines;
                        Text(g, NumFmt.Compact(ax2.Min + ax2.Step * i, s.Unit, a.S[1].Percent), th.Font, axisPx, false, C(th.Muted), new RectangleF(plot.Right + u * 0.4f, y - u, right, u * 2f), StringAlignment.Near, StringAlignment.Center);
                    }
                }
            }
            float slot = plot.Width / n;
            float zero = plot.Bottom - (float)ax.Where(0) * plot.Height;
            int mode = DashLook.LabelMode(c);
            if (line)
            {
                PointF[] pts = new PointF[n];
                for (int i = 0; i < n; i++)
                {
                    float x = plot.X + slot * (i + 0.5f);
                    float y = plot.Bottom - (float)ax.Where(a.S[0].V[i]) * plot.Height;
                    pts[i] = new PointF(x, y);
                }
                Color lc = C(colors[0]);
                if (c.Kind == ChartKind.Area && n > 1)
                {
                    PointF[] poly = new PointF[n + 2];
                    for (int i = 0; i < n; i++) poly[i] = pts[i];
                    poly[n] = new PointF(pts[n - 1].X, zero);
                    poly[n + 1] = new PointF(pts[0].X, zero);
                    using (SolidBrush fb = new SolidBrush(Color.FromArgb(70, lc)))
                    {
                        g.FillPolygon(fb, poly);
                    }
                }
                using (Pen lp = new Pen(lc, Math.Max(1.5f, u * 0.2f)))
                {
                    lp.LineJoin = LineJoin.Round;
                    if (n > 1)
                    {
                        if (c.Smooth && n > 2) g.DrawCurve(lp, pts, 0.4f);
                        else g.DrawLines(lp, pts);
                    }
                }
                if (c.Trend > 0 && n > 2) Trend(g, plot, a.S[0].V, ax, c.Trend, C(th.Muted), u);
                float dot = Math.Max(2f, u * 0.5f);
                if (slot > dot * 1.6f && c.Kind != ChartKind.Area)
                {
                    using (SolidBrush db = new SolidBrush(lc))
                    {
                        for (int i = 0; i < n; i++) g.FillEllipse(db, pts[i].X - dot / 2f, pts[i].Y - dot / 2f, dot, dot);
                    }
                }
                if (mode == 2)
                {
                    int every = LabelEvery(g, a.S[0].V, s, percent, Px(7.5, u), slot);
                    for (int i = 0; i < n; i += every)
                    {
                        Text(g, NumFmt.Compact(a.S[0].V[i], s.Unit, percent), th.Font, Px(7.5, u), false, C(th.Text), new RectangleF(pts[i].X - u * 4f, pts[i].Y - u * 1.9f, u * 8f, u * 1.5f), StringAlignment.Center, StringAlignment.Center);
                    }
                }
            }
            else
            {
                bool stacked = ChartKind.IsStacked(c.Kind);
                float groupW = slot * 0.66f;
                float bw = stacked || combo ? groupW : groupW / barSeries;
                int every = mode == 2 && barSeries == 1 ? LabelEvery(g, a.S[0].V, s, percent, Px(7.5, u), slot) : 1;
                for (int i = 0; i < n; i++)
                {
                    float x0 = plot.X + slot * i + (slot - groupW) / 2f;
                    double up = 0, down = 0;
                    for (int j = 0; j < barSeries; j++)
                    {
                        double val = a.S[j].V[i];
                        double start0 = 0;
                        if (stacked)
                        {
                            start0 = val >= 0 ? up : down;
                            if (val >= 0) up += val;
                            else down += val;
                        }
                        float y1 = plot.Bottom - (float)ax.Where(start0 + val) * plot.Height;
                        float y0 = plot.Bottom - (float)ax.Where(start0) * plot.Height;
                        float top = Math.Min(y0, y1), h = Math.Abs(y1 - y0);
                        float x = stacked || combo ? x0 : x0 + bw * j;
                        int col = barSeries > 1 ? colors[j % colors.Count] : colors[0];
                        if (h >= 0.5f)
                        {
                            using (SolidBrush bb = new SolidBrush(C(col)))
                            {
                                g.FillRectangle(bb, x, top, Math.Max(1f, bw - (stacked || combo ? 0f : 1f)), h);
                            }
                        }
                        if (mode == 2 && barSeries == 1 && i % every == 0)
                            Text(g, NumFmt.Compact(val, s.Unit, percent), th.Font, Px(7.5, u), false, C(th.Text), new RectangleF(x + bw / 2f - u * 4f, (val >= 0 ? top : top + h + u * 1.7f) - u * 1.6f, u * 8f, u * 1.5f), StringAlignment.Center, StringAlignment.Center);
                    }
                }
                if (ax2 != null)
                {
                    PointF[] pts = new PointF[n];
                    for (int i = 0; i < n; i++)
                    {
                        float x = plot.X + slot * (i + 0.5f);
                        float y = plot.Bottom - (float)ax2.Where(a.S[1].V[i]) * plot.Height;
                        pts[i] = new PointF(x, y);
                    }
                    Color lc = C(colors.Count > 1 ? colors[1] : th.Text);
                    using (Pen lp = new Pen(lc, Math.Max(1.5f, u * 0.2f)))
                    {
                        lp.LineJoin = LineJoin.Round;
                        if (n > 1) g.DrawLines(lp, pts);
                    }
                    float dot = Math.Max(2f, u * 0.5f);
                    using (SolidBrush db = new SolidBrush(lc))
                    {
                        for (int i = 0; i < n; i++) g.FillEllipse(db, pts[i].X - dot / 2f, pts[i].Y - dot / 2f, dot, dot);
                    }
                }
            }
            if (ax.Min < 0)
            {
                using (Pen zp = new Pen(C(th.Muted), 1f))
                {
                    g.DrawLine(zp, plot.X, zero, plot.Right, zero);
                }
            }
            // category names: every k-th one when they would run into each other
            float widest = 0f;
            int probe = Math.Min(n, 12);
            for (int i = 0; i < probe; i++) widest = Math.Max(widest, TextWidth(g, AxisName(c, a.Cats[i]), th.Font, axisPx, false));
            if (widest > u * 8f) widest = u * 8f;
            int step = Math.Max(1, (int)Math.Ceiling((widest + u * 0.6f) / slot));
            for (int i = 0; i < n; i += step)
            {
                float w = slot * step;
                RectangleF lr = new RectangleF(plot.X + slot * i + slot / 2f - w / 2f, plot.Bottom + u * 0.3f, w, labelH - u * 0.3f);
                Text(g, AxisName(c, a.Cats[i]), th.Font, axisPx, false, C(th.Muted), lr, StringAlignment.Center, StringAlignment.Near);
            }
        }

        // the name of a group under a chart: date groups get their short names
        public static string AxisName(ChartSpec c, string label)
        {
            return DatePart.Is(c.Dim) ? DataTab.ShortLabel(c.Dim, label) : label;
        }

        // a straight trend line (least squares) or a moving average of three
        static void Trend(Graphics g, RectangleF plot, double[] v, AxisScale ax, int kind, Color c, float u)
        {
            int n = v.Length;
            double[] y = new double[n];
            if (kind == 1)
            {
                double sx = 0, sy = 0, sxx = 0, sxy = 0;
                for (int i = 0; i < n; i++)
                {
                    sx += i;
                    sy += v[i];
                    sxx += (double)i * i;
                    sxy += i * v[i];
                }
                double den = n * sxx - sx * sx;
                double slope = den == 0 ? 0 : (n * sxy - sx * sy) / den;
                double icpt = (sy - slope * sx) / n;
                for (int i = 0; i < n; i++) y[i] = icpt + slope * i;
            }
            else
            {
                for (int i = 0; i < n; i++)
                {
                    double sum = 0;
                    int cnt = 0;
                    for (int k = i - 2; k <= i; k++)
                    {
                        if (k < 0) continue;
                        sum += v[k];
                        cnt++;
                    }
                    y[i] = sum / cnt;
                }
            }
            float slot = plot.Width / n;
            PointF[] pts = new PointF[n];
            for (int i = 0; i < n; i++)
            {
                float py = plot.Bottom - (float)ax.Where(y[i]) * plot.Height;
                if (py < plot.Y) py = plot.Y;
                if (py > plot.Bottom) py = plot.Bottom;
                pts[i] = new PointF(plot.X + slot * (i + 0.5f), py);
            }
            using (Pen tp = new Pen(c, Math.Max(1f, u * 0.12f)))
            {
                tp.DashStyle = DashStyle.Dash;
                g.DrawLines(tp, pts);
            }
        }

        static void Bars(Graphics g, RectangleF r, DashSpec s, ChartSpec c, AggResult a, List<int> colors, float u)
        {
            ThemeSpec th = s.Theme;
            int n = a.Cats.Count;
            bool stacked = ChartKind.IsStacked(c.Kind);
            int nser = a.S.Count;
            bool percent = a.S[0].Percent;
            AxisScale ax = AxisFor(c, a, 0, nser - 1);
            float axisPx = Px(DashLook.AxisPt, u);
            float widest = 0f;
            for (int i = 0; i < n; i++) widest = Math.Max(widest, TextWidth(g, a.Cats[i], th.Font, axisPx, false));
            float gutter = Math.Min(widest + u * 0.8f, r.Width * 0.38f);
            int mode = DashLook.LabelMode(c);
            float numW = AxisTextWidth(g, ax, s, percent, axisPx);
            float valueW = mode == 2 && !stacked ? numW + u : numW / 2f + u * 0.3f;
            RectangleF plot = new RectangleF(r.X + gutter, r.Y + u * 0.2f, r.Width - gutter - valueW, r.Height - u * 1.9f);
            if (plot.Width < 6f || plot.Height < 6f) return;
            using (Pen gp = new Pen(C(th.Grid), 1f))
            {
                for (int i = 0; i <= ax.Lines; i++)
                {
                    float x = plot.X + plot.Width * i / ax.Lines;
                    g.DrawLine(gp, x, plot.Y, x, plot.Bottom);
                    Text(g, NumFmt.Compact(ax.Min + ax.Step * i, s.Unit, percent), th.Font, axisPx, false, C(th.Muted), new RectangleF(x - u * 4f, plot.Bottom + u * 0.2f, u * 8f, u * 1.6f), StringAlignment.Center, StringAlignment.Center);
                }
            }
            float slot = plot.Height / n;
            float bh = slot * 0.68f;
            float zero = plot.X + (float)ax.Where(0) * plot.Width;
            for (int i = 0; i < n; i++)
            {
                float y = plot.Y + slot * i + (slot - bh) / 2f;
                Text(g, a.Cats[i], th.Font, axisPx, false, C(th.Muted), new RectangleF(r.X, plot.Y + slot * i, gutter - u * 0.5f, slot), StringAlignment.Far, StringAlignment.Center);
                double up = 0, down = 0;
                for (int j = 0; j < nser; j++)
                {
                    double val = a.S[j].V[i];
                    double start0 = 0;
                    if (stacked)
                    {
                        start0 = val >= 0 ? up : down;
                        if (val >= 0) up += val;
                        else down += val;
                    }
                    float x0 = plot.X + (float)ax.Where(start0) * plot.Width;
                    float x1 = plot.X + (float)ax.Where(start0 + val) * plot.Width;
                    float left = Math.Min(x0, x1), w = Math.Abs(x1 - x0);
                    float h = stacked ? bh : bh / nser;
                    float yy = stacked ? y : y + h * j;
                    int col = nser > 1 ? colors[j % colors.Count] : colors[0];
                    if (w >= 0.5f)
                    {
                        using (SolidBrush bb = new SolidBrush(C(col)))
                        {
                            g.FillRectangle(bb, left, yy, w, Math.Max(1f, h - (stacked ? 0f : 0.5f)));
                        }
                    }
                    if (mode == 2 && nser == 1)
                        Text(g, NumFmt.Compact(val, s.Unit, percent), th.Font, axisPx, false, C(th.Text), new RectangleF(left + w + u * 0.3f, yy, valueW + u, h), StringAlignment.Near, StringAlignment.Center);
                }
            }
            if (ax.Min < 0)
            {
                using (Pen zp = new Pen(C(th.Muted), 1f))
                {
                    g.DrawLine(zp, zero, plot.Y, zero, plot.Bottom);
                }
            }
        }

        static void Table(Graphics g, RectangleF r, DashSpec s, ChartSpec c, AggResult a, int index, float u)
        {
            ThemeSpec th = s.Theme;
            int n = a.Cats.Count;
            float rowH = u * 2f;
            int rows = Math.Min(n, (int)Math.Floor((r.Height - rowH) / rowH));
            if (rows < 1) return;
            int nv = a.S.Count;                       // value columns, the last one is the share
            float nameW = r.Width * (nv >= 3 ? 0.34f : 0.42f);
            float colW = (r.Width - nameW) / nv;
            float px = Px(DashLook.TablePt, u);
            Text(g, DashText.DimName(c.Dim).ToUpperInvariant(), th.Font, Px(8, u), true, C(th.Muted), new RectangleF(r.X, r.Y, nameW, rowH), StringAlignment.Near, StringAlignment.Center);
            for (int j = 0; j < nv; j++)
                Text(g, a.S[j].Name.ToUpperInvariant(), th.Font, Px(8, u), true, C(th.Muted), new RectangleF(r.X + nameW + colW * j, r.Y, colW - u * 0.3f, rowH), StringAlignment.Far, StringAlignment.Center);
            double mx = 0;
            for (int i = 0; i < rows; i++) if (Math.Abs(a.S[0].V[i]) > mx) mx = Math.Abs(a.S[0].V[i]);
            int accent = DashLook.ChartColor(s, c, index);
            using (Pen lp = new Pen(C(th.Grid), 1f))
            {
                for (int i = 0; i < rows; i++)
                {
                    float y = r.Y + rowH * (i + 1);
                    g.DrawLine(lp, r.X, y, r.Right, y);
                    Text(g, a.Cats[i], th.Font, px, false, C(th.Text), new RectangleF(r.X, y, nameW - u * 0.4f, rowH), StringAlignment.Near, StringAlignment.Center);
                    for (int j = 0; j < nv; j++)
                    {
                        RectangleF cell = new RectangleF(r.X + nameW + colW * j, y, colW - u * 0.3f, rowH);
                        double val = a.S[j].V[i];
                        if (j == 0 && mx > 0)
                        {
                            float bw = (float)(Math.Abs(val) / mx) * (colW - u * 0.6f);
                            FillBox(g, C(ColorMath.Mix(th.Card, accent, th.Dark ? 0.5 : 0.28)), new RectangleF(cell.X + u * 0.3f, y + rowH * 0.2f, bw, rowH * 0.6f), u * 0.15f);
                        }
                        Text(g, NumFmt.Compact(val, s.Unit, a.S[j].Percent), th.Font, px, j == 0, C(th.Text), cell, StringAlignment.Far, StringAlignment.Center);
                    }
                }
            }
        }
    }
}
