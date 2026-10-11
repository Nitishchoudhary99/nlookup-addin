using System;
using System.Collections.Generic;

namespace SmartPro
{
    // =====================================================================
    //  WHERE EVERYTHING GOES
    //  The dashboard sheet is a grid of small square cells (12 x 12 pixels). Every card, chart and filter is a box
    //  of whole cells, with one cell between two boxes. The same boxes are used for the sheet in Excel and for the
    //  preview in Dashboard Studio, so the preview shows what will be built. No Excel calls in this file.
    // =====================================================================
    internal sealed class Box
    {
        public const int Header = 0, Card = 1, Chart = 2, Slicer = 3, Insights = 4, Footer = 5;
        public int Kind, Index;
        public int X, Y, W, H;            // in grid cells from the top-left corner of the sheet; X and Y start at 0

        public Box(int kind, int index, int x, int y, int w, int h)
        {
            Kind = kind;
            Index = index;
            X = x;
            Y = y;
            W = w;
            H = h;
        }
    }

    internal sealed class DashGrid
    {
        public int Cols, Rows;            // size of the whole canvas in cells, margins included
        public int Margin, Width;         // left margin and width of the content
        public int HeaderRows, CardRows;
        public List<Box> Boxes = new List<Box>();

        public Box Find(int kind, int index)
        {
            foreach (Box b in Boxes) if (b.Kind == kind && b.Index == index) return b;
            return null;
        }
    }

    internal static class DashLayout
    {
        public const int CellPx = 12;             // one grid cell on the sheet is 12 pixels wide and high
        public const double CellPt = 9.0;         // which is 9 points
        public const int Panel = 14;              // width of the filter panel in cells

        public static readonly string[] ArrangeNames =
        {
            "Balanced grid (two charts in a row)",
            "Trend focus (one wide chart, the others below)",
            "Story (wide charts under each other)",
            "Three across",
            "Feature (one tall chart, two beside it)",
            "Mosaic (wide and narrow charts mixed)",
            "Compact (three short charts in a row)",
            "Poster (one large chart, small ones below)"
        };
        public static readonly string[] HeightNames = { "Automatic", "Short", "Standard", "Tall", "Extra tall" };
        public static readonly string[] SpanNames = { "Automatic", "Quarter width", "Third", "Half", "Two thirds", "Three quarters", "Full width" };
        public static readonly int[] SpanValues = { 0, 3, 4, 6, 8, 9, 12 };
        public static readonly string[] SlicerPosNames = { "Panel on the left", "Strip on top", "Panel on the right", "No filters" };
        public static readonly string[] DensityNames = { "Compact", "Comfortable", "Spacious" };
        public static readonly string[] WidthNames = { "Standard (fits a laptop screen)", "Wide (large monitor)" };

        static readonly int[] HeightUnits = { 4, 3, 4, 6, 8 };       // chart height classes in units
        const int MaxUnitRows = 12;               // a unit of chart height never gets taller than this many rows

        public static int ContentWidth(DashSpec s) { return s.Width == 1 ? 132 : 96; }

        // height of a filter in the side panel, in cells: a header and up to 8 buttons
        public static int SlicerCells(int items)
        {
            int k = items < 1 ? 1 : (items > 8 ? 8 : items);
            return (int)Math.Ceiling((34 + 20 * k) / CellPt);
        }

        // the automatic size of chart i of n: width in twelfths * 100 + height class
        public static int AutoSize(int arrange, int n, int i)
        {
            int span = 6, h = 2;
            int m = n - 1;
            if (n <= 1) return 1200 + 2;
            switch (arrange)
            {
                case 1:
                    if (i == 0) span = 12;
                    else span = m == 1 ? 12 : (m == 2 || m == 4 ? 6 : 4);
                    break;
                case 2:
                    span = 12;
                    break;
                case 3:
                    span = n == 2 ? 6 : 4;
                    break;
                case 4:
                    if (n >= 3 && i == 0) h = 4;
                    span = 6;
                    break;
                case 5:
                    if (n % 2 == 1 && i == n - 1) span = 12;
                    else
                    {
                        int p = i % 6;
                        span = p == 0 || p == 3 ? 8 : (p == 1 || p == 2 ? 4 : 6);
                    }
                    break;
                case 6:
                    span = n == 2 || n == 4 ? 6 : 4;
                    h = 1;
                    break;
                case 7:
                    if (i == 0)
                    {
                        span = 12;
                        h = 3;
                    }
                    else span = m == 1 ? 12 : (m == 2 || m == 4 ? 6 : 4);
                    break;
                default:
                    span = n % 2 == 1 && i == 0 ? 12 : 6;
                    break;
            }
            return span * 100 + h;
        }

        static bool IsFree(List<bool[]> occ, int row, int col, int span, int units)
        {
            if (col + span > 12) return false;
            for (int r = row; r < row + units; r++)
            {
                if (r >= occ.Count) continue;
                for (int c = col; c < col + span; c++) if (occ[r][c]) return false;
            }
            return true;
        }

        static void Mark(List<bool[]> occ, int row, int col, int span, int units)
        {
            while (occ.Count < row + units) occ.Add(new bool[12]);
            for (int r = row; r < row + units; r++)
                for (int c = col; c < col + span; c++) occ[r][c] = true;
        }

        // slicerItems: how many buttons each filter has. insightLines: how many lines of key insights.
        public static DashGrid Compute(DashSpec s, int[] slicerItems, int insightLines)
        {
            DashGrid g = new DashGrid();
            int M = 2, W = ContentWidth(s);
            int dens = s.Density < 0 ? 0 : (s.Density > 2 ? 2 : s.Density);
            int headerH = 5 + dens;
            int cardH = 7 + dens;
            int pitch = 5 + dens;                 // rows of one unit of chart height, the gap included
            g.Margin = M;
            g.Width = W;
            g.Cols = M + W + M;
            g.HeaderRows = headerH;
            g.CardRows = cardH;
            g.Boxes.Add(new Box(Box.Header, 0, 0, 0, g.Cols, headerH));
            int y = headerH + 1;

            // ---- KPI cards: one or two rows across the whole width ----
            int nCards = s.Cards.Count;
            if (nCards > 0)
            {
                int most = W >= 120 ? 8 : 6;
                int per = s.CardsPerRow > 0 ? s.CardsPerRow : (nCards <= most ? nCards : (nCards + 1) / 2);
                if (per > (W + 1) / 13) per = (W + 1) / 13;
                if (per < 1) per = 1;
                int idx = 0;
                while (idx < nCards)
                {
                    int k = Math.Min(per, nCards - idx);
                    for (int i = 0; i < k; i++)
                    {
                        int x0 = M + (int)Math.Round(i * (W + 1) / (double)k);
                        int x1 = M + (int)Math.Round((i + 1) * (W + 1) / (double)k) - 1;
                        g.Boxes.Add(new Box(Box.Card, idx + i, x0, y, x1 - x0, cardH));
                    }
                    idx += k;
                    y += cardH + 1;
                }
            }

            // ---- filters ----
            int ns = slicerItems == null ? 0 : slicerItems.Length;
            int pos = s.SlicerPos;
            if (pos == 3) ns = 0;
            int cx = M, cw = W;
            int panelBottom = y - 1;
            if (ns > 0 && pos == 1)
            {
                int stripH = 7;
                for (int i = 0; i < ns; i++)
                {
                    int x0 = M + (int)Math.Round(i * (W + 1) / (double)ns);
                    int x1 = M + (int)Math.Round((i + 1) * (W + 1) / (double)ns) - 1;
                    g.Boxes.Add(new Box(Box.Slicer, i, x0, y, x1 - x0, stripH));
                }
                y += stripH + 1;
                panelBottom = y - 1;
            }
            else if (ns > 0)
            {
                int px = pos == 2 ? M + W - Panel : M;
                if (pos != 2) cx = M + Panel + 1;
                cw = W - Panel - 1;
                int sy = y;
                for (int i = 0; i < ns; i++)
                {
                    int h = SlicerCells(slicerItems[i]);
                    g.Boxes.Add(new Box(Box.Slicer, i, px, sy, Panel, h));
                    sy += h + 1;
                }
                panelBottom = sy - 2;
            }

            // ---- charts: twelve columns of units, each chart takes the first free place that fits ----
            int nCharts = s.Charts.Count;
            List<bool[]> occ = new List<bool[]>();
            int[] uRow = new int[nCharts], uCol = new int[nCharts], uSpan = new int[nCharts], uH = new int[nCharts];
            for (int i = 0; i < nCharts; i++)
            {
                ChartSpec c = s.Charts[i];
                int auto = AutoSize(s.Arrange, nCharts, i);
                int span = c.Span > 0 ? c.Span : auto / 100;
                int hc = c.Height > 0 ? c.Height : auto % 100;
                if (span < 3) span = 3;
                if (span > 12) span = 12;
                if (hc < 1 || hc > 4) hc = 2;
                int hUnits = HeightUnits[hc];
                bool placed = false;
                for (int r = 0; r < 400 && !placed; r++)
                {
                    for (int col = 0; col + span <= 12 && !placed; col++)
                    {
                        if (!IsFree(occ, r, col, span, hUnits)) continue;
                        Mark(occ, r, col, span, hUnits);
                        uRow[i] = r;
                        uCol[i] = col;
                        uSpan[i] = span;
                        uH[i] = hUnits;
                        placed = true;
                    }
                }
            }
            // a chart without an own width grows into free space on its right, so no row ends in a hole
            for (int i = 0; i < nCharts; i++)
            {
                if (s.Charts[i].Span > 0) continue;
                while (uCol[i] + uSpan[i] < 12 && IsFree(occ, uRow[i], uCol[i] + uSpan[i], 1, uH[i]))
                {
                    Mark(occ, uRow[i], uCol[i] + uSpan[i], 1, uH[i]);
                    uSpan[i]++;
                }
            }
            // Beside a filter panel the charts grow to the height of the panel: the spare rows are shared out
            // over the rows of charts, so the charts and the panel end on the same line.
            bool side = ns > 0 && pos != 1;
            int units = occ.Count;
            int rowsEach = pitch, spare = 0;
            if (side && units > 0)
            {
                int target = panelBottom - y + 2;
                if (target >= units * MaxUnitRows) rowsEach = MaxUnitRows;
                else if (target > units * pitch)
                {
                    rowsEach = target / units;
                    spare = target % units;
                }
            }
            int[] unitTop = new int[units + 1];
            unitTop[0] = y;
            for (int k = 0; k < units; k++) unitTop[k + 1] = unitTop[k] + rowsEach + ((k + 1) * spare / units - k * spare / units);
            int chartsBottom = y - 1;
            for (int i = 0; i < nCharts; i++)
            {
                int x0 = cx + (int)Math.Round(uCol[i] * (cw + 1) / 12.0);
                int x1 = cx + (int)Math.Round((uCol[i] + uSpan[i]) * (cw + 1) / 12.0) - 1;
                int y0 = unitTop[uRow[i]];
                int h = unitTop[uRow[i] + uH[i]] - y0 - 1;
                g.Boxes.Add(new Box(Box.Chart, i, x0, y0, x1 - x0, h));
                if (y0 + h - 1 > chartsBottom) chartsBottom = y0 + h - 1;
            }
            if (side && nCharts > 0 && chartsBottom > panelBottom)
            {
                // the charts are taller: the last filter grows and shows more buttons
                Box last = g.Find(Box.Slicer, ns - 1);
                if (last != null) last.H += chartsBottom - panelBottom;
                panelBottom = chartsBottom;
            }

            // ---- key insights: under the charts when the filter panel is much taller, otherwise across the page ----
            if (insightLines > 0)
            {
                int h = 3 + 2 * insightLines;
                if (side && panelBottom - chartsBottom >= h + 1)
                {
                    int top = nCharts > 0 ? chartsBottom + 2 : y;
                    g.Boxes.Add(new Box(Box.Insights, 0, cx, top, cw, h));
                    chartsBottom = top + h - 1;
                    y = Math.Max(panelBottom, chartsBottom) + 2;
                }
                else
                {
                    y = Math.Max(panelBottom, chartsBottom) + 2;
                    g.Boxes.Add(new Box(Box.Insights, 0, M, y, W, h));
                    y += h + 1;
                }
            }
            else y = Math.Max(panelBottom, chartsBottom) + 2;

            // ---- the footer line ----
            g.Boxes.Add(new Box(Box.Footer, 0, M, y, W, 2));
            y += 2;
            g.Rows = y + 1;
            return g;
        }
    }
}
