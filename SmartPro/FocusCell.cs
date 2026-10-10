using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace SmartPro
{
    // FOCUS CELL: a soft cross over the row and the column of the active cell, so the eye never loses the line in a wide table.
    // It is a see-through window that lies over Excel's grid and lets every click pass. Nothing is written to the
    // workbook (no conditional formatting, no names), so Undo keeps working and the file is not changed.
    // On / off with the ribbon button; the choice is remembered and comes back when Excel starts.
    internal static class FocusCell
    {
        static Timer timer;
        static FocusBand band;
        static string lastKey;

        [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();

        public static bool On { get { return timer != null; } }

        public static void Toggle()
        {
            if (On)
            {
                Stop();
                Settings.Set("FocusOn", "0");
                Xl.Status("Focus Cell is off.");
            }
            else
            {
                Start();
                Settings.Set("FocusOn", "1");
                Xl.Status("Focus Cell is on: the row and the column of the active cell are marked. Click Focus Cell again to switch it off.");
            }
        }

        // at start-up, when it was on last time
        public static void Restore()
        {
            if (Settings.Get("FocusOn") == "1" && !On) Start();
        }

        static void Start()
        {
            band = new FocusBand(ColorOf(Settings.Get("FocusColor")));
            timer = new Timer();
            timer.Interval = 120;
            timer.Tick += delegate { Tick(); };
            timer.Start();
        }

        public static void Stop()
        {
            if (timer != null) { timer.Stop(); timer.Dispose(); timer = null; }
            if (band != null) { try { band.Close(); band.Dispose(); } catch { } band = null; }
            lastKey = null;
        }

        static Color ColorOf(string s)
        {
            int rgb;
            if (int.TryParse(s, out rgb)) return Color.FromArgb(rgb & 255, (rgb >> 8) & 255, (rgb >> 16) & 255);
            return Color.FromArgb(37, 99, 235);
        }

        static void Tick()
        {
            try
            {
                IntPtr xl = ExcelDnaUtil.WindowHandle;
                if (GetForegroundWindow() != xl)
                {
                    band.Off();
                    lastKey = null;
                    return;
                }
                dynamic app = Xl.App;
                dynamic win = app.ActiveWindow;
                if (win == null) { band.Off(); return; }
                dynamic cell = app.ActiveCell;
                if (cell == null) { band.Off(); return; }
                dynamic pane = win.ActivePane;
                dynamic vr = pane.VisibleRange;
                double cl = (double)cell.Left, ct = (double)cell.Top, cw = (double)cell.Width, ch = (double)cell.Height;
                double vl = (double)vr.Left, vt = (double)vr.Top, vw = (double)vr.Width, vh = (double)vr.Height;
                string key = cl + "|" + ct + "|" + cw + "|" + ch + "|" + vl + "|" + vt + "|" + (double)win.Zoom + "|" + (int)win.Left + "|" + (int)win.Top + "|" + (int)win.Width + "|" + (int)win.Height;
                if (key == lastKey) return;
                lastKey = key;
                int gx1 = (int)pane.PointsToScreenPixelsX(vl), gx2 = (int)pane.PointsToScreenPixelsX(vl + vw);
                int gy1 = (int)pane.PointsToScreenPixelsY(vt), gy2 = (int)pane.PointsToScreenPixelsY(vt + vh);
                int x1 = (int)pane.PointsToScreenPixelsX(cl), x2 = (int)pane.PointsToScreenPixelsX(cl + cw);
                int y1 = (int)pane.PointsToScreenPixelsY(ct), y2 = (int)pane.PointsToScreenPixelsY(ct + ch);
                band.Place(new Rectangle(gx1, gy1, gx2 - gx1, gy2 - gy1), new Rectangle(x1, y1, x2 - x1, y2 - y1), xl);
            }
            catch
            {
                // Excel is busy (a cell in edit mode, a dialog open): try again on the next tick
                lastKey = null;
            }
        }
    }

    // One see-through window over the grid; its shape is the row band plus the column band, without the active cell
    internal sealed class FocusBand : Form
    {
        [DllImport("user32.dll", EntryPoint = "SetWindowLong")] static extern int SetWindowLong32(IntPtr hWnd, int nIndex, int dwNewLong);
        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr")] static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);
        bool owned;

        protected override bool ShowWithoutActivation { get { return true; } }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                // WS_EX_LAYERED | WS_EX_TRANSPARENT (clicks pass through) | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE
                cp.ExStyle |= 0x00080000 | 0x00000020 | 0x00000080 | 0x08000000;
                return cp;
            }
        }

        public FocusBand(Color color)
        {
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            BackColor = color;
            Opacity = 0.16;
            AutoScaleMode = AutoScaleMode.None;
        }

        public void Off()
        {
            if (Visible) Hide();
        }

        public void Place(Rectangle grid, Rectangle cell, IntPtr excel)
        {
            if (grid.Width <= 0 || grid.Height <= 0)
            {
                Off();
                return;
            }
            Rectangle rowBand = new Rectangle(grid.Left, cell.Top, grid.Width, cell.Height);
            Rectangle colBand = new Rectangle(cell.Left, grid.Top, cell.Width, grid.Height);
            rowBand.Intersect(grid);
            colBand.Intersect(grid);
            if (rowBand.IsEmpty && colBand.IsEmpty)
            {
                Off();
                return;
            }
            Bounds = grid;
            Region old = Region;
            Region r = new Region(Rectangle.Empty);
            Point o = grid.Location;
            if (!rowBand.IsEmpty) r.Union(new Rectangle(rowBand.Left - o.X, rowBand.Top - o.Y, rowBand.Width, rowBand.Height));
            if (!colBand.IsEmpty) r.Union(new Rectangle(colBand.Left - o.X, colBand.Top - o.Y, colBand.Width, colBand.Height));
            r.Exclude(new Rectangle(cell.Left - o.X, cell.Top - o.Y, cell.Width, cell.Height));
            Region = r;
            if (old != null) old.Dispose();
            if (!Visible)
            {
                Show();
                if (!owned)
                {
                    // owned by Excel: always just above Excel, never above other programs, hidden with it
                    if (IntPtr.Size == 8) SetWindowLongPtr64(Handle, -8, excel);
                    else SetWindowLong32(Handle, -8, excel.ToInt32());
                    owned = true;
                }
            }
        }
    }
}
