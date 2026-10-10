using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using ExcelDna.Integration;

namespace SmartPro
{
    // Excel constants (no Office interop reference is needed)
    internal static class C
    {
        public const int CalcManual = -4135;
        public const int xlDatabase = 1;
        public const int xlSum = -4157, xlAverage = -4106;
        public const int xlRowField = 1;
        public const int xlAscending = 1, xlDescending = 2;
        public const int xlAutomatic = -4105, xlTop = 1;
        public const int xlColumnClustered = 51, xlBarClustered = 57, xlLineMarkers = 65, xlLine = 4, xlDoughnut = -4120;
        public const int xlCategory = 1, xlValue = 2;
        public const int xlLegendRight = -4152;
        public const int xlSheetHidden = 0, xlSheetVisible = -1;
        public const int xlCenter = -4108, xlLeft = -4131;
        public const int xlEdgeRight = 10;
        public const int xlContinuous = 1, xlThick = 4;
        public const int xlShiftUp = -4162;
        public const int xlPasteValues = -4163, xlPasteFormats = -4122, xlPasteColumnWidths = 8;
        public const int xlWBATWorksheet = -4167;
        public const int xlTypePDF = 0;
        public const int xlOpenXMLWorkbook = 51;
        public const int xlR1C1 = -4150;

        // ---- used by the dashboard builder ----
        public const int xlCount = -4112, xlMax = -4136, xlMin = -4139;
        public const int xlHidden = 0, xlColumnField = 2, xlPageField = 3;
        public const int xlManual = -4135;
        public const int xlPercentOfTotal = 8, xlRunningTotal = 5, xlPercentRunningTotal = 13;
        public const int xlPivotTableVersion14 = 4;
        public const int xlColumnStacked = 52, xlColumnStacked100 = 53, xlBarStacked = 58, xlArea = 1, xlPie = 5;
        public const int xlLegendBottom = -4107, xlLegendTop = -4160;
        public const int xlPrimary = 1, xlSecondary = 2;
        public const int xlHundredThousands = -5, xlTenMillions = -7;
        public const int xlLinear = -4132, xlMovingAvg = 6;
        public const int xlLabelCenter = -4108, xlLabelOutsideEnd = 2, xlLabelAbove = 0;
        public const int xlHorizontal = -4128;
        public const int xlNone = -4142;
        public const int xlAlignRight = -4152, xlAlignTop = -4160, xlAlignBottom = -4107;
        public const int xlEdgeLeft = 7, xlEdgeTop = 8, xlEdgeBottom = 9;
        public const int xlThin = 2, xlMedium = -4138;
        public const int xlLandscape = 2;
        public const int xlMarkerCircle = 8;
        public const int xlWholeTable = 0, xlHeaderRow = 1;
        public const int xlSlicerUnselectedItemWithData = 28, xlSlicerUnselectedItemWithNoData = 29, xlSlicerSelectedItemWithData = 30,
                         xlSlicerSelectedItemWithNoData = 31, xlSlicerHoveredUnselectedItemWithData = 32, xlSlicerHoveredSelectedItemWithData = 33,
                         xlSlicerHoveredUnselectedItemWithNoData = 34, xlSlicerHoveredSelectedItemWithNoData = 35;
        public const int xlConditionValueNumber = 0, xlConditionValueAutomaticMax = 7;

        // ---- Office shapes ----
        public const int msoTrue = -1, msoFalse = 0;
        public const int msoShapeRectangle = 1, msoShapeRoundedRectangle = 5;
        public const int msoTextHorizontal = 1;
        public const int msoAnchorTop = 1, msoAnchorMiddle = 3, msoAnchorBottom = 4;
        public const int msoAlignLeft = 1, msoAlignCenter = 2, msoAlignRight = 3;
        public const int msoGradientVertical = 2;       // the colour changes from left to right
        public const int msoLineDash = 4;
    }

    // Parent window for message boxes
    internal sealed class ExcelWin : IWin32Window
    {
        public IntPtr Handle { get { return ExcelDnaUtil.WindowHandle; } }
    }

    // using (new Speed()) { ... }  switches off screen updates, events and recalculation, and always restores them
    internal sealed class Speed : IDisposable
    {
        readonly object calc;

        public Speed()
        {
            dynamic app = Xl.App;
            try { calc = app.Calculation; } catch { calc = null; }
            try
            {
                app.ScreenUpdating = false;
                app.EnableEvents = false;
                app.DisplayAlerts = false;
                app.Calculation = C.CalcManual;
            }
            catch { }
        }

        public void Dispose()
        {
            dynamic app = Xl.App;
            try { if (calc != null) app.Calculation = calc; } catch { }
            Restore();
        }

        public static void Restore()
        {
            dynamic app = Xl.App;
            try
            {
                app.ScreenUpdating = true;
                app.EnableEvents = true;
                app.DisplayAlerts = true;
                app.StatusBar = false;
            }
            catch { }
        }
    }

    internal static class Xl
    {
        public static readonly object M = Type.Missing;

        public static dynamic App { get { return ExcelDnaUtil.Application; } }

        public static void Status(string text)
        {
            try { App.StatusBar = Design.AppName + ": " + text; } catch { }
        }

        // ---------- arrays (always 1-based, like the arrays Excel returns) ----------
        public static object[,] New2D(int rows, int cols)
        {
            return (object[,])Array.CreateInstance(typeof(object), new[] { rows, cols }, new[] { 1, 1 });
        }

        public static object[,] To2D(object val)
        {
            object[,] arr = val as object[,];
            if (arr != null) return arr;
            object[,] one = New2D(1, 1);
            one[1, 1] = val;
            return one;
        }

        // dates arrive as DateTime, currency as decimal, errors as int.
        // Excel refuses to hand over a whole block when one cell holds a date or currency value it cannot
        // express (for example a very large number in a date-formatted cell): "Out of present range".
        // Then the block is read column by column and only the column that fails falls back to raw numbers.
        public static object[,] ReadTyped(dynamic rng)
        {
            object val;
            try { val = rng.Value; }
            catch (System.Runtime.InteropServices.COMException) { return ReadTypedSafe(rng); }
            catch (ArgumentException) { return ReadTypedSafe(rng); }
            catch (OverflowException) { return ReadTypedSafe(rng); }
            return To2D(val);
        }

        static object[,] ReadTypedSafe(dynamic rng)
        {
            int nr = (int)rng.Rows.Count, nc = (int)rng.Columns.Count;
            object[,] all = New2D(nr, nc);
            for (int c = 1; c <= nc; c++)
            {
                dynamic col = rng.Columns[c];
                object[,] part;
                try
                {
                    object typed = col.Value;
                    part = To2D(typed);
                }
                catch
                {
                    object raw = col.Value2;                 // Value2 never fails: dates come as plain numbers
                    part = To2D(raw);
                    bool isDate = false;
                    try
                    {
                        string f = col.NumberFormat as string;   // text only when the whole column shares one format
                        if (f == null && nr > 1) f = Convert.ToString((object)col.Cells[2, 1].NumberFormat, CultureInfo.InvariantCulture);
                        isDate = LooksLikeDateFormat(f);
                    }
                    catch { }
                    if (isDate)
                    {
                        for (int r = 1; r <= nr; r++)
                        {
                            object x = part[r, 1];
                            if (!(x is double)) continue;
                            double d = (double)x;
                            if (d >= 1 && d < 2958466) part[r, 1] = DateTime.FromOADate(d);
                        }
                    }
                }
                for (int r = 1; r <= nr; r++) all[r, c] = part[r, 1];
            }
            return all;
        }

        // true for number formats such as dd-mm-yyyy, d mmm yy, hh:mm
        static bool LooksLikeDateFormat(string f)
        {
            if (string.IsNullOrEmpty(f)) return false;
            if (f.IndexOf("General", StringComparison.OrdinalIgnoreCase) >= 0) return false;
            bool quoted = false, bracket = false;
            for (int i = 0; i < f.Length; i++)
            {
                char ch = char.ToLowerInvariant(f[i]);
                if (quoted) { if (ch == '"') quoted = false; continue; }
                if (bracket) { if (ch == ']') bracket = false; continue; }
                if (ch == '"') { quoted = true; continue; }
                if (ch == '[') { bracket = true; continue; }
                if (ch == '\\' || ch == '_' || ch == '*') { i++; continue; }   // the next character is literal
                if (ch == 'd' || ch == 'y' || ch == 'm' || ch == 'h' || ch == 's') return true;
            }
            return false;
        }
        public static object[,] ReadFormulas(dynamic rng) { object val = rng.Formula; return To2D(val); }
        public static object[,] ReadFormulasR1C1(dynamic rng) { object val = rng.FormulaR1C1; return To2D(val); }

        // ---------- cell value helpers ----------
        public static bool IsErr(object v) { return v is int && (int)v < -2146000000; }

        public static bool IsBlank(object v)
        {
            if (v == null) return true;
            string s = v as string;
            return s != null && s.Trim().Length == 0;
        }

        public static bool IsNum(object v) { return v is double || v is decimal || v is float; }

        public static double Num(object v) { return Convert.ToDouble(v, CultureInfo.InvariantCulture); }

        public static string ErrText(int code)
        {
            switch (code)
            {
                case -2146826281: return "#DIV/0!";
                case -2146826246: return "#N/A";
                case -2146826259: return "#NAME?";
                case -2146826288: return "#NULL!";
                case -2146826252: return "#NUM!";
                case -2146826265: return "#REF!";
                case -2146826273: return "#VALUE!";
                default: return "#ERROR";
            }
        }

        public static string Str(object v)
        {
            if (v == null) return "";
            string s = v as string;
            if (s != null) return s;
            if (v is double) return ((double)v).ToString(CultureInfo.InvariantCulture);
            if (v is DateTime)
            {
                DateTime d = (DateTime)v;
                return d.TimeOfDay.Ticks == 0 ? d.ToString("dd-MMM-yyyy", CultureInfo.InvariantCulture) : d.ToString("dd-MMM-yyyy HH:mm:ss", CultureInfo.InvariantCulture);
            }
            if (v is bool) return (bool)v ? "TRUE" : "FALSE";
            if (v is int) return IsErr(v) ? ErrText((int)v) : ((int)v).ToString(CultureInfo.InvariantCulture);
            return Convert.ToString(v, CultureInfo.InvariantCulture);
        }

        public static string Key(object v) { return Str(v).Trim().ToLowerInvariant(); }

        public static string GroupKey(object v)
        {
            if (IsBlank(v)) return "(blank)";
            return Str(v).Trim();
        }

        public static string HeaderText(object v, int j)
        {
            if (IsBlank(v) || IsErr(v)) return "Column " + j;
            return Str(v).Trim();
        }

        public static bool RowBlank(object[,] v, int r)
        {
            int nc = v.GetLength(1);
            for (int c = 1; c <= nc; c++) if (!IsBlank(v[r, c])) return false;
            return true;
        }

        public static string ColLetter(int col)
        {
            string s = "";
            while (col > 0)
            {
                int m = (col - 1) % 26;
                s = (char)('A' + m) + s;
                col = (col - 1) / 26;
            }
            return s;
        }

        public static string A1(int row, int col) { return ColLetter(col) + row.ToString(CultureInfo.InvariantCulture); }

        // ---------- asking the user ----------
        // Lets the user confirm / pick a range. Returns null when cancelled.
        public static dynamic PickRange(string title, string prompt)
        {
            dynamic app = App;
            string def = "";
            try
            {
                dynamic sel = app.Selection;
                long cnt = (long)(int)sel.Rows.Count * (int)sel.Columns.Count;
                dynamic use = cnt > 1 ? sel.Areas[1] : sel.CurrentRegion;
                def = (string)use.Address;
            }
            catch { }
            object res;
            try { res = app.InputBox(prompt ?? "Select the data (include the header row):", title, def, M, M, M, M, 8); }
            catch { return null; }
            if (res == null || res is bool) return null;
            dynamic r = res;
            try { r = app.Intersect(r.Areas[1], r.Worksheet.UsedRange); }
            catch { return null; }
            return r;
        }

        public static dynamic PickRange(string title) { return PickRange(title, null); }

        public static long CellCount(dynamic rng) { return (long)(int)rng.Rows.Count * (int)rng.Columns.Count; }

        // ---------- sheets ----------
        public static string SafeSheetName(string s)
        {
            foreach (char ch in new[] { ':', '\\', '/', '?', '*', '[', ']' }) s = s.Replace(ch, '-');
            s = s.Trim().Trim('\'');
            if (s.Length == 0) s = "(blank)";
            if (s.Length > 28) s = s.Substring(0, 28);
            return s;
        }

        public static string SafeFileName(string s)
        {
            foreach (char ch in Path.GetInvalidFileNameChars()) s = s.Replace(ch, '-');
            s = s.Trim();
            if (s.Length == 0) s = "blank";
            if (s.Length > 80) s = s.Substring(0, 80);
            return s;
        }

        public static bool SheetExists(dynamic wb, string name)
        {
            try { object sh = wb.Sheets[name]; return sh != null; }
            catch { return false; }
        }

        public static dynamic AddSheet(dynamic wb, string name)
        {
            string stem = SafeSheetName(name);
            string nm = stem;
            int i = 1;
            bool exists = SheetExists(wb, nm);
            while (exists)
            {
                i++;
                nm = stem + " " + i;
                exists = SheetExists(wb, nm);
            }
            int last = (int)wb.Sheets.Count;
            dynamic ws = wb.Worksheets.Add(M, wb.Sheets[last]);
            ws.Name = nm;
            return ws;
        }

        // ---------- writing ----------
        // Writes a table in one call. Columns holding number-like TEXT (00123, 1-2, =x) are formatted as Text first
        // so Excel does not convert them; all-date columns get a date format; errors are written as their text.
        public static void WriteTable(dynamic topLeft, object[,] arr, int hdrRows)
        {
            int nr = arr.GetLength(0), nc = arr.GetLength(1);
            int r0 = arr.GetLowerBound(0), c0 = arr.GetLowerBound(1);
            dynamic dest = topLeft.Resize[nr, nc];
            for (int j = 0; j < nc; j++)
            {
                bool risky = false, anyDate = false, anyOther = false;
                for (int i = 0; i < nr; i++)
                {
                    object x = arr[r0 + i, c0 + j];
                    if (x == null) continue;
                    if (x is DateTime)
                    {
                        arr[r0 + i, c0 + j] = ((DateTime)x).ToOADate();
                        if (i >= hdrRows) anyDate = true;
                    }
                    else if (x is decimal) { arr[r0 + i, c0 + j] = (double)(decimal)x; if (i >= hdrRows) anyOther = true; }
                    else if (x is int)
                    {
                        if (IsErr(x)) arr[r0 + i, c0 + j] = ErrText((int)x);
                        else { arr[r0 + i, c0 + j] = (double)(int)x; if (i >= hdrRows) anyOther = true; }
                    }
                    else if (i >= hdrRows)
                    {
                        string s = x as string;
                        if (s == null) anyOther = true;
                        else if (!risky && Logic.RiskyText(s)) risky = true;
                    }
                }
                if (nr > hdrRows)
                {
                    dynamic body = dest.Columns[j + 1].Offset[hdrRows, 0].Resize[nr - hdrRows, 1];
                    if (risky) body.NumberFormat = "@";
                    else if (anyDate && !anyOther) body.NumberFormat = "dd-mmm-yyyy";
                }
            }
            if (hdrRows > 0) dest.Rows[1].Resize[hdrRows, nc].NumberFormat = "@";
            dest.Value2 = arr;
        }

        public static void StyleHeader(dynamic rng)
        {
            rng.Font.Bold = true;
            rng.Font.Color = Design.White;
            rng.Interior.Color = Design.Navy;
        }

        public static long Cell(int r, int c) { return ((long)c << 32) | (uint)r; }

        // Colours many cells fast: consecutive rows of one column are painted with a single call
        public static void Paint(dynamic rng, List<long> cells, int color)
        {
            if (cells.Count == 0) return;
            cells.Sort();
            dynamic ws = rng.Worksheet;
            int i = 0;
            while (i < cells.Count)
            {
                int c = (int)(cells[i] >> 32);
                int r1 = (int)(cells[i] & 0xFFFFFFFFL);
                int r2 = r1;
                while (i + 1 < cells.Count && (int)(cells[i + 1] >> 32) == c && (int)(cells[i + 1] & 0xFFFFFFFFL) == r2 + 1)
                {
                    i++;
                    r2++;
                }
                ws.Range[rng.Cells[r1, c], rng.Cells[r2, c]].Interior.Color = color;
                i++;
            }
        }

        public static void Try(Action act)
        {
            try { act(); } catch { }
        }

        // ---------- backups ----------
        public static string BackupFolder()
        {
            string f = Path.Combine(Path.GetTempPath(), "SmartPro_Backups");
            Directory.CreateDirectory(f);
            return f;
        }

        public static string Backup(dynamic wb)
        {
            try
            {
                string nm = (string)wb.Name;
                string path = (string)wb.Path;
                string ext = ".xlsx";
                int p = nm.LastIndexOf('.');
                if (p > 0 && !string.IsNullOrEmpty(path))
                {
                    ext = nm.Substring(p);
                    nm = nm.Substring(0, p);
                }
                string f = Path.Combine(BackupFolder(), SafeFileName(nm) + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) + ext);
                wb.SaveCopyAs(f);
                return f;
            }
            catch { return null; }
        }
    }

    internal static class Ui
    {
        public static void Info(string text)
        {
            if (string.IsNullOrEmpty(text)) return;
            MessageBox.Show(new ExcelWin(), text, Design.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static void Warn(string text)
        {
            MessageBox.Show(new ExcelWin(), text, Design.AppName, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        public static bool Ok(string text)
        {
            return MessageBox.Show(new ExcelWin(), text, Design.AppName, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK;
        }

        public static DialogResult YesNoCancel(string text)
        {
            return MessageBox.Show(new ExcelWin(), text, Design.AppName, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
        }

        // text box; null when cancelled
        public static string AskText(string prompt, string title, string def)
        {
            object res;
            try { res = Xl.App.InputBox(prompt, title, def, Xl.M, Xl.M, Xl.M, Xl.M, 2); }
            catch { return null; }
            if (res == null || res is bool) return null;
            return Convert.ToString(res, CultureInfo.InvariantCulture);
        }

        // searchable list; returns the index picked or -1
        public static int Pick(string title, string hint, string[] items, string[] keys)
        {
            using (PickForm f = new PickForm(title, hint, items, keys))
            {
                return f.ShowDialog(new ExcelWin()) == DialogResult.OK ? f.Picked : -1;
            }
        }

        // column picker: shows the headers, returns the 1-based column or 0
        public static int PickColumn(string title, string hint, string[] headers, bool[] allowed)
        {
            List<string> items = new List<string>();
            List<int> map = new List<int>();
            for (int c = 1; c < headers.Length; c++)
            {
                if (allowed != null && !allowed[c]) continue;
                items.Add(Xl.ColLetter(c) + "   " + headers[c]);
                map.Add(c);
            }
            if (items.Count == 0) return 0;
            int pick = Pick(title, hint, items.ToArray(), null);
            return pick < 0 ? 0 : map[pick];
        }

        public static string PickFolder(string title)
        {
            try
            {
                dynamic fd = Xl.App.FileDialog[4];
                fd.Title = title;
                fd.AllowMultiSelect = false;
                int res = (int)fd.Show();
                if (res != -1) return null;
                return (string)fd.SelectedItems.Item(1);
            }
            catch { return null; }
        }
    }

    // Small search-as-you-type list used by Quick Find and the column pickers
    internal sealed class PickForm : Form
    {
        readonly TextBox box = new TextBox();
        readonly ListBox list = new ListBox();
        readonly string[] items;
        readonly string[] keys;
        readonly List<int> shown = new List<int>();
        public int Picked = -1;

        public PickForm(string title, string hint, string[] items, string[] keys)
        {
            this.items = items;
            this.keys = keys ?? items;
            Text = title;
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new System.Drawing.Size(560, 430);
            Font = new System.Drawing.Font("Segoe UI", 10f);

            Label lbl = new Label();
            lbl.Text = hint;
            lbl.SetBounds(12, 10, 536, 22);
            box.SetBounds(12, 36, 536, 28);
            list.SetBounds(12, 72, 536, 306);
            list.IntegralHeight = false;
            Button ok = new Button();
            ok.Text = "OK";
            ok.SetBounds(368, 388, 85, 30);
            Button cancel = new Button();
            cancel.Text = "Cancel";
            cancel.SetBounds(463, 388, 85, 30);
            cancel.DialogResult = DialogResult.Cancel;

            Controls.Add(lbl);
            Controls.Add(box);
            Controls.Add(list);
            Controls.Add(ok);
            Controls.Add(cancel);
            AcceptButton = ok;
            CancelButton = cancel;

            box.TextChanged += delegate { Fill(); };
            box.KeyDown += OnBoxKey;
            list.DoubleClick += delegate { Accept(); };
            ok.Click += delegate { Accept(); };
            Fill();
        }

        void Accept()
        {
            if (list.SelectedIndex < 0) return;
            Picked = shown[list.SelectedIndex];
            DialogResult = DialogResult.OK;
            Close();
        }

        void OnBoxKey(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                if (list.SelectedIndex < list.Items.Count - 1) list.SelectedIndex++;
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Up)
            {
                if (list.SelectedIndex > 0) list.SelectedIndex--;
                e.Handled = true;
            }
        }

        void Fill()
        {
            string[] words = box.Text.Trim().ToLowerInvariant().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            list.BeginUpdate();
            list.Items.Clear();
            shown.Clear();
            for (int i = 0; i < items.Length; i++)
            {
                string hay = keys[i].ToLowerInvariant();
                bool match = true;
                foreach (string w in words)
                {
                    if (hay.IndexOf(w, StringComparison.Ordinal) < 0) { match = false; break; }
                }
                if (match)
                {
                    list.Items.Add(items[i]);
                    shown.Add(i);
                }
            }
            list.EndUpdate();
            if (list.Items.Count > 0) list.SelectedIndex = 0;
        }
    }
}
