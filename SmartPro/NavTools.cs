using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace SmartPro
{
    // NAVIGATION and SHEET MANAGEMENT for big workbooks:
    //   Sheet Navigator (Ctrl+Shift+J) ... type a few letters, jump to any sheet of any open workbook; "back" to the last one
    //   Workbook Index ................... a clickable index that is updated in place (not a new sheet every time),
    //                                      with hidden sheets, sizes, tab colours and an optional "Index" button on every sheet
    //   Sheet Manager .................... show / hide / very hide / rename / colour / move / sort / delete many sheets at once
    //   Swap Ranges ...................... swaps two blocks of the same size, formulas included
    internal static class NavTools
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        const string IndexTitle = "Workbook index: ";
        const string BackButton = "DB_IndexButton";
        static string prevBook, prevSheet;      // for "back to the previous sheet"

        static string N(int n) { return n.ToString("#,##0", Inv); }

        static string Quote(string sheet) { return "'" + sheet.Replace("'", "''") + "'"; }

        static string Status(dynamic sh)
        {
            int v = (int)sh.Visible;
            return v == C.xlSheetVisible ? "Visible" : v == C.xlSheetHidden ? "Hidden" : "Very hidden";
        }

        // ---------------------------------------------------------------
        //  SHEET NAVIGATOR
        // ---------------------------------------------------------------
        public static void Navigator()
        {
            dynamic app = Xl.App;
            dynamic active = app.ActiveWorkbook;
            string activeBook = active == null ? "" : (string)active.Name;
            string activeSheet = "";
            try { activeSheet = (string)app.ActiveSheet.Name; } catch { }

            List<string> items = new List<string>(), keys = new List<string>();
            List<string[]> target = new List<string[]>();          // book, sheet
            bool hasBack = prevBook != null && !(prevBook == activeBook && prevSheet == activeSheet);
            if (hasBack)
            {
                items.Add("<-  Back to " + prevSheet + (prevBook != activeBook ? "   [" + prevBook + "]" : ""));
                keys.Add("back previous last " + prevSheet);
                target.Add(new[] { prevBook, prevSheet });
            }
            int books = 0;
            foreach (dynamic wb in app.Workbooks)
            {
                string bn = (string)wb.Name;
                bool isActive = bn == activeBook;
                try { if (!(bool)wb.Windows[1].Visible) continue; } catch { }   // the hidden Personal workbook
                books++;
                int i = 0;
                foreach (dynamic sh in wb.Sheets)
                {
                    i++;
                    string sn = (string)sh.Name;
                    string st = Status(sh);
                    string size = "";
                    try { size = "   " + N((int)sh.UsedRange.Rows.Count) + " rows"; } catch { size = "   chart"; }
                    string line = (isActive ? "" : "[" + bn + "]  ") + i + ".  " + sn + (st == "Visible" ? "" : "   (" + st.ToLowerInvariant() + ")") + size;
                    // the active workbook first
                    if (isActive)
                    {
                        int at = hasBack ? 1 : 0;
                        int pos = at;
                        while (pos < items.Count && target[pos][0] == activeBook) pos++;
                        items.Insert(pos, line);
                        keys.Insert(pos, sn + " " + bn + " " + st);
                        target.Insert(pos, new[] { bn, sn });
                    }
                    else
                    {
                        items.Add(line);
                        keys.Add(sn + " " + bn + " " + st);
                        target.Add(new[] { bn, sn });
                    }
                }
            }
            if (items.Count == 0)
            {
                Ui.Info("Open a workbook first.");
                return;
            }
            int pick = Ui.Pick(Design.AppName + " - Sheet Navigator   (" + N(items.Count - (hasBack ? 1 : 0)) + " sheets in " + books + (books == 1 ? " workbook)" : " workbooks)"),
                               "Type part of a sheet name and press Enter.  Shortcut: Ctrl+Shift+J", items.ToArray(), keys.ToArray());
            if (pick < 0) return;
            GoTo(target[pick][0], target[pick][1], activeBook, activeSheet);
        }

        static void GoTo(string book, string sheet, string fromBook, string fromSheet)
        {
            dynamic wb = Xl.App.Workbooks[book];
            dynamic sh = wb.Sheets[sheet];
            if ((int)sh.Visible != C.xlSheetVisible)
            {
                if (!Ui.Ok("'" + sheet + "' is hidden. Show it and go there?")) return;
                sh.Visible = C.xlSheetVisible;
            }
            wb.Activate();
            sh.Activate();
            prevBook = fromBook;
            prevSheet = fromSheet;
        }

        // ---------------------------------------------------------------
        //  WORKBOOK INDEX
        // ---------------------------------------------------------------
        public static void WorkbookIndex()
        {
            dynamic wb = Xl.App.ActiveWorkbook;
            dynamic ws = null;
            // an index made before is rebuilt in place, so the links that point to it keep working
            foreach (dynamic sh in wb.Worksheets)
            {
                string a1 = "";
                try { a1 = Convert.ToString((object)sh.Range["A1"].Value2, Inv) ?? ""; } catch { }
                if (a1.StartsWith(IndexTitle, StringComparison.Ordinal) || a1.StartsWith("Sheet index: ", StringComparison.Ordinal)) { ws = sh; break; }
            }
            bool buttons = Ui.Ok("Also put a small 'Index' button on every sheet, to come back here in one click?\n\n" +
                                 "OK = yes (a button at the top right, it does not touch your cells)\nCancel = only the index");
            int count = 0, hidden = 0;
            using (new Speed())
            {
                if (ws == null)
                {
                    ws = Xl.AddSheet(wb, "Index");
                    ws.Move(wb.Sheets[1]);
                }
                else ws.Cells.Clear();
                string me = (string)ws.Name;
                ws.Range["A1"].Value2 = IndexTitle + (string)wb.Name;
                ws.Range["A1"].Font.Size = 14;
                ws.Range["A1"].Font.Bold = true;
                ws.Range["A2"].Value2 = "Updated " + DateTime.Now.ToString("dd-MMM-yyyy HH:mm", Inv) + ".  Run Workbook Index again to update it.";
                ws.Range["A2"].Font.Color = Design.Slate;
                string[] hd = { "No.", "Sheet (click to open)", "Status", "Rows", "Columns", "Last cell", "Charts", "Tab" };
                object[,] h = Xl.New2D(1, hd.Length);
                for (int j = 0; j < hd.Length; j++) h[1, j + 1] = hd[j];
                ws.Range["A4"].Resize[1, hd.Length].Value2 = h;
                Xl.StyleHeader(ws.Range["A4"].Resize[1, hd.Length]);
                int r = 4;
                foreach (dynamic sh in wb.Sheets)
                {
                    string nm = (string)sh.Name;
                    if (nm == me) continue;
                    r++;
                    count++;
                    string st = Status(sh);
                    if (st != "Visible") hidden++;
                    ws.Cells[r, 1].Value2 = count;
                    ws.Cells[r, 3].Value2 = st;
                    bool isSheet = true;
                    try
                    {
                        dynamic ur = sh.UsedRange;
                        ws.Cells[r, 4].Value2 = (int)ur.Rows.Count;
                        ws.Cells[r, 5].Value2 = (int)ur.Columns.Count;
                        ws.Cells[r, 6].Value2 = "'" + ((string)ur.Cells[(int)ur.Rows.Count, (int)ur.Columns.Count].Address).Replace("$", "");
                        int ch = 0;
                        try { ch = (int)sh.ChartObjects().Count; } catch { }
                        ws.Cells[r, 7].Value2 = ch;
                    }
                    catch { isSheet = false; }
                    if (isSheet && st == "Visible") ws.Hyperlinks.Add(ws.Cells[r, 2], "", Quote(nm) + "!A1", Xl.M, nm);
                    else
                    {
                        ws.Cells[r, 2].Value2 = "'" + nm;
                        if (!isSheet) ws.Cells[r, 3].Value2 = st + " (chart sheet)";
                    }
                    if (st != "Visible") ws.Cells[r, 3].Font.Color = Design.Rgb(185, 28, 28);
                    try
                    {
                        object tc = sh.Tab.Color;
                        if (tc is double || tc is int) ws.Cells[r, 8].Interior.Color = Convert.ToInt32(tc, Inv);
                    }
                    catch { }
                    if (buttons && isSheet) AddIndexButton(sh, me);
                    else if (!buttons && isSheet) RemoveIndexButton(sh);
                }
                try
                {
                    dynamic tbl = ws.Range["A4"].Resize[r - 3, hd.Length];
                    tbl.Borders.Color = Design.GridLine;
                    ws.Columns["A:H"].AutoFit();
                    ws.Columns["B"].ColumnWidth = Math.Max(24.0, (double)ws.Columns["B"].ColumnWidth);
                    ws.Activate();
                    ws.Range["A5"].Select();
                    Xl.App.ActiveWindow.FreezePanes = false;
                    Xl.App.ActiveWindow.SplitRow = 4;
                    Xl.App.ActiveWindow.SplitColumn = 0;
                    Xl.App.ActiveWindow.FreezePanes = true;
                    Xl.App.ActiveWindow.DisplayGridlines = false;
                }
                catch { }
            }
            Xl.Status("Index of " + count + " sheets" + (hidden > 0 ? " (" + hidden + " hidden)" : "") + (buttons ? ", with an Index button on every sheet." : "."));
        }

        static void RemoveIndexButton(dynamic sh)
        {
            try { sh.Shapes.Item(BackButton).Delete(); } catch { }
        }

        static void AddIndexButton(dynamic sh, string indexSheet)
        {
            RemoveIndexButton(sh);
            try
            {
                dynamic ur = sh.UsedRange;
                int col = (int)ur.Column + (int)ur.Columns.Count + 1;
                if (col > 16000) col = 2;
                double left = (double)sh.Cells[1, col].Left;
                dynamic b = sh.Shapes.AddShape(C.msoShapeRoundedRectangle, left, 2.0, 56.0, 18.0);
                b.Name = BackButton;
                b.Fill.ForeColor.RGB = Design.Navy;
                b.Line.Visible = C.msoFalse;
                dynamic tf = b.TextFrame2;
                tf.TextRange.Text = "Index";
                tf.TextRange.Font.Size = 9;
                tf.TextRange.Font.Bold = C.msoTrue;
                tf.TextRange.Font.Fill.ForeColor.RGB = Design.White;
                tf.TextRange.ParagraphFormat.Alignment = C.msoAlignCenter;
                tf.VerticalAnchor = C.msoAnchorMiddle;
                tf.MarginTop = 0;
                tf.MarginBottom = 0;
                b.Placement = 3;                              // xlFreeFloating: stays put when rows or columns change
                sh.Hyperlinks.Add(b, "", Quote(indexSheet) + "!A1", "Back to the index");
            }
            catch { }
        }

        // ---------------------------------------------------------------
        //  SWAP RANGES
        // ---------------------------------------------------------------
        public static void SwapRanges()
        {
            dynamic a = Xl.PickRange("Swap Ranges (1 of 2)", "Select the FIRST block (for example a column):");
            if (a == null) return;
            dynamic b = Xl.PickRange("Swap Ranges (2 of 2)", "Select the SECOND block, the same size (" + (int)a.Rows.Count + " x " + (int)a.Columns.Count + "):");
            if (b == null) return;
            if ((int)a.Rows.Count != (int)b.Rows.Count || (int)a.Columns.Count != (int)b.Columns.Count)
            {
                Ui.Info("The two blocks must be the same size.");
                return;
            }
            if (Xl.App.Intersect(a, b) != null)
            {
                Ui.Info("The two blocks overlap. Pick two blocks that do not share cells.");
                return;
            }
            Xl.Backup(a.Worksheet.Parent);
            using (new Speed())
            {
                object fa = a.Formula, fb = b.Formula;
                object na = a.NumberFormat, nb = b.NumberFormat;
                a.Formula = fb;
                b.Formula = fa;
                try { if (na is string && nb is string) { a.NumberFormat = nb; b.NumberFormat = na; } } catch { }
            }
            Xl.Status("Swapped " + ((string)a.Address).Replace("$", "") + " and " + ((string)b.Address).Replace("$", "") + " (values and formulas).");
        }

        // ---------------------------------------------------------------
        //  SHEET MANAGER
        // ---------------------------------------------------------------
        public static void SheetManager()
        {
            using (SheetManagerForm f = new SheetManagerForm(Xl.App.ActiveWorkbook))
            {
                f.ShowDialog(new ExcelWin());
            }
        }
    }

    internal sealed class SheetManagerForm : Form
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        readonly dynamic wb;
        readonly ListView list = new ListView();
        readonly Label msg = new Label();
        readonly float z = 1f;

        int S(int px) { return (int)Math.Round(px * z); }

        public SheetManagerForm(dynamic workbook)
        {
            wb = workbook;
            try { using (Graphics g = Graphics.FromHwnd(IntPtr.Zero)) { z = Math.Max(1f, Math.Min(2.5f, g.DpiX / 96f)); } } catch { }
            AutoScaleMode = AutoScaleMode.None;
            Text = Design.AppName + " - Sheet Manager   (" + (string)wb.Name + ")";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            BackColor = Color.White;
            Font = new Font("Segoe UI", 9.5f);
            ClientSize = new Size(S(760), S(520));

            Label hint = new Label();
            hint.Text = "Tick sheets (or select rows), then press a button. Changes happen at once; Ctrl+A ticks all.";
            hint.SetBounds(S(14), S(10), S(560), S(22));
            Controls.Add(hint);
            list.SetBounds(S(14), S(36), S(560), S(420));
            list.View = View.Details;
            list.CheckBoxes = true;
            list.FullRowSelect = true;
            list.HideSelection = false;
            list.Columns.Add("Sheet", S(250));
            list.Columns.Add("Status", S(100));
            list.Columns.Add("Rows", S(80));
            list.Columns.Add("Tab colour", S(100));
            list.KeyDown += delegate (object s, KeyEventArgs e)
            {
                if (e.Control && e.KeyCode == Keys.A)
                {
                    bool all = true;
                    foreach (ListViewItem it in list.Items) if (!it.Checked) all = false;
                    foreach (ListViewItem it in list.Items) it.Checked = !all;
                    e.Handled = true;
                }
            };
            Controls.Add(list);

            int y = 36;
            Btn("Show", y, delegate { SetVisible(C.xlSheetVisible); }); y += 38;
            Btn("Hide", y, delegate { SetVisible(C.xlSheetHidden); }); y += 38;
            Btn("Very hide (not in Unhide)", y, delegate { SetVisible(2); }); y += 46;
            Btn("Rename (add / replace)", y, delegate { Rename(); }); y += 38;
            Btn("Tab colour", y, delegate { TabColour(); }); y += 38;
            Btn("No tab colour", y, delegate { Each(delegate (dynamic sh) { sh.Tab.ColorIndex = C.xlNone; }); }); y += 46;
            Btn("Move up", y, delegate { MoveSheets(-1); }); y += 38;
            Btn("Move down", y, delegate { MoveSheets(1); }); y += 38;
            Btn("Sort A to Z", y, delegate { SortAll(); }); y += 46;
            Btn("Delete", y, delegate { DeleteSheets(); });

            msg.SetBounds(S(14), S(462), S(560), S(44));
            msg.ForeColor = Color.FromArgb(71, 85, 105);
            Controls.Add(msg);
            Button close = new Button();
            close.Text = "Close";
            close.SetBounds(S(590), S(474), S(156), S(32));
            close.DialogResult = DialogResult.Cancel;
            Controls.Add(close);
            CancelButton = close;
            Fill(null);
        }

        void Btn(string text, int y, EventHandler click)
        {
            Button b = new Button();
            b.Text = text;
            b.SetBounds(S(590), S(y), S(156), S(32));
            b.Click += delegate (object s, EventArgs e)
            {
                try { click(s, e); }
                catch (Exception ex) { Say("Excel refused: " + ex.Message); }
                Fill(Picked());
            };
            Controls.Add(b);
        }

        void Say(string s) { msg.Text = s; }

        void Fill(List<string> keep)
        {
            list.BeginUpdate();
            list.Items.Clear();
            foreach (dynamic sh in wb.Sheets)
            {
                string nm = (string)sh.Name;
                int v = (int)sh.Visible;
                string rows = "";
                try { rows = ((int)sh.UsedRange.Rows.Count).ToString("#,##0", Inv); } catch { rows = "chart"; }
                string tab = "";
                Color tc = Color.Empty;
                try
                {
                    object c = sh.Tab.Color;
                    if (c is double || c is int)
                    {
                        int rgb = Convert.ToInt32(c, Inv);
                        tc = Color.FromArgb(rgb & 255, (rgb >> 8) & 255, (rgb >> 16) & 255);
                        tab = "■■■■";
                    }
                }
                catch { }
                ListViewItem it = new ListViewItem(new[] { nm, v == C.xlSheetVisible ? "Visible" : v == C.xlSheetHidden ? "Hidden" : "Very hidden", rows, tab });
                it.UseItemStyleForSubItems = false;
                if (v != C.xlSheetVisible) it.SubItems[1].ForeColor = Color.FromArgb(185, 28, 28);
                if (tc != Color.Empty) it.SubItems[3].ForeColor = tc;
                if (keep != null && keep.Contains(nm)) it.Checked = true;
                list.Items.Add(it);
            }
            list.EndUpdate();
        }

        List<string> Picked()
        {
            List<string> l = new List<string>();
            foreach (ListViewItem it in list.Items) if (it.Checked || it.Selected) l.Add(it.Text);
            return l;
        }

        delegate void SheetAction(dynamic sh);

        int Each(SheetAction act)
        {
            List<string> names = Picked();
            if (names.Count == 0)
            {
                Say("Tick at least one sheet first.");
                return 0;
            }
            int n = 0;
            foreach (string nm in names)
            {
                act(wb.Sheets[nm]);
                n++;
            }
            Say(n + (n == 1 ? " sheet" : " sheets") + " changed.");
            return n;
        }

        void SetVisible(int state)
        {
            List<string> names = Picked();
            int visible = 0;
            foreach (dynamic sh in wb.Sheets) if ((int)sh.Visible == C.xlSheetVisible && !names.Contains((string)sh.Name)) visible++;
            if (state != C.xlSheetVisible && visible == 0)
            {
                Say("At least one sheet must stay visible. Untick one sheet.");
                return;
            }
            Each(delegate (dynamic sh) { sh.Visible = state; });
        }

        void Rename()
        {
            if (Picked().Count == 0) { Say("Tick at least one sheet first."); return; }
            string front = Ui.AskText("Text to add IN FRONT of each name (empty for none):", "Rename sheets (1 of 3)", "");
            if (front == null) return;
            string back = Ui.AskText("Text to add AT THE END of each name (empty for none):", "Rename sheets (2 of 3)", "");
            if (back == null) return;
            string find = Ui.AskText("Replace this text inside the names (empty to skip):", "Rename sheets (3 of 3)", "");
            if (find == null) return;
            string with = "";
            if (find.Length > 0)
            {
                with = Ui.AskText("... with this text:", "Rename sheets", "");
                if (with == null) return;
            }
            int skipped = 0;
            Each(delegate (dynamic sh)
            {
                string nm = (string)sh.Name;
                if (find.Length > 0) nm = nm.Replace(find, with);
                nm = Xl.SafeSheetName(front + nm + back);
                if (nm.Length > 31) nm = nm.Substring(0, 31);
                if (nm == (string)sh.Name) return;
                if (Xl.SheetExists(wb, nm)) { skipped++; return; }
                sh.Name = nm;
            });
            if (skipped > 0) Say(skipped + " sheet(s) kept their name because the new name was already taken.");
        }

        void TabColour()
        {
            using (ColorDialog cd = new ColorDialog())
            {
                cd.FullOpen = true;
                if (cd.ShowDialog(this) != DialogResult.OK) return;
                int rgb = Design.Rgb(cd.Color.R, cd.Color.G, cd.Color.B);
                Each(delegate (dynamic sh) { sh.Tab.Color = rgb; });
            }
        }

        void MoveSheets(int dir)
        {
            List<string> names = Picked();
            if (names.Count == 0) { Say("Tick at least one sheet first."); return; }
            int count = (int)wb.Sheets.Count;
            if (dir < 0)
            {
                foreach (string nm in names)
                {
                    dynamic sh = wb.Sheets[nm];
                    int i = (int)sh.Index;
                    if (i > 1 && !names.Contains((string)wb.Sheets[i - 1].Name)) sh.Move(wb.Sheets[i - 1]);
                }
            }
            else
            {
                for (int k = names.Count - 1; k >= 0; k--)
                {
                    dynamic sh = wb.Sheets[names[k]];
                    int i = (int)sh.Index;
                    if (i < count && !names.Contains((string)wb.Sheets[i + 1].Name)) sh.Move(Type.Missing, wb.Sheets[i + 1]);
                }
            }
            Say("Moved.");
        }

        void SortAll()
        {
            List<string> names = new List<string>();
            foreach (dynamic sh in wb.Sheets) names.Add((string)sh.Name);
            names.Sort(StringComparer.OrdinalIgnoreCase);
            for (int i = names.Count - 1; i >= 0; i--) wb.Sheets[names[i]].Move(wb.Sheets[1]);
            Say("Sheets sorted A to Z.");
        }

        void DeleteSheets()
        {
            List<string> names = Picked();
            if (names.Count == 0) { Say("Tick at least one sheet first."); return; }
            if (names.Count >= (int)wb.Sheets.Count) { Say("A workbook needs at least one sheet. Untick one."); return; }
            if (MessageBox.Show(this, "Delete " + names.Count + (names.Count == 1 ? " sheet" : " sheets") + "?\n\n" + string.Join(", ", names.ToArray()) +
                                "\n\nA backup of the workbook is saved first (Help > Open Backups).", Design.AppName, MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) != DialogResult.OK) return;
            Xl.Backup(wb);
            dynamic app = Xl.App;
            bool alerts = (bool)app.DisplayAlerts;
            app.DisplayAlerts = false;
            try
            {
                foreach (string nm in names)
                {
                    dynamic sh = wb.Sheets[nm];
                    if ((int)sh.Visible != C.xlSheetVisible) sh.Visible = C.xlSheetVisible;
                    sh.Delete();
                }
            }
            finally { app.DisplayAlerts = alerts; }
            Say(names.Count + " deleted. A backup was saved first.");
        }
    }
}
