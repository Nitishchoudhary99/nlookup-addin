using System;
using System.Collections.Generic;
using System.Globalization;

namespace SmartPro
{
    // Everyday jobs on whole tables, sheets and workbooks.
    internal static class WorkbookTools
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        static string N(int n) { return n.ToString("#,##0", Inv); }

        static bool[] NumericCols(object[,] v)
        {
            int nr = v.GetLength(0), nc = v.GetLength(1);
            bool[] ok = new bool[nc + 1];
            int last = Math.Min(nr, 3000);
            for (int c = 1; c <= nc; c++)
            {
                int nb = 0, nn = 0;
                for (int r = 2; r <= last; r++)
                {
                    double d;
                    object x = v[r, c];
                    if (Xl.IsBlank(x) || Xl.IsErr(x) || x is DateTime) continue;
                    nb++;
                    if (Logic.TryNum(x, out d)) nn++;
                }
                ok[c] = nb > 0 && nn >= 0.8 * nb;
            }
            return ok;
        }

        // =================================================================
        //  MAKE TABLE NEAT
        // =================================================================
        public static void MakeNeat()
        {
            dynamic src = Xl.PickRange("Make Table Neat", "Select the table with its header row:");
            if (src == null) return;
            object[,] v = Xl.ReadTyped(src);
            int nr = v.GetLength(0), nc = v.GetLength(1);
            if (nr < 2)
            {
                Ui.Info("Select a table with a header row and data rows.");
                return;
            }
            string[] h = EverydayTools.Heads(v);
            bool[] numeric = NumericCols(v);
            Xl.Backup(src.Worksheet.Parent);
            using (new Speed())
            {
                dynamic app = Xl.App;
                dynamic ws = src.Worksheet;
                dynamic head = src.Rows[1];
                Xl.StyleHeader(head);
                head.VerticalAlignment = C.xlCenter;
                dynamic b = src.Borders;
                b.LineStyle = C.xlContinuous;
                b.Weight = 2;                                     // thin
                b.Color = 14800331;                               // light grey-blue
                for (int c = 1; c <= nc; c++)
                {
                    if (!numeric[c]) continue;
                    bool dec = false;
                    for (int r = 2; r <= nr && !dec; r++)
                    {
                        if (Xl.IsNum(v[r, c]) && Xl.Num(v[r, c]) != Math.Floor(Xl.Num(v[r, c]))) dec = true;
                    }
                    // codes such as invoice numbers must not get commas: only amount-like and quantity-like columns are formatted
                    if (!dec && Logic.MeasureRank(h[c]) > 2) continue;
                    dynamic body = src.Columns[c].Offset[1, 0].Resize[nr - 1, 1];
                    string nf = Convert.ToString((object)src.Cells[2, c].NumberFormat, Inv);
                    if (nf == "General") body.NumberFormat = dec ? "#,##0.00" : "#,##0";
                    body.HorizontalAlignment = -4152;             // right
                }
                src.EntireColumn.AutoFit();
                for (int c = 1; c <= nc; c++)
                {
                    try
                    {
                        double w = Convert.ToDouble((object)src.Columns[c].ColumnWidth, Inv);
                        if (w > 60) src.Columns[c].ColumnWidth = 60;
                    }
                    catch { }
                }
                bool isTable = false;
                try { isTable = src.Cells[1, 1].ListObject != null; } catch { }
                if (!isTable)
                {
                    object af = ws.AutoFilterMode;
                    if (!(af is bool && (bool)af))
                    {
                        try { src.AutoFilter(); } catch { }
                    }
                }
                try
                {
                    ws.Activate();
                    app.ActiveWindow.FreezePanes = false;
                    ws.Cells[(int)src.Row + 1, 1].Select();
                    app.ActiveWindow.FreezePanes = true;
                }
                catch { }
            }
            Ui.Info("Done: header styled, thin borders, number formats, column widths, filter buttons and a frozen header row.");
        }

        // =================================================================
        //  QUICK SUMMARY
        // =================================================================
        public static void QuickSummary()
        {
            dynamic src = Xl.PickRange("Quick Summary", "Select the table with its header row:");
            if (src == null) return;
            object[,] v = Xl.ReadTyped(src);
            int nr = v.GetLength(0), nc = v.GetLength(1);
            if (nr < 2)
            {
                Ui.Info("Select a table with a header row and data rows.");
                return;
            }
            string[] h = EverydayTools.Heads(v);
            int gc = Ui.PickColumn("Quick Summary (1 of 2)", "Summarise by which column?", h, null);
            if (gc == 0) return;
            bool[] numeric = NumericCols(v);
            numeric[gc] = false;
            bool any = false;
            foreach (bool x in numeric) if (x) any = true;
            int mc = 0;
            if (any)
            {
                mc = Ui.PickColumn("Quick Summary (2 of 2)", "Add up which number column?", h, numeric);
                if (mc == 0) return;
            }
            Dictionary<string, double[]> agg = new Dictionary<string, double[]>(StringComparer.OrdinalIgnoreCase);   // rows, sum, min, max
            List<string> names = new List<string>();
            for (int r = 2; r <= nr; r++)
            {
                if (Xl.RowBlank(v, r)) continue;
                string k = Xl.GroupKey(v[r, gc]);
                double[] a;
                if (!agg.TryGetValue(k, out a))
                {
                    a = new double[] { 0, 0, double.MaxValue, double.MinValue };
                    agg[k] = a;
                    names.Add(k);
                }
                a[0] += 1;
                double d;
                if (mc > 0 && Logic.TryNum(v[r, mc], out d))
                {
                    a[1] += d;
                    if (d < a[2]) a[2] = d;
                    if (d > a[3]) a[3] = d;
                }
            }
            if (names.Count == 0)
            {
                Ui.Info("The table has no data rows.");
                return;
            }
            int key = mc > 0 ? 1 : 0;
            names.Sort(delegate (string p, string q) { return agg[q][key].CompareTo(agg[p][key]); });
            int n = names.Count, cols = mc > 0 ? 6 : 2;
            object[,] t = Xl.New2D(n + 2, cols);
            t[1, 1] = h[gc];
            t[1, 2] = "Rows";
            if (mc > 0)
            {
                t[1, 3] = "Total " + h[mc];
                t[1, 4] = "Average";
                t[1, 5] = "Smallest";
                t[1, 6] = "Largest";
            }
            double allRows = 0, allSum = 0;
            for (int i = 0; i < n; i++)
            {
                double[] a = agg[names[i]];
                t[i + 2, 1] = names[i];
                t[i + 2, 2] = a[0];
                allRows += a[0];
                if (mc > 0)
                {
                    t[i + 2, 3] = a[1];
                    t[i + 2, 4] = a[0] > 0 ? a[1] / a[0] : 0;
                    if (a[2] != double.MaxValue) t[i + 2, 5] = a[2];
                    if (a[3] != double.MinValue) t[i + 2, 6] = a[3];
                    allSum += a[1];
                }
            }
            t[n + 2, 1] = "Total";
            t[n + 2, 2] = allRows;
            if (mc > 0)
            {
                t[n + 2, 3] = allSum;
                t[n + 2, 4] = allRows > 0 ? allSum / allRows : 0;
            }
            using (new Speed())
            {
                dynamic ws = Xl.AddSheet(src.Worksheet.Parent, "Summary");
                dynamic top = ws.Range["A1"];
                Xl.WriteTable(top, t, 1);
                Xl.StyleHeader(top.Resize[1, cols]);
                if (mc > 0) ws.Range["C2"].Resize[n + 1, 4].NumberFormat = "#,##0.00";
                ws.Range["B2"].Resize[n + 1, 1].NumberFormat = "#,##0";
                top.Resize[n + 2, cols].Rows[n + 2].Font.Bold = true;
                top.Resize[1, cols].EntireColumn.AutoFit();
                ws.Activate();
            }
            Ui.Info(N(n) + " groups summarised on the new 'Summary' sheet, biggest first." + (mc > 0 ? "\nTotal " + h[mc] + ": " + Logic.ShortINR(allSum) : ""));
        }

        // =================================================================
        //  COMPARE TWO LISTS
        // =================================================================
        static void Collect(object[,] v, Dictionary<string, string> map, List<string> order)
        {
            foreach (object x in v)
            {
                if (Xl.IsBlank(x) || Xl.IsErr(x)) continue;
                string k = Xl.Key(x);
                if (map.ContainsKey(k)) continue;
                map[k] = Xl.Str(x).Trim();
                order.Add(k);
            }
        }

        public static void CompareLists()
        {
            dynamic ra = Xl.PickRange("Compare Two Lists (1 of 2)", "Select the FIRST list (one column of names, codes or numbers):");
            if (ra == null) return;
            dynamic rb = Xl.PickRange("Compare Two Lists (2 of 2)", "Select the SECOND list. You can click another sheet or workbook:");
            if (rb == null) return;
            object[,] av = Xl.ReadTyped(ra);
            object[,] bv = Xl.ReadTyped(rb);
            Dictionary<string, string> a = new Dictionary<string, string>(), b = new Dictionary<string, string>();
            List<string> ao = new List<string>(), bo = new List<string>();
            Collect(av, a, ao);
            Collect(bv, b, bo);
            List<string> onlyA = new List<string>(), onlyB = new List<string>(), both = new List<string>();
            foreach (string k in ao)
            {
                if (b.ContainsKey(k)) both.Add(a[k]);
                else onlyA.Add(a[k]);
            }
            foreach (string k in bo) if (!a.ContainsKey(k)) onlyB.Add(b[k]);
            int n = Math.Max(1, Math.Max(onlyA.Count, Math.Max(onlyB.Count, both.Count)));
            if (n > 1048000)
            {
                Ui.Info("The lists are too long for one sheet.");
                return;
            }
            object[,] t = Xl.New2D(n + 1, 3);
            t[1, 1] = "Only in the first list (" + N(onlyA.Count) + ")";
            t[1, 2] = "Only in the second list (" + N(onlyB.Count) + ")";
            t[1, 3] = "In both lists (" + N(both.Count) + ")";
            for (int i = 0; i < onlyA.Count; i++) t[i + 2, 1] = onlyA[i];
            for (int i = 0; i < onlyB.Count; i++) t[i + 2, 2] = onlyB[i];
            for (int i = 0; i < both.Count; i++) t[i + 2, 3] = both[i];
            using (new Speed())
            {
                dynamic ws = Xl.AddSheet(rb.Worksheet.Parent, "List Compare");
                Xl.WriteTable(ws.Range["A1"], t, 1);
                Xl.StyleHeader(ws.Range["A1:C1"]);
                ws.Columns["A:C"].AutoFit();
                ws.Activate();
            }
            Ui.Info("Only in the first list: " + N(onlyA.Count) + "\nOnly in the second list: " + N(onlyB.Count) + "\nIn both: " + N(both.Count) +
                    "\n\nCapital letters and spaces at the ends are ignored. The three lists are on the new 'List Compare' sheet.");
        }

        // =================================================================
        //  COLOUR COUNT
        // =================================================================
        public static void ColourCount()
        {
            dynamic src = Xl.PickRange("Colour Count", "Select the coloured cells to count and add up:");
            if (src == null) return;
            long cells = Xl.CellCount(src);
            if (cells > 50000)
            {
                Ui.Info("Select up to 50,000 cells at a time. Reading a colour takes a moment for every cell.");
                return;
            }
            object[,] v = Xl.ReadTyped(src);
            int nr = v.GetLength(0), nc = v.GetLength(1);
            Dictionary<int, double[]> agg = new Dictionary<int, double[]>();     // cells, numbers, sum
            List<int> order = new List<int>();
            const int NoFill = -1;
            Xl.Status("reading the colours...");
            for (int r = 1; r <= nr; r++)
            {
                for (int c = 1; c <= nc; c++)
                {
                    dynamic cell = src.Cells[r, c];
                    int color = NoFill;
                    try
                    {
                        dynamic inside;
                        try { inside = cell.DisplayFormat.Interior; }      // also sees conditional-formatting colours
                        catch { inside = cell.Interior; }
                        int idx = Convert.ToInt32((object)inside.ColorIndex, Inv);
                        if (idx != -4142) color = Convert.ToInt32((object)inside.Color, Inv);
                    }
                    catch { }
                    double[] a;
                    if (!agg.TryGetValue(color, out a))
                    {
                        a = new double[3];
                        agg[color] = a;
                        order.Add(color);
                    }
                    a[0] += 1;
                    if (Xl.IsNum(v[r, c]))
                    {
                        a[1] += 1;
                        a[2] += Xl.Num(v[r, c]);
                    }
                }
            }
            try { Xl.App.StatusBar = false; } catch { }
            using (new Speed())
            {
                dynamic ws = Xl.AddSheet(src.Worksheet.Parent, "Colour Count");
                ws.Range["A1"].Value2 = "Colour";
                ws.Range["B1"].Value2 = "Cells";
                ws.Range["C1"].Value2 = "Cells with numbers";
                ws.Range["D1"].Value2 = "Sum of the numbers";
                Xl.StyleHeader(ws.Range["A1:D1"]);
                int row = 1;
                foreach (int color in order)
                {
                    row++;
                    double[] a = agg[color];
                    if (color == NoFill) ws.Cells[row, 1].Value2 = "No fill";
                    else ws.Cells[row, 1].Interior.Color = color;
                    ws.Cells[row, 2].Value2 = a[0];
                    ws.Cells[row, 3].Value2 = a[1];
                    ws.Cells[row, 4].Value2 = a[2];
                }
                ws.Range["D2"].Resize[order.Count, 1].NumberFormat = "#,##0.00";
                ws.Columns["A:D"].AutoFit();
                ws.Columns[1].ColumnWidth = 14;
                ws.Activate();
            }
            Ui.Info(N(order.Count) + " different colours found. The count and the sum for each colour are on the new 'Colour Count' sheet.");
        }

        // =================================================================
        //  VALUES-ONLY COPY
        // =================================================================
        public static void ValuesCopy()
        {
            dynamic app = Xl.App;
            dynamic wb = app.ActiveWorkbook;
            int pick = Ui.Pick("Values-Only Copy", "Make a copy without formulas (safe to email) of:", new[] { "This sheet only", "All visible sheets" }, null);
            if (pick < 0) return;
            int n = 0;
            using (new Speed())
            {
                dynamic nb = null;
                List<string> names = new List<string>();
                if (pick == 0) names.Add((string)app.ActiveSheet.Name);
                else
                {
                    foreach (dynamic sh in wb.Worksheets)
                    {
                        if ((int)sh.Visible == C.xlSheetVisible) names.Add((string)sh.Name);
                    }
                }
                foreach (string nm in names)
                {
                    dynamic sh = wb.Worksheets[nm];
                    if (nb == null)
                    {
                        sh.Copy();                                    // a new workbook with this one sheet
                        nb = app.ActiveWorkbook;
                    }
                    else sh.Copy(Xl.M, nb.Sheets[(int)nb.Sheets.Count]);
                    n++;
                }
                if (nb != null)
                {
                    foreach (dynamic sh in nb.Worksheets)
                    {
                        dynamic ur = sh.UsedRange;
                        try
                        {
                            ur.Copy();
                            ur.PasteSpecial(C.xlPasteValues);
                        }
                        catch
                        {
                            try { ur.Value2 = ur.Value2; } catch { }
                        }
                    }
                    app.CutCopyMode = false;
                    try
                    {
                        object ls = nb.LinkSources(1);
                        Array arr = ls as Array;
                        if (arr != null) foreach (object o in arr) nb.BreakLink(Convert.ToString(o, Inv), 1);
                    }
                    catch { }
                }
            }
            Ui.Info(N(n) + " sheet(s) copied into a new workbook with values only: no formulas and no links.\nIt is open now and not saved yet. Your own workbook was not changed.");
        }

        // =================================================================
        //  EMAIL THIS SHEET (a draft in Outlook, nothing is sent until the user presses Send)
        // =================================================================
        public static void EmailSheet()
        {
            dynamic app = Xl.App;
            dynamic ws = app.ActiveSheet;
            int pick = Ui.Pick("Email This Sheet", "Attach this sheet to a new Outlook email as:", new[] { "PDF", "Excel file with values only (no formulas)", "Both" }, null);
            if (pick < 0) return;
            Type outlook = Type.GetTypeFromProgID("Outlook.Application");
            if (outlook == null)
            {
                Ui.Info("The Outlook desktop app was not found on this PC.\nThis tool needs classic Outlook; it does not work with webmail or the new Outlook app.");
                return;
            }
            string sheet = (string)ws.Name;
            string stamp = DateTime.Now.ToString("dd-MMM-yyyy", Inv);
            string folder = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "DocBrisk_Mail");
            System.IO.Directory.CreateDirectory(folder);
            string stem = Xl.SafeFileName(sheet) + " " + stamp;
            List<string> files = new List<string>();
            using (new Speed())
            {
                if (pick == 0 || pick == 2)
                {
                    string pdf = System.IO.Path.Combine(folder, stem + ".pdf");
                    ws.ExportAsFixedFormat(C.xlTypePDF, pdf);
                    files.Add(pdf);
                }
                if (pick == 1 || pick == 2)
                {
                    string xl = System.IO.Path.Combine(folder, stem + ".xlsx");
                    ws.Copy();                                        // a new workbook with only this sheet
                    dynamic nb = app.ActiveWorkbook;
                    dynamic ur = nb.Worksheets[1].UsedRange;
                    try
                    {
                        ur.Copy();
                        ur.PasteSpecial(C.xlPasteValues);
                    }
                    catch
                    {
                        try { ur.Value2 = ur.Value2; } catch { }
                    }
                    app.CutCopyMode = false;
                    nb.SaveAs(xl, C.xlOpenXMLWorkbook);
                    nb.Close(false);
                    files.Add(xl);
                }
            }
            dynamic ol = Activator.CreateInstance(outlook);
            dynamic mail = ol.CreateItem(0);
            mail.Subject = sheet + " - " + stamp;
            mail.Body = "Please find attached: " + sheet + " (" + stamp + ").";
            foreach (string f in files) mail.Attachments.Add(f);
            mail.Display();
        }

        // =================================================================
        //  UNHIDE EVERYTHING
        // =================================================================
        public static void UnhideAll()
        {
            dynamic wb = Xl.App.ActiveWorkbook;
            if (!Ui.Ok("Show every hidden sheet, row and column in this workbook?")) return;
            int sheets = 0, failed = 0;
            using (new Speed())
            {
                foreach (dynamic sh in wb.Sheets)
                {
                    try
                    {
                        if ((int)sh.Visible != C.xlSheetVisible)
                        {
                            sh.Visible = C.xlSheetVisible;
                            sheets++;
                        }
                    }
                    catch { failed++; }
                }
                foreach (dynamic ws in wb.Worksheets)
                {
                    try
                    {
                        ws.Cells.EntireRow.Hidden = false;
                        ws.Cells.EntireColumn.Hidden = false;
                    }
                    catch { failed++; }
                }
            }
            Ui.Info(N(sheets) + " hidden sheets shown, and all rows and columns unhidden." + (failed > 0 ? "\n" + N(failed) + " sheets are protected and were skipped." : "") +
                    "\nFiltered rows stay filtered: clear the filter to see them.");
        }

        // =================================================================
        //  PROTECT / UNPROTECT ALL SHEETS
        // =================================================================
        public static void ProtectAll(bool protect)
        {
            dynamic wb = Xl.App.ActiveWorkbook;
            string title = protect ? "Protect All Sheets" : "Unprotect All Sheets";
            string pw = Ui.AskText(protect
                ? "Password for every sheet (leave empty for no password).\nKeep it safe: a lost password cannot be recovered."
                : "Password of the sheets (leave empty if there is none):", Design.AppName + " - " + title, "");
            if (pw == null) return;
            int ok = 0, failed = 0;
            using (new Speed())
            {
                foreach (dynamic ws in wb.Worksheets)
                {
                    try
                    {
                        if (protect) ws.Protect(pw);
                        else ws.Unprotect(pw);
                        ok++;
                    }
                    catch { failed++; }
                }
            }
            Ui.Info(N(ok) + " sheets " + (protect ? "protected." : "unprotected.") + (failed > 0 ? "\n" + N(failed) + " could not be changed" + (protect ? "." : " (wrong password).") : ""));
        }

        // =================================================================
        //  LINKS TO OTHER WORKBOOKS
        // =================================================================
        public static void Links()
        {
            dynamic wb = Xl.App.ActiveWorkbook;
            object ls = null;
            try { ls = wb.LinkSources(1); } catch { }
            Array arr = ls as Array;
            if (arr == null || arr.Length == 0)
            {
                Ui.Info("This workbook has no links to other workbooks.");
                return;
            }
            List<string> names = new List<string>();
            foreach (object o in arr) names.Add(Convert.ToString(o, Inv));
            string list = string.Join("\n", names.GetRange(0, Math.Min(names.Count, 12)).ToArray()) + (names.Count > 12 ? "\n..." : "");
            if (!Ui.Ok(N(names.Count) + " linked workbook(s):\n\n" + list + "\n\nBreak these links? Every formula that reads from them becomes a fixed value. A backup is saved first.")) return;
            Xl.Backup(wb);
            int n = 0;
            using (new Speed())
            {
                foreach (string s in names)
                {
                    try
                    {
                        wb.BreakLink(s, 1);
                        n++;
                    }
                    catch { }
                }
            }
            Ui.Info(N(n) + " of " + N(names.Count) + " links broken. The values stay, the formulas are gone.");
        }

        // =================================================================
        //  FILE DIET
        // =================================================================
        public static void FileDiet()
        {
            dynamic wb = Xl.App.ActiveWorkbook;
            List<string> sheet = new List<string>();
            List<int[]> job = new List<int[]>();                  // last row with content, last used row, last column with content, last used column
            string plan = "";
            foreach (dynamic ws in wb.Worksheets)
            {
                try
                {
                    dynamic ur = ws.UsedRange;
                    int usedR = (int)ur.Row + (int)ur.Rows.Count - 1, usedC = (int)ur.Column + (int)ur.Columns.Count - 1;
                    int lastR = 0, lastC = 0;
                    dynamic f1 = ws.Cells.Find("*", ws.Cells[1, 1], -4123, 2, 1, 2);      // formulas, part, by rows, previous
                    if (f1 != null)
                    {
                        lastR = (int)f1.Row;
                        dynamic f2 = ws.Cells.Find("*", ws.Cells[1, 1], -4123, 2, 2, 2);  // by columns
                        lastC = (int)f2.Column;
                    }
                    try
                    {
                        foreach (dynamic shp in ws.Shapes)                                  // pictures, charts and buttons count as content
                        {
                            int br = (int)shp.BottomRightCell.Row, bc = (int)shp.BottomRightCell.Column;
                            if (br > lastR) lastR = br;
                            if (bc > lastC) lastC = bc;
                        }
                    }
                    catch { }
                    bool rows = usedR - lastR > 100, cols = usedC - lastC > 30;
                    if (!rows && !cols) continue;
                    string nm = (string)ws.Name;
                    sheet.Add(nm);
                    job.Add(new[] { rows ? lastR : -1, usedR, cols ? lastC : -1, usedC });
                    if (sheet.Count <= 10) plan += "\n" + nm + ": " + (rows ? N(usedR - lastR) + " empty rows" : "") + (rows && cols ? ", " : "") + (cols ? N(usedC - lastC) + " empty columns" : "");
                }
                catch { }
            }
            if (sheet.Count == 0)
            {
                Ui.Info("No sheet carries unused rows or columns. There is nothing to trim.");
                return;
            }
            if (!Ui.Ok("These sheets carry formatted but empty rows or columns below and to the right of the data. They make the file big and slow:\n" + plan +
                       (sheet.Count > 10 ? "\n..." : "") + "\n\nRemove them? A backup is saved first.")) return;
            Xl.Backup(wb);
            int done = 0, failed = 0;
            using (new Speed())
            {
                for (int i = 0; i < sheet.Count; i++)
                {
                    try
                    {
                        dynamic ws = wb.Worksheets[sheet[i]];
                        int[] j = job[i];
                        if (j[0] >= 0) ws.Range[ws.Rows[j[0] + 1], ws.Rows[j[1]]].Delete();
                        if (j[2] >= 0) ws.Range[ws.Columns[j[2] + 1], ws.Columns[j[3]]].Delete();
                        done++;
                    }
                    catch { failed++; }
                }
            }
            Ui.Info(N(done) + " sheets trimmed." + (failed > 0 ? " " + N(failed) + " are protected and were skipped." : "") + "\n\nSave the workbook now: the file size drops when it is saved.");
        }

        // =================================================================
        //  SORT SHEETS / SHEETS FROM A LIST
        // =================================================================
        public static void SortSheets()
        {
            dynamic wb = Xl.App.ActiveWorkbook;
            List<string> names = new List<string>();
            foreach (dynamic sh in wb.Sheets) names.Add((string)sh.Name);
            if (names.Count < 2) return;
            if (!Ui.Ok("Put the " + names.Count + " sheet tabs of this workbook in A to Z order?")) return;
            names.Sort(StringComparer.OrdinalIgnoreCase);
            using (new Speed())
            {
                foreach (string nm in names)
                {
                    int count = (int)wb.Sheets.Count;
                    try { wb.Sheets[nm].Move(Xl.M, wb.Sheets[count]); } catch { }      // each one goes to the end, in order
                }
                wb.Sheets[1].Activate();
            }
            Ui.Info("The sheet tabs are now in A to Z order.");
        }

        public static void SheetsFromList()
        {
            dynamic src = Xl.PickRange("Sheets from a List", "Select the cells that hold the sheet names (one name per cell):");
            if (src == null) return;
            object[,] v = Xl.ReadTyped(src);
            List<string> names = new List<string>();
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (object x in v)
            {
                if (Xl.IsBlank(x) || Xl.IsErr(x)) continue;
                string s = Xl.SafeSheetName(Xl.Str(x).Trim());
                if (s.Length > 0 && seen.Add(s)) names.Add(s);
            }
            if (names.Count == 0)
            {
                Ui.Info("The selection has no names.");
                return;
            }
            if (names.Count > 200)
            {
                Ui.Info("The selection has " + N(names.Count) + " names. Create up to 200 sheets at a time.");
                return;
            }
            if (!Ui.Ok("Create " + names.Count + " new sheets, one for each name?")) return;
            dynamic wb = src.Worksheet.Parent;
            dynamic back = src.Worksheet;
            using (new Speed())
            {
                foreach (string s in names) Xl.AddSheet(wb, s);
                back.Activate();
            }
            Ui.Info(N(names.Count) + " sheets created. A name that already existed got a number added.");
        }
    }
}
