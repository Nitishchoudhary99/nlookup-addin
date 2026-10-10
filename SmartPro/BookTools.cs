using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace SmartPro
{
    internal static class BookTools
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        // ---------------------------------------------------------------
        //  backups
        // ---------------------------------------------------------------
        public static void BackupNow()
        {
            string f = Xl.Backup(Xl.App.ActiveWorkbook);
            if (f != null) Ui.Info("Backup saved:\n" + f);
            else Ui.Warn("The backup could not be saved. Save the workbook once, then try again.");
        }

        public static void OpenBackups()
        {
            System.Diagnostics.Process.Start("explorer.exe", "\"" + Xl.BackupFolder() + "\"");
        }

        // ---------------------------------------------------------------
        //  split
        // ---------------------------------------------------------------
        // copies the header look, column widths and number formats from the source table
        static void CopyLook(dynamic src, dynamic dest)
        {
            try
            {
                src.Rows[1].Copy();
                dest.Rows[1].PasteSpecial(C.xlPasteFormats);
                dest.Rows[1].PasteSpecial(C.xlPasteColumnWidths);
                Xl.App.CutCopyMode = false;
                int nc = (int)src.Columns.Count, nr = (int)dest.Rows.Count, srcRows = (int)src.Rows.Count;
                if (srcRows > 1 && nr > 1)
                {
                    for (int j = 1; j <= nc; j++)
                    {
                        string nf = Convert.ToString((object)dest.Cells[2, j].NumberFormat);
                        if (nf != "@") dest.Cells[2, j].Resize[nr - 1, 1].NumberFormat = src.Cells[2, j].NumberFormat;
                    }
                }
            }
            catch { }
        }

        public static void SplitByColumn(bool toFiles)
        {
            string title = toFiles ? "Split to Files" : "Split to Sheets";
            dynamic src = Xl.PickRange(title);
            if (src == null) return;
            object[,] v = Xl.ReadTyped(src);
            int nr = v.GetLength(0), nc = v.GetLength(1);
            if (nr < 2)
            {
                Ui.Info("Select a table with a header row and data rows.");
                return;
            }
            string[] heads = new string[nc + 1];
            for (int j = 1; j <= nc; j++) heads[j] = Xl.HeaderText(v[1, j], j);
            int kc = Ui.PickColumn(title, "Split by which column?", heads, null);
            if (kc == 0) return;

            Dictionary<string, List<int>> groups = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);
            List<string> order = new List<string>();
            for (int r = 2; r <= nr; r++)
            {
                string key = Xl.GroupKey(v[r, kc]);
                List<int> lst;
                if (!groups.TryGetValue(key, out lst))
                {
                    lst = new List<int>();
                    groups[key] = lst;
                    order.Add(key);
                }
                lst.Add(r);
            }
            if (order.Count > 60 && !Ui.Ok("This will create " + order.Count + (toFiles ? " files" : " sheets") + ". Continue?")) return;
            string folder = null;
            if (toFiles)
            {
                folder = Ui.PickFolder("Choose the folder for the split files");
                if (folder == null) return;
            }
            dynamic wb = src.Worksheet.Parent;
            using (new Speed())
            {
                int done = 0;
                foreach (string g in order)
                {
                    done++;
                    Xl.Status("writing " + done + " of " + order.Count);
                    List<int> rows = groups[g];
                    object[,] o = Xl.New2D(rows.Count + 1, nc);
                    for (int j = 1; j <= nc; j++) o[1, j] = heads[j];
                    for (int i = 0; i < rows.Count; i++)
                        for (int j = 1; j <= nc; j++) o[i + 2, j] = v[rows[i], j];
                    dynamic nb = null;
                    dynamic ws;
                    if (toFiles)
                    {
                        nb = Xl.App.Workbooks.Add(C.xlWBATWorksheet);
                        ws = nb.Worksheets[1];
                        ws.Name = Xl.SafeSheetName(g);
                    }
                    else ws = Xl.AddSheet(wb, g);
                    Xl.WriteTable(ws.Range["A1"], o, 1);
                    CopyLook(src, ws.Range["A1"].Resize[rows.Count + 1, nc]);
                    if (toFiles)
                    {
                        nb.SaveAs(Path.Combine(folder, Xl.SafeFileName(g) + ".xlsx"), C.xlOpenXMLWorkbook);
                        nb.Close(false);
                    }
                }
                if (!toFiles) src.Worksheet.Activate();
            }
            if (toFiles) Ui.Info(order.Count + " files saved in:\n" + folder);
            else Ui.Info(order.Count + " sheets created, one for each value of '" + heads[kc] + "'.");
        }

        // ---------------------------------------------------------------
        //  combine
        // ---------------------------------------------------------------
        // blocks: source name + its values. Columns are matched by header name.
        static string WriteCombined(List<KeyValuePair<string, object[,]>> blocks, dynamic ws, string srcCaption)
        {
            Dictionary<string, int> hd = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            List<string> hdrs = new List<string>();
            long total = 0;
            foreach (KeyValuePair<string, object[,]> b in blocks)
            {
                object[,] v = b.Value;
                total += v.GetLength(0) - 1;
                for (int j = 1; j <= v.GetLength(1); j++)
                {
                    string key = Xl.HeaderText(v[1, j], j);
                    if (hd.ContainsKey(key)) continue;
                    hdrs.Add(key);
                    hd[key] = hdrs.Count;
                }
            }
            if (total > 1048000) throw new InvalidOperationException("The combined data has " + total.ToString("#,##0", Inv) + " rows, more than one sheet can hold.");
            if (total < 1) throw new InvalidOperationException("No data rows were found under the headers.");
            int nh = hdrs.Count;
            object[,] o = Xl.New2D((int)total + 1, nh + 1);
            o[1, 1] = srcCaption;
            for (int j = 1; j <= nh; j++) o[1, j + 1] = hdrs[j - 1];
            int n = 1;
            foreach (KeyValuePair<string, object[,]> b in blocks)
            {
                object[,] v = b.Value;
                int vc = v.GetLength(1);
                int[] map = new int[vc + 1];
                for (int j = 1; j <= vc; j++) map[j] = hd[Xl.HeaderText(v[1, j], j)];
                for (int r = 2; r <= v.GetLength(0); r++)
                {
                    if (Xl.RowBlank(v, r)) continue;
                    n++;
                    o[n, 1] = b.Key;
                    for (int j = 1; j <= vc; j++) o[n, map[j] + 1] = v[r, j];
                }
            }
            Xl.WriteTable(ws.Range["A1"], o, 1);
            Xl.StyleHeader(ws.Range["A1"].Resize[1, nh + 1]);
            ws.Range["A1"].Resize[1, nh + 1].EntireColumn.AutoFit();
            ws.Activate();
            return (n - 1).ToString("#,##0", Inv) + " rows from " + blocks.Count + " sources combined into " + nh + " columns.\nColumn A shows where each row came from.";
        }

        public static void CombineSheets()
        {
            dynamic app = Xl.App;
            dynamic wb = app.ActiveWorkbook;
            List<KeyValuePair<string, object[,]>> blocks = new List<KeyValuePair<string, object[,]>>();
            foreach (dynamic sh in wb.Worksheets)
            {
                if ((int)sh.Visible != C.xlSheetVisible) continue;
                double cnt = (double)app.WorksheetFunction.CountA(sh.UsedRange);
                if (cnt <= 0) continue;
                string nm = (string)sh.Name;
                object[,] arr = Xl.ReadTyped(sh.UsedRange);
                blocks.Add(new KeyValuePair<string, object[,]>(nm, arr));
            }
            if (blocks.Count < 2)
            {
                Ui.Info("This workbook needs at least 2 visible sheets with data.");
                return;
            }
            if (!Ui.Ok("Combine " + blocks.Count + " visible sheets into one new sheet?\n\nThe first used row of each sheet is its header. Columns are matched by header name, so the column order may differ between sheets.")) return;
            string msg;
            using (new Speed())
            {
                dynamic ws = Xl.AddSheet(wb, "Combined");
                msg = WriteCombined(blocks, ws, "Source Sheet");
            }
            Ui.Info(msg);
        }

        public static void CombineFolder()
        {
            string folder = Ui.PickFolder("Choose the folder with the Excel / CSV files");
            if (folder == null) return;
            List<string> files = new List<string>();
            foreach (string f in Directory.GetFiles(folder))
            {
                string name = Path.GetFileName(f);
                string ext = Path.GetExtension(f).ToLowerInvariant();
                if (name.StartsWith("~$", StringComparison.Ordinal)) continue;
                if (ext == ".xlsx" || ext == ".xlsm" || ext == ".xlsb" || ext == ".xls" || ext == ".csv") files.Add(f);
            }
            files.Sort(StringComparer.OrdinalIgnoreCase);
            if (files.Count == 0)
            {
                Ui.Info("No Excel or CSV files were found in that folder.");
                return;
            }
            if (!Ui.Ok("Combine the first sheet of " + files.Count + " files into one new workbook?\n\nColumns are matched by header name. Macros inside those files are not run.")) return;
            string msg = null;
            int nFail = 0;
            using (new Speed())
            {
                dynamic app = Xl.App;
                object sec = null;
                try
                {
                    sec = app.AutomationSecurity;
                    app.AutomationSecurity = 3;
                }
                catch { }
                List<KeyValuePair<string, object[,]>> blocks = new List<KeyValuePair<string, object[,]>>();
                try
                {
                    foreach (string f in files)
                    {
                        string name = Path.GetFileName(f);
                        Xl.Status("reading " + name);
                        dynamic wbk = null;
                        bool wasOpen = false;
                        try
                        {
                            foreach (dynamic w in app.Workbooks)
                            {
                                if (string.Equals((string)w.Name, name, StringComparison.OrdinalIgnoreCase))
                                {
                                    wbk = w;
                                    wasOpen = true;
                                }
                            }
                            if (wbk == null) wbk = app.Workbooks.Open(f, 0, true);
                            dynamic sh = wbk.Worksheets[1];
                            double cnt = (double)app.WorksheetFunction.CountA(sh.UsedRange);
                            if (cnt > 0)
                            {
                                object[,] arr = Xl.ReadTyped(sh.UsedRange);
                                blocks.Add(new KeyValuePair<string, object[,]>(name, arr));
                            }
                        }
                        catch { nFail++; }
                        if (wbk != null && !wasOpen)
                        {
                            try { wbk.Close(false); } catch { }
                        }
                    }
                }
                finally
                {
                    try { if (sec != null) app.AutomationSecurity = sec; } catch { }
                }
                if (blocks.Count > 0)
                {
                    dynamic nb = app.Workbooks.Add(C.xlWBATWorksheet);
                    dynamic ws = nb.Worksheets[1];
                    ws.Name = "Combined";
                    msg = WriteCombined(blocks, ws, "Source File");
                }
            }
            if (msg == null) Ui.Warn("None of the files could be read.");
            else Ui.Info(msg + (nFail > 0 ? "\n\n" + nFail + " files could not be opened and were skipped." : ""));
        }

        // ---------------------------------------------------------------
        //  index and PDF
        // ---------------------------------------------------------------
        public static void SheetIndex()
        {
            dynamic wb = Xl.App.ActiveWorkbook;
            using (new Speed())
            {
                dynamic ws = Xl.AddSheet(wb, "Index");
                ws.Move(wb.Sheets[1]);
                string me = (string)ws.Name;
                ws.Range["A1"].Value2 = "Sheet index: " + (string)wb.Name;
                ws.Range["A1"].Font.Size = 14;
                ws.Range["A1"].Font.Bold = true;
                ws.Range["A3"].Value2 = "No.";
                ws.Range["B3"].Value2 = "Sheet (click to open)";
                ws.Range["C3"].Value2 = "Rows used";
                ws.Range["D3"].Value2 = "Columns used";
                Xl.StyleHeader(ws.Range["A3:D3"]);
                int r = 3;
                foreach (dynamic sh in wb.Sheets)
                {
                    string nm = (string)sh.Name;
                    if (nm == me || (int)sh.Visible != C.xlSheetVisible) continue;
                    r++;
                    ws.Cells[r, 1].Value2 = r - 3;
                    int ur = 0, uc = 0;
                    bool isSheet = true;
                    try
                    {
                        ur = (int)sh.UsedRange.Rows.Count;
                        uc = (int)sh.UsedRange.Columns.Count;
                    }
                    catch { isSheet = false; }
                    if (isSheet)
                    {
                        ws.Hyperlinks.Add(ws.Cells[r, 2], "", "'" + nm.Replace("'", "''") + "'!A1", Xl.M, nm);
                        ws.Cells[r, 3].Value2 = ur;
                        ws.Cells[r, 4].Value2 = uc;
                    }
                    else
                    {
                        ws.Cells[r, 2].Value2 = "'" + nm;
                        ws.Cells[r, 3].Value2 = "(chart sheet)";
                    }
                }
                ws.Columns["A:D"].AutoFit();
                ws.Activate();
            }
        }

        public static void SheetsToPdf()
        {
            dynamic app = Xl.App;
            dynamic wb = app.ActiveWorkbook;
            string folder = Ui.PickFolder("Choose the folder for the PDF files");
            if (folder == null) return;
            int n = 0;
            using (new Speed())
            {
                foreach (dynamic ws in wb.Worksheets)
                {
                    if ((int)ws.Visible != C.xlSheetVisible) continue;
                    double cnt = (double)app.WorksheetFunction.CountA(ws.UsedRange);
                    int charts = 0;
                    try { charts = (int)ws.ChartObjects().Count; } catch { }
                    if (cnt <= 0 && charts == 0) continue;
                    string nm = (string)ws.Name;
                    Xl.Status("saving " + nm);
                    ws.ExportAsFixedFormat(C.xlTypePDF, Path.Combine(folder, Xl.SafeFileName(nm) + ".pdf"));
                    n++;
                }
            }
            Ui.Info(n + " PDF files saved in:\n" + folder);
        }

        // ---------------------------------------------------------------
        //  India pack
        // ---------------------------------------------------------------
        public static void IndianFormat()
        {
            dynamic app = Xl.App;
            dynamic sel;
            bool hasDec = false;
            try
            {
                sel = app.Selection;
                dynamic chk = app.Intersect(sel.Areas[1], sel.Worksheet.UsedRange);
                if (chk != null)
                {
                    long cells = Xl.CellCount(chk);
                    if (cells <= 500000)
                    {
                        object[,] v = Xl.ReadTyped(chk);
                        foreach (object x in v)
                        {
                            if (!Xl.IsNum(x)) continue;
                            double d = Xl.Num(x);
                            if (d != Math.Floor(d))
                            {
                                hasDec = true;
                                break;
                            }
                        }
                    }
                }
            }
            catch
            {
                Ui.Info("Select the cells with numbers first.");
                return;
            }
            sel.NumberFormat = hasDec
                ? "[>=10000000]##\\,##\\,##\\,##0.00;[>=100000]##\\,##\\,##0.00;##,##0.00"
                : "[>=10000000]##\\,##\\,##\\,##0;[>=100000]##\\,##\\,##0;##,##0";
        }

        public static void AddFiscalYear()
        {
            dynamic app = Xl.App;
            string def = "";
            try { def = (string)app.ActiveCell.Address; } catch { }
            object res;
            try { res = app.InputBox("Click any cell in the DATE column of your table:", "Add Fiscal Year", def, Xl.M, Xl.M, Xl.M, Xl.M, 8); }
            catch { return; }
            if (res == null || res is bool) return;
            dynamic src = res;
            dynamic ws = src.Worksheet;
            int c = (int)src.Cells[1, 1].Column;
            dynamic reg = src.Cells[1, 1].CurrentRegion;
            int r1 = (int)reg.Row;
            int r2 = r1 + (int)reg.Rows.Count - 1;
            if (r2 <= r1)
            {
                Ui.Info("Click a cell inside a table that has a header row and data rows.");
                return;
            }
            Xl.Backup(ws.Parent);
            using (new Speed())
            {
                ws.Range[ws.Columns[c + 1], ws.Columns[c + 2]].Insert();
                ws.Cells[r1, c + 1].Value2 = "Fiscal Year";
                ws.Cells[r1, c + 2].Value2 = "FY Quarter";
                dynamic both = ws.Range[ws.Cells[r1 + 1, c + 1], ws.Cells[r2, c + 2]];
                both.NumberFormat = "General";
                both.HorizontalAlignment = C.xlLeft;
                ws.Range[ws.Cells[r1 + 1, c + 1], ws.Cells[r2, c + 1]].FormulaR1C1 =
                    "=IF(ISNUMBER(RC[-1]),\"FY \"&(YEAR(RC[-1])-(MONTH(RC[-1])<4))&\"-\"&TEXT(MOD(YEAR(RC[-1])-(MONTH(RC[-1])<4)+1,100),\"00\"),\"\")";
                ws.Range[ws.Cells[r1 + 1, c + 2], ws.Cells[r2, c + 2]].FormulaR1C1 =
                    "=IF(ISNUMBER(RC[-2]),\"Q\"&(INT(MOD(MONTH(RC[-2])-4,12)/3)+1),\"\")";
                ws.Range[ws.Columns[c + 1], ws.Columns[c + 2]].AutoFit();
            }
            Ui.Info("Two columns added: Fiscal Year (April to March) and FY Quarter.\nThey are normal Excel formulas, so the file works on any PC.\n\n" +
                    "Blank results mean the date is stored as text: run Text to Dates on that column first.");
        }

        public static void RupeesInWords()
        {
            dynamic app = Xl.App;
            dynamic src;
            object[,] v, right;
            try
            {
                src = app.Intersect(app.Selection, app.ActiveSheet.UsedRange);
                if (src == null) return;
                src = src.Areas[1];
                long cells = Xl.CellCount(src);
                if (cells > 5000)
                {
                    Ui.Info("Select up to 5,000 amount cells at a time.");
                    return;
                }
                v = Xl.ReadTyped(src);
                right = Xl.ReadTyped(src.Offset[0, 1]);
            }
            catch
            {
                Ui.Info("Select the cells with amounts first.");
                return;
            }
            int n = 0, skipped = 0;
            using (new Speed())
            {
                int nr = v.GetLength(0), nc = v.GetLength(1);
                for (int c = nc; c >= 1; c--)
                    for (int r = 1; r <= nr; r++)
                    {
                        if (!Xl.IsNum(v[r, c])) continue;
                        if (right[r, c] != null)
                        {
                            skipped++;
                            continue;
                        }
                        src.Cells[r, c + 1].Value2 = Logic.RupeeWords(Xl.Num(v[r, c]));
                        n++;
                    }
            }
            Ui.Info(n + " amounts written in words in the cell to the right.\n" +
                    (skipped > 0 ? skipped + " skipped because the cell to the right was not empty.\n" : "") +
                    "\nFor a live formula type  =RUPEEWORDS(A2)");
        }

        public static void CheckGstin()
        {
            dynamic src = Xl.PickRange("Check GSTIN", "Select the cells that contain GSTIN numbers:");
            if (src == null) return;
            object[,] v = Xl.ReadTyped(src);
            int nr = v.GetLength(0), nc = v.GetLength(1);
            List<long> good = new List<long>(), bad = new List<long>();
            for (int c = 1; c <= nc; c++)
            {
                int start = 1;
                string head = v[1, c] as string;
                if (nr > 1 && head != null && !Logic.GstinOk(head.Trim().ToUpperInvariant())) start = 2;
                for (int r = start; r <= nr; r++)
                {
                    string s = v[r, c] as string;
                    if (s == null) continue;
                    s = s.Trim().ToUpperInvariant();
                    if (s.Length == 0) continue;
                    if (Logic.GstinOk(s)) good.Add(Xl.Cell(r, c));
                    else bad.Add(Xl.Cell(r, c));
                }
            }
            using (new Speed())
            {
                Xl.Paint(src, good, Design.Good);
                Xl.Paint(src, bad, Design.Bad);
            }
            Ui.Info(good.Count + " valid GSTIN (green)\n" + bad.Count + " not valid (red)\n\n" +
                    "The check covers the format, the state code and the check digit. It does not confirm that the GSTIN is active on the GST portal.");
        }

        // ---------------------------------------------------------------
        //  help
        // ---------------------------------------------------------------
        public static void FunctionList()
        {
            dynamic app = Xl.App;
            dynamic wb = app.ActiveWorkbook;
            if (wb == null) wb = app.Workbooks.Add();
            string[][] rows =
            {
                new[] { "RUPEEWORDS", "Amount in words, Indian style", "=RUPEEWORDS(125000.5)  ->  Rupees One Lakh Twenty Five Thousand and Fifty Paise Only" },
                new[] { "INRFORMAT", "Number as text with lakh / crore commas", "=INRFORMAT(1234567.5)  ->  12,34,567.50" },
                new[] { "INRSHORT", "Short amount in K / L / Cr", "=INRSHORT(12500000)  ->  1.25 Cr" },
                new[] { "FYEAR", "Fiscal year of a date (April start; change with a 2nd argument)", "=FYEAR(A2)  ->  FY 2026-27" },
                new[] { "FYQUARTER", "Fiscal quarter of a date", "=FYQUARTER(A2)  ->  Q3" },
                new[] { "GSTINVALID", "TRUE when the GSTIN format, state code and check digit are right", "=GSTINVALID(A2)" },
                new[] { "GSTSTATE", "State name from a GSTIN or a 2-digit state code", "=GSTSTATE(A2)  ->  Haryana" },
                new[] { "PANVALID", "TRUE when the PAN format is right", "=PANVALID(A2)" },
                new[] { "SIMILARITY", "How alike two texts are, 0 to 1", "=SIMILARITY(A2,B2)  ->  0.93" },
                new[] { "FUZZYLOOKUP", "Lookup that tolerates typos and Pvt/Ltd differences", "=FUZZYLOOKUP(A2, Master!A2:A5000, Master!B2:B5000, 0.85)" },
                new[] { "EXTRACTNUM", "The 1st (or nth) number inside a text", "=EXTRACTNUM(\"Qty 120 kg\")  ->  120" },
                new[] { "TEXTBETWEEN", "Text between two markers", "=TEXTBETWEEN(A2,\"(\",\")\")" },
                new[] { "NUMWORDS", "Whole number in words, Indian style", "=NUMWORDS(125000)  ->  One Lakh Twenty Five Thousand" },
                new[] { "GSTADD", "Amount with GST added", "=GSTADD(1000,18)  ->  1180" },
                new[] { "GSTBASE", "Amount before GST from an inclusive amount", "=GSTBASE(1180,18)  ->  1000" },
                new[] { "GSTTAX", "The GST part of an amount (3rd argument TRUE if GST is included)", "=GSTTAX(1180,18,TRUE)  ->  180" },
                new[] { "AADHAARVALID", "TRUE when the 12-digit Aadhaar check digit is right", "=AADHAARVALID(A2)" },
                new[] { "IFSCVALID", "TRUE when the IFSC format is right", "=IFSCVALID(A2)" },
                new[] { "MOBILEVALID", "TRUE for a 10-digit Indian mobile number", "=MOBILEVALID(A2)" },
                new[] { "PINVALID", "TRUE for a 6-digit PIN code", "=PINVALID(A2)" },
                new[] { "EMAILVALID", "TRUE when the text looks like an email address", "=EMAILVALID(A2)" },
                new[] { "NTHWORD", "The nth word of a text (-1 is the last word)", "=NTHWORD(A2,-1)" },
                new[] { "WORDCOUNT", "Number of words", "=WORDCOUNT(A2)" },
                new[] { "ONLYDIGITS", "Keeps only the digits", "=ONLYDIGITS(\"PO-4500/123\")  ->  4500123" },
                new[] { "ONLYLETTERS", "Keeps only letters and spaces", "=ONLYLETTERS(A2)" },
                new[] { "CLEANSPACES", "Removes extra and non-breaking spaces", "=CLEANSPACES(A2)" },
                new[] { "PROPERNAME", "Tidy capitals for names", "=PROPERNAME(\"TATA STEEL PVT LTD\")  ->  Tata Steel Pvt Ltd" },
                new[] { "AGEDAYS", "Days from a date to today", "=AGEDAYS(A2)" },
                new[] { "AGEBUCKET", "Ageing bucket: 0-30, 31-60, 61-90, 91-180, 180+", "=AGEBUCKET(A2)" },
                new[] { "MONTHKEY", "Month label that sorts correctly", "=MONTHKEY(A2)  ->  2026-04 (Apr)" },
                new[] { "PCTCHANGE", "Change from old to new", "=PCTCHANGE(100,125)  ->  0.25" }
            };
            using (new Speed())
            {
                dynamic ws = Xl.AddSheet(wb, Design.AppName + " Functions");
                ws.Range["A1"].Value2 = Design.AppName + " functions (type them in any cell like normal Excel formulas)";
                ws.Range["A1"].Font.Size = 14;
                ws.Range["A1"].Font.Bold = true;
                object[,] t = Xl.New2D(rows.Length + 1, 3);
                t[1, 1] = "Function";
                t[1, 2] = "What it gives";
                t[1, 3] = "Example";
                for (int i = 0; i < rows.Length; i++)
                    for (int j = 0; j < 3; j++) t[i + 2, j + 1] = rows[i][j];
                Xl.WriteTable(ws.Range["A3"], t, 1);
                Xl.StyleHeader(ws.Range["A3:C3"]);
                ws.Cells[rows.Length + 5, 1].Value2 = "These functions need " + Design.AppName + " installed on the PC that opens the file. Use Formulas to Values before sending a file to someone who does not have it.";
                ws.Columns["A:C"].AutoFit();
                ws.Activate();
            }
        }
    }
}
