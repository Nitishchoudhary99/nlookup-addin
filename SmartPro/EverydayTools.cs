using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace SmartPro
{
    // Small jobs that eat time every day and that Excel has no single button for.
    // Rows and cells, text, numbers and formulas. All of them save a backup before changing data.
    internal static class EverydayTools
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        const int xlShiftDown = -4121;

        static dynamic PickCell(string title, string prompt)
        {
            object res;
            try { res = Xl.App.InputBox(prompt, title, Xl.M, Xl.M, Xl.M, Xl.M, Xl.M, 8); }
            catch { return null; }
            if (res == null || res is bool) return null;
            dynamic r = res;
            return r.Cells[1, 1];
        }

        public static string[] Heads(object[,] v)
        {
            int nc = v.GetLength(1);
            string[] h = new string[nc + 1];
            for (int c = 1; c <= nc; c++) h[c] = Xl.HeaderText(v[1, c], c);
            return h;
        }

        static string N(int n) { return n.ToString("#,##0", Inv); }

        static bool IsFormula(object f)
        {
            string s = f as string;
            return s != null && s.Length > 1 && s[0] == '=';
        }

        static bool IsFiltered(dynamic ws)
        {
            object f = ws.FilterMode;
            return f is bool && (bool)f;
        }

        // =================================================================
        //  PASTE TO VISIBLE CELLS
        // =================================================================
        static bool IsHidden(dynamic ws, int row)
        {
            object h = ws.Rows[row].Hidden;
            return h is bool && (bool)h;
        }

        public static void PasteToVisible()
        {
            dynamic src = Xl.PickRange("Paste to Visible Cells (1 of 2)", "Select the cells to copy:");
            if (src == null) return;
            object[,] v = Xl.ReadTyped(src);
            int nr = v.GetLength(0), nc = v.GetLength(1);
            if ((long)nr * nc > 50000)
            {
                Ui.Info("Select up to 50,000 cells at a time.");
                return;
            }
            dynamic first = PickCell("Paste to Visible Cells (2 of 2)", "Click the FIRST cell where the values should go.\nHidden and filtered-out rows are skipped.");
            if (first == null) return;
            dynamic ws = first.Worksheet;
            int row = (int)first.Row, col = (int)first.Column, maxRow = (int)ws.Rows.Count;
            Xl.Backup(ws.Parent);
            int done = 0;
            using (new Speed())
            {
                for (int i = 1; i <= nr; i++)
                {
                    while (row <= maxRow)
                    {
                        bool hid = IsHidden(ws, row);
                        if (!hid) break;
                        row++;
                    }
                    if (row > maxRow) break;
                    for (int j = 1; j <= nc; j++)
                    {
                        object x = v[i, j];
                        dynamic cell = ws.Cells[row, col + j - 1];
                        if (x == null) cell.ClearContents();
                        else if (!Xl.IsErr(x)) CleanTools.PutValue(cell, x);
                    }
                    row++;
                    done++;
                }
            }
            Ui.Info(N(done) + " rows pasted into visible rows only. Hidden rows were not touched.");
        }

        // =================================================================
        //  DELETE ROWS WHERE ...
        // =================================================================
        static bool RowMatch(object x, int mode, string arg, double lim)
        {
            double d;
            switch (mode)
            {
                case 0: return !Xl.IsBlank(x) && Xl.Str(x).IndexOf(arg, StringComparison.OrdinalIgnoreCase) >= 0;
                case 1: return !Xl.IsBlank(x) && string.Equals(Xl.Str(x).Trim(), arg, StringComparison.OrdinalIgnoreCase);
                case 2: return Xl.IsBlank(x);
                case 3: return Logic.TryNum(x, out d) && d == 0;
                case 4: return Logic.TryNum(x, out d) && d < lim;
                default: return Logic.TryNum(x, out d) && d > lim;
            }
        }

        public static void DeleteRowsWhere()
        {
            dynamic src = Xl.PickRange("Delete Rows Where", "Select the table with its header row:");
            if (src == null) return;
            if (IsFiltered(src.Worksheet))
            {
                Ui.Info("This sheet is filtered. Clear the filter first, then run the tool again.");
                return;
            }
            object[,] v = Xl.ReadTyped(src);
            int nr = v.GetLength(0), nc = v.GetLength(1);
            if (nr < 2)
            {
                Ui.Info("Select a table with a header row and data rows.");
                return;
            }
            string[] h = Heads(v);
            int col = Ui.PickColumn("Delete Rows Where", "Check which column?", h, null);
            if (col == 0) return;
            int mode = Ui.Pick("Delete Rows Where", "Delete every row where '" + h[col] + "':", new[]
            {
                "contains a text ...", "is exactly a text ...", "is blank", "is zero", "is a number below ...", "is a number above ..."
            }, null);
            if (mode < 0) return;
            string arg = "";
            double lim = 0;
            if (mode == 0 || mode == 1)
            {
                arg = Ui.AskText("Type the text to look for in '" + h[col] + "':", Design.AppName + " - Delete Rows Where", "");
                if (arg == null || arg.Trim().Length == 0) return;
                arg = arg.Trim();
            }
            else if (mode >= 4)
            {
                string a = Ui.AskText("Type the number:", Design.AppName + " - Delete Rows Where", "0");
                if (a == null || !Logic.TryNumText(a.Trim(), out lim)) return;
            }
            List<int> hit = new List<int>();
            for (int r = 2; r <= nr; r++) if (RowMatch(v[r, col], mode, arg, lim)) hit.Add(r);
            if (hit.Count == 0)
            {
                Ui.Info("No rows match. Nothing was deleted.");
                return;
            }
            if (!Ui.Ok(N(hit.Count) + " of " + N(nr - 1) + " rows will be deleted. A backup is saved first.\n\nContinue?")) return;
            Xl.Backup(src.Worksheet.Parent);
            using (new Speed())
            {
                dynamic ws = src.Worksheet;
                int i = hit.Count - 1;
                while (i >= 0)                                     // bottom-up, one block of neighbouring rows at a time
                {
                    int r2 = hit[i], r1 = r2;
                    while (i > 0 && hit[i - 1] == r1 - 1)
                    {
                        i--;
                        r1--;
                    }
                    ws.Range[src.Cells[r1, 1], src.Cells[r2, nc]].Delete(C.xlShiftUp);
                    i--;
                }
            }
            Ui.Info(N(hit.Count) + " rows deleted.");
        }

        // =================================================================
        //  BLANK ROW AT EACH CHANGE
        // =================================================================
        public static void BlankRowsAtChange()
        {
            dynamic src = Xl.PickRange("Blank Row at Each Change", "Select the table with its header row (sort it by the column first):");
            if (src == null) return;
            if (IsFiltered(src.Worksheet))
            {
                Ui.Info("This sheet is filtered. Clear the filter first, then run the tool again.");
                return;
            }
            object[,] v = Xl.ReadTyped(src);
            int nr = v.GetLength(0), nc = v.GetLength(1);
            if (nr < 3)
            {
                Ui.Info("Select a table with a header row and at least 2 data rows.");
                return;
            }
            string[] h = Heads(v);
            int col = Ui.PickColumn("Blank Row at Each Change", "Add an empty row whenever which column changes?", h, null);
            if (col == 0) return;
            Xl.Backup(src.Worksheet.Parent);
            int n = 0;
            using (new Speed())
            {
                dynamic ws = src.Worksheet;
                for (int r = nr; r >= 3; r--)
                {
                    if (Xl.Key(v[r, col]) == Xl.Key(v[r - 1, col])) continue;
                    ws.Range[src.Cells[r, 1], src.Cells[r, nc]].Insert(xlShiftDown);
                    n++;
                }
            }
            Ui.Info(N(n) + " empty rows added, one before each new '" + h[col] + "'.");
        }

        // =================================================================
        //  MERGE SAME VALUES
        // =================================================================
        public static void MergeSameValues()
        {
            dynamic src = Xl.PickRange("Merge Same Values", "Select the column(s) where equal cells under each other should become one merged cell:");
            if (src == null) return;
            object[,] v = Xl.ReadTyped(src);
            int nr = v.GetLength(0), nc = v.GetLength(1);
            Xl.Backup(src.Worksheet.Parent);
            int n = 0;
            using (new Speed())
            {
                dynamic ws = src.Worksheet;
                for (int c = 1; c <= nc; c++)
                {
                    int r = 1;
                    while (r <= nr)
                    {
                        if (Xl.IsBlank(v[r, c]))
                        {
                            r++;
                            continue;
                        }
                        string k = Xl.Key(v[r, c]);
                        int r2 = r;
                        while (r2 < nr && !Xl.IsBlank(v[r2 + 1, c]) && Xl.Key(v[r2 + 1, c]) == k) r2++;
                        if (r2 > r)
                        {
                            dynamic m = ws.Range[src.Cells[r, c], src.Cells[r2, c]];
                            m.Merge();
                            m.VerticalAlignment = C.xlCenter;
                            n++;
                        }
                        r = r2 + 1;
                    }
                }
            }
            Ui.Info(N(n) + " blocks merged.\n\nMerged cells look good in a printed report but they break sorting and filters. Unmerge and Fill (Quick Fixes) reverses this.");
        }

        // =================================================================
        //  JOIN COLUMNS
        // =================================================================
        public static void JoinColumns()
        {
            dynamic src = Xl.PickRange("Join Columns", "Select the columns to join into one (2 or more columns):");
            if (src == null) return;
            object[,] v = Xl.ReadTyped(src);
            int nr = v.GetLength(0), nc = v.GetLength(1);
            if (nc < 2)
            {
                Ui.Info("Select at least 2 columns.");
                return;
            }
            string sep = Ui.AskText("Put what between the parts?\nExamples: a space, a comma, - or /   (leave empty for nothing)", Design.AppName + " - Join Columns", " ");
            if (sep == null) return;
            dynamic ws = src.Worksheet;
            int top = (int)src.Row, outCol = (int)src.Column + nc;
            dynamic target = ws.Range[ws.Cells[top, outCol], ws.Cells[top + nr - 1, outCol]];
            double used = (double)Xl.App.WorksheetFunction.CountA(target);
            if (used > 0 && !Ui.Ok("The column to the right already has data in " + used + " cells. Overwrite it?")) return;
            object[,] o = Xl.New2D(nr, 1);
            for (int r = 1; r <= nr; r++)
            {
                StringBuilder sb = new StringBuilder();
                for (int c = 1; c <= nc; c++)
                {
                    if (Xl.IsBlank(v[r, c]) || Xl.IsErr(v[r, c])) continue;
                    if (sb.Length > 0) sb.Append(sep);
                    sb.Append(Xl.Str(v[r, c]).Trim());
                }
                o[r, 1] = sb.ToString();
            }
            Xl.Backup(ws.Parent);
            using (new Speed())
            {
                target.NumberFormat = "@";
                target.Value2 = o;
            }
            Ui.Info(N(nr) + " rows joined into the column to the right of your selection.");
        }

        // =================================================================
        //  TEXT TOOLS
        // =================================================================
        static string Sentence(string s)
        {
            char[] a = s.ToLowerInvariant().ToCharArray();
            bool start = true;
            for (int i = 0; i < a.Length; i++)
            {
                if (start && char.IsLetter(a[i]))
                {
                    a[i] = char.ToUpperInvariant(a[i]);
                    start = false;
                }
                else if (a[i] == '.' || a[i] == '!' || a[i] == '?') start = true;
            }
            return new string(a);
        }

        static string Keep(string s, bool digits)
        {
            StringBuilder sb = new StringBuilder();
            foreach (char ch in s)
            {
                if (digits ? (ch >= '0' && ch <= '9') : (char.IsLetter(ch) || ch == ' ')) sb.Append(ch);
            }
            return sb.ToString();
        }

        public static void TextTools()
        {
            dynamic src = Xl.PickRange("Text Tools", "Select the cells to change:");
            if (src == null) return;
            long cells = Xl.CellCount(src);
            if (cells > 500000)
            {
                Ui.Info("Select up to 5 lakh cells at a time.");
                return;
            }
            int op = Ui.Pick("Text Tools", "What should happen to the text in the selected cells?", new[]
            {
                "UPPER CASE", "lower case", "Proper Case (Each Word)", "Sentence case",
                "Add text in front ...", "Add text at the end ...",
                "Remove the first N characters ...", "Remove the last N characters ...",
                "Keep only the digits", "Keep only letters and spaces",
                "Remove every space", "Remove line breaks inside cells"
            }, null);
            if (op < 0) return;
            string arg = "";
            int count = 0;
            if (op == 4 || op == 5)
            {
                arg = Ui.AskText("Type the text to add:", Design.AppName + " - Text Tools", "");
                if (string.IsNullOrEmpty(arg)) return;
            }
            else if (op == 6 || op == 7)
            {
                string a = Ui.AskText("How many characters?", Design.AppName + " - Text Tools", "1");
                if (a == null || !int.TryParse(a.Trim(), NumberStyles.Integer, Inv, out count) || count < 1) return;
            }
            Xl.Backup(src.Worksheet.Parent);
            int n = 0;
            using (new Speed())
            {
                object[,] v = Xl.ReadTyped(src);
                object[,] f = Xl.ReadFormulas(src);
                int nr = v.GetLength(0), nc = v.GetLength(1);
                bool numbersToo = op >= 4 && op <= 7;               // INV- in front of 1024 should work as well
                for (int c = 1; c <= nc; c++)
                {
                    for (int r = 1; r <= nr; r++)
                    {
                        object x = v[r, c];
                        if (x == null || IsFormula(f[r, c]) || Xl.IsErr(x) || x is DateTime || x is bool) continue;
                        if (!(x is string) && !numbersToo) continue;
                        string s = Xl.Str(x), t;
                        switch (op)
                        {
                            case 0: t = s.ToUpperInvariant(); break;
                            case 1: t = s.ToLowerInvariant(); break;
                            case 2: t = Inv.TextInfo.ToTitleCase(s.ToLowerInvariant()); break;
                            case 3: t = Sentence(s); break;
                            case 4: t = arg + s; break;
                            case 5: t = s + arg; break;
                            case 6: t = s.Length > count ? s.Substring(count) : ""; break;
                            case 7: t = s.Length > count ? s.Substring(0, s.Length - count) : ""; break;
                            case 8: t = Keep(s, true); break;
                            case 9: t = Logic.CleanText(Keep(s, false)); break;
                            case 10: t = s.Replace(" ", "").Replace("\u00A0", ""); break;
                            default: t = Logic.CleanText(s.Replace("\r\n", " ").Replace("\n", " ").Replace("\r", " ")); break;
                        }
                        if (t == s && x is string) continue;
                        CleanTools.PutText(src.Cells[r, c], t);
                        n++;
                    }
                }
            }
            Ui.Info(N(n) + " cells changed. Formulas, dates and blank cells were left alone.");
        }

        // =================================================================
        //  MATH ON CELLS
        // =================================================================
        public static void MathOnCells()
        {
            dynamic src = Xl.PickRange("Math on Cells", "Select the numbers to change:");
            if (src == null) return;
            long cells = Xl.CellCount(src);
            if (cells > 500000)
            {
                Ui.Info("Select up to 5 lakh cells at a time.");
                return;
            }
            string ans = Ui.AskText("Type the change. Examples:\n   *1.18      add 18% GST\n   /100000   show in lakhs\n   +500      add 500\n   -10%      reduce by 10%", Design.AppName + " - Math on Cells", "*1.18");
            if (ans == null) return;
            string a = ans.Replace(" ", "");
            if (a.Length < 2 || "+-*/".IndexOf(a[0]) < 0)
            {
                Ui.Info("Start with +, -, * or / and then a number, for example *1.18");
                return;
            }
            char op = a[0];
            bool pct = a.EndsWith("%", StringComparison.Ordinal);
            double num;
            if (!double.TryParse(a.Substring(1).TrimEnd('%'), NumberStyles.Float, Inv, out num))
            {
                Ui.Info("The number could not be read. Example: *1.18");
                return;
            }
            // every change becomes "multiply by mul, then add add"
            double mul = 1, add = 0;
            if (op == '*') mul = pct ? num / 100 : num;
            else if (op == '/')
            {
                double dv = pct ? num / 100 : num;
                if (dv == 0)
                {
                    Ui.Info("Dividing by zero is not possible.");
                    return;
                }
                mul = 1 / dv;
            }
            else if (pct) mul = op == '+' ? 1 + num / 100 : 1 - num / 100;
            else add = op == '+' ? num : -num;
            string tail = add != 0 ? (add > 0 ? "+" : "-") + Math.Abs(add).ToString("R", Inv) : (op == '/' && !pct ? "/" + num.ToString("R", Inv) : "*" + mul.ToString("R", Inv));
            Xl.Backup(src.Worksheet.Parent);
            int nVal = 0, nFor = 0;
            using (new Speed())
            {
                object[,] v = Xl.ReadTyped(src);
                object[,] f = Xl.ReadFormulas(src);
                int nr = v.GetLength(0), nc = v.GetLength(1);
                for (int c = 1; c <= nc; c++)
                {
                    for (int r = 1; r <= nr; r++)
                    {
                        if (IsFormula(f[r, c]))
                        {
                            if (!Xl.IsNum(v[r, c])) continue;
                            try
                            {
                                src.Cells[r, c].Formula = "=(" + ((string)f[r, c]).Substring(1) + ")" + tail;
                                nFor++;
                            }
                            catch { }
                        }
                        else if (Xl.IsNum(v[r, c]))
                        {
                            src.Cells[r, c].Value2 = Xl.Num(v[r, c]) * mul + add;
                            nVal++;
                        }
                    }
                }
            }
            Ui.Info(N(nVal) + " typed numbers changed and " + N(nFor) + " formulas extended with " + tail + ".\nText, dates and blank cells were left alone.");
        }

        // =================================================================
        //  WRAP WITH IFERROR
        // =================================================================
        public static void WrapIfError()
        {
            dynamic src = Xl.PickRange("Hide Formula Errors", "Select the formulas that should stop showing #N/A, #DIV/0! and other errors:");
            if (src == null) return;
            string rep = Ui.AskText("Show what when the formula gives an error?\nLeave empty for a blank cell, or type 0 or a text such as Not found.", Design.AppName + " - Hide Formula Errors", "");
            if (rep == null) return;
            double d;
            string lit = double.TryParse(rep.Trim(), NumberStyles.Float, Inv, out d) ? d.ToString("R", Inv) : "\"" + rep.Replace("\"", "\"\"") + "\"";
            Xl.Backup(src.Worksheet.Parent);
            int n = 0, skipped = 0;
            using (new Speed())
            {
                object[,] f = Xl.ReadFormulas(src);
                int nr = f.GetLength(0), nc = f.GetLength(1);
                for (int c = 1; c <= nc; c++)
                {
                    for (int r = 1; r <= nr; r++)
                    {
                        if (!IsFormula(f[r, c])) continue;
                        string s = (string)f[r, c];
                        if (s.StartsWith("=IFERROR(", StringComparison.OrdinalIgnoreCase)) continue;
                        try
                        {
                            src.Cells[r, c].Formula = "=IFERROR(" + s.Substring(1) + "," + lit + ")";
                            n++;
                        }
                        catch { skipped++; }
                    }
                }
            }
            Ui.Info(N(n) + " formulas now show " + (rep.Trim().Length == 0 ? "a blank" : rep.Trim()) + " instead of an error." + (skipped > 0 ? "\n" + N(skipped) + " could not be changed (array formulas)." : ""));
        }

        // =================================================================
        //  LOCK / UNLOCK REFERENCES ($)
        // =================================================================
        public static void ToggleDollar()
        {
            dynamic src = Xl.PickRange("Lock References ($)", "Select the formulas to change:");
            if (src == null) return;
            int pick = Ui.Pick("Lock References ($)", "Change the cell references in the selected formulas to:", new[]
            {
                "Fully locked        $A$1", "Row locked            A$1", "Column locked      $A1", "Not locked            A1"
            }, null);
            if (pick < 0) return;
            Xl.Backup(src.Worksheet.Parent);
            int n = 0, skipped = 0;
            using (new Speed())
            {
                dynamic app = Xl.App;
                object[,] f = Xl.ReadFormulas(src);
                int nr = f.GetLength(0), nc = f.GetLength(1);
                for (int c = 1; c <= nc; c++)
                {
                    for (int r = 1; r <= nr; r++)
                    {
                        if (!IsFormula(f[r, c])) continue;
                        try
                        {
                            dynamic cell = src.Cells[r, c];
                            object res = app.ConvertFormula((string)f[r, c], 1, 1, pick + 1, cell);     // 1 = A1 style; 1..4 = $A$1, A$1, $A1, A1
                            string s = res as string;
                            if (s == null || s == (string)f[r, c]) continue;
                            cell.Formula = s;
                            n++;
                        }
                        catch { skipped++; }
                    }
                }
            }
            Ui.Info(N(n) + " formulas changed." + (skipped > 0 ? "\n" + N(skipped) + " were too long for Excel to convert and were left as they are." : ""));
        }
    }
}
