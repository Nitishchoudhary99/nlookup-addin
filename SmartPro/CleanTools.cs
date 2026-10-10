using System;
using System.Collections.Generic;
using System.Globalization;

namespace SmartPro
{
    internal static class CleanTools
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        // Writes text without letting Excel turn 00123 / 1-2 / =x into something else
        public static void PutText(dynamic cell, string s)
        {
            cell.Value2 = Logic.RiskyText(s) ? "'" + s : s;
        }

        public static void PutValue(dynamic cell, object x)
        {
            string s = x as string;
            if (s != null) PutText(cell, s);
            else if (x is DateTime)
            {
                cell.NumberFormat = "dd-mmm-yyyy";
                cell.Value2 = ((DateTime)x).ToOADate();
            }
            else if (x is decimal) cell.Value2 = (double)(decimal)x;
            else cell.Value2 = x;
        }

        // force = false: Smart Clean. Row 1 is the header and a column is converted only when
        //                at least 80% of it already looks like numbers / dates.
        // force = true : convert every selected cell that can be converted.
        public static void Clean(bool doTrim, bool doNum, bool doDate, bool force, string title)
        {
            dynamic src = Xl.PickRange(title, "Select the cells to clean:");
            if (src == null) return;
            long cells = Xl.CellCount(src);
            if (cells > 3000000)
            {
                Ui.Info("That selection is very large. Select up to 30 lakh cells at a time.");
                return;
            }
            string bk = Xl.Backup(src.Worksheet.Parent);
            int nNum = 0, nDate = 0, nTrim = 0;
            using (new Speed())
            {
                Xl.Status("cleaning...");
                object[,] v = Xl.ReadTyped(src);
                object[,] f = Xl.ReadFormulas(src);
                int nr = v.GetLength(0), nc = v.GetLength(1);
                int startRow = (!force && nr > 1) ? 2 : 1;
                for (int c = 1; c <= nc; c++)
                {
                    bool colNum = force && doNum, colDate = force && doDate;
                    if (!force)
                    {
                        int nN = 0, nD = 0, nNB = 0;
                        for (int r = startRow; r <= nr; r++)
                        {
                            object x = v[r, c];
                            double tn;
                            DateTime td;
                            if (Xl.IsBlank(x) || Xl.IsErr(x)) continue;
                            nNB++;
                            if (x is DateTime) nD++;
                            else if (Logic.TryNum(x, out tn)) nN++;
                            else if (Logic.TryDate(x, out td)) nD++;
                        }
                        if (nNB > 0)
                        {
                            colNum = doNum && nN >= 0.8 * nNB;
                            colDate = doDate && nD >= 0.8 * nNB;
                        }
                    }
                    for (int r = 1; r <= nr; r++)
                    {
                        string s = v[r, c] as string;
                        if (s == null) continue;
                        string fs = f[r, c] as string;
                        if (fs != null && fs.StartsWith("=", StringComparison.Ordinal)) continue;
                        string s2 = doTrim ? Logic.CleanText(s) : s;
                        bool done = false;
                        if (r >= startRow)
                        {
                            double d;
                            DateTime dt;
                            if (colNum && Logic.TryNumText(s2, out d))
                            {
                                dynamic cel = src.Cells[r, c];
                                if (Convert.ToString((object)cel.NumberFormat) == "@") cel.NumberFormat = "General";
                                cel.Value2 = d;
                                nNum++;
                                done = true;
                            }
                            else if (colDate && Logic.TryDateText(s2, out dt))
                            {
                                dynamic cel = src.Cells[r, c];
                                cel.NumberFormat = "dd-mmm-yyyy";
                                cel.Value2 = dt.ToOADate();
                                nDate++;
                                done = true;
                            }
                        }
                        if (!done && s2 != s)
                        {
                            PutText(src.Cells[r, c], s2);
                            nTrim++;
                        }
                    }
                }
            }
            string msg = title + " finished.\n\n";
            if (doNum) msg += nNum.ToString("#,##0", Inv) + " text values turned into real numbers\n";
            if (doDate) msg += nDate.ToString("#,##0", Inv) + " text values turned into real dates\n";
            if (doTrim) msg += nTrim.ToString("#,##0", Inv) + " cells trimmed\n";
            if (bk != null) msg += "\nA backup was saved first (Help > Open Backups).";
            Ui.Info(msg);
        }

        public static void RemoveBlankRows()
        {
            dynamic src = Xl.PickRange("Remove Blank Rows", "Select the table. Rows that are completely empty inside it will be deleted:");
            if (src == null) return;
            Xl.Backup(src.Worksheet.Parent);
            int n = 0;
            using (new Speed())
            {
                object[,] v = Xl.ReadTyped(src);
                int nc = v.GetLength(1);
                dynamic ws = src.Worksheet;
                int r = v.GetLength(0);
                // bottom-up, one block of consecutive blank rows at a time
                while (r >= 1)
                {
                    if (!Xl.RowBlank(v, r))
                    {
                        r--;
                        continue;
                    }
                    int r2 = r;
                    while (r > 1 && Xl.RowBlank(v, r - 1)) r--;
                    dynamic block = ws.Range[src.Cells[r, 1], src.Cells[r2, nc]];
                    block.Delete(C.xlShiftUp);
                    n += r2 - r + 1;
                    r--;
                }
            }
            Ui.Info(n.ToString("#,##0", Inv) + " blank rows removed.");
        }

        public static void FillBlanksDown()
        {
            dynamic src = Xl.PickRange("Fill Blanks Down", "Select the column(s) where blank cells should take the value above:");
            if (src == null) return;
            Xl.Backup(src.Worksheet.Parent);
            int n = 0;
            using (new Speed())
            {
                object[,] v = Xl.ReadTyped(src);
                int nr = v.GetLength(0), nc = v.GetLength(1);
                for (int c = 1; c <= nc; c++)
                {
                    for (int r = 2; r <= nr; r++)
                    {
                        if (v[r, c] != null) continue;
                        object above = v[r - 1, c];
                        if (above == null || Xl.IsErr(above)) continue;
                        v[r, c] = above;
                        PutValue(src.Cells[r, c], above);
                        n++;
                    }
                }
            }
            Ui.Info(n.ToString("#,##0", Inv) + " blank cells filled.");
        }

        public static void UnmergeAndFill()
        {
            dynamic src = Xl.PickRange("Unmerge and Fill", "Select the area with merged cells:");
            if (src == null) return;
            long cells = Xl.CellCount(src);
            if (cells > 300000)
            {
                Ui.Info("Select a smaller area (up to 3 lakh cells) for this tool.");
                return;
            }
            Xl.Backup(src.Worksheet.Parent);
            int n = 0;
            using (new Speed())
            {
                int nr = (int)src.Rows.Count, nc = (int)src.Columns.Count;
                dynamic ws = src.Worksheet;
                HashSet<string> seen = new HashSet<string>();
                List<string> areas = new List<string>();
                for (int r = 1; r <= nr; r++)
                {
                    for (int c = 1; c <= nc; c++)
                    {
                        dynamic cel = src.Cells[r, c];
                        object merged = cel.MergeCells;
                        if (merged is bool && (bool)merged)
                        {
                            string a = (string)cel.MergeArea.Address;
                            if (seen.Add(a)) areas.Add(a);
                        }
                    }
                }
                foreach (string a in areas)
                {
                    dynamic ma = ws.Range[a];
                    ma.UnMerge();
                    long cnt = Xl.CellCount(ma);
                    if (cnt > 1) ma.Cells[1, 1].Copy(ma);
                    n++;
                }
                Xl.App.CutCopyMode = false;
            }
            Ui.Info(n + " merged blocks were unmerged and filled.");
        }

        public static void FormulasToValues()
        {
            dynamic src = Xl.PickRange("Formulas to Values", "Select the cells whose formulas should become fixed values:");
            if (src == null) return;
            object filtered = src.Worksheet.FilterMode;
            if (filtered is bool && (bool)filtered)
            {
                Ui.Info("This sheet is filtered. Clear the filter first, then run the tool again.");
                return;
            }
            Xl.Backup(src.Worksheet.Parent);
            using (new Speed())
            {
                src.Copy();
                src.PasteSpecial(C.xlPasteValues);
                Xl.App.CutCopyMode = false;
            }
            Ui.Info("The formulas in the selection are now fixed values.\nA backup was saved first (Help > Open Backups).");
        }
    }
}
