using System;
using System.Collections.Generic;
using System.Globalization;

namespace SmartPro
{
    // SMART VALIDATION: entry rules for the selected cells in one go, with a clear message when a wrong value is typed.
    // Dropdowns (typed list, or the different values of a column), no repeats, numbers or dates in a range,
    // and the Indian formats: mobile, PIN, PAN, GSTIN, IFSC, Aadhaar (format), email.
    // Everything is plain Excel data validation: it works for everyone who opens the file, without DocBrisk.
    internal static class ValidationTools
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        const int xlValidateWholeNumber = 1, xlValidateDecimal = 2, xlValidateList = 3, xlValidateDate = 4, xlValidateTextLength = 6, xlValidateCustom = 7;
        const int xlValidAlertStop = 1, xlBetween = 1, xlEqual = 3, xlGreaterEqual = 7;

        static readonly string[] Kinds =
        {
            "Dropdown: type the choices (Yes, No, Pending ...)",
            "Dropdown: the different values of a column (a master list)",
            "No repeats in this column (stops a second entry of the same value)",
            "Whole number between two values",
            "Decimal number between two values",
            "Date between two dates",
            "Not in the future (date up to today)",
            "Indian mobile number (10 digits, starts with 6-9)",
            "PIN code (6 digits)",
            "PAN (ABCDE1234F)",
            "GSTIN (15 characters: 2 digits, PAN, 1, Z, 1)",
            "IFSC (4 letters, 0, 6 letters or digits)",
            "Aadhaar format (12 digits)",
            "Email address",
            "Text length up to N characters",
            "Remove the rules from the selected cells"
        };

        public static void Run()
        {
            dynamic app = Xl.App;
            dynamic sel;
            try { sel = app.Selection; sel = sel.Areas[1]; } catch { Ui.Info("Select the cells that need the rule first."); return; }
            if (Xl.CellCount(sel) == 1)
            {
                // one cell: offer the column under it, from the cell to the end of the data (or 1,000 rows)
                dynamic ws = sel.Worksheet;
                int last = Math.Max((int)ws.UsedRange.Row + (int)ws.UsedRange.Rows.Count - 1, (int)sel.Row + 999);
                if (Ui.Ok("Apply the rule to " + ((string)sel.Address).Replace("$", "") + " and the cells below it, down to row " + last.ToString("#,##0", Inv) + "?\n\nCancel = only this cell"))
                    sel = ws.Range[sel, ws.Cells[last, (int)sel.Column]];
            }
            int kind = Ui.Pick(Design.AppName + " - Smart Validation", "Which rule for " + ((string)sel.Address).Replace("$", "") + "?  (type to search)", Kinds, null);
            if (kind < 0) return;
            string top = ((string)sel.Cells[1, 1].Address).Replace("$", "");               // A2
            string col = top.TrimEnd('0', '1', '2', '3', '4', '5', '6', '7', '8', '9');        // A
            string first = ((string)sel.Cells[1, 1].Address);                                    // $A$2
            string colRange = "$" + col + ":$" + col;
            string title = "", message = "";
            dynamic v = sel.Validation;
            switch (kind)
            {
                case 0:
                    {
                        string items = Ui.AskText("Type the choices, separated by commas:", "Dropdown", "Yes,No,Pending");
                        if (string.IsNullOrEmpty(items)) return;
                        string sep = ListSeparator();
                        string list = string.Join(sep, Split(items).ToArray());
                        if (list.Length > 255) { Ui.Info("The list is too long for a typed dropdown (255 characters). Put the choices in cells and use the second option."); return; }
                        v.Delete();
                        v.Add(xlValidateList, xlValidAlertStop, xlBetween, list);
                        title = "Pick from the list";
                        message = "Choose one of: " + items;
                        break;
                    }
                case 1:
                    {
                        dynamic src = Xl.PickRange("Dropdown from a column", "Select the cells that hold the choices, WITHOUT the header. Repeats and blanks are removed:");
                        if (src == null) return;
                        object[,] vals = Xl.ReadTyped(src);
                        List<string> uniq = new List<string>();
                        HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        for (int r = 1; r <= vals.GetLength(0); r++)
                            for (int c = 1; c <= vals.GetLength(1); c++)
                            {
                                if (Xl.IsBlank(vals[r, c]) || Xl.IsErr(vals[r, c])) continue;
                                string s = Xl.Str(vals[r, c]).Trim();
                                if (seen.Add(s)) uniq.Add(s);
                            }
                        if (uniq.Count == 0) { Ui.Info("No values were found in those cells."); return; }
                        uniq.Sort(StringComparer.OrdinalIgnoreCase);
                        dynamic wb = sel.Worksheet.Parent;
                        // the list lives on a hidden sheet, so it stays sorted and free of repeats
                        dynamic ls = ListSheet(wb);
                        int lc = 1;
                        while (!Xl.IsBlank(ls.Cells[1, lc].Value2)) lc++;
                        object[,] o = Xl.New2D(uniq.Count + 1, 1);
                        o[1, 1] = (string)src.Worksheet.Name + "!" + ((string)src.Address).Replace("$", "");
                        for (int i = 0; i < uniq.Count; i++) o[i + 2, 1] = uniq[i];
                        Xl.WriteTable(ls.Cells[1, lc], o, 1);
                        string refers = "='" + ((string)ls.Name).Replace("'", "''") + "'!" + (string)ls.Cells[2, lc].Resize[uniq.Count, 1].Address;
                        v.Delete();
                        v.Add(xlValidateList, xlValidAlertStop, xlBetween, refers);
                        title = "Pick from the list";
                        message = uniq.Count + " choices from " + (string)o[1, 1];
                        break;
                    }
                case 2:
                    v.Delete();
                    v.Add(xlValidateCustom, xlValidAlertStop, xlBetween, "=COUNTIF(" + colRange + "," + top + ")<=1");
                    title = "Already used";
                    message = "This value is already in the column. Each value may appear only once.";
                    break;
                case 3:
                case 4:
                    {
                        string lo = Ui.AskText("Smallest allowed number:", "Number between", "0");
                        if (lo == null) return;
                        string hi = Ui.AskText("Largest allowed number:", "Number between", "100000");
                        if (hi == null) return;
                        v.Delete();
                        v.Add(kind == 3 ? xlValidateWholeNumber : xlValidateDecimal, xlValidAlertStop, xlBetween, lo, hi);
                        title = "Number out of range";
                        message = "Enter a " + (kind == 3 ? "whole " : "") + "number from " + lo + " to " + hi + ".";
                        break;
                    }
                case 5:
                    {
                        string lo = Ui.AskText("First allowed date (dd-mm-yyyy):", "Date between", "01-04-" + FyStart());
                        if (lo == null) return;
                        string hi = Ui.AskText("Last allowed date (dd-mm-yyyy):", "Date between", "31-03-" + (FyStart() + 1));
                        if (hi == null) return;
                        DateTime d1, d2;
                        if (!Logic.TryDateText(lo, out d1) || !Logic.TryDateText(hi, out d2)) { Ui.Info("Those dates could not be read. Use dd-mm-yyyy."); return; }
                        v.Delete();
                        v.Add(xlValidateDate, xlValidAlertStop, xlBetween, "=DATE(" + d1.Year + "," + d1.Month + "," + d1.Day + ")", "=DATE(" + d2.Year + "," + d2.Month + "," + d2.Day + ")");
                        title = "Date out of range";
                        message = "Enter a date from " + d1.ToString("dd-MMM-yyyy", Inv) + " to " + d2.ToString("dd-MMM-yyyy", Inv) + ".";
                        break;
                    }
                case 6:
                    v.Delete();
                    v.Add(xlValidateCustom, xlValidAlertStop, xlBetween, "=AND(ISNUMBER(" + top + ")," + top + "<=TODAY())");
                    title = "Date in the future";
                    message = "Enter a date up to today.";
                    break;
                case 7: Custom(v, Pattern(top, 10, "MID(" + top + ",1,1)>=\"6\""), "Mobile number", "Enter 10 digits starting with 6, 7, 8 or 9, without +91 or spaces."); title = null; break;
                case 8: Custom(v, Pattern(top, 6, "LEFT(" + top + ",1)<>\"0\""), "PIN code", "Enter a 6-digit PIN code that does not start with 0."); title = null; break;
                case 9:
                    Custom(v, "=AND(LEN(" + top + ")=10," + Letters(top, 1, 5) + "," + Digits(top, 6, 4) + "," + Letters(top, 10, 1) + ")", "PAN", "Enter the PAN as 5 letters, 4 digits and 1 letter: ABCDE1234F.");
                    title = null;
                    break;
                case 10:
                    Custom(v, "=AND(LEN(" + top + ")=15," + Digits(top, 1, 2) + "," + Letters(top, 3, 5) + "," + Digits(top, 8, 4) + "," + Letters(top, 12, 1) + ",MID(" + top + ",14,1)=\"Z\")",
                           "GSTIN", "Enter the 15-character GSTIN: 2-digit state code, the PAN, 1 entity number, Z and a check character.");
                    title = null;
                    break;
                case 11:
                    Custom(v, "=AND(LEN(" + top + ")=11," + Letters(top, 1, 4) + ",MID(" + top + ",5,1)=\"0\")", "IFSC", "Enter the IFSC: 4 letters, the digit 0, then 6 letters or digits.");
                    title = null;
                    break;
                case 12: Custom(v, Pattern(top, 12, "LEFT(" + top + ",1)>=\"2\""), "Aadhaar number", "Enter 12 digits (the first digit is 2 to 9), without spaces."); title = null; break;
                case 13:
                    Custom(v, "=AND(ISERROR(FIND(\" \"," + top + ")),LEN(" + top + ")-LEN(SUBSTITUTE(" + top + ",\"@\",\"\"))=1,ISNUMBER(FIND(\".\"," + top + ",FIND(\"@\"," + top + ")+2)),RIGHT(" + top + ",1)<>\".\")",
                           "Email address", "Enter an email address like name@company.com, without spaces.");
                    title = null;
                    break;
                case 14:
                    {
                        string n = Ui.AskText("Most characters allowed:", "Text length", "50");
                        if (n == null) return;
                        v.Delete();
                        v.Add(xlValidateTextLength, xlValidAlertStop, xlBetween, "0", n);
                        title = "Too long";
                        message = "Use at most " + n + " characters.";
                        break;
                    }
                case 15:
                    v.Delete();
                    Xl.Status("Validation rules removed from " + ((string)sel.Address).Replace("$", "") + ".");
                    return;
            }
            if (title != null)
            {
                v.ErrorTitle = title;
                v.ErrorMessage = message;
                v.InputMessage = message;
                v.ShowInput = true;
                v.ShowError = true;
                v.IgnoreBlank = true;
            }
            int bad = CountBad(sel);
            Ui.Info("Rule set on " + ((string)sel.Address).Replace("$", "") + ": " + Kinds[kind] + "." +
                    (bad > 0 ? "\n\n" + bad + " cell(s) that are already filled do not follow the rule. They are circled in red (Data > Data Validation > Clear Validation Circles to remove the circles)." : "") +
                    "\n\nThe rule is plain Excel data validation: it works for everyone who opens the file.");
        }

        static void Custom(dynamic v, string formula, string title, string message)
        {
            v.Delete();
            v.Add(xlValidateCustom, xlValidAlertStop, xlBetween, formula);
            v.ErrorTitle = title;
            v.ErrorMessage = message;
            v.InputTitle = title;
            v.InputMessage = message;
            v.ShowInput = true;
            v.ShowError = true;
            v.IgnoreBlank = true;
        }

        // exactly n digits (text or number) and an extra test
        static string Pattern(string a, int n, string extra)
        {
            return "=AND(LEN(" + a + ")=" + n + ",ISNUMBER(--" + a + "),ISERROR(FIND(\".\"," + a + ")),ISERROR(FIND(\"-\"," + a + ")),ISERROR(FIND(\" \"," + a + "))," + extra + ")";
        }

        // characters start..start+len-1 are all A-Z (works in every Excel version, no LET/REGEX needed)
        static string Letters(string a, int start, int len)
        {
            return "SUMPRODUCT(--(ABS(CODE(MID(UPPER(" + a + ")," + start + "+ROW($1:$" + len + ")-1,1))-77.5)<13))=" + len;
        }

        static string Digits(string a, int start, int len)
        {
            return "SUMPRODUCT(--ISNUMBER(--MID(" + a + "," + start + "+ROW($1:$" + len + ")-1,1)))=" + len;
        }

        static int FyStart() { DateTime t = DateTime.Today; return t.Month >= 4 ? t.Year : t.Year - 1; }

        static List<string> Split(string items)
        {
            List<string> l = new List<string>();
            foreach (string p in items.Split(new[] { ',', ';' }))
            {
                string t = p.Trim();
                if (t.Length > 0 && !l.Contains(t)) l.Add(t);
            }
            return l;
        }

        // the list separator Excel expects in a typed list (comma, or semicolon in many European settings)
        static string ListSeparator()
        {
            try
            {
                object s = Xl.App.International[5];          // xlListSeparator
                string t = s as string;
                if (!string.IsNullOrEmpty(t)) return t;
            }
            catch { }
            return ",";
        }

        static dynamic ListSheet(dynamic wb)
        {
            const string name = "DocBrisk Lists";
            if (Xl.SheetExists(wb, name)) return wb.Sheets[name];
            dynamic active = Xl.App.ActiveSheet;
            dynamic ws = Xl.AddSheet(wb, name);
            ws.Visible = C.xlSheetHidden;
            try { active.Activate(); } catch { }
            return ws;
        }

        // circles the filled cells that break the new rule, and counts them
        static int CountBad(dynamic sel)
        {
            int bad = 0;
            try
            {
                dynamic ws = sel.Worksheet;
                dynamic filled = null;
                try { filled = sel.SpecialCells(2); } catch { }            // xlCellTypeConstants
                if (filled == null) return 0;
                if ((long)Xl.CellCount(filled) > 20000) return 0;
                foreach (dynamic c in filled.Cells)
                {
                    bool ok = true;
                    try { ok = (bool)c.Validation.Value; } catch { }
                    if (!ok) bad++;
                }
                if (bad > 0) ws.CircleInvalid();
            }
            catch { }
            return bad;
        }
    }
}
