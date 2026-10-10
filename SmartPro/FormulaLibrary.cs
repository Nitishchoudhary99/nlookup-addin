using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace SmartPro
{
    // FORMULA LIBRARY: the formulas office users look up again and again, as recipes. Pick one, click the cells it
    // needs, and the finished formula is written into the active cell (and filled down if you like). Works offline,
    // in every Excel version. For anything not in the list, AI Formula writes it from your own words.
    //   Placeholders in a recipe:  {c:prompt} one cell (relative, A2)   {r:prompt} a column or range (absolute, $A$2:$A$500)
    //                              {n:prompt:default} a number or text typed in
    internal static class FormulaLibrary
    {
        sealed class Recipe
        {
            public readonly string Group, Title, Keys, Template;
            public Recipe(string group, string title, string keys, string template) { Group = group; Title = title; Keys = keys; Template = template; }
        }

        static readonly Recipe[] All =
        {
            new Recipe("Lookup", "Look up a value (works in every Excel)", "vlookup index match find bring", "=IFERROR(INDEX({r:the column to RETURN},MATCH({c:the KEY in this row},{r:the column to SEARCH, same size},0)),\"Not found\")"),
            new Recipe("Lookup", "Look up without #N/A from spaces or numbers as text (DocBrisk)", "smartlookup vlookup na spaces", "=SMARTLOOKUP({c:the KEY in this row},{r:the column to SEARCH},{r:the column to RETURN, same size},\"Not found\")"),
            new Recipe("Lookup", "Look up with TWO conditions", "two criteria multiple conditions lookup", "=IFERROR(LOOKUP(2,1/(({r:first column to search}={c:first value in this row})*({r:second column to search, same size}={c:second value in this row})),{r:the column to RETURN, same size}),\"Not found\")"),
            new Recipe("Lookup", "Closest name match (typos, Pvt/Ltd)", "fuzzy similar name lookup", "=FUZZYLOOKUP({c:the NAME in this row},{r:the list of names},{r:the column to RETURN, same size})"),
            new Recipe("Sum and count", "Total where a column equals a value (SUMIFS)", "sumif sumifs total condition", "=SUMIFS({r:the column to ADD},{r:the column to CHECK, same size},{c:the value to match})"),
            new Recipe("Sum and count", "Total with two conditions", "sumifs two conditions total", "=SUMIFS({r:the column to ADD},{r:first column to check},{c:first value},{r:second column to check},{c:second value})"),
            new Recipe("Sum and count", "Total between two dates", "sum between dates period month", "=SUMIFS({r:the column to ADD},{r:the DATE column, same size},\">=\"&{c:the FROM date},{r:the DATE column, same size},\"<=\"&{c:the TO date})"),
            new Recipe("Sum and count", "Count where a column equals a value (COUNTIFS)", "countif countifs how many", "=COUNTIFS({r:the column to CHECK},{c:the value to count})"),
            new Recipe("Sum and count", "Average where a column equals a value", "averageif average condition", "=IFERROR(AVERAGEIFS({r:the column to AVERAGE},{r:the column to CHECK, same size},{c:the value to match}),0)"),
            new Recipe("Sum and count", "Running total", "cumulative running total balance", "=SUM({a:the FIRST cell of the amount column}:{c:the amount in this row})"),
            new Recipe("Sum and count", "Share of the total (%)", "percent share of total contribution", "={c:the amount in this row}/SUM({r:the whole amount column})"),
            new Recipe("Sum and count", "Count of different values", "unique distinct count", "=SUMPRODUCT(({r:the column}<>\"\")/COUNTIF({r:the column},{r:the column}&\"\"))"),
            new Recipe("Sum and count", "Rank (1 = largest)", "rank position top", "=RANK.EQ({c:the value in this row},{r:the whole column},0)"),
            new Recipe("Dates", "Age in years", "age years datedif birthday", "=DATEDIF({c:the date of birth},TODAY(),\"y\")"),
            new Recipe("Dates", "Days between two dates", "days difference between dates", "={c:the LATER date}-{c:the EARLIER date}"),
            new Recipe("Dates", "Working days between two dates (Mon-Fri)", "networkdays working days business days", "=NETWORKDAYS({c:the START date},{c:the END date})"),
            new Recipe("Dates", "Days overdue (0 when not due)", "overdue days due late", "=MAX(0,TODAY()-{c:the DUE date})"),
            new Recipe("Dates", "Due date after N days", "due date credit days add days", "={c:the INVOICE date}+{n:Credit days:30}"),
            new Recipe("Dates", "Ageing bucket (0-30, 31-60 ...) (DocBrisk)", "ageing aging bucket receivable", "=AGEBUCKET({c:the DATE})"),
            new Recipe("Dates", "Month name (Apr)", "month name text", "=TEXT({c:the DATE},\"mmm\")"),
            new Recipe("Dates", "Month for sorting (2026-04)", "month key yyyy-mm", "=TEXT({c:the DATE},\"yyyy-mm\")"),
            new Recipe("Dates", "Fiscal year April-March (FY 2026-27) (DocBrisk)", "fiscal financial year fy", "=FYEAR({c:the DATE})"),
            new Recipe("Dates", "Last day of the month", "eomonth month end", "=EOMONTH({c:the DATE},0)"),
            new Recipe("Text", "First name", "first name split name", "=TRIM(LEFT({c:the FULL NAME},FIND(\" \",{c:the FULL NAME}&\" \")-1))"),
            new Recipe("Text", "Last name", "last name surname split", "=TRIM(RIGHT(SUBSTITUTE(TRIM({c:the FULL NAME}),\" \",REPT(\" \",100)),100))"),
            new Recipe("Text", "Join two cells with a space", "join concatenate combine", "=TRIM({c:the FIRST cell}&\" \"&{c:the SECOND cell})"),
            new Recipe("Text", "Clean spaces (also the invisible web space)", "trim clean spaces nbsp", "=TRIM(SUBSTITUTE({c:the TEXT},CHAR(160),\" \"))"),
            new Recipe("Text", "Number stored as text to a real number", "text to number convert value", "=IFERROR(VALUE(TRIM(SUBSTITUTE({c:the TEXT number},\",\",\"\"))),{c:the TEXT number})"),
            new Recipe("Text", "The first number inside a text (DocBrisk)", "extract number digits", "=EXTRACTNUM({c:the TEXT})"),
            new Recipe("Text", "Text between two markers (DocBrisk)", "text between extract middle", "=TEXTBETWEEN({c:the TEXT},\"{n:Text before:(}\",\"{n:Text after:)}\")"),
            new Recipe("Checks", "Mark repeats (2nd time onwards)", "duplicate repeat flag", "=IF(COUNTIF({a:the FIRST cell of the column}:{c:the cell in this row},{c:the cell in this row})>1,\"Repeat\",\"\")"),
            new Recipe("Checks", "Is this value in the other list?", "exists in list match found missing", "=IF(COUNTIF({r:the OTHER list},{c:the value in this row})>0,\"Yes\",\"No\")"),
            new Recipe("Checks", "Grade from marks or score", "grade if nested score band", "=IF({c:the SCORE}>={n:Score for A:90},\"A\",IF({c:the SCORE}>={n:Score for B:75},\"B\",IF({c:the SCORE}>={n:Score for C:50},\"C\",\"D\")))"),
            new Recipe("Checks", "GSTIN is valid? (DocBrisk)", "gstin gst valid check", "=GSTINVALID({c:the GSTIN})"),
            new Recipe("Money and GST", "Amount with GST added", "gst add inclusive tax", "=ROUND({c:the AMOUNT}*(1+{n:GST rate %:18}/100),2)"),
            new Recipe("Money and GST", "GST inside a GST-inclusive amount", "gst part tax inclusive", "=ROUND({c:the AMOUNT incl. GST}*{n:GST rate %:18}/(100+{n:GST rate %:18}),2)"),
            new Recipe("Money and GST", "Amount in words, Indian style (DocBrisk)", "rupees words amount", "=RUPEEWORDS({c:the AMOUNT})"),
            new Recipe("Money and GST", "Percentage change", "growth change percent increase", "=IFERROR(({c:the NEW value}-{c:the OLD value})/ABS({c:the OLD value}),0)"),
            new Recipe("Money and GST", "Monthly EMI of a loan", "emi loan pmt instalment", "=ROUND(PMT({n:Interest % a year:9}/1200,{n:Months:60},-{c:the LOAN amount}),0)"),
            new Recipe("Money and GST", "CAGR (growth a year)", "cagr compound growth", "=({c:the END value}/{c:the START value})^(1/{n:Years:5})-1"),
            new Recipe("Money and GST", "Round to the nearest 10 / 100 / 1000", "round mround nearest", "=MROUND({c:the AMOUNT},{n:Round to:100})")
        };

        public static void Run()
        {
            dynamic cell = Xl.App.ActiveCell;
            if (cell == null) { Ui.Info("Click the cell where the formula should go."); return; }
            string[] items = new string[All.Length + 1], keys = new string[All.Length + 1];
            items[0] = "AI Formula (Pro): write any formula from your own words ...";
            keys[0] = "ai any custom own words other write";
            for (int i = 0; i < All.Length; i++)
            {
                items[i + 1] = All[i].Group + "   |   " + All[i].Title;
                keys[i + 1] = All[i].Group + " " + All[i].Title + " " + All[i].Keys;
            }
            string target = ((string)cell.Address).Replace("$", "");
            int pick = Ui.Pick(Design.AppName + " - Formula Library   (" + All.Length + " formulas, for " + target + ")", "Type what you need: lookup, sumifs, age, gst, overdue, first name ...", items, keys);
            if (pick < 0) return;
            if (pick == 0)
            {
                Runner.Run("aiformula");
                return;
            }
            Recipe rc = All[pick - 1];
            string formula = Fill(rc, cell);
            if (formula == null) return;
            try { cell.Formula = formula; }
            catch (Exception ex)
            {
                Ui.Warn("Excel did not accept the formula:\n" + formula + "\n\n" + ex.Message);
                return;
            }
            bool fillDown = false;
            try
            {
                dynamic region = cell.CurrentRegion;
                int last = (int)region.Row + (int)region.Rows.Count - 1;
                if (last > (int)cell.Row && Ui.Ok("Formula written in " + target + ":\n" + formula + "\n\nFill it down to row " + last + "?"))
                {
                    cell.Worksheet.Range[cell, cell.Worksheet.Cells[last, (int)cell.Column]].FillDown();
                    fillDown = true;
                }
            }
            catch { }
            Xl.Status(rc.Title + ": " + formula + (fillDown ? "  (filled down)" : ""));
        }

        // asks for each placeholder in order (the same placeholder twice is asked once) and returns the formula, or null when cancelled
        static string Fill(Recipe rc, dynamic cell)
        {
            string t = rc.Template;
            StringBuilder sb = new StringBuilder();
            Dictionary<string, string> asked = new Dictionary<string, string>();
            int i = 0;
            while (i < t.Length)
            {
                int open = t.IndexOf('{', i);
                int close = open < 0 ? -1 : t.IndexOf('}', open);
                if (open < 0 || close < 0) { sb.Append(t.Substring(i)); break; }
                sb.Append(t.Substring(i, open - i));
                string ph = t.Substring(open + 1, close - open - 1);
                string val;
                if (!asked.TryGetValue(ph, out val))
                {
                    string kind = ph.Substring(0, 1), prompt = ph.Substring(2), def = "";
                    int colon = prompt.IndexOf(':');
                    if (kind == "n" && colon > 0) { def = prompt.Substring(colon + 1); prompt = prompt.Substring(0, colon); }
                    val = Ask(kind, prompt, def, rc.Title, cell);
                    if (val == null) return null;
                    asked[ph] = val;
                }
                sb.Append(val);
                i = close + 1;
            }
            return sb.ToString();
        }

        static string Ask(string kind, string prompt, string def, string title, dynamic cell)
        {
            if (kind == "n")
            {
                string s = Ui.AskText(prompt + ":", title, def);
                if (s == null) return null;
                return s.Trim();
            }
            object res;
            try { res = Xl.App.InputBox("Click " + prompt + (kind == "r" ? " (the whole column of data)" : "") + ":", title, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8); }
            catch { return null; }
            if (res == null || res is bool) return null;
            dynamic r = res;
            bool otherSheet = (string)r.Worksheet.Name != (string)cell.Worksheet.Name;
            bool otherBook = (string)r.Worksheet.Parent.Name != (string)cell.Worksheet.Parent.Name;
            string addr;
            if (kind == "c") addr = ((string)r.Cells[1, 1].Address).Replace("$", "");
            else if (kind == "a") addr = (string)r.Cells[1, 1].Address;
            else addr = (string)r.Address;
            if (otherBook) addr = "'[" + ((string)r.Worksheet.Parent.Name).Replace("'", "''") + "]" + ((string)r.Worksheet.Name).Replace("'", "''") + "'!" + addr;
            else if (otherSheet) addr = "'" + ((string)r.Worksheet.Name).Replace("'", "''") + "'!" + addr;
            return addr;
        }

        public static int Count { get { return All.Length; } }
    }
}
