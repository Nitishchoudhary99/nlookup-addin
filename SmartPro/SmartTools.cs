using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Windows.Forms;

namespace SmartPro
{
    // SMART TOOLS
    //   Explain Formula ....... free: a formula in plain words, step by step, with the value of every part
    //   Smart Analyst ......... Pro: a written analysis of a table (who drives the total, trend, unusual rows)
    //   Duplicate Payments .... Pro: invoices that may have been entered or paid twice
    //   Mail to Each Party .... Pro: one Outlook email for every party with only its own rows
    // Everything is worked out on this PC by fixed rules. No data is sent to DocBrisk or to any AI service.
    internal static class SmartTools
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        static string N(double n) { return n.ToString("#,##0", Inv); }
        static string F(double x) { return Logic.ShortINR(Math.Round(x, 2)); }
        static string P(double x) { return x.ToString(Math.Abs(x) < 0.1 ? "0.0%" : "0%", Inv); }
        static string Esc(string s) { return SecurityElement.Escape(s ?? "") ?? ""; }

        // =================================================================
        //  EXPLAIN FORMULA
        // =================================================================
        // NAME | what it does | name of part 1 | name of part 2 ...
        static readonly string[] FnText =
        {
            "SUM|adds up the numbers",
            "SUMIF|adds up the numbers of the rows that meet one condition|look in|condition|add up",
            "SUMIFS|adds up the numbers of the rows that meet every condition|add up|look in|condition|look in|condition|look in|condition|look in|condition|look in|condition",
            "SUMPRODUCT|multiplies the lists row by row and adds up the results",
            "COUNT|counts the cells that hold numbers",
            "COUNTA|counts the cells that are not empty",
            "COUNTBLANK|counts the empty cells",
            "COUNTIF|counts the cells that meet one condition|look in|condition",
            "COUNTIFS|counts the rows that meet every condition|look in|condition|look in|condition|look in|condition|look in|condition",
            "AVERAGE|gives the average of the numbers",
            "AVERAGEIF|gives the average of the rows that meet one condition|look in|condition|average of",
            "AVERAGEIFS|gives the average of the rows that meet every condition|average of|look in|condition|look in|condition|look in|condition",
            "MIN|gives the smallest number",
            "MAX|gives the largest number",
            "MINIFS|gives the smallest number of the rows that meet every condition|smallest of|look in|condition|look in|condition",
            "MAXIFS|gives the largest number of the rows that meet every condition|largest of|look in|condition|look in|condition",
            "IF|checks a condition and gives one result when it is true and another when it is false|condition|when true|when false",
            "IFS|checks the conditions in order and gives the result of the first one that is true|condition|result|condition|result|condition|result|condition|result",
            "IFERROR|gives the first part, or the second part when the first part is an error|try this|when it is an error",
            "IFNA|gives the first part, or the second part when the first part is #N/A|try this|when it is #N/A",
            "AND|is TRUE only when every condition is true",
            "OR|is TRUE when at least one condition is true",
            "NOT|turns TRUE into FALSE and FALSE into TRUE",
            "SWITCH|compares a value with a list and gives the result that belongs to the first match|value|match|result|match|result|match|result",
            "VLOOKUP|looks for a value in the first column of a table and returns the value from another column of that row|look for|in this table|return column number|match type",
            "HLOOKUP|looks for a value in the first row of a table and returns the value from another row of that column|look for|in this table|return row number|match type",
            "XLOOKUP|looks for a value in one list and returns the value at the same place of another list|look for|in this list|return from|when not found|match type|search direction",
            "LOOKUP|looks for a value in a sorted list and returns the value at the same place|look for|in this list|return from",
            "INDEX|returns the value at a given row and column of a range|range|row number|column number",
            "MATCH|gives the position of a value in a list|look for|in this list|match type",
            "XMATCH|gives the position of a value in a list|look for|in this list|match type|search direction",
            "OFFSET|gives a range that starts some rows and columns away from a cell|start cell|rows down|columns right|height|width",
            "INDIRECT|turns a text such as \"A1\" into a real reference|reference as text",
            "CHOOSE|picks one value from a list by its number|which one|value 1|value 2|value 3|value 4",
            "LEFT|takes the first characters of a text|text|how many",
            "RIGHT|takes the last characters of a text|text|how many",
            "MID|takes characters from the middle of a text|text|start at|how many",
            "LEN|counts the characters of a text",
            "TRIM|removes extra spaces from a text",
            "CLEAN|removes characters that cannot be printed",
            "UPPER|changes a text to capital letters",
            "LOWER|changes a text to small letters",
            "PROPER|writes every word with a capital first letter",
            "CONCATENATE|joins texts into one",
            "CONCAT|joins texts into one",
            "TEXTJOIN|joins texts into one with a separator between them|separator|skip empty cells|text 1|text 2|text 3",
            "TEXT|shows a number or a date as text in a given format|value|format",
            "VALUE|turns a text that looks like a number into a number",
            "SUBSTITUTE|replaces one text with another inside a text|text|replace this|with this|which occurrence",
            "REPLACE|replaces characters at a position of a text|text|start at|how many|with this",
            "FIND|gives the position of a text inside another text (capital letters matter)|look for|inside|start at",
            "SEARCH|gives the position of a text inside another text|look for|inside|start at",
            "EXACT|is TRUE when two texts are exactly the same, capital letters included",
            "REPT|repeats a text a number of times|text|times",
            "ROUND|rounds a number to a number of decimals|number|decimals",
            "ROUNDUP|rounds a number up|number|decimals",
            "ROUNDDOWN|rounds a number down|number|decimals",
            "MROUND|rounds a number to the nearest multiple|number|multiple",
            "CEILING|rounds a number up to the nearest multiple|number|multiple",
            "FLOOR|rounds a number down to the nearest multiple|number|multiple",
            "INT|drops the decimals of a number",
            "ABS|gives a number without its minus sign",
            "MOD|gives the remainder of a division|number|divide by",
            "POWER|raises a number to a power|number|power",
            "SQRT|gives the square root of a number",
            "TODAY|gives today's date",
            "NOW|gives the date and time of this moment",
            "DATE|makes a date from a year, a month and a day|year|month|day",
            "DATEVALUE|turns a text that looks like a date into a date",
            "YEAR|gives the year of a date",
            "MONTH|gives the month of a date",
            "DAY|gives the day of a date",
            "EOMONTH|gives the last day of a month, a number of months away|start date|months away",
            "EDATE|gives the same day a number of months away|start date|months away",
            "DATEDIF|gives the time between two dates in years, months or days|start date|end date|unit",
            "DAYS|counts the days between two dates|end date|start date",
            "NETWORKDAYS|counts the working days between two dates|start date|end date|holidays",
            "WORKDAY|gives the date a number of working days away|start date|working days|holidays",
            "WEEKDAY|gives the day of the week as a number|date|numbering",
            "WEEKNUM|gives the week number of a date|date|numbering",
            "ISBLANK|is TRUE when the cell is empty",
            "ISNUMBER|is TRUE when the value is a number",
            "ISTEXT|is TRUE when the value is a text",
            "ISERROR|is TRUE when the value is an error",
            "ISNA|is TRUE when the value is #N/A",
            "SUBTOTAL|adds up, counts or averages only the rows that are visible after filtering|what to do (9 = sum, 3 = count, 1 = average)|range",
            "AGGREGATE|adds up, counts or averages and can skip errors and hidden rows|what to do|what to skip|range",
            "UNIQUE|lists every different value once",
            "FILTER|lists the rows that meet a condition|range|condition|when nothing is found",
            "SORT|lists the range in sorted order|range|sort by column|order",
            "SORTBY|lists the range sorted by another list|range|sort by|order",
            "RANK|gives the rank of a number in a list|number|list|order",
            "LARGE|gives the k-th largest number|list|k",
            "SMALL|gives the k-th smallest number|list|k",
            "MEDIAN|gives the middle value of the numbers",
            "TRANSPOSE|turns rows into columns and columns into rows",
            "ROW|gives the row number of a cell",
            "COLUMN|gives the column number of a cell",
            "ROWS|counts the rows of a range",
            "COLUMNS|counts the columns of a range",
            "LET|gives names to values and uses them in the last part",
            "PMT|gives the payment of a loan for one period|rate per period|number of periods|loan amount",
            "RUPEEWORDS|writes an amount in words (a DocBrisk function)",
            "INRFORMAT|shows a number with Indian commas (a DocBrisk function)",
            "FUZZYLOOKUP|finds the most similar name in a list (a DocBrisk function)",
            "SIMILARITY|says how alike two texts are, from 0 to 1 (a DocBrisk function)",
            "NLOOKUP|looks up or adds up rows by one or more conditions (the DocBrisk NLOOKUP add-in)"
        };

        static Dictionary<string, string[]> fnMap;

        static Dictionary<string, string[]> Fn
        {
            get
            {
                if (fnMap == null)
                {
                    Dictionary<string, string[]> m = new Dictionary<string, string[]>(StringComparer.Ordinal);
                    foreach (string line in FnText)
                    {
                        string[] p = line.Split('|');
                        string[] rest = new string[p.Length - 1];
                        Array.Copy(p, 1, rest, 0, rest.Length);
                        m[p[0]] = rest;
                    }
                    fnMap = m;
                }
                return fnMap;
            }
        }

        // one function inside a formula
        sealed class Call
        {
            public string Name = "", Text = "";
            public List<string> Args = new List<string>();
            public int Depth;
        }

        // index of the quote that closes the one at position i ("" and '' inside belong to the text); -1 when it never closes
        static int EndQuote(string f, int i)
        {
            char q = f[i];
            for (int k = i + 1; k < f.Length; k++)
            {
                if (f[k] != q) continue;
                if (k + 1 < f.Length && f[k + 1] == q)
                {
                    k++;
                    continue;
                }
                return k;
            }
            return -1;
        }

        // index of the bracket that closes the ( at position open; -1 when it never closes
        static int CloseOf(string f, int open)
        {
            int depth = 0;
            for (int i = open; i < f.Length; i++)
            {
                char ch = f[i];
                if (ch == '"' || ch == '\'')
                {
                    i = EndQuote(f, i);
                    if (i < 0) return -1;
                    continue;
                }
                if (ch == '(') depth++;
                else if (ch == ')')
                {
                    depth--;
                    if (depth == 0) return i;
                }
            }
            return -1;
        }

        // the parts of a function, split at the commas that belong to it
        static List<string> SplitArgs(string inner)
        {
            List<string> parts = new List<string>();
            int depth = 0, start = 0;
            for (int i = 0; i < inner.Length; i++)
            {
                char ch = inner[i];
                if (ch == '"' || ch == '\'')
                {
                    int e = EndQuote(inner, i);
                    if (e < 0) break;
                    i = e;
                    continue;
                }
                if (ch == '(' || ch == '{' || ch == '[') depth++;
                else if (ch == ')' || ch == '}' || ch == ']') depth--;
                else if (ch == ',' && depth == 0)
                {
                    parts.Add(inner.Substring(start, i - start).Trim());
                    start = i + 1;
                }
            }
            parts.Add(inner.Substring(start).Trim());
            if (parts.Count == 1 && parts[0].Length == 0) parts.Clear();
            return parts;
        }

        // every function of the formula: the outside one first, then the ones inside it
        static void FindCalls(string f, int depth, List<Call> into)
        {
            int i = 0;
            while (i < f.Length)
            {
                char ch = f[i];
                if (ch == '"' || ch == '\'')
                {
                    int e = EndQuote(f, i);
                    if (e < 0) return;
                    i = e + 1;
                    continue;
                }
                if (!(char.IsLetter(ch) || ch == '_'))
                {
                    i++;
                    continue;
                }
                int s = i;
                while (i < f.Length && (char.IsLetterOrDigit(f[i]) || f[i] == '_' || f[i] == '.')) i++;
                if (i >= f.Length || f[i] != '(') continue;
                int close = CloseOf(f, i);
                if (close < 0) return;
                string nm = f.Substring(s, i - s).ToUpperInvariant();
                if (nm.StartsWith("_XLFN.", StringComparison.Ordinal)) nm = nm.Substring(6);
                if (nm.StartsWith("_XLWS.", StringComparison.Ordinal)) nm = nm.Substring(6);
                Call c = new Call();
                c.Name = nm;
                c.Depth = depth;
                c.Text = nm + f.Substring(i, close - i + 1);
                c.Args = SplitArgs(f.Substring(i + 1, close - i - 1));
                into.Add(c);
                if (into.Count < 60 && depth < 12)
                {
                    foreach (string a in c.Args) FindCalls(a, depth + 1, into);
                }
                i = close + 1;
            }
        }

        // a typed number, a typed text or TRUE / FALSE: its value does not need to be shown
        static bool IsConstant(string arg)
        {
            if (arg.Length == 0) return true;
            double d;
            if (double.TryParse(arg, NumberStyles.Float, Inv, out d)) return true;
            if (arg == "TRUE" || arg == "FALSE") return true;
            return arg.Length >= 2 && arg[0] == '"' && EndQuote(arg, 0) == arg.Length - 1;
        }

        // a value as short text for the explanation
        static string Shorten(object res)
        {
            if (res != null && Marshal.IsComObject(res))
            {
                dynamic rng = res;                                // a reference: show its size or its value
                int rr = (int)rng.Rows.Count, cc = (int)rng.Columns.Count;
                if ((long)rr * cc > 1) return "a range, " + rr.ToString("#,##0", Inv) + (rr == 1 ? " row x " : " rows x ") + cc.ToString("#,##0", Inv) + (cc == 1 ? " column" : " columns");
                res = rng.Value;
            }
            object[,] arr = res as object[,];
            if (arr != null) return "a list of " + arr.Length.ToString("#,##0", Inv) + " values";
            if (res == null) return "empty";
            string s = res as string;
            if (s != null) return "\"" + (s.Length > 40 ? s.Substring(0, 40) + "..." : s) + "\"";
            if (res is bool) return (bool)res ? "TRUE" : "FALSE";
            if (res is double)
            {
                double d = (double)res;
                return d.ToString(Math.Abs(d) >= 1000 ? "#,##0.##" : "0.####", Inv);
            }
            return Xl.Str(res);
        }

        // what a piece of the formula gives right now; null when Excel cannot work it out on its own
        static string ValueNow(string expr)
        {
            if (expr.Length == 0 || expr.Length > 250) return null;
            try
            {
                object res = Xl.App.Evaluate(expr);
                return Shorten(res);
            }
            catch { return null; }
        }

        static void Note(List<string> notes, string text)
        {
            if (!notes.Contains(text)) notes.Add(text);
        }

        static string ErrorMeaning(string err)
        {
            switch (err)
            {
                case "#N/A": return "#N/A means: the value that was looked for was not found.";
                case "#DIV/0!": return "#DIV/0! means: a number is divided by zero or by an empty cell.";
                case "#VALUE!": return "#VALUE! means: a text is used where a number is needed (often a number stored as text, or a space in a cell).";
                case "#REF!": return "#REF! means: a cell, a column or a sheet the formula points to was deleted.";
                case "#NAME?": return "#NAME? means: Excel does not know a name in the formula: a misspelt function, a missing add-in or text without quotes.";
                case "#NUM!": return "#NUM! means: a number is too large, too small or not possible for this calculation.";
                case "#NULL!": return "#NULL! means: two ranges that do not cross are joined by a space.";
                default: return "";
            }
        }

        public static void ExplainFormula()
        {
            dynamic cell;
            string f, place;
            object val = null;
            try
            {
                cell = Xl.App.ActiveCell;
                f = Convert.ToString((object)cell.Formula, Inv);
                place = "'" + (string)cell.Worksheet.Name + "'!" + (string)cell.Address[false, false];
                try { val = cell.Value; } catch { val = null; }
            }
            catch
            {
                Ui.Info("Click a cell that holds a formula, then press Explain Formula.");
                return;
            }
            if (f == null || f.Length < 2 || f[0] != '=')
            {
                Ui.Info("The active cell has no formula.\n\nClick a cell that holds a formula, then press Explain Formula.");
                return;
            }
            string result = Xl.IsBlank(val) ? "(empty)" : Shorten(val);
            List<Call> calls = new List<Call>();
            FindCalls(f.Substring(1), 0, calls);

            const string NL = "\r\n";
            StringBuilder sb = new StringBuilder();
            sb.Append("Cell:      ").Append(place).Append(NL);
            sb.Append("Formula:   ").Append(f).Append(NL);
            sb.Append("Result:    ").Append(result).Append(NL).Append(NL);
            int errStep = -1;
            if (calls.Count == 0)
            {
                sb.Append("This formula uses no function. It only calculates with the cells and the numbers you see in it.").Append(NL).Append(NL);
            }
            else
            {
                sb.Append("STEP BY STEP   (the outside function first, then the functions inside it)").Append(NL).Append(NL);
                int evals = 0;
                for (int i = 0; i < calls.Count; i++)
                {
                    Call c = calls[i];
                    string pad = new string(' ', c.Depth * 4);
                    string[] info;
                    if (!Fn.TryGetValue(c.Name, out info)) info = null;
                    sb.Append(pad).Append(i + 1).Append(". ").Append(c.Name).Append("   ").Append(info != null ? info[0] : "a function of Excel or of an add-in").Append(NL);
                    for (int a = 0; a < c.Args.Count; a++)
                    {
                        string arg = c.Args[a];
                        string label = info != null && a + 1 < info.Length ? info[a + 1] : "part " + (a + 1);
                        sb.Append(pad).Append("      ").Append(label).Append(":  ").Append(arg.Length > 90 ? arg.Substring(0, 90) + "..." : arg);
                        if (evals < 60 && !IsConstant(arg))
                        {
                            evals++;
                            string now = ValueNow(arg);
                            if (now != null) sb.Append("     [now: ").Append(now).Append("]");
                        }
                        sb.Append(NL);
                    }
                    if (evals < 60)
                    {
                        evals++;
                        string gives = ValueNow(c.Text);
                        if (gives != null)
                        {
                            sb.Append(pad).Append("      gives now:  ").Append(gives).Append(NL);
                            if (gives.StartsWith("#", StringComparison.Ordinal)) errStep = i;
                        }
                    }
                    sb.Append(NL);
                }
            }

            // ---- what to watch out for ----
            List<string> watch = new List<string>();
            int ifCount = 0;
            foreach (Call k in calls)
            {
                if (k.Name == "IF") ifCount++;
                if (k.Name == "VLOOKUP" || k.Name == "HLOOKUP")
                {
                    string last = k.Args.Count >= 4 ? k.Args[3].ToUpperInvariant() : "";
                    if (last == "" || last == "TRUE" || last == "1")
                        Note(watch, k.Name + " uses nearest match here (no FALSE at the end). On a list that is not sorted it can return a wrong row without any error. Put FALSE as the last part for an exact match.");
                    int colNo;
                    if (k.Args.Count >= 3 && int.TryParse(k.Args[2], NumberStyles.Integer, Inv, out colNo))
                        Note(watch, "The number " + colNo + " in " + k.Name + " is typed in. When someone adds a column or row inside the table, the formula returns a different one.");
                }
                if (k.Name == "MATCH" && (k.Args.Count < 3 || k.Args[2] == "1"))
                    Note(watch, "MATCH uses nearest match here (the last part is missing or 1). The list must be sorted from small to large; use 0 for an exact match.");
                if (k.Name == "NOW" || k.Name == "TODAY" || k.Name == "RAND" || k.Name == "RANDBETWEEN" || k.Name == "OFFSET" || k.Name == "INDIRECT")
                    Note(watch, k.Name + " is worked out again on every change in the workbook. Many such formulas make a file slow.");
            }
            if (ifCount >= 4) Note(watch, ifCount + " IFs inside one formula are hard to read and easy to break. IFS, or a small table with a lookup, does the same job.");
            if (f.IndexOf('[') >= 0 && f.IndexOf(".xls", StringComparison.OrdinalIgnoreCase) >= 0)
                Note(watch, "This formula reads another workbook. It shows old values when that file is moved, renamed or not updated.");
            string meaning = result.Length > 0 && result[0] == '#' ? ErrorMeaning(result) : "";
            if (meaning.Length > 0)
            {
                sb.Append("WHY THE ERROR").Append(NL).Append("  ").Append(meaning).Append(NL);
                if (errStep >= 0) sb.Append("  It starts in step ").Append(errStep + 1).Append(" (").Append(calls[errStep].Name).Append("): that is the innermost step that gives an error.").Append(NL);
                sb.Append(NL);
            }
            if (watch.Count > 0)
            {
                sb.Append("WATCH OUT").Append(NL);
                foreach (string w in watch) sb.Append("  - ").Append(w).Append(NL);
                sb.Append(NL);
            }
            sb.Append("Worked out on this PC by ").Append(Design.AppName).Append(". The values in [ ] are the values right now.");
            using (TextForm tf = new TextForm("Explain Formula   " + place, sb.ToString()))
            {
                tf.ShowDialog(new ExcelWin());
            }
        }

        // =================================================================
        //  SMART ANALYST
        // =================================================================
        sealed class Report
        {
            public string Title = "", Period, NumFmt = "#,##0";
            public List<string> Says = new List<string>(), Next = new List<string>(), TableTitles = new List<string>();
            public List<object[,]> Tables = new List<object[,]>();
            public List<int> Kinds = new List<int>();          // 0 = by a column (bar chart), 1 = by month (line chart)
        }

        static string Move(double cur, double prev)
        {
            if (prev == 0) return cur == 0 ? "no change" : "up from nothing";
            double pct = (cur - prev) / Math.Abs(prev);
            if (Math.Abs(pct) < 0.0005) return "no change";
            return (pct > 0 ? "up " : "down ") + P(Math.Abs(pct));
        }

        static string MonthName(int key)
        {
            return new DateTime(key / 12, key % 12 + 1, 1).ToString("MMM-yy", Inv);
        }

        // Pure arithmetic: no Excel calls. mc = the number column; firstRow = sheet row of the first selected row.
        static Report Analyse(object[,] v, int hr, string[] h, ColInfo[] ci, int mc, int firstRow)
        {
            int nr = v.GetLength(0), nc = v.GetLength(1);
            Report rep = new Report();
            string mName = h[mc];
            // up to three columns to group by, and the first date column
            List<int> cats = new List<int>();
            int dc = 0;
            for (int c = 1; c <= nc; c++)
            {
                if (c == mc) continue;
                if (ci[c].Kind == "cat" && ci[c].Distinct >= 2 && cats.Count < 3) cats.Add(c);
                if (ci[c].Kind == "date" && dc == 0) dc = c;
            }
            int nCat = cats.Count;
            Dictionary<string, double>[] sums = new Dictionary<string, double>[nCat];
            int[] blanks = new int[nCat];
            for (int k = 0; k < nCat; k++) sums[k] = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            SortedDictionary<int, double> months = new SortedDictionary<int, double>();
            List<double> vals = new List<double>();
            double total = 0, negSum = 0, maxVal = double.MinValue;
            int rows = 0, noNum = 0, negRows = 0, maxRow = 0;
            DateTime dMin = DateTime.MaxValue, dMax = DateTime.MinValue;
            for (int r = hr + 1; r <= nr; r++)
            {
                if (Xl.RowBlank(v, r) || !Profiler.RowIsData(v, r, nc)) continue;
                double d;
                if (!Logic.TryNum(v[r, mc], out d))
                {
                    noNum++;
                    continue;
                }
                rows++;
                total += d;
                vals.Add(d);
                if (d < 0)
                {
                    negRows++;
                    negSum += d;
                }
                if (d > maxVal)
                {
                    maxVal = d;
                    maxRow = r;
                }
                for (int k = 0; k < nCat; k++)
                {
                    object x = v[r, cats[k]];
                    if (Xl.IsBlank(x)) blanks[k]++;
                    string key = Xl.GroupKey(x);
                    double cur;
                    sums[k][key] = sums[k].TryGetValue(key, out cur) ? cur + d : d;
                }
                DateTime dt;
                if (dc != 0 && Logic.TryDate(v[r, dc], out dt))
                {
                    if (dt < dMin) dMin = dt;
                    if (dt > dMax) dMax = dt;
                    int mk = dt.Year * 12 + dt.Month - 1;
                    double cm;
                    months[mk] = months.TryGetValue(mk, out cm) ? cm + d : d;
                }
            }
            if (rows < 2) return null;

            rep.Title = "Smart Analyst: " + mName;
            rep.NumFmt = Math.Abs(total) >= 1000 ? "#,##0" : "#,##0.00";
            if (dc != 0 && dMin <= dMax) rep.Period = dMin.ToString("dd-MMM-yyyy", Inv) + " to " + dMax.ToString("dd-MMM-yyyy", Inv);
            rep.Says.Add("Total " + mName + " is " + F(total) + " from " + N(rows) + " rows. The average row is " + F(total / rows) + ".");

            // ---- who drives the total ----
            for (int k = 0; k < nCat; k++)
            {
                string cName = h[cats[k]];
                List<KeyValuePair<string, double>> list = new List<KeyValuePair<string, double>>(sums[k]);
                list.Sort(delegate (KeyValuePair<string, double> p, KeyValuePair<string, double> q) { return q.Value.CompareTo(p.Value); });
                int n = list.Count;
                if (n == 0) continue;
                if (total > 0 && list[0].Value > 0)
                {
                    double top3 = 0, run = 0;
                    int n80 = 0;
                    for (int i = 0; i < n; i++)
                    {
                        if (i < 3) top3 += list[i].Value;
                        if (n80 == 0)
                        {
                            run += list[i].Value;
                            if (run >= 0.8 * total) n80 = i + 1;
                        }
                    }
                    string line = "By " + cName + ": '" + list[0].Key + "' leads with " + F(list[0].Value) + " (" + P(list[0].Value / total) + " of the total).";
                    if (n > 4) line += " The top 3 of " + N(n) + " make " + P(top3 / total) + ".";
                    if (n > 4 && n80 > 0) line += " " + N(n80) + " of them give 80%.";
                    rep.Says.Add(line);
                    if (n > 2 && list[0].Value / total >= 0.5) rep.Says.Add("Watch: more than half of " + mName + " depends on one " + cName + " ('" + list[0].Key + "').");
                }
                else rep.Says.Add("By " + cName + ": the largest is '" + list[0].Key + "' with " + F(list[0].Value) + ".");
                if (blanks[k] > 0) rep.Says.Add(N(blanks[k]) + " rows have no " + cName + ". They are counted as (blank).");

                // table: the first 15 and one line for the rest
                int show = Math.Min(n, 15);
                bool rest = n > show;
                object[,] t = Xl.New2D(show + (rest ? 2 : 1), 4);
                t[1, 1] = cName;
                t[1, 2] = mName;
                t[1, 3] = "Share";
                t[1, 4] = "Running share";
                double cum = 0;
                for (int i = 0; i < show; i++)
                {
                    cum += list[i].Value;
                    t[i + 2, 1] = list[i].Key;
                    t[i + 2, 2] = list[i].Value;
                    if (total != 0)
                    {
                        t[i + 2, 3] = list[i].Value / total;
                        t[i + 2, 4] = cum / total;
                    }
                }
                if (rest)
                {
                    t[show + 2, 1] = "All other (" + N(n - show) + ")";
                    t[show + 2, 2] = total - cum;
                    if (total != 0)
                    {
                        t[show + 2, 3] = (total - cum) / total;
                        t[show + 2, 4] = 1.0;
                    }
                }
                rep.TableTitles.Add(mName + " by " + cName);
                rep.Tables.Add(t);
                rep.Kinds.Add(0);
            }

            // ---- over time ----
            int mCount = months.Count;
            if (mCount >= 2)
            {
                int[] mKeys = new int[mCount];
                double[] mVals = new double[mCount];
                int at = 0;
                foreach (KeyValuePair<int, double> kv in months)
                {
                    mKeys[at] = kv.Key;
                    mVals[at] = kv.Value;
                    at++;
                }
                int best = 0, worst = 0, lastI = mCount - 1;
                for (int i = 1; i < mCount; i++)
                {
                    if (mVals[i] > mVals[best]) best = i;
                    if (mVals[i] < mVals[worst]) worst = i;
                }
                rep.Says.Add("Latest month " + MonthName(mKeys[lastI]) + ": " + F(mVals[lastI]) + ", " + Move(mVals[lastI], mVals[lastI - 1]) + " against " + MonthName(mKeys[lastI - 1]) + ".");
                double lastYear;
                if (months.TryGetValue(mKeys[lastI] - 12, out lastYear))
                    rep.Says.Add("Against the same month last year (" + MonthName(mKeys[lastI] - 12) + ", " + F(lastYear) + ") it is " + Move(mVals[lastI], lastYear) + ".");
                if (mCount >= 3)
                {
                    rep.Says.Add("Best month: " + MonthName(mKeys[best]) + " with " + F(mVals[best]) + ". Weakest month: " + MonthName(mKeys[worst]) + " with " + F(mVals[worst]) + ".");
                    double mean = 0, sxy = 0, sxx = 0, mid = (mCount - 1) / 2.0;
                    for (int i = 0; i < mCount; i++) mean += mVals[i];
                    mean /= mCount;
                    for (int i = 0; i < mCount; i++)
                    {
                        sxy += (i - mid) * (mVals[i] - mean);
                        sxx += (i - mid) * (i - mid);
                    }
                    double slope = sxx > 0 ? sxy / sxx : 0;
                    if (Math.Abs(slope) * (mCount - 1) < 0.1 * Math.Abs(mean)) rep.Says.Add("Over " + mCount + " months the level is steady: no clear rise or fall.");
                    else if (slope != 0) rep.Says.Add("Over " + mCount + " months the trend is " + (slope > 0 ? "rising" : "falling") + ": about " + F(Math.Abs(slope)) + (slope > 0 ? " more" : " less") + " each month.");
                }
                object[,] mt = Xl.New2D(mCount + 1, 4);
                mt[1, 1] = "Month";
                mt[1, 2] = mName;
                mt[1, 3] = "Change";
                mt[1, 4] = "Change %";
                for (int i = 0; i < mCount; i++)
                {
                    mt[i + 2, 1] = MonthName(mKeys[i]);
                    mt[i + 2, 2] = mVals[i];
                    if (i == 0) continue;
                    mt[i + 2, 3] = mVals[i] - mVals[i - 1];
                    if (mVals[i - 1] != 0) mt[i + 2, 4] = (mVals[i] - mVals[i - 1]) / Math.Abs(mVals[i - 1]);
                }
                rep.TableTitles.Add(mName + " by month");
                rep.Tables.Add(mt);
                rep.Kinds.Add(1);
            }

            // ---- unusual rows ----
            int nVals = vals.Count;
            if (nVals >= 8)
            {
                double[] sorted = vals.ToArray();
                Array.Sort(sorted);
                double q1 = Logic.Percentile(sorted, 0.25), q3 = Logic.Percentile(sorted, 0.75);
                double fence = q3 + 3 * (q3 - q1);
                int big = 0;
                if (q3 > q1)
                {
                    for (int i = nVals - 1; i >= 0 && sorted[i] > fence; i--) big++;
                }
                if (big > 0)
                {
                    string whose = nCat > 0 ? " for '" + Xl.GroupKey(v[maxRow, cats[0]]) + "'" : "";
                    rep.Says.Add(N(big) + (big == 1 ? " row is" : " rows are") + " unusually large (above " + F(fence) + "). The largest is " + F(maxVal) + whose + ", in sheet row " + N(firstRow + maxRow - 1) + ".");
                }
            }
            if (negRows > 0) rep.Says.Add(N(negRows) + (negRows == 1 ? " row is" : " rows are") + " negative, " + F(negSum) + " in total: returns, credit notes or corrections.");
            if (noNum > 0) rep.Says.Add(N(noNum) + " rows have no number in " + mName + " and were left out.");
            if (ci[mc].TextNum > 0) rep.Says.Add(N(ci[mc].TextNum) + " numbers in " + mName + " are stored as text. Smart Clean turns them into real numbers.");

            // ---- what to do next ----
            if (mCount >= 6) rep.Next.Add("Forecast  -  the next 3, 6 or 12 months from this history, with a likely range.");
            if (mCount >= 2 && nCat > 0) rep.Next.Add("Why It Changed  -  which " + h[cats[0]] + " caused the move between two months.");
            if (nCat > 0) rep.Next.Add("ABC Analysis  -  every " + h[cats[0]] + " classed A, B or C by its share.");
            rep.Next.Add("One-Click Dashboard  -  these numbers as KPI cards, charts and filter buttons (free).");
            return rep;
        }

        static int Block(dynamic ws, int row, string title, List<string> lines)
        {
            if (lines.Count == 0) return row;
            ws.Cells[row, 1].Value2 = title;
            ws.Cells[row, 1].Font.Bold = true;
            ws.Cells[row, 1].Font.Size = 12;
            object[,] a = Xl.New2D(lines.Count, 1);
            for (int i = 0; i < lines.Count; i++) a[i + 1, 1] = "•  " + lines[i];
            ws.Cells[row + 1, 1].Resize[lines.Count, 1].Value2 = a;
            return row + lines.Count + 2;
        }

        static void AddChart(dynamic ws, dynamic top, int tr, int kind, string title, int row)
        {
            try
            {
                int n = Math.Min(tr - 1, kind == 0 ? 10 : 36);
                int skip = kind == 0 ? 1 : tr - n;               // months: the latest ones
                dynamic anchor = ws.Range[ws.Cells[row + 1, 6], ws.Cells[row + 15, 13]];
                dynamic co = ws.ChartObjects().Add((double)anchor.Left, (double)anchor.Top, (double)anchor.Width, (double)anchor.Height);
                dynamic ch = co.Chart;
                ch.ChartType = kind == 0 ? C.xlBarClustered : C.xlLineMarkers;
                dynamic s1 = ch.SeriesCollection().NewSeries();
                s1.Name = title;
                s1.Values = top.Offset[skip, 1].Resize[n, 1];
                s1.XValues = top.Offset[skip, 0].Resize[n, 1];
                ch.HasTitle = true;
                ch.ChartTitle.Text = title + (kind == 0 && tr - 1 > n ? "  (top " + n + ")" : "");
                ch.HasLegend = false;
                if (kind == 0)
                {
                    try { ch.Axes(C.xlCategory).ReversePlotOrder = true; } catch { }
                }
            }
            catch { }
        }

        static void WriteReport(dynamic wb, Report rep, string srcName)
        {
            dynamic ws = Xl.AddSheet(wb, "Smart Analyst");
            ws.Range["A1"].Value2 = rep.Title;
            ws.Range["A1"].Font.Size = 14;
            ws.Range["A1"].Font.Bold = true;
            ws.Range["A2"].Value2 = "Source: " + srcName + (rep.Period != null ? "   |   " + rep.Period : "") + "   |   worked out on this PC by " + Design.AppName;
            ws.Range["A2"].Font.Color = Design.Slate;
            ws.Columns[1].ColumnWidth = 34;
            ws.Range["B:D"].ColumnWidth = 16;
            int row = 4;
            row = Block(ws, row, "What the numbers say", rep.Says);
            row = Block(ws, row, "What to do next", rep.Next);
            for (int i = 0; i < rep.Tables.Count; i++)
            {
                object[,] t = rep.Tables[i];
                int tr = t.GetLength(0);
                ws.Cells[row, 1].Value2 = rep.TableTitles[i];
                ws.Cells[row, 1].Font.Bold = true;
                ws.Cells[row, 1].Font.Size = 12;
                dynamic top = ws.Cells[row + 1, 1];
                Xl.WriteTable(top, t, 1);
                Xl.StyleHeader(top.Resize[1, 4]);
                if (tr > 1)
                {
                    top.Offset[1, 1].Resize[tr - 1, 1].NumberFormat = rep.NumFmt;
                    if (rep.Kinds[i] == 0) top.Offset[1, 2].Resize[tr - 1, 2].NumberFormat = "0.0%";
                    else
                    {
                        top.Offset[1, 2].Resize[tr - 1, 1].NumberFormat = rep.NumFmt;
                        top.Offset[1, 3].Resize[tr - 1, 1].NumberFormat = "0.0%";
                    }
                    AddChart(ws, top, tr, rep.Kinds[i], rep.TableTitles[i], row);
                }
                row += Math.Max(tr + 3, 18);
            }
            ws.Activate();
        }

        public static void SmartAnalyst()
        {
            dynamic src = Xl.PickRange("Smart Analyst");
            if (src == null) return;
            object[,] v = Xl.ReadTyped(src);
            int nr = v.GetLength(0), nc = v.GetLength(1);
            if (nr < 3 || nc < 2)
            {
                Ui.Info("Select a table with a header row, at least two columns and a few rows.");
                return;
            }
            int hr = Profiler.FindHeaderRow(v);
            string[] h = Profiler.MakeHeaders(v, hr);
            ColInfo[] ci = Profiler.Profile(v, hr, h, 20000);
            // the number to analyse: amount / value first, then quantity, then any other number
            List<int> nums = new List<int>();
            for (int rank = 1; rank <= 3; rank++)
            {
                for (int c = 1; c <= nc; c++)
                {
                    if (ci[c].Kind == "num" && Logic.MeasureRank(ci[c].Hdr) == rank && !Logic.MeasureIsAvg(ci[c].Hdr)) nums.Add(c);
                }
            }
            for (int c = 1; c <= nc; c++)
            {
                if (ci[c].Kind == "num" && !nums.Contains(c)) nums.Add(c);
            }
            if (nums.Count == 0)
            {
                Ui.Info("No number column was found. Run Text to Numbers first if the values are stored as text.");
                return;
            }
            int mc = nums[0];
            if (nums.Count > 1)
            {
                string[] items = new string[nums.Count];
                for (int i = 0; i < nums.Count; i++) items[i] = Xl.ColLetter(nums[i]) + "   " + h[nums[i]] + (i == 0 ? "     (suggested)" : "");
                int pick = Ui.Pick("Smart Analyst", "Which number should be analysed? Press Enter for the suggested one.", items, null);
                if (pick < 0) return;
                mc = nums[pick];
            }
            Report rep = Analyse(v, hr, h, ci, mc, (int)src.Row);
            if (rep == null)
            {
                Ui.Info("There are not enough rows with a number in '" + h[mc] + "'.");
                return;
            }
            string srcName = "'" + (string)src.Worksheet.Name + "'!" + (string)src.Address[false, false];
            using (new Speed())
            {
                WriteReport(src.Worksheet.Parent, rep, srcName);
            }
            StringBuilder sb = new StringBuilder(rep.Title);
            sb.Append("\r\n");
            foreach (string s in rep.Says) sb.Append("- ").Append(s).Append("\r\n");
            if (Ui.Ok("Smart Analyst wrote " + rep.Says.Count + " findings and " + rep.Tables.Count + " tables with charts on a new sheet.\n\n" +
                      "OK  =  also copy the findings, ready to paste into an email or a chat\nCancel  =  no, thanks"))
            {
                try { Clipboard.SetText(sb.ToString()); } catch { }
            }
        }

        // =================================================================
        //  DUPLICATE PAYMENT CHECK
        // =================================================================
        sealed class PayRow
        {
            public int Row;                      // row on the sheet
            public string PartyText = "", Party = "", Inv = "", InvKey = "";
            public long Paise;
            public double Amount;
            public DateTime Date;
            public bool HasDate;
        }

        sealed class Finding
        {
            public int Level;                    // 3 = high, 2 = medium, 1 = low
            public string Why = "";
            public double Amount;
            public List<int> Rows = new List<int>();
        }

        // invoice number without case, punctuation and leading zeros:  "INV-0012/A" -> "INV12A"
        static string InvKey(string s)
        {
            StringBuilder sb = new StringBuilder(s.Length);
            int i = 0;
            while (i < s.Length)
            {
                char ch = char.ToUpperInvariant(s[i]);
                if (ch >= '0' && ch <= '9')
                {
                    int e = i;
                    while (e < s.Length && s[e] >= '0' && s[e] <= '9') e++;
                    int z = i;
                    while (z < e - 1 && s[z] == '0') z++;
                    sb.Append(s, z, e - z);
                    i = e;
                    continue;
                }
                if (ch >= 'A' && ch <= 'Z') sb.Append(ch);
                i++;
            }
            return sb.ToString();
        }

        // true when two different numbers differ by one character, or by two neighbours that changed places
        static bool Near(string a, string b)
        {
            int la = a.Length, lb = b.Length;
            if (la < 4 || lb < 4 || Math.Abs(la - lb) > 1) return false;
            int i = 0;
            while (i < la && i < lb && a[i] == b[i]) i++;
            if (la == lb)
            {
                if (i >= la) return true;
                if (string.CompareOrdinal(a, i + 1, b, i + 1, la - i - 1) == 0) return true;
                return i + 1 < la && a[i] == b[i + 1] && a[i + 1] == b[i] && string.CompareOrdinal(a, i + 2, b, i + 2, la - i - 2) == 0;
            }
            if (la > lb) return string.CompareOrdinal(a, i + 1, b, i, lb - i) == 0;
            return string.CompareOrdinal(a, i, b, i + 1, la - i) == 0;
        }

        static void Push(Dictionary<string, List<int>> map, string key, int value)
        {
            List<int> lst;
            if (!map.TryGetValue(key, out lst))
            {
                lst = new List<int>();
                map[key] = lst;
            }
            lst.Add(value);
        }

        // Pure logic: no Excel calls. days = how close two dates must be for a "same amount" finding.
        static List<Finding> FindDuplicates(List<PayRow> rows, int days, out int tooBig)
        {
            List<Finding> found = new List<Finding>();
            tooBig = 0;
            Dictionary<string, List<int>> exact = new Dictionary<string, List<int>>(StringComparer.Ordinal);
            Dictionary<string, List<int>> bucket = new Dictionary<string, List<int>>(StringComparer.Ordinal);
            Dictionary<string, List<int>> byInv = new Dictionary<string, List<int>>(StringComparer.Ordinal);
            for (int i = 0; i < rows.Count; i++)
            {
                PayRow p = rows[i];
                string pa = p.Party + "\u0001" + p.Paise.ToString(Inv);
                Push(bucket, pa, i);
                if (p.InvKey.Length == 0) continue;
                Push(exact, pa + "\u0001" + p.InvKey, i);
                Push(byInv, p.Party + "\u0001" + p.InvKey, i);
            }
            // high: same party, same amount, same invoice number
            foreach (KeyValuePair<string, List<int>> kv in exact)
            {
                List<int> g = kv.Value;
                if (g.Count < 2) continue;
                bool sameText = true;
                for (int i = 1; i < g.Count; i++)
                {
                    if (!string.Equals(rows[g[i]].Inv, rows[g[0]].Inv, StringComparison.OrdinalIgnoreCase)) sameText = false;
                }
                Finding f = new Finding();
                f.Level = 3;
                f.Why = sameText ? "Same party, same invoice number, same amount" : "Same party and amount; the invoice number is only written differently";
                f.Amount = (g.Count - 1) * Math.Abs(rows[g[0]].Amount);
                f.Rows.AddRange(g);
                found.Add(f);
            }
            // medium and low: same party and amount, invoice numbers nearly the same, or dates close together
            foreach (KeyValuePair<string, List<int>> kv in bucket)
            {
                List<int> g = kv.Value;
                if (g.Count < 2) continue;
                if (g.Count > 400)
                {
                    tooBig++;
                    continue;
                }
                for (int a = 0; a < g.Count; a++)
                {
                    for (int b = a + 1; b < g.Count; b++)
                    {
                        PayRow x = rows[g[a]], y = rows[g[b]];
                        if (x.InvKey.Length > 0 && x.InvKey == y.InvKey) continue;          // already a high finding
                        int lvl = 0;
                        string why = "";
                        if (x.InvKey.Length > 0 && y.InvKey.Length > 0 && Near(x.InvKey, y.InvKey))
                        {
                            bool lastOnly = x.InvKey.Length == y.InvKey.Length && string.CompareOrdinal(x.InvKey, 0, y.InvKey, 0, x.InvKey.Length - 1) == 0;
                            lvl = lastOnly ? 1 : 2;
                            why = lastOnly ? "Same party and amount; the invoice numbers differ only in the last character (two bills in a row?)"
                                           : "Same party and amount; the invoice numbers are nearly the same (typing mistake?)";
                        }
                        else if (x.HasDate && y.HasDate && Math.Abs((x.Date - y.Date).TotalDays) <= days)
                        {
                            lvl = 1;
                            why = "Same party and amount within " + days + " days, different invoice numbers";
                        }
                        if (lvl == 0) continue;
                        Finding f = new Finding();
                        f.Level = lvl;
                        f.Why = why;
                        f.Amount = Math.Abs(x.Amount);
                        f.Rows.Add(g[a]);
                        f.Rows.Add(g[b]);
                        found.Add(f);
                        if (found.Count >= 5000) return found;
                    }
                }
            }
            // low: one invoice number of a party with different amounts
            foreach (KeyValuePair<string, List<int>> kv in byInv)
            {
                List<int> g = kv.Value;
                if (g.Count < 2 || g.Count > 50) continue;
                bool differ = false;
                for (int i = 1; i < g.Count; i++)
                {
                    if (rows[g[i]].Paise != rows[g[0]].Paise) differ = true;
                }
                if (!differ) continue;
                Finding f = new Finding();
                f.Level = 1;
                f.Why = "Same party and invoice number with different amounts (part payment, or one entry is wrong?)";
                f.Amount = Math.Abs(rows[g[0]].Amount);
                f.Rows.AddRange(g);
                found.Add(f);
            }
            return found;
        }

        public static void DuplicatePayments()
        {
            const string title = "Duplicate Payments";
            dynamic src = Xl.PickRange(title, "Select the list of invoices or payments (include the header row):");
            if (src == null) return;
            object[,] v = Xl.ReadTyped(src);
            int nr = v.GetLength(0), nc = v.GetLength(1);
            if (nr < 3 || nc < 2)
            {
                Ui.Info("Select a table with a header row, an amount column and a party or invoice column.");
                return;
            }
            int hr = Profiler.FindHeaderRow(v);
            string[] h = Profiler.MakeHeaders(v, hr);
            ColInfo[] ci = Profiler.Profile(v, hr, h, 5000);
            bool[] numeric = new bool[nc + 1];
            bool anyNum = false;
            int dcol = 0;
            for (int c = 1; c <= nc; c++)
            {
                numeric[c] = ci[c].NumCount > 0 && ci[c].NumCount >= 0.8 * ci[c].NonBlank;
                if (numeric[c]) anyNum = true;
                if (dcol == 0 && ci[c].Kind == "date") dcol = c;
            }
            if (!anyNum)
            {
                Ui.Info("No number column was found for the amount. Run Text to Numbers first if the amounts are stored as text.");
                return;
            }
            int ac = Ui.PickColumn(title + " (1 of 3)", "Which column holds the AMOUNT?", h, numeric);
            if (ac == 0) return;
            int pc = Ui.PickColumn(title + " (2 of 3)", "Which column holds the PARTY (vendor, customer)?", h, null);
            if (pc == 0) return;
            int ic = Ui.PickColumn(title + " (3 of 3)", "Which column holds the INVOICE / BILL / REFERENCE number?", h, null);
            if (ic == 0) return;
            if (dcol == ic || dcol == pc || dcol == ac) dcol = 0;

            const int days = 7;
            int firstRow = (int)src.Row, firstCol = (int)src.Column;
            List<PayRow> rows = new List<PayRow>();
            for (int r = hr + 1; r <= nr; r++)
            {
                double amt;
                if (Xl.RowBlank(v, r) || !Profiler.RowIsData(v, r, nc) || !Logic.TryNum(v[r, ac], out amt) || amt == 0) continue;
                PayRow p = new PayRow();
                p.Row = firstRow + r - 1;
                p.PartyText = Xl.Str(v[r, pc]).Trim();
                p.Party = Logic.NormName(p.PartyText);
                if (p.Party.Length == 0) p.Party = p.PartyText.ToLowerInvariant();
                p.Inv = Xl.Str(v[r, ic]).Trim();
                p.InvKey = InvKey(p.Inv);
                p.Amount = amt;
                p.Paise = (long)Math.Round(amt * 100.0);
                DateTime dt;
                if (dcol != 0 && Logic.TryDate(v[r, dcol], out dt))
                {
                    p.Date = dt;
                    p.HasDate = true;
                }
                rows.Add(p);
            }
            if (rows.Count < 2)
            {
                Ui.Info("There are not enough rows with an amount to compare.");
                return;
            }
            int tooBig;
            List<Finding> found = FindDuplicates(rows, days, out tooBig);
            if (found.Count == 0)
            {
                Ui.Info(N(rows.Count) + " rows checked. No possible duplicate was found.");
                return;
            }
            found.Sort(delegate (Finding p, Finding q)
            {
                int k = q.Level.CompareTo(p.Level);
                return k != 0 ? k : q.Amount.CompareTo(p.Amount);
            });
            string[] levelName = { "", "Low", "Medium", "High" };
            int[] count = new int[4], lines = new int[4];
            double risk = 0;
            List<object[]> outRows = new List<object[]>();
            for (int i = 0; i < found.Count; i++)
            {
                Finding f = found[i];
                count[f.Level]++;
                if (f.Level == 3) risk += f.Amount;
                foreach (int idx in f.Rows)
                {
                    PayRow p = rows[idx];
                    lines[f.Level]++;
                    outRows.Add(new object[] { (double)(i + 1), levelName[f.Level], f.Why, (double)p.Row, p.PartyText, p.Inv, p.HasDate ? (object)p.Date : null, p.Amount });
                }
            }
            string sheet = (string)src.Worksheet.Name;
            string[] heads = { "Finding", "Confidence", "Why it is listed", "Sheet row", h[pc], h[ic], dcol != 0 ? h[dcol] : "Date", h[ac] };
            using (new Speed())
            {
                dynamic ws = Xl.AddSheet(src.Worksheet.Parent, "Duplicate Check");
                ws.Range["A1"].Value2 = "Duplicate Payment Check: " + sheet;
                ws.Range["A1"].Font.Size = 14;
                ws.Range["A1"].Font.Bold = true;
                ws.Range["A2"].Value2 = N(rows.Count) + " rows checked.  " + N(found.Count) + " findings:  " + count[3] + " high,  " + count[2] + " medium,  " + count[1] + " low." +
                                        (count[3] > 0 ? "   Amount that may have been entered or paid twice (high): " + F(risk) : "");
                ws.Range["A3"].Value2 = "High = same party, amount and invoice number.  Medium = invoice numbers nearly the same.  Low = worth a look.  Dates within " + days + " days count as close." +
                                        (tooBig > 0 ? "  " + tooBig + " very large groups of equal amounts were only checked for equal invoice numbers." : "");
                ws.Range["A3"].Font.Color = Design.Slate;
                object[,] t = AnalyseTools.Table(heads, outRows);
                Xl.WriteTable(ws.Range["A5"], t, 1);
                Xl.StyleHeader(ws.Range["A5:H5"]);
                int at = 6;
                int[] colour = { 0, Design.Changed, Design.Unusual, Design.Bad };
                for (int lv = 3; lv >= 1; lv--)
                {
                    if (lines[lv] == 0) continue;
                    ws.Range["B" + at].Resize[lines[lv], 1].Interior.Color = colour[lv];
                    at += lines[lv];
                }
                ws.Range["H6"].Resize[outRows.Count, 1].NumberFormat = "#,##0.00";
                string link = "'" + sheet.Replace("'", "''") + "'!";
                int links = Math.Min(outRows.Count, 500);
                for (int i = 0; i < links; i++)
                {
                    int sheetRow = (int)(double)outRows[i][3];
                    try { ws.Hyperlinks.Add(ws.Cells[6 + i, 4], "", link + Xl.A1(sheetRow, firstCol + ic - 1), Xl.M, sheetRow.ToString(Inv)); } catch { }
                }
                ws.Columns["A:H"].AutoFit();
                ws.Columns[3].ColumnWidth = 60;
                ws.Activate();
            }
            Ui.Info(N(rows.Count) + " rows checked.\n\n" + count[3] + " high, " + count[2] + " medium and " + count[1] + " low findings are on the new 'Duplicate Check' sheet." +
                    (count[3] > 0 ? "\nAmount that may have been entered or paid twice: " + F(risk) : "") + "\n\nClick a row number there to jump to the entry. Your data was not changed.");
        }

        // =================================================================
        //  MAIL TO EACH PARTY (Outlook)
        // =================================================================
        static string Tokens(string text, string party, string total, int rows)
        {
            return text.Replace("{party}", party).Replace("{total}", total).Replace("{rows}", rows.ToString(Inv));
        }

        static string Money(double d)
        {
            return Logic.IndianGroup(d, d == Math.Floor(d) ? 0 : 2);
        }

        // the rows of one party as an HTML table; skipCol (the email column) is left out
        static string HtmlTable(object[,] v, string[] h, ColInfo[] ci, List<int> rws, int skipCol, int totalCol, double sum)
        {
            int nc = v.GetLength(1);
            const string cell = "border:1px solid #cbd5e1;padding:6px 10px;";
            StringBuilder sb = new StringBuilder();
            sb.Append("<table style=\"border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:13px\"><tr>");
            for (int c = 1; c <= nc; c++)
            {
                if (c == skipCol) continue;
                sb.Append("<th style=\"").Append(cell).Append("background:#0f172a;color:#ffffff;text-align:left\">").Append(Esc(h[c])).Append("</th>");
            }
            sb.Append("</tr>");
            int shown = 0;
            foreach (int r in rws)
            {
                if (shown == 300) break;
                shown++;
                sb.Append("<tr>");
                for (int c = 1; c <= nc; c++)
                {
                    if (c == skipCol) continue;
                    object x = v[r, c];
                    bool num = ci[c].Kind == "num" && (x is double || x is decimal);
                    string txt = num ? Money(Xl.Num(x)) : Xl.Str(x);
                    sb.Append("<td style=\"").Append(cell).Append(num ? "text-align:right" : "").Append("\">").Append(Esc(txt)).Append("</td>");
                }
                sb.Append("</tr>");
            }
            if (totalCol != 0 && shown == rws.Count)
            {
                bool first = true;
                sb.Append("<tr>");
                for (int c = 1; c <= nc; c++)
                {
                    if (c == skipCol) continue;
                    string txt = c == totalCol ? Money(sum) : (first ? "Total" : "");
                    first = false;
                    sb.Append("<td style=\"").Append(cell).Append("font-weight:bold;").Append(c == totalCol ? "text-align:right" : "").Append("\">").Append(Esc(txt)).Append("</td>");
                }
                sb.Append("</tr>");
            }
            sb.Append("</table>");
            if (rws.Count > shown) sb.Append("<p>... and ").Append(rws.Count - shown).Append(" more rows.</p>");
            return sb.ToString();
        }

        // puts our text in front of the signature Outlook has already written into the email
        static string Merge(string old, string html)
        {
            if (string.IsNullOrEmpty(old)) return "<html><body>" + html + "</body></html>";
            int p = old.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
            int e = p < 0 ? -1 : old.IndexOf('>', p);
            if (e < 0) return html + old;
            return old.Substring(0, e + 1) + html + old.Substring(e + 1);
        }

        // the rows of one party as a file to attach: how 2 = PDF, otherwise an Excel file
        static List<string> MakeFiles(object[,] v, string[] h, ColInfo[] ci, List<int> rws, int skipCol, string folder, string party, int how)
        {
            List<string> files = new List<string>();
            int nc = v.GetLength(1), cols = skipCol != 0 ? nc - 1 : nc;
            object[,] o = Xl.New2D(rws.Count + 1, cols);
            string[] fmt = new string[cols + 1];
            int at = 0;
            for (int c = 1; c <= nc; c++)
            {
                if (c == skipCol) continue;
                at++;
                o[1, at] = h[c];
                if (ci[c].Kind == "num") fmt[at] = ci[c].AllInt ? "#,##0" : "#,##0.00";
                for (int i = 0; i < rws.Count; i++) o[i + 2, at] = v[rws[i], c];
            }
            dynamic nb = Xl.App.Workbooks.Add(C.xlWBATWorksheet);
            try
            {
                dynamic ws = nb.Worksheets[1];
                ws.Name = Xl.SafeSheetName(party);
                Xl.WriteTable(ws.Range["A1"], o, 1);
                dynamic all = ws.Range["A1"].Resize[rws.Count + 1, cols];
                Xl.StyleHeader(all.Rows[1]);
                for (int j = 1; j <= cols; j++)
                {
                    if (fmt[j] != null) all.Columns[j].Offset[1, 0].Resize[rws.Count, 1].NumberFormat = fmt[j];
                }
                all.EntireColumn.AutoFit();
                string stem = Path.Combine(folder, Xl.SafeFileName(party));
                if (how == 2)
                {
                    try
                    {
                        ws.PageSetup.Orientation = C.xlLandscape;
                        ws.PageSetup.Zoom = false;
                        ws.PageSetup.FitToPagesWide = 1;
                        ws.PageSetup.FitToPagesTall = false;
                    }
                    catch { }
                    ws.ExportAsFixedFormat(C.xlTypePDF, stem + ".pdf");
                    files.Add(stem + ".pdf");
                }
                else
                {
                    nb.SaveAs(stem + ".xlsx", C.xlOpenXMLWorkbook);
                    files.Add(stem + ".xlsx");
                }
            }
            finally
            {
                try { nb.Close(false); } catch { }
            }
            return files;
        }

        public static void MailToEachParty()
        {
            const string title = "Mail to Each Party";
            Type outlook = Type.GetTypeFromProgID("Outlook.Application");
            if (outlook == null)
            {
                Ui.Info("The Outlook desktop app was not found on this PC.\nThis tool needs classic Outlook; it does not work with webmail or the new Outlook app.");
                return;
            }
            dynamic src = Xl.PickRange(title, "Select the table (include the header row). Every party gets an email with only its own rows:");
            if (src == null) return;
            object[,] v = Xl.ReadTyped(src);
            int nr = v.GetLength(0), nc = v.GetLength(1);
            if (nr < 2 || nc < 2)
            {
                Ui.Info("Select a table with a header row, a party column and at least one more column.");
                return;
            }
            int hr = Profiler.FindHeaderRow(v);
            string[] h = Profiler.MakeHeaders(v, hr);
            ColInfo[] ci = Profiler.Profile(v, hr, h, 5000);
            int pc = Ui.PickColumn(title + " (1 of 2)", "One email for each value of which column (party, customer, vendor, salesman)?", h, null);
            if (pc == 0) return;

            // the column with the email addresses: the one where most cells hold an @
            int ec = 0, bestAt = 0;
            for (int c = 1; c <= nc; c++)
            {
                int hits = 0;
                for (int r = hr + 1; r <= nr && r <= hr + 500; r++)
                {
                    string s = v[r, c] as string;
                    if (s != null && s.IndexOf('@') > 0) hits++;
                }
                if (hits > bestAt)
                {
                    bestAt = hits;
                    ec = c;
                }
            }
            if (ec == 0 && !Ui.Ok("No column with email addresses was found.\n\nOK  =  make the emails anyway and type the addresses in Outlook\nCancel  =  stop, and add a column with the email addresses first")) return;
            // the amount for {total} and the total line: the first amount-like or quantity-like column
            int tc = 0;
            for (int rank = 1; rank <= 2 && tc == 0; rank++)
            {
                for (int c = 1; c <= nc && tc == 0; c++)
                {
                    if (ci[c].Kind == "num" && Logic.MeasureRank(ci[c].Hdr) == rank && !Logic.MeasureIsAvg(ci[c].Hdr)) tc = c;
                }
            }
            int how = Ui.Pick(title + " (2 of 2)", "How should each party get its rows?", new[] { "As a table inside the email", "As an Excel file attached", "As a PDF attached", "As a table inside the email and an Excel file attached" }, null);
            if (how < 0) return;

            string lastSubject = Settings.Get("MailSubject"), lastText = Settings.Get("MailText");
            string subject = Ui.AskText("Subject of the emails.\n{party} is replaced by the name of the party.", title, string.IsNullOrEmpty(lastSubject) ? "Statement of account - {party}" : lastSubject);
            if (subject == null) return;
            string message = Ui.AskText("First lines of the email.\n{party} = the party,  {total} = its total" + (tc != 0 ? " of " + h[tc] : "") + ",  {rows} = its number of rows.", title,
                                        string.IsNullOrEmpty(lastText) ? "Dear Sir / Madam, please find the details for {party} below." + (tc != 0 ? " Total: {total}." : "") : lastText);
            if (message == null) return;
            Settings.Set("MailSubject", subject);
            Settings.Set("MailText", message);

            Dictionary<string, List<int>> groups = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);
            List<string> order = new List<string>();
            for (int r = hr + 1; r <= nr; r++)
            {
                if (Xl.RowBlank(v, r) || !Profiler.RowIsData(v, r, nc)) continue;
                string key = Xl.GroupKey(v[r, pc]);
                if (!groups.ContainsKey(key)) order.Add(key);
                Push(groups, key, r);
            }
            if (order.Count == 0)
            {
                Ui.Info("No data rows were found under the header.");
                return;
            }
            int send = Ui.Pick(title, order.Count + " emails are ready to be made. What should happen to them?", new[] { "Save them in Outlook Drafts  (check them, then send)", "Send them now" }, null);
            if (send < 0) return;
            if (send == 1 && !Ui.Ok("Send " + order.Count + " emails now?\nParties without an email address are skipped.\n\nSent emails cannot be taken back.")) return;
            if (send == 0 && order.Count > 300 && !Ui.Ok(order.Count + " drafts will be made. Continue?")) return;

            bool withFiles = how != 0;
            string folder = Path.Combine(Path.GetTempPath(), "DocBrisk_Mail", DateTime.Now.ToString("yyyyMMdd_HHmmss", Inv));
            if (withFiles) Directory.CreateDirectory(folder);
            dynamic ol = Activator.CreateInstance(outlook);
            List<object[]> log = new List<object[]>();
            int made = 0, noAddress = 0, failed = 0;
            using (new Speed())
            {
                int done = 0;
                foreach (string party in order)
                {
                    done++;
                    Xl.Status("email " + done + " of " + order.Count);
                    List<int> rws = groups[party];
                    string addr = "";
                    double sum = 0;
                    foreach (int r in rws)
                    {
                        double d;
                        if (tc != 0 && Logic.TryNum(v[r, tc], out d)) sum += d;
                        if (ec != 0 && addr.Length == 0)
                        {
                            string s = Xl.Str(v[r, ec]).Trim();
                            if (s.IndexOf('@') > 0) addr = s;
                        }
                    }
                    string totalText = tc != 0 ? Money(sum) : "";
                    string status;
                    try
                    {
                        if (addr.Length == 0) noAddress++;
                        if (send == 1 && addr.Length == 0) status = "Skipped: no email address";
                        else
                        {
                            dynamic mail = ol.CreateItem(0);
                            if (addr.Length > 0) mail.To = addr;
                            mail.Subject = Tokens(subject, party, totalText, rws.Count);
                            string html = "<p style=\"font-family:Segoe UI,Arial,sans-serif;font-size:14px\">" + Esc(Tokens(message, party, totalText, rws.Count)) + "</p>";
                            if (how == 0 || how == 3) html += HtmlTable(v, h, ci, rws, ec, tc, sum);
                            string old = "";
                            try
                            {
                                object inspector = mail.GetInspector;       // makes Outlook write the user's signature into the email
                                old = Convert.ToString((object)mail.HTMLBody, Inv);
                                if (inspector == null) old = "";
                            }
                            catch { old = ""; }
                            mail.HTMLBody = Merge(old, html);
                            if (withFiles)
                            {
                                List<string> files = MakeFiles(v, h, ci, rws, ec, folder, party, how);
                                foreach (string file in files) mail.Attachments.Add(file);
                            }
                            if (send == 1)
                            {
                                mail.Send();
                                status = "Sent";
                            }
                            else
                            {
                                mail.Save();
                                status = addr.Length > 0 ? "Saved in Drafts" : "Saved in Drafts (no address yet)";
                            }
                            made++;
                        }
                    }
                    catch (Exception ex)
                    {
                        failed++;
                        status = "Failed: " + ex.Message;
                    }
                    log.Add(new object[] { party, addr, (double)rws.Count, tc != 0 ? (object)sum : null, status });
                }
                dynamic ws = Xl.AddSheet(src.Worksheet.Parent, "Mail Log");
                object[,] t = AnalyseTools.Table(new[] { h[pc], "Email", "Rows", tc != 0 ? h[tc] : "Total", "Status" }, log);
                Xl.WriteTable(ws.Range["A1"], t, 1);
                Xl.StyleHeader(ws.Range["A1:E1"]);
                ws.Range["D2"].Resize[log.Count, 1].NumberFormat = "#,##0.00";
                ws.Columns["A:E"].AutoFit();
                ws.Activate();
            }
            Ui.Info(made + (send == 1 ? " emails were sent." : " emails are in the Drafts folder of Outlook. Check them, then send.") +
                    (noAddress > 0 ? "\n" + noAddress + " parties have no email address." : "") +
                    (failed > 0 ? "\n" + failed + " could not be made: see the Status column." : "") +
                    "\n\nThe list is on the new 'Mail Log' sheet." + (withFiles ? "\nThe attached files are also here:\n" + folder : ""));
        }
    }

    // A window with a text that can be read and copied (Explain Formula)
    internal sealed class TextForm : Form
    {
        public TextForm(string title, string text)
        {
            float z = 1f;
            try
            {
                using (Graphics g = Graphics.FromHwnd(IntPtr.Zero)) { z = g.DpiX / 96f; }
            }
            catch { z = 1f; }
            if (z < 1f) z = 1f;
            int w = (int)(900 * z), hgt = (int)(600 * z), pad = (int)(12 * z), bh = (int)(32 * z), bw = (int)(96 * z);
            try
            {
                Rectangle area = Screen.PrimaryScreen.WorkingArea;       // never larger than the screen
                if (w > area.Width - 40) w = area.Width - 40;
                if (hgt > area.Height - 60) hgt = area.Height - 60;
            }
            catch { }
            AutoScaleMode = AutoScaleMode.None;
            Text = title;
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(w, hgt);
            Font = new Font("Segoe UI", 9.5f);

            TextBox box = new TextBox();
            box.Multiline = true;
            box.ReadOnly = true;
            box.WordWrap = false;
            box.ScrollBars = ScrollBars.Both;
            box.BackColor = Color.White;
            box.Font = new Font("Consolas", 10f);
            box.SetBounds(pad, pad, w - 2 * pad, hgt - 3 * pad - bh);
            box.Text = text;
            Button copy = new Button();
            copy.Text = "Copy";
            copy.SetBounds(w - 2 * bw - 2 * pad, hgt - pad - bh, bw, bh);
            Button close = new Button();
            close.Text = "Close";
            close.SetBounds(w - bw - pad, hgt - pad - bh, bw, bh);
            close.DialogResult = DialogResult.Cancel;
            Controls.Add(box);
            Controls.Add(copy);
            Controls.Add(close);
            AcceptButton = close;
            CancelButton = close;
            copy.Click += delegate
            {
                try
                {
                    Clipboard.SetText(text);
                    copy.Text = "Copied";
                }
                catch { }
            };
            ActiveControl = close;
            box.Select(0, 0);
        }
    }
}
