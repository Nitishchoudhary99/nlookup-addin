using System;
using System.Globalization;
using System.Text;
using ExcelDna.Integration;

namespace SmartPro
{
    // Worksheet functions. They appear in Insert Function under the "DocBrisk" category.
    public static class Functions
    {
        const string Cat = Design.FunctionCategory;
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;

        // ---- argument helpers (Excel sends numbers and dates as double) ----
        static bool ArgNum(object x, out double d)
        {
            d = 0;
            if (x is double)
            {
                d = (double)x;
                return true;
            }
            string s = x as string;
            return s != null && Logic.TryNumText(s, out d);
        }

        static string ArgStr(object x)
        {
            if (x == null || x is ExcelMissing || x is ExcelEmpty || x is ExcelError) return "";
            string s = x as string;
            if (s != null) return s;
            if (x is double) return ((double)x).ToString(Inv);
            if (x is bool) return (bool)x ? "TRUE" : "FALSE";
            return Convert.ToString(x, Inv);
        }

        static int ArgInt(object x, int def)
        {
            double d;
            if (x is ExcelMissing || x is ExcelEmpty || !ArgNum(x, out d)) return def;
            return (int)Math.Round(d);
        }

        static bool ArgDate(object x, out DateTime dt)
        {
            dt = DateTime.MinValue;
            if (x is double)
            {
                double d = (double)x;
                if (d <= 0 || d >= 2958466) return false;
                dt = DateTime.FromOADate(d);
                return true;
            }
            string s = x as string;
            return s != null && Logic.TryDateText(s, out dt);
        }

        // Excel-DNA passes ranges as 0-based arrays; a single cell arrives as a single value
        static object[,] ArgArray(object x)
        {
            object[,] arr = x as object[,];
            if (arr != null) return arr;
            return new object[,] { { x } };
        }

        // ------------------------------------------------------------------
        [ExcelFunction(Name = "RUPEEWORDS", Category = Cat, Description = "Amount in words, Indian style: Rupees One Lakh Twenty Five Thousand Only")]
        public static object RupeeWords([ExcelArgument(Name = "amount", Description = "the amount")] object amount)
        {
            double d;
            if (!ArgNum(amount, out d) || Math.Abs(d) >= 1e15) return ExcelError.ExcelErrorValue;
            return Logic.RupeeWords(d);
        }

        [ExcelFunction(Name = "INRFORMAT", Category = Cat, Description = "Number as text with lakh / crore commas: 12,34,567.50")]
        public static object InrFormat([ExcelArgument(Name = "amount", Description = "the number")] object amount,
                                       [ExcelArgument(Name = "decimals", Description = "decimal places, 2 if left out")] object decimals)
        {
            double d;
            if (!ArgNum(amount, out d)) return ExcelError.ExcelErrorValue;
            return Logic.IndianGroup(d, ArgInt(decimals, 2));
        }

        [ExcelFunction(Name = "INRSHORT", Category = Cat, Description = "Short amount in K / L / Cr: 1.25 Cr")]
        public static object InrShort([ExcelArgument(Name = "amount", Description = "the number")] object amount,
                                      [ExcelArgument(Name = "decimals", Description = "decimal places, 2 if left out")] object decimals)
        {
            double d;
            if (!ArgNum(amount, out d)) return ExcelError.ExcelErrorValue;
            int dec = ArgInt(decimals, 2);
            double a = Math.Abs(d);
            if (a >= 10000000) return Logic.IndianGroup(d / 10000000, dec) + " Cr";
            if (a >= 100000) return Logic.IndianGroup(d / 100000, dec) + " L";
            if (a >= 1000) return Logic.IndianGroup(d / 1000, dec) + " K";
            return Logic.IndianGroup(d, dec);
        }

        [ExcelFunction(Name = "FYEAR", Category = Cat, Description = "Fiscal year of a date: FY 2026-27")]
        public static object FYear([ExcelArgument(Name = "date", Description = "the date")] object date,
                                   [ExcelArgument(Name = "start_month", Description = "first month of the year, 4 (April) if left out")] object startMonth)
        {
            DateTime dt;
            int sm = ArgInt(startMonth, 4);
            if (sm < 1 || sm > 12 || !ArgDate(date, out dt)) return ExcelError.ExcelErrorValue;
            return Logic.FYLabel(dt, sm);
        }

        [ExcelFunction(Name = "FYQUARTER", Category = Cat, Description = "Fiscal quarter of a date: Q3")]
        public static object FYQuarter([ExcelArgument(Name = "date", Description = "the date")] object date,
                                       [ExcelArgument(Name = "start_month", Description = "first month of the year, 4 (April) if left out")] object startMonth)
        {
            DateTime dt;
            int sm = ArgInt(startMonth, 4);
            if (sm < 1 || sm > 12 || !ArgDate(date, out dt)) return ExcelError.ExcelErrorValue;
            return Logic.FYQuarter(dt, sm);
        }

        [ExcelFunction(Name = "GSTINVALID", Category = Cat, Description = "TRUE when the GSTIN format, state code and check digit are right")]
        public static object GstinValid([ExcelArgument(Name = "gstin", Description = "the GSTIN")] object gstin)
        {
            return Logic.GstinOk(ArgStr(gstin).Trim().ToUpperInvariant());
        }

        [ExcelFunction(Name = "GSTSTATE", Category = Cat, Description = "State name from a GSTIN or a 2-digit state code")]
        public static object GstState([ExcelArgument(Name = "gstin_or_code", Description = "a GSTIN or a state code such as 06")] object gstin)
        {
            string s = ArgStr(gstin).Trim();
            if (s.Length == 1) s = "0" + s;
            if (s.Length < 2) return ExcelError.ExcelErrorNA;
            string name = Logic.GstState(s.Substring(0, 2));
            if (name == null) return ExcelError.ExcelErrorNA;
            return name;
        }

        [ExcelFunction(Name = "PANVALID", Category = Cat, Description = "TRUE when the PAN format is right")]
        public static object PanValid([ExcelArgument(Name = "pan", Description = "the PAN")] object pan)
        {
            return Logic.PanOk(ArgStr(pan).Trim().ToUpperInvariant());
        }

        [ExcelFunction(Name = "SIMILARITY", Category = Cat, Description = "How alike two texts are, from 0 to 1")]
        public static object Similarity([ExcelArgument(Name = "text1", Description = "first text")] object text1,
                                        [ExcelArgument(Name = "text2", Description = "second text")] object text2)
        {
            string a = ArgStr(text1).Trim().ToLowerInvariant(), b = ArgStr(text2).Trim().ToLowerInvariant();
            if (a.Length == 0 && b.Length == 0) return 1.0;
            if (a == b) return 1.0;
            return Math.Round(Logic.JaroW(a, b), 4);
        }

        [ExcelFunction(Name = "FUZZYLOOKUP", Category = Cat, Description = "Lookup that tolerates typos and Pvt/Ltd differences. Returns the closest match.")]
        public static object FuzzyLookup([ExcelArgument(Name = "lookup_value", Description = "the text to find")] object lookupValue,
                                         [ExcelArgument(Name = "lookup_range", Description = "where to look, for example Master!A2:A5000")] object lookupRange,
                                         [ExcelArgument(Name = "return_range", Description = "same size as lookup_range; the matched text is returned if left out")] object returnRange,
                                         [ExcelArgument(Name = "min_score", Description = "0 to 1, 0.85 if left out")] object minScore)
        {
            string s = Logic.NormName(ArgStr(lookupValue));
            if (s.Length == 0) return ExcelError.ExcelErrorNA;
            object[,] look = ArgArray(lookupRange);
            double min = 0.85, m;
            if (!(minScore is ExcelMissing) && !(minScore is ExcelEmpty) && ArgNum(minScore, out m)) min = m;
            double best = 0;
            int bi = -1, bj = -1;
            int nr = look.GetLength(0), nc = look.GetLength(1);
            for (int i = 0; i < nr; i++)
            {
                for (int j = 0; j < nc; j++)
                {
                    string t = look[i, j] as string;
                    if (t == null) continue;
                    t = Logic.NormName(t);
                    if (t.Length == 0) continue;
                    double sc = t == s ? 1.0 : Logic.JaroW(s, t);
                    if (sc > best)
                    {
                        best = sc;
                        bi = i;
                        bj = j;
                    }
                }
            }
            if (bi < 0 || best < min) return ExcelError.ExcelErrorNA;
            if (returnRange is ExcelMissing || returnRange is ExcelEmpty) return look[bi, bj];
            object[,] ret = ArgArray(returnRange);
            if (bi >= ret.GetLength(0) || bj >= ret.GetLength(1)) return ExcelError.ExcelErrorRef;
            object res = ret[bi, bj];
            if (res == null || res is ExcelEmpty) return "";
            return res;
        }

        [ExcelFunction(Name = "EXTRACTNUM", Category = Cat, Description = "The first (or nth) number inside a text")]
        public static object ExtractNum([ExcelArgument(Name = "text", Description = "the text")] object text,
                                        [ExcelArgument(Name = "instance", Description = "which number to return, 1 if left out")] object instance)
        {
            string s = ArgStr(text) + " ";
            int want = ArgInt(instance, 1), k = 0;
            StringBuilder cur = new StringBuilder();
            bool dot = false;
            for (int i = 0; i < s.Length; i++)
            {
                char ch = s[i];
                bool nextDigit = i + 1 < s.Length && s[i + 1] >= '0' && s[i + 1] <= '9';
                if (ch >= '0' && ch <= '9') cur.Append(ch);
                else if (ch == '.' && cur.Length > 0 && !dot && nextDigit)
                {
                    cur.Append(ch);
                    dot = true;
                }
                else if (ch == ',' && cur.Length > 0 && nextDigit)
                {
                    // a comma inside a number: skip it
                }
                else if (cur.Length > 0)
                {
                    k++;
                    if (k == want) return double.Parse(cur.ToString(), Inv);
                    cur.Length = 0;
                    dot = false;
                }
            }
            return ExcelError.ExcelErrorNA;
        }

        [ExcelFunction(Name = "TEXTBETWEEN", Category = Cat, Description = "Text between two markers")]
        public static object TextBetween([ExcelArgument(Name = "text", Description = "the text")] object text,
                                         [ExcelArgument(Name = "start_text", Description = "marker before the part you want")] object startText,
                                         [ExcelArgument(Name = "end_text", Description = "marker after the part you want")] object endText)
        {
            string s = ArgStr(text), a = ArgStr(startText), b = ArgStr(endText);
            int p1 = 0;
            if (a.Length > 0)
            {
                p1 = s.IndexOf(a, StringComparison.OrdinalIgnoreCase);
                if (p1 < 0) return ExcelError.ExcelErrorNA;
                p1 += a.Length;
            }
            int p2 = s.Length;
            if (b.Length > 0)
            {
                p2 = s.IndexOf(b, p1, StringComparison.OrdinalIgnoreCase);
                if (p2 < 0) return ExcelError.ExcelErrorNA;
            }
            return s.Substring(p1, p2 - p1);
        }

        // ==================================================================
        //  More functions: GST maths, Indian ID checks, text helpers, ageing
        // ==================================================================
        static double RateOf(double r) { return Math.Abs(r) > 1 ? r / 100.0 : r; }   // 18 and 0.18 both mean 18%

        [ExcelFunction(Name = "NUMWORDS", Category = Cat, Description = "Whole number in words, Indian style: One Lakh Twenty Five Thousand")]
        public static object NumWords([ExcelArgument(Name = "number", Description = "the number")] object number)
        {
            double d;
            if (!ArgNum(number, out d) || Math.Abs(d) >= 1e15) return ExcelError.ExcelErrorValue;
            string s = Logic.RupeeWords(Math.Round(Math.Abs(d)));
            if (s.StartsWith("Rupees ", StringComparison.Ordinal)) s = s.Substring(7);
            if (s.EndsWith(" Only", StringComparison.Ordinal)) s = s.Substring(0, s.Length - 5);
            return d < 0 ? "Minus " + s : s;
        }

        [ExcelFunction(Name = "GSTADD", Category = Cat, Description = "Amount with GST added: =GSTADD(1000,18) gives 1180")]
        public static object GstAdd([ExcelArgument(Name = "amount", Description = "amount before GST")] object amount,
                                    [ExcelArgument(Name = "rate", Description = "GST rate, 18 or 18%")] object rate)
        {
            double a, r;
            if (!ArgNum(amount, out a) || !ArgNum(rate, out r)) return ExcelError.ExcelErrorValue;
            return Math.Round(a * (1 + RateOf(r)), 2);
        }

        [ExcelFunction(Name = "GSTBASE", Category = Cat, Description = "Amount before GST from a GST-inclusive amount: =GSTBASE(1180,18) gives 1000")]
        public static object GstBase([ExcelArgument(Name = "amount_incl", Description = "amount including GST")] object amount,
                                     [ExcelArgument(Name = "rate", Description = "GST rate, 18 or 18%")] object rate)
        {
            double a, r;
            if (!ArgNum(amount, out a) || !ArgNum(rate, out r) || RateOf(r) <= -1) return ExcelError.ExcelErrorValue;
            return Math.Round(a / (1 + RateOf(r)), 2);
        }

        [ExcelFunction(Name = "GSTTAX", Category = Cat, Description = "The GST part of an amount. Third argument TRUE when the amount already includes GST.")]
        public static object GstTax([ExcelArgument(Name = "amount", Description = "the amount")] object amount,
                                    [ExcelArgument(Name = "rate", Description = "GST rate, 18 or 18%")] object rate,
                                    [ExcelArgument(Name = "inclusive", Description = "TRUE if the amount includes GST; FALSE if left out")] object inclusive)
        {
            double a, r;
            if (!ArgNum(amount, out a) || !ArgNum(rate, out r) || RateOf(r) <= -1) return ExcelError.ExcelErrorValue;
            bool incl = inclusive is bool && (bool)inclusive;
            double rr = RateOf(r);
            return Math.Round(incl ? a - a / (1 + rr) : a * rr, 2);
        }

        // Verhoeff check digit (used by Aadhaar)
        static readonly int[,] VD =
        {
            {0,1,2,3,4,5,6,7,8,9}, {1,2,3,4,0,6,7,8,9,5}, {2,3,4,0,1,7,8,9,5,6}, {3,4,0,1,2,8,9,5,6,7}, {4,0,1,2,3,9,5,6,7,8},
            {5,9,8,7,6,0,4,3,2,1}, {6,5,9,8,7,1,0,4,3,2}, {7,6,5,9,8,2,1,0,4,3}, {8,7,6,5,9,3,2,1,0,4}, {9,8,7,6,5,4,3,2,1,0}
        };
        static readonly int[,] VP =
        {
            {0,1,2,3,4,5,6,7,8,9}, {1,5,7,6,2,8,3,0,9,4}, {5,8,0,3,7,9,6,1,4,2}, {8,9,1,6,0,4,3,5,2,7},
            {9,4,5,3,1,2,6,8,7,0}, {4,2,8,6,5,7,3,9,0,1}, {2,7,9,3,8,0,6,4,1,5}, {7,0,4,6,9,1,3,2,5,8}
        };

        static bool VerhoeffOk(string digits)
        {
            int c = 0;
            for (int i = 0; i < digits.Length; i++)
            {
                int dgt = digits[digits.Length - 1 - i] - '0';
                c = VD[c, VP[i % 8, dgt]];
            }
            return c == 0;
        }

        static string DigitsOf(string s)
        {
            StringBuilder sb = new StringBuilder();
            foreach (char ch in s) if (ch >= '0' && ch <= '9') sb.Append(ch);
            return sb.ToString();
        }

        [ExcelFunction(Name = "AADHAARVALID", Category = Cat, Description = "TRUE when a 12-digit Aadhaar number has the right format and check digit")]
        public static object AadhaarValid([ExcelArgument(Name = "aadhaar", Description = "the number, spaces allowed")] object aadhaar)
        {
            string raw = ArgStr(aadhaar).Trim();
            foreach (char ch in raw) if (!(ch >= '0' && ch <= '9') && ch != ' ' && ch != '-') return false;
            string s = DigitsOf(raw);
            if (s.Length != 12 || s[0] < '2') return false;
            return VerhoeffOk(s);
        }

        [ExcelFunction(Name = "IFSCVALID", Category = Cat, Description = "TRUE when a bank IFSC code has the right format (4 letters, 0, 6 characters)")]
        public static object IfscValid([ExcelArgument(Name = "ifsc", Description = "the IFSC code")] object ifsc)
        {
            return System.Text.RegularExpressions.Regex.IsMatch(ArgStr(ifsc).Trim().ToUpperInvariant(), "^[A-Z]{4}0[A-Z0-9]{6}$");
        }

        [ExcelFunction(Name = "MOBILEVALID", Category = Cat, Description = "TRUE for a 10-digit Indian mobile number (+91, 91 or 0 in front is allowed)")]
        public static object MobileValid([ExcelArgument(Name = "mobile", Description = "the mobile number")] object mobile)
        {
            string raw = ArgStr(mobile).Trim();
            foreach (char ch in raw) if (!(ch >= '0' && ch <= '9') && ch != ' ' && ch != '-' && ch != '+') return false;
            string s = DigitsOf(raw);
            if (s.Length == 12 && s.StartsWith("91", StringComparison.Ordinal)) s = s.Substring(2);
            else if (s.Length == 11 && s[0] == '0') s = s.Substring(1);
            return s.Length == 10 && s[0] >= '6';
        }

        [ExcelFunction(Name = "PINVALID", Category = Cat, Description = "TRUE for a 6-digit Indian PIN code that does not start with 0")]
        public static object PinValid([ExcelArgument(Name = "pin", Description = "the PIN code")] object pin)
        {
            return System.Text.RegularExpressions.Regex.IsMatch(ArgStr(pin).Trim().Replace(" ", ""), "^[1-9][0-9]{5}$");
        }

        [ExcelFunction(Name = "EMAILVALID", Category = Cat, Description = "TRUE when the text looks like an email address")]
        public static object EmailValid([ExcelArgument(Name = "email", Description = "the email address")] object email)
        {
            return System.Text.RegularExpressions.Regex.IsMatch(ArgStr(email).Trim(), @"^[^\s@]+@[^\s@]+\.[^\s@]{2,}$");
        }

        static string[] WordsOf(string s)
        {
            return s.Split(new[] { ' ', '\t', '\r', '\n', '\u00A0' }, StringSplitOptions.RemoveEmptyEntries);
        }

        [ExcelFunction(Name = "NTHWORD", Category = Cat, Description = "The nth word of a text; a negative n counts from the end (-1 is the last word)")]
        public static object NthWord([ExcelArgument(Name = "text", Description = "the text")] object text,
                                     [ExcelArgument(Name = "n", Description = "which word; 1 if left out")] object n)
        {
            string[] w = WordsOf(ArgStr(text));
            int k = ArgInt(n, 1);
            if (k < 0) k = w.Length + k + 1;
            if (k < 1 || k > w.Length) return "";
            return w[k - 1];
        }

        [ExcelFunction(Name = "WORDCOUNT", Category = Cat, Description = "Number of words in a text")]
        public static object WordCount([ExcelArgument(Name = "text", Description = "the text")] object text)
        {
            return (double)WordsOf(ArgStr(text)).Length;
        }

        [ExcelFunction(Name = "ONLYDIGITS", Category = Cat, Description = "Keeps only the digits of a text, as text")]
        public static object OnlyDigits([ExcelArgument(Name = "text", Description = "the text")] object text)
        {
            return DigitsOf(ArgStr(text));
        }

        [ExcelFunction(Name = "ONLYLETTERS", Category = Cat, Description = "Keeps only letters and spaces of a text")]
        public static object OnlyLetters([ExcelArgument(Name = "text", Description = "the text")] object text)
        {
            StringBuilder sb = new StringBuilder();
            foreach (char ch in ArgStr(text)) if (char.IsLetter(ch) || ch == ' ') sb.Append(ch);
            return Logic.CleanText(sb.ToString());
        }

        [ExcelFunction(Name = "CLEANSPACES", Category = Cat, Description = "Removes leading, trailing, double and non-breaking spaces")]
        public static object CleanSpaces([ExcelArgument(Name = "text", Description = "the text")] object text)
        {
            return Logic.CleanText(ArgStr(text));
        }

        static readonly string[] KeepUpper = { "LLP", "GST", "HUF", "SEZ", "MSME", "PAN", "TDS", "USA", "UAE", "UK", "II", "III", "IV" };

        [ExcelFunction(Name = "PROPERNAME", Category = Cat, Description = "Tidy capitals for names: TATA STEEL PVT LTD becomes Tata Steel Pvt Ltd")]
        public static object ProperName([ExcelArgument(Name = "text", Description = "the name")] object text)
        {
            string[] w = WordsOf(Logic.CleanText(ArgStr(text)));
            for (int i = 0; i < w.Length; i++)
            {
                string up = w[i].ToUpperInvariant();
                string bare = up.Trim('.', ',', '(', ')');
                if (Array.IndexOf(KeepUpper, bare) >= 0) w[i] = up;
                else w[i] = Inv.TextInfo.ToTitleCase(w[i].ToLowerInvariant());
            }
            return string.Join(" ", w);
        }

        [ExcelFunction(Name = "AGEDAYS", Category = Cat, Description = "Days from a date to today (or to the second date). Recalculates when the sheet recalculates.")]
        public static object AgeDays([ExcelArgument(Name = "date", Description = "the start date")] object date,
                                     [ExcelArgument(Name = "as_of", Description = "end date; today if left out")] object asOf)
        {
            DateTime d, e = DateTime.Today;
            if (!ArgDate(date, out d)) return ExcelError.ExcelErrorValue;
            if (!(asOf is ExcelMissing) && !(asOf is ExcelEmpty) && !ArgDate(asOf, out e)) return ExcelError.ExcelErrorValue;
            return Math.Floor((e.Date - d.Date).TotalDays);
        }

        [ExcelFunction(Name = "AGEBUCKET", Category = Cat, Description = "Ageing bucket of a date: 0-30, 31-60, 61-90, 91-180, 180+ or Not due")]
        public static object AgeBucket([ExcelArgument(Name = "date", Description = "the invoice or due date")] object date,
                                       [ExcelArgument(Name = "as_of", Description = "end date; today if left out")] object asOf)
        {
            DateTime d, e = DateTime.Today;
            if (!ArgDate(date, out d)) return ExcelError.ExcelErrorValue;
            if (!(asOf is ExcelMissing) && !(asOf is ExcelEmpty) && !ArgDate(asOf, out e)) return ExcelError.ExcelErrorValue;
            double days = Math.Floor((e.Date - d.Date).TotalDays);
            return days < 0 ? "Not due" : days <= 30 ? "0-30" : days <= 60 ? "31-60" : days <= 90 ? "61-90" : days <= 180 ? "91-180" : "180+";
        }

        [ExcelFunction(Name = "MONTHKEY", Category = Cat, Description = "Month label that sorts correctly: 2026-04 (Apr)")]
        public static object MonthKey([ExcelArgument(Name = "date", Description = "the date")] object date)
        {
            DateTime d;
            if (!ArgDate(date, out d)) return ExcelError.ExcelErrorValue;
            return d.ToString("yyyy-MM", Inv) + " (" + d.ToString("MMM", Inv) + ")";
        }

        [ExcelFunction(Name = "PCTCHANGE", Category = Cat, Description = "Change from old to new as a fraction: =PCTCHANGE(100,125) gives 0.25")]
        public static object PctChange([ExcelArgument(Name = "old_value", Description = "the earlier number")] object oldValue,
                                       [ExcelArgument(Name = "new_value", Description = "the later number")] object newValue)
        {
            double a, b;
            if (!ArgNum(oldValue, out a) || !ArgNum(newValue, out b)) return ExcelError.ExcelErrorValue;
            if (a == 0) return ExcelError.ExcelErrorDiv0;
            return (b - a) / Math.Abs(a);
        }
    }
}
