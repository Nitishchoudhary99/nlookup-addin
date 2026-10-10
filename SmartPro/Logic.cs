using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace SmartPro
{
    // Pure logic: no Excel calls in this file, so everything here is fast and thread-safe.
    internal static class Logic
    {
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        static readonly Regex CurrencyWords = new Regex(@"rs\.?|inr", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        // ------------------------------------------------------------------
        //  numbers stored as text:  1,234.50   1,234.50-   (1,234)   Rs. 500   12%
        // ------------------------------------------------------------------
        public static bool TryNum(object x, out double d)
        {
            d = 0;
            if (x == null) return false;
            if (x is double) { d = (double)x; return true; }
            if (x is decimal) { d = (double)(decimal)x; return true; }
            if (x is float) { d = (float)x; return true; }
            string s = x as string;
            return s != null && TryNumText(s, out d);
        }

        public static bool TryNumText(string raw, out double d)
        {
            d = 0;
            string s = raw.Replace('\u00A0', ' ').Trim();
            if (s.Length == 0 || s.Length > 30) return false;
            bool neg = false, pct = false, letters = false;
            if (s[s.Length - 1] == '-') { neg = true; s = s.Substring(0, s.Length - 1); }
            if (s.Length >= 2 && s[0] == '(' && s[s.Length - 1] == ')') { neg = true; s = s.Substring(1, s.Length - 2); }
            if (s.Length > 0 && s[s.Length - 1] == '%') { pct = true; s = s.Substring(0, s.Length - 1); }
            s = s.Replace(",", "").Replace("\u20B9", "");
            foreach (char ch in s) { if (char.IsLetter(ch)) { letters = true; break; } }
            if (letters) s = CurrencyWords.Replace(s, "");
            s = s.Trim();
            if (s.Length == 0) return false;
            if (s[0] == '-') { neg = !neg; s = s.Substring(1); }
            else if (s[0] == '+') s = s.Substring(1);
            if (s.Length == 0) return false;
            int dots = 0, digits = 0;
            foreach (char ch in s)
            {
                if (ch == '.') dots++;
                else if (ch >= '0' && ch <= '9') digits++;
                else return false;
            }
            if (dots > 1 || digits == 0 || digits > 15) return false;
            if (s.Length > 1 && s[0] == '0' && s[1] != '.') return false;   // keep codes such as 00123 as text
            if (!double.TryParse(s, NumberStyles.AllowDecimalPoint, Inv, out d)) return false;
            if (pct) d /= 100.0;
            if (neg) d = -d;
            return true;
        }

        // ------------------------------------------------------------------
        //  dates stored as text: 01.04.2025  01-04-2025  01/04/25  2025-04-01  01-Apr-2025  Apr 1, 2025
        // ------------------------------------------------------------------
        public static bool TryDate(object x, out DateTime dt)
        {
            dt = DateTime.MinValue;
            if (x is DateTime) { dt = (DateTime)x; return true; }
            string s = x as string;
            return s != null && TryDateText(s, out dt);
        }

        static bool DPart(string s)
        {
            if (s.Length < 1 || s.Length > 4) return false;
            foreach (char ch in s) if (ch < '0' || ch > '9') return false;
            return true;
        }

        static int MonthNum(string s)
        {
            if (s.Length < 3) return 0;
            foreach (char ch in s) if (!char.IsLetter(ch)) return 0;
            int p = "janfebmaraprmayjunjulaugsepoctnovdec".IndexOf(s.Substring(0, 3).ToLowerInvariant(), StringComparison.Ordinal);
            return (p >= 0 && p % 3 == 0) ? p / 3 + 1 : 0;
        }

        public static bool TryDateText(string raw, out DateTime dt)
        {
            dt = DateTime.MinValue;
            string s = raw.Trim();
            if (s.IndexOf(':') >= 0)
            {
                int sp = s.IndexOf(' ');
                if (sp > 0) s = s.Substring(0, sp);
            }
            if (s.Length < 6 || s.Length > 18) return false;
            s = s.Replace(",", "").Replace('.', '-').Replace('/', '-').Replace(' ', '-');
            while (s.Contains("--")) s = s.Replace("--", "-");
            string[] p = s.Split('-');
            if (p.Length != 3) return false;
            int yy, mm, dd, yLen;
            bool d0 = DPart(p[0]), d1 = DPart(p[1]), d2 = DPart(p[2]);
            if (d0 && d1 && d2)
            {
                if (p[0].Length == 4)
                {
                    yy = int.Parse(p[0], Inv); mm = int.Parse(p[1], Inv); dd = int.Parse(p[2], Inv); yLen = 4;
                }
                else
                {
                    dd = int.Parse(p[0], Inv); mm = int.Parse(p[1], Inv); yy = int.Parse(p[2], Inv); yLen = p[2].Length;
                    if (mm > 12 && dd <= 12) { int t = dd; dd = mm; mm = t; }
                }
            }
            else if (d0 && d2) { dd = int.Parse(p[0], Inv); mm = MonthNum(p[1]); yy = int.Parse(p[2], Inv); yLen = p[2].Length; }
            else if (d1 && d2) { mm = MonthNum(p[0]); dd = int.Parse(p[1], Inv); yy = int.Parse(p[2], Inv); yLen = p[2].Length; }
            else return false;
            if (yLen == 2) yy += yy < 50 ? 2000 : 1900;
            else if (yLen != 4) return false;
            if (yy < 1900 || yy > 2200 || mm < 1 || mm > 12 || dd < 1) return false;
            if (dd > DateTime.DaysInMonth(yy, mm)) return false;
            dt = new DateTime(yy, mm, dd);
            return true;
        }

        public static string CleanText(string s)
        {
            s = s.Replace('\u00A0', ' ').Replace('\t', ' ').Replace("\u200B", "").Replace("\uFEFF", "").Trim();
            while (s.Contains("  ")) s = s.Replace("  ", " ");
            return s;
        }

        // Text that Excel would silently turn into a number, date or formula when written to a General cell
        public static bool RiskyText(string s)
        {
            if (string.IsNullOrEmpty(s)) return false;
            char c = s[0];
            if (c == '=' || c == '+' || c == '-' || c == '@') return true;
            if (s.Length > 40) return false;
            bool hasDigit = false;
            foreach (char ch in s) { if (ch >= '0' && ch <= '9') { hasDigit = true; break; } }
            if (!hasDigit)
            {
                string u = s.Trim().ToUpperInvariant();
                return u == "TRUE" || u == "FALSE";
            }
            double d;
            DateTime dt;
            if (double.TryParse(s, NumberStyles.Any, Inv, out d)) return true;
            if (DateTime.TryParse(s, CultureInfo.CurrentCulture, DateTimeStyles.None, out dt)) return true;
            return s.Trim().EndsWith("%", StringComparison.Ordinal);
        }

        // ------------------------------------------------------------------
        //  header words
        // ------------------------------------------------------------------
        public const string MeasureWords = "qty,quantity,amount,amt,value,sales,revenue,turnover,price, rate ,cost,total, net ,gross,weight, wt , mt ,tonn, kg ,profit,margin,stock,balance,units,pcs,volume,tax,discount,freight,salary,score,marks,hours, days ,income,expense,paid,outstanding, count ";
        public const string IdWords = " id , no , sr , sno ,code,number, pin ,pincode,phone,mobile,invoice,order, po ,serial,year,month, week ,plant,hsn,gstin, pan , zip ,account,batch, lot , doc ,voucher, bill ,challan, ref ";

        public static string PadName(string s)
        {
            StringBuilder sb = new StringBuilder(" ");
            foreach (char raw in s)
            {
                char ch = char.ToLowerInvariant(raw);
                if ((ch >= 'a' && ch <= 'z') || (ch >= '0' && ch <= '9')) sb.Append(ch);
                else if (ch == '#') sb.Append(" no ");
                else sb.Append(' ');
            }
            sb.Append(' ');
            return sb.ToString();
        }

        public static bool HasWord(string paddedName, string list)
        {
            foreach (string w in list.Split(','))
            {
                if (w.Length > 0 && paddedName.IndexOf(w, StringComparison.OrdinalIgnoreCase) >= 0) return true;
            }
            return false;
        }

        public static int MeasureRank(string hdr)
        {
            string nm = PadName(hdr);
            if (HasWord(nm, "amount,amt,value,sales,revenue,turnover,total, net ,gross,income")) return 1;
            if (HasWord(nm, "qty,quantity,weight, wt , mt ,tonn, kg ,units,volume,stock,pcs")) return 2;
            return 3;
        }

        public static bool MeasureIsAvg(string hdr)
        {
            return hdr.IndexOf('%') >= 0 || HasWord(PadName(hdr), " rate ,price,percent, pct , avg ,average,ratio, days , age ,score");
        }

        // ------------------------------------------------------------------
        //  India: fiscal year, lakh / crore, rupees in words, GSTIN, PAN
        // ------------------------------------------------------------------
        public static string FYLabel(DateTime dt, int startMonth)
        {
            int y = dt.Year;
            if (dt.Month < startMonth) y--;
            if (startMonth == 1) return "FY " + y;
            return "FY " + y + "-" + ((y + 1) % 100).ToString("00", Inv);
        }

        public static string FYQuarter(DateTime dt, int startMonth)
        {
            return "Q" + (((dt.Month - startMonth + 12) % 12) / 3 + 1);
        }

        public static string IndianGroup(double n, int decimals)
        {
            if (decimals < 0) decimals = 0;
            if (decimals > 6) decimals = 6;
            bool neg = n < 0;
            n = Math.Abs(n);
            string s = n.ToString("F" + decimals, Inv);
            string ip = s, fp = "";
            int dot = s.IndexOf('.');
            if (dot >= 0) { ip = s.Substring(0, dot); fp = s.Substring(dot); }
            if (ip.Length > 3)
            {
                string o = ip.Substring(ip.Length - 3);
                ip = ip.Substring(0, ip.Length - 3);
                while (ip.Length > 2)
                {
                    o = ip.Substring(ip.Length - 2) + "," + o;
                    ip = ip.Substring(0, ip.Length - 2);
                }
                ip = ip + "," + o;
            }
            return (neg ? "-" : "") + ip + fp;
        }

        public static string ShortINR(double n)
        {
            double a = Math.Abs(n);
            if (a >= 1e7) return IndianGroup(n / 1e7, 2) + " Cr";
            if (a >= 1e5) return IndianGroup(n / 1e5, 2) + " L";
            return IndianGroup(n, n == Math.Floor(n) ? 0 : 2);
        }

        static readonly string[] Ones = { "", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve",
                                          "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen" };
        static readonly string[] Tens = { "", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety" };

        static string TwoDigits(int n)
        {
            if (n < 20) return Ones[n];
            if (n % 10 == 0) return Tens[n / 10];
            return Tens[n / 10] + " " + Ones[n % 10];
        }

        static string WordsIndian(double n)
        {
            double cr = Math.Floor(n / 1e7);
            double rest = n - cr * 1e7;
            int lk = (int)Math.Floor(rest / 1e5);
            rest -= lk * 1e5;
            int th = (int)Math.Floor(rest / 1e3);
            rest -= th * 1e3;
            int hu = (int)Math.Floor(rest / 100.0);
            int tn = (int)Math.Round(rest - hu * 100.0);
            StringBuilder sb = new StringBuilder();
            if (cr > 0) sb.Append(WordsIndian(cr)).Append(" Crore ");
            if (lk > 0) sb.Append(TwoDigits(lk)).Append(" Lakh ");
            if (th > 0) sb.Append(TwoDigits(th)).Append(" Thousand ");
            if (hu > 0) sb.Append(TwoDigits(hu)).Append(" Hundred ");
            if (tn > 0) sb.Append(TwoDigits(tn));
            return sb.ToString().Trim();
        }

        // 125000.5 -> Rupees One Lakh Twenty Five Thousand and Fifty Paise Only   (null when the number is too large)
        public static string RupeeWords(double amount)
        {
            double n = Math.Abs(amount);
            if (n >= 1e15) return null;
            double whole = Math.Floor(n);
            int p = (int)Math.Round((n - whole) * 100.0, MidpointRounding.AwayFromZero);
            if (p >= 100) { whole += 1; p = 0; }
            string s = "Rupees " + (whole == 0 ? "Zero" : WordsIndian(whole));
            if (p > 0) s += " and " + TwoDigits(p) + " Paise";
            if (amount < 0) s = "Minus " + s;
            return s + " Only";
        }

        static readonly Dictionary<string, string> States = new Dictionary<string, string>
        {
            {"01","Jammu and Kashmir"},{"02","Himachal Pradesh"},{"03","Punjab"},{"04","Chandigarh"},{"05","Uttarakhand"},{"06","Haryana"},
            {"07","Delhi"},{"08","Rajasthan"},{"09","Uttar Pradesh"},{"10","Bihar"},{"11","Sikkim"},{"12","Arunachal Pradesh"},
            {"13","Nagaland"},{"14","Manipur"},{"15","Mizoram"},{"16","Tripura"},{"17","Meghalaya"},{"18","Assam"},
            {"19","West Bengal"},{"20","Jharkhand"},{"21","Odisha"},{"22","Chhattisgarh"},{"23","Madhya Pradesh"},{"24","Gujarat"},
            {"25","Daman and Diu"},{"26","Dadra and Nagar Haveli and Daman and Diu"},{"27","Maharashtra"},{"28","Andhra Pradesh (old)"},
            {"29","Karnataka"},{"30","Goa"},{"31","Lakshadweep"},{"32","Kerala"},{"33","Tamil Nadu"},{"34","Puducherry"},
            {"35","Andaman and Nicobar Islands"},{"36","Telangana"},{"37","Andhra Pradesh"},{"38","Ladakh"},
            {"97","Other Territory"},{"99","Centre Jurisdiction"}
        };

        public static string GstState(string code)
        {
            string name;
            return States.TryGetValue(code, out name) ? name : null;
        }

        const string CS = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        static readonly Regex GstRx = new Regex("^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][0-9A-Z]{3}$", RegexOptions.Compiled);
        static readonly Regex PanRx = new Regex("^[A-Z]{5}[0-9]{4}[A-Z]$", RegexOptions.Compiled);

        // format + state code + check digit (expects upper case, trimmed)
        public static bool GstinOk(string s)
        {
            if (s == null || s.Length != 15 || !GstRx.IsMatch(s)) return false;
            if (GstState(s.Substring(0, 2)) == null) return false;
            int total = 0;
            for (int i = 0; i < 14; i++)
            {
                int x = CS.IndexOf(s[i]);
                if ((i + 1) % 2 == 0) x *= 2;
                total += x / 36 + x % 36;
            }
            int p = (36 - total % 36) % 36;
            return CS[p] == s[14];
        }

        public static bool PanOk(string s)
        {
            return s != null && s.Length == 10 && PanRx.IsMatch(s) && "PCHFATBLJG".IndexOf(s[3]) >= 0;
        }

        // ------------------------------------------------------------------
        //  fuzzy matching
        // ------------------------------------------------------------------
        static readonly HashSet<string> NameNoise = new HashSet<string>
        {
            "private", "pvt", "limited", "ltd", "llp", "inc", "co", "company", "corporation", "corp", "and", "the", "messrs"
        };

        // name without case, punctuation and company suffixes:  "M/s. Tata Steel Pvt. Ltd." -> "tata steel"
        public static string NormName(string s)
        {
            StringBuilder sb = new StringBuilder(s.Length + 8);
            foreach (char raw in s)
            {
                char ch = char.ToLowerInvariant(raw);
                if ((ch >= 'a' && ch <= 'z') || (ch >= '0' && ch <= '9')) sb.Append(ch);
                else if (ch == '&') sb.Append(" and ");
                else sb.Append(' ');
            }
            string[] parts = sb.ToString().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            StringBuilder o = new StringBuilder();
            int start = (parts.Length > 2 && parts[0] == "m" && parts[1] == "s") ? 2 : 0;
            for (int i = start; i < parts.Length; i++)
            {
                if (NameNoise.Contains(parts[i])) continue;
                if (o.Length > 0) o.Append(' ');
                o.Append(parts[i]);
            }
            return o.ToString();
        }

        // Jaro-Winkler similarity, 0 to 1. No allocations, so it is safe and fast in parallel loops.
        public static double JaroW(string s1, string s2)
        {
            if (s1.Length > 64) s1 = s1.Substring(0, 64);
            if (s2.Length > 64) s2 = s2.Substring(0, 64);
            int l1 = s1.Length, l2 = s2.Length;
            if (l1 == 0 || l2 == 0) return 0;
            if (string.Equals(s1, s2, StringComparison.Ordinal)) return 1;
            int win = Math.Max(l1, l2) / 2 - 1;
            if (win < 0) win = 0;
            ulong f1 = 0, f2 = 0;
            int m = 0;
            for (int i = 0; i < l1; i++)
            {
                int lo = i - win;
                if (lo < 0) lo = 0;
                int hi = i + win;
                if (hi > l2 - 1) hi = l2 - 1;
                char ch = s1[i];
                for (int j = lo; j <= hi; j++)
                {
                    if ((f2 & (1UL << j)) == 0 && s2[j] == ch)
                    {
                        f1 |= 1UL << i;
                        f2 |= 1UL << j;
                        m++;
                        break;
                    }
                }
            }
            if (m == 0) return 0;
            int t = 0, k = 0;
            for (int i = 0; i < l1; i++)
            {
                if ((f1 & (1UL << i)) != 0)
                {
                    while ((f2 & (1UL << k)) == 0) k++;
                    if (s1[i] != s2[k]) t++;
                    k++;
                }
            }
            double jr = ((double)m / l1 + (double)m / l2 + (m - t / 2.0) / m) / 3.0;
            int pre = 0;
            for (int i = 0; i < 4 && i < l1 && i < l2; i++)
            {
                if (s1[i] == s2[i]) pre++;
                else break;
            }
            return jr + pre * 0.1 * (1 - jr);
        }

        // quartile with linear interpolation (same as Excel QUARTILE.INC); the array must be sorted
        public static double Percentile(double[] sorted, double p)
        {
            int n = sorted.Length;
            if (n == 1) return sorted[0];
            double pos = p * (n - 1);
            int lo = (int)Math.Floor(pos);
            int hi = Math.Min(lo + 1, n - 1);
            return sorted[lo] + (sorted[hi] - sorted[lo]) * (pos - lo);
        }
    }

    // ----------------------------------------------------------------------
    //  Column profiling: what is in each column of a table
    // ----------------------------------------------------------------------
    internal sealed class ColInfo
    {
        public string Hdr = "";
        public string Kind = "empty";      // date / num / cat / id / text / empty
        public int NonBlank, Blank, Distinct, NumCount, DateCount, TextNum, TextDate, Spaces, Errors;
        public bool AllInt = true;
        public double MinV, MaxV, SumV;

        public string KindLabel
        {
            get
            {
                switch (Kind)
                {
                    case "date": return "Date";
                    case "num": return "Number (measure)";
                    case "cat": return "Category";
                    case "id": return "ID / code";
                    case "empty": return "Empty";
                    default: return "Text";
                }
            }
        }
    }

    internal static class Profiler
    {
        // the first row (within the top 10) that is mostly filled and has data under it
        public static int FindHeaderRow(object[,] v)
        {
            int nr = v.GetLength(0), nc = v.GetLength(1);
            int lim = Math.Min(nr - 1, 10);
            for (int r = 1; r <= lim; r++)
            {
                int filled = 0, nextFilled = 0;
                for (int c = 1; c <= nc; c++)
                {
                    if (!Xl.IsBlank(v[r, c])) filled++;
                    if (!Xl.IsBlank(v[r + 1, c])) nextFilled++;
                }
                if (filled >= 0.6 * nc && nextFilled >= 0.5 * nc) return r;
            }
            return 1;
        }

        // clean, unique header names; index 1..nc (index 0 is not used)
        public static string[] MakeHeaders(object[,] v, int hr)
        {
            int nc = v.GetLength(1);
            string[] h = new string[nc + 1];
            h[0] = "";
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int c = 1; c <= nc; c++)
            {
                string s = Xl.IsErr(v[hr, c]) ? "" : Xl.Str(v[hr, c]);
                s = Logic.CleanText(s.Replace('\r', ' ').Replace('\n', ' ').Replace('"', '\''));
                if (s.Length == 0) s = "Column " + c;
                if (s.Length > 60) s = s.Substring(0, 60);
                string t = s;
                int k = 1;
                while (!seen.Add(t)) { k++; t = s + " " + k; }
                h[c] = t;
            }
            return h;
        }

        public static string FreeName(string[] h, string want)
        {
            bool clash = true;
            while (clash)
            {
                clash = false;
                for (int i = 1; i < h.Length; i++)
                {
                    if (string.Equals(h[i], want, StringComparison.OrdinalIgnoreCase)) clash = true;
                }
                if (clash) want += " (auto)";
            }
            return want;
        }

        // maxRows = 0 reads every row; otherwise only the first maxRows data rows are profiled
        public static ColInfo[] Profile(object[,] v, int hr, string[] h, int maxRows)
        {
            int nr = v.GetLength(0), nc = v.GetLength(1);
            if (maxRows > 0 && nr > hr + maxRows) nr = hr + maxRows;
            ColInfo[] ci = new ColInfo[nc + 1];
            ci[0] = new ColInfo();
            for (int c = 1; c <= nc; c++)
            {
                ColInfo k = new ColInfo();
                k.Hdr = h[c];
                HashSet<string> distinct = new HashSet<string>(StringComparer.Ordinal);
                bool first = true;
                for (int r = hr + 1; r <= nr; r++)
                {
                    object x = v[r, c];
                    if (Xl.IsErr(x)) { k.Errors++; k.NonBlank++; continue; }
                    if (Xl.IsBlank(x)) { k.Blank++; continue; }
                    k.NonBlank++;
                    string s = x as string;
                    distinct.Add(s ?? Xl.Str(x));
                    if (s != null && (s.Length > 0 && (s[0] == ' ' || s[s.Length - 1] == ' '))) k.Spaces++;
                    double d;
                    DateTime dt;
                    if (x is DateTime) k.DateCount++;
                    else if (Logic.TryNum(x, out d))
                    {
                        k.NumCount++;
                        if (s != null) k.TextNum++;
                        if (d != Math.Floor(d)) k.AllInt = false;
                        if (first) { k.MinV = d; k.MaxV = d; first = false; }
                        else { if (d < k.MinV) k.MinV = d; if (d > k.MaxV) k.MaxV = d; }
                        k.SumV += d;
                    }
                    else if (s != null && Logic.TryDateText(s, out dt)) { k.DateCount++; k.TextDate++; }
                }
                k.Distinct = distinct.Count;
                k.Kind = Classify(k);
                ci[c] = k;
            }
            return ci;
        }

        static string Classify(ColInfo k)
        {
            int n = k.NonBlank;
            string nm = Logic.PadName(k.Hdr);
            if (n == 0) return "empty";
            if (k.DateCount >= 0.8 * n) return "date";
            if (k.NumCount >= 0.8 * n)
            {
                if (Logic.HasWord(nm, Logic.MeasureWords)) return "num";
                if (Logic.HasWord(nm, Logic.IdWords)) return k.Distinct <= 50 ? "cat" : "id";
                if (k.AllInt && n >= 20 && k.Distinct >= 0.9 * n) return "id";
                return "num";
            }
            if (k.Distinct <= 500 && (k.Distinct <= 0.5 * n || n <= 20)) return "cat";
            return "text";
        }

        // false for empty rows and for total / subtotal rows
        public static bool RowIsData(object[,] v, int r, int nc)
        {
            int filled = 0;
            for (int c = 1; c <= nc; c++)
            {
                object x = v[r, c];
                if (Xl.IsBlank(x)) continue;
                filled++;
                string s = x as string;
                if (s == null) continue;
                s = s.Trim().ToLowerInvariant();
                if (s == "total" || s == "grand total" || s == "subtotal" || s == "sub total") return false;
                if (s.StartsWith("total ", StringComparison.Ordinal) || s.EndsWith(" total", StringComparison.Ordinal)) return false;
            }
            return filled > 0;
        }
    }
}
