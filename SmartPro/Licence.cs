using System;
using System.Collections.Generic;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Win32;

namespace SmartPro
{
    internal sealed class LicenceInfo
    {
        public bool Valid, Pro;
        public string Name = "", Reason = "";
        public DateTime Expires = DateTime.MaxValue;              // MaxValue = never expires
        public bool Never { get { return Expires == DateTime.MaxValue; } }
    }

    // Checks DocBrisk keys (format DBK1, the keys the DocBrisk website issues) offline with the PUBLIC key.
    // A Pro key (tier P) with an expiry date is the subscription: when it expires, Pro tools lock again.
    // The window the user sees is in ProCentre.cs; the prices and the words of the offer are in ProPlan.cs.
    internal static class Licence
    {
        // ---- settings -------------------------------------------------------------
        const string PubX = LicenceKeys.PubX;                     // your keys and the buy link are in LicenceKeys.cs
        const string PubY = LicenceKeys.PubY;
        public const string BuyUrl = LicenceKeys.BuyUrl;
        public const int FreeRuns = 3;                            // free runs of each Pro tool on a PC
        const int WarnDays = 7;                                   // remind this many days before expiry
        const string RegPath = @"Software\DocBrisk\SmartPro";
        // ---------------------------------------------------------------------------

        const string B32 = "0123456789ABCDEFGHJKMNPQRSTVWXYZ";
        static readonly DateTime Epoch = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        static readonly CultureInfo Inv = CultureInfo.InvariantCulture;
        static LicenceInfo cached;
        static bool warned;
        static string note;                                       // one line for the status bar after a free trial run

        static byte[] B64(string s)
        {
            s = s.Trim().Replace('-', '+').Replace('_', '/');
            while (s.Length % 4 != 0) s += "=";
            return Convert.FromBase64String(s);
        }

        static byte[] Decode(string text)
        {
            string t = (text ?? "").Trim().ToUpperInvariant();
            if (!t.StartsWith("DBK1-", StringComparison.Ordinal)) return null;
            List<byte> bytes = new List<byte>();
            int bits = 0, value = 0;
            foreach (char raw in t.Substring(5))
            {
                char ch = raw;
                if (ch == 'O') ch = '0';
                if (ch == 'I' || ch == 'L') ch = '1';
                if (!((ch >= '0' && ch <= '9') || (ch >= 'A' && ch <= 'Z'))) continue;
                int v = B32.IndexOf(ch);
                if (v < 0) return null;
                value = (value << 5) | v;
                bits += 5;
                if (bits >= 8)
                {
                    bytes.Add((byte)((value >> (bits - 8)) & 255));
                    bits -= 8;
                    value &= (1 << bits) - 1;
                }
            }
            return bytes.Count >= 12 + 64 ? bytes.ToArray() : null;
        }

        // Full check of one key: signature, version, tier and expiry
        public static LicenceInfo Check(string text)
        {
            LicenceInfo li = new LicenceInfo();
            try
            {
                if (PubX.StartsWith("PASTE", StringComparison.Ordinal))
                {
                    li.Reason = "The public key is not set in LicenceKeys.cs yet.";
                    return li;
                }
                byte[] all = Decode(text);
                if (all == null)
                {
                    li.Reason = "This is not a DocBrisk key (it should start with DBK1-).";
                    return li;
                }
                int n = all.Length - 64;
                byte[] payload = new byte[n], sig = new byte[64];
                Array.Copy(all, 0, payload, 0, n);
                Array.Copy(all, n, sig, 0, 64);
                ECParameters p = new ECParameters();
                p.Curve = ECCurve.NamedCurves.nistP256;
                p.Q = new ECPoint { X = B64(PubX), Y = B64(PubY) };
                bool ok;
                using (ECDsa ec = ECDsa.Create(p))
                {
                    ok = ec.VerifyData(payload, sig, HashAlgorithmName.SHA256);
                }
                if (!ok || payload[0] != 1 || (payload[1] != 0x50 && payload[1] != 0x46))
                {
                    li.Reason = "The key was mistyped or changed. Copy it again from the DocBrisk page or email.";
                    return li;
                }
                li.Pro = payload[1] == 0x50;
                li.Name = Encoding.UTF8.GetString(payload, 12, n - 12);
                int exp = (payload[4] << 8) | payload[5];
                if (exp != 0) li.Expires = Epoch.AddDays(exp);
                if (!li.Never && DateTime.UtcNow.Date > li.Expires.Date)
                {
                    li.Reason = "This key expired on " + li.Expires.ToString("dd-MMM-yyyy", Inv) + ".";
                    return li;
                }
                li.Valid = true;
            }
            catch (Exception ex)
            {
                li.Reason = "The key could not be checked: " + ex.Message;
            }
            return li;
        }

        // ---- registry (per Windows user) ----
        static string ReadReg(string name)
        {
            try
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(RegPath))
                {
                    return k == null ? null : Convert.ToString(k.GetValue(name), Inv);
                }
            }
            catch { return null; }
        }

        static void WriteReg(string name, string value)
        {
            try
            {
                using (RegistryKey k = Registry.CurrentUser.CreateSubKey(RegPath))
                {
                    k.SetValue(name, value);
                }
            }
            catch { }
        }

        public static LicenceInfo Current()
        {
            if (cached != null)
            {
                // an expiry date can pass while Excel stays open
                if (cached.Valid && !cached.Never && DateTime.UtcNow.Date > cached.Expires.Date) cached = null;
                else return cached;
            }
            string key = ReadReg("Key");
            cached = string.IsNullOrEmpty(key) ? new LicenceInfo() : Check(key);
            return cached;
        }

        public static bool IsProNow()
        {
            LicenceInfo li = Current();
            return li.Valid && li.Pro;
        }

        static int DaysLeft(LicenceInfo li)
        {
            return (int)(li.Expires.Date - DateTime.UtcNow.Date).TotalDays;
        }

        // one line that says which plan is active: top of the DocBrisk Pro window and the About box
        public static string PlanText()
        {
            LicenceInfo li = Current();
            if (li.Valid && li.Pro)
            {
                if (li.Never) return "Pro plan   |   licensed to " + li.Name + "   |   no expiry";
                return "Pro plan   |   licensed to " + li.Name + "   |   valid until " + li.Expires.ToString("dd-MMM-yyyy", Inv) + "  (" + DaysLeft(li) + " days left)";
            }
            if (li.Valid) return "Free plan   |   licensed to " + li.Name + "   |   Pro tools: " + FreeRuns + " free runs each";
            if (!string.IsNullOrEmpty(li.Reason) && !string.IsNullOrEmpty(ReadReg("Key"))) return "Free plan   |   " + li.Reason;
            return "Free plan   |   every dashboard and everyday tool is free   |   Pro tools: " + FreeRuns + " free runs each";
        }

        // ---- free trial runs of the Pro tools ----
        public static int TrialUsed(string id)
        {
            int used;
            if (!int.TryParse(ReadReg("Trial_" + id), NumberStyles.Integer, Inv, out used) || used < 0) used = 0;
            return used;
        }

        public static int TrialLeft(string id)
        {
            int used = TrialUsed(id);
            return used >= FreeRuns ? 0 : FreeRuns - used;
        }

        // the status-bar line left by the last free trial run (read once)
        public static string TakeNote()
        {
            string s = note;
            note = null;
            return s;
        }

        // ---- the key ----
        // Checks a key and keeps it when it is good. Returns null when it was accepted, otherwise the reason.
        public static string Accept(string text)
        {
            LicenceInfo li = Check(text);
            if (!li.Valid) return li.Reason;
            WriteReg("Key", text.Trim());
            cached = li;
            warned = false;
            return null;
        }

        static string Squash(string s)
        {
            StringBuilder sb = new StringBuilder(s.Length);
            foreach (char raw in s)
            {
                char ch = char.ToUpperInvariant(raw);
                if ((ch >= '0' && ch <= '9') || (ch >= 'A' && ch <= 'Z')) sb.Append(ch);
            }
            return sb.ToString();
        }

        // A DocBrisk key that sits on the clipboard (copied from the website or from an email), or null.
        // Only text that starts with DBK1- is looked at; nothing is sent anywhere.
        public static string ClipboardKey()
        {
            try
            {
                if (!System.Windows.Forms.Clipboard.ContainsText()) return null;
                string t = System.Windows.Forms.Clipboard.GetText();
                if (t == null || t.Length > 4000) return null;
                int p = t.IndexOf("DBK1-", StringComparison.OrdinalIgnoreCase);
                if (p < 0) return null;
                int e = p + 5;
                while (e < t.Length && (char.IsLetterOrDigit(t[e]) || t[e] == '-')) e++;
                string key = t.Substring(p, e - p);
                return key.Length >= 60 ? key : null;
            }
            catch { return null; }
        }

        // the same, but only a valid key that is not the one already in use
        public static string NewClipboardKey()
        {
            string k = ClipboardKey();
            if (k == null) return null;
            string have = ReadReg("Key");
            if (!string.IsNullOrEmpty(have) && string.Equals(Squash(have), Squash(k), StringComparison.Ordinal)) return null;
            return Check(k).Valid ? k : null;
        }

        // Ribbon button "DocBrisk Pro": the plan, what Pro adds, the box for the key. True when Pro is active afterwards.
        public static bool Activate()
        {
            return ProCentre.Show(null);
        }

        // opens the page where Pro keys are sold, in the default browser
        public static void OpenBuyPage()
        {
            try { System.Diagnostics.Process.Start(BuyUrl); }
            catch { Ui.Info("Open this page in your browser to get a Pro key:\n" + BuyUrl); }
        }

        // Called before every Pro tool. True = run the tool.
        public static bool AllowPro(string id, string caption)
        {
            LicenceInfo li = Current();
            if (li.Valid && li.Pro)
            {
                if (!warned && !li.Never)
                {
                    int left = DaysLeft(li);
                    if (left <= WarnDays)
                    {
                        warned = true;
                        if (Ui.Ok("Your Pro plan ends " + (left <= 0 ? "today" : "in " + left + (left == 1 ? " day" : " days")) + ", on " + li.Expires.ToString("dd-MMM-yyyy", Inv) + ".\n\n" +
                                  "OK  =  open the DocBrisk page and renew now\nCancel  =  later")) OpenBuyPage();
                    }
                }
                return true;
            }
            int used = TrialUsed(id);
            if (used < FreeRuns)
            {
                WriteReg("Trial_" + id, (used + 1).ToString(Inv));
                int more = FreeRuns - used - 1;
                note = "'" + caption + "' is a Pro tool: " + (more > 0 ? more + (more == 1 ? " free run" : " free runs") + " left on this PC" : "that was its last free run on this PC") +
                       ".   " + ProPlan.Short + "   (button: DocBrisk Pro)";
                return true;
            }
            string why = "'" + caption + "' is a Pro tool and its " + FreeRuns + " free runs on this PC are used.";
            if (li.Valid) why += "   Your key is a Free key.";
            else if (!string.IsNullOrEmpty(li.Reason) && !string.IsNullOrEmpty(ReadReg("Key"))) why += "   " + li.Reason;
            return ProCentre.Show(why);
        }
    }
}
