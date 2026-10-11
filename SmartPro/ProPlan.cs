namespace SmartPro
{
    // =====================================================================
    //  THE PRO PACKAGE: the prices and the words shown in the DocBrisk Pro window.
    //  Keep the two prices the same as PRO_PLANS in the website Worker (seo-worker.js).
    //  Which tools are Pro is the list ProIds in Design.cs.
    // =====================================================================
    internal static class ProPlan
    {
        public const int YearRupees = 999;       // Pro for 1 year    (Worker: PRO_PLANS.year  = 99900 paise)
        public const int MonthRupees = 99;       // Pro for 1 month   (Worker: PRO_PLANS.month =  9900 paise)
        public const string Rupee = "₹";

        // PRO WEEK: every Pro tool is open for this many days from the first day DocBrisk is used on a PC,
        // without using the free runs. People who have used a tool for a week miss it when it locks.
        // 0 switches the Pro Week off (then only the free runs per tool remain).
        public static readonly int WelcomeProDays = 7;

        // What one hour of office time is worth, for the line "your time saved is worth about ..." in the
        // DocBrisk Pro and My Impact windows. Shown as an estimate; 0 hides the line.
        public static readonly int HourValueRupees = 400;

        // Up to five short lines. Write only what is true for your offer; change a line when the offer changes.
        public static readonly string[] Promise =
        {
            "Every Pro tool, unlimited runs",
            "Dashboards stay free for ever",
            "Works offline, data stays on your PC",
            "One payment, no auto-debit",
            "New versions work with the same key"
        };

        public static string YearLine { get { return "Pro Yearly   " + Rupee + YearRupees + " / year"; } }

        public static string MonthLine { get { return "Pro Monthly   " + Rupee + MonthRupees + " / month"; } }

        // what the yearly plan saves against twelve monthly payments; empty when it saves nothing
        public static string SavingLine
        {
            get
            {
                int save = MonthRupees * 12 - YearRupees;
                if (save <= 0) return "";
                return "Saves " + Rupee + save + " a year  (" + PerDay + ")";
            }
        }

        // the yearly price per day, rounded up to 10 paise: "₹2.8 a day"
        public static string PerDay
        {
            get
            {
                double d = System.Math.Ceiling(YearRupees / 365.0 * 10) / 10;
                return Rupee + d.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture) + " a day";
            }
        }

        // one short line for the status bar and the About box
        public static string Short { get { return "Pro: " + Rupee + MonthRupees + " a month or " + Rupee + YearRupees + " a year"; } }
    }
}
